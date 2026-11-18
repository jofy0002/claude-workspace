#!/usr/bin/env python3
"""Twitch Auto-Clipper: liest den Chat eines Live-Streams mit und erstellt
automatisch einen Twitch-Clip, wenn der Chat plötzlich ausrastet.

Nur Python-Standardbibliothek, keine Installation nötig.
Konfiguration über config.json (siehe config.example.json).
"""
import json, socket, ssl, sys, time, csv, os, re, urllib.request, urllib.error
from collections import deque

HERE = os.path.dirname(os.path.abspath(__file__))

DEFAULTS = {
    "channel": "",              # Twitch-Kanalname, z. B. "meinkanal"
    "client_id": "",            # von dev.twitch.tv
    "access_token": "",         # User-Token mit Scope clips:edit
    "window_seconds": 10,       # wie viele Sekunden Chat werden gezählt
    "spike_factor": 3.0,        # x-mal mehr Nachrichten als normal = Highlight
    "min_messages": 8,          # mindestens so viele Nachrichten im Fenster
    "cooldown_seconds": 90,     # Pause zwischen zwei Clips
    "keywords": ["lol", "lmao", "omg", "wtf", "kekw", "lul", "pog", "poggers",
                 "w", "clip", "!clip", "holy", "xd", "😂", "💀", "🔥"],
    "keyword_weight": 1.5,      # Nachrichten mit Keyword zählen stärker
    "chat_command": "!clip",    # Chat-Befehl für manuellen Clip ("" = aus)
}


def load_config():
    path = os.path.join(HERE, "config.json")
    if not os.path.exists(path):
        sys.exit("config.json fehlt. Kopiere config.example.json nach config.json und trag deine Daten ein.")
    with open(path, encoding="utf-8") as f:
        cfg = {**DEFAULTS, **json.load(f)}
    cfg["channel"] = cfg["channel"].lower().lstrip("#")
    cfg["access_token"] = cfg["access_token"].removeprefix("oauth:")
    for k in ("channel", "client_id", "access_token"):
        if not cfg[k]:
            sys.exit(f"'{k}' fehlt in config.json")
    return cfg


def helix(cfg, method, path, params=""):
    req = urllib.request.Request(
        f"https://api.twitch.tv/helix/{path}?{params}", method=method,
        headers={"Client-Id": cfg["client_id"],
                 "Authorization": f"Bearer {cfg['access_token']}"})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        print(f"[API-Fehler] {e.code}: {e.read().decode(errors='ignore')}")
        return None


class SpikeDetector:
    """Vergleicht die aktuelle Chat-Aktivität mit dem gleitenden Durchschnitt."""

    def __init__(self, cfg):
        self.cfg = cfg
        self.events = deque()          # (zeit, gewicht)
        self.baseline = None           # normale Aktivität pro Fenster
        self.last_clip = 0.0
        self.kw = [k.lower() for k in cfg["keywords"]]

    def weight(self, msg):
        words = re.findall(r"\S+", msg.lower())
        hit = any(k in words or (len(k) > 2 and k in msg.lower()) for k in self.kw)
        return self.cfg["keyword_weight"] if hit else 1.0

    def add(self, msg, now):
        self.events.append((now, self.weight(msg)))

    def check(self, now):
        win = self.cfg["window_seconds"]
        while self.events and self.events[0][0] < now - win:
            self.events.popleft()
        score = sum(w for _, w in self.events)
        if self.baseline is None:
            self.baseline = max(score, 1.0)
            return False, score
        is_spike = (score >= self.cfg["min_messages"]
                    and score >= self.baseline * self.cfg["spike_factor"]
                    and now - self.last_clip >= self.cfg["cooldown_seconds"])
        # Baseline langsam anpassen (Spitzen kaum einrechnen)
        alpha = 0.01 if score > self.baseline * 2 else 0.05
        self.baseline = max(1.0, (1 - alpha) * self.baseline + alpha * score)
        return is_spike, score


def make_clip(cfg, broadcaster_id, reason):
    data = helix(cfg, "POST", "clips", f"broadcaster_id={broadcaster_id}")
    if not data or not data.get("data"):
        print("[!] Clip fehlgeschlagen (Stream offline oder Clips deaktiviert?)")
        return
    clip_id = data["data"][0]["id"]
    url = f"https://clips.twitch.tv/{clip_id}"
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[CLIP] {stamp}  {url}  ({reason})")
    with open(os.path.join(HERE, "clips.csv"), "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([stamp, cfg["channel"], url, reason])


def chat_lines(channel):
    """Liest den Chat anonym (kein Login nötig) und liefert Nachrichten."""
    while True:
        try:
            raw = socket.create_connection(("irc.chat.twitch.tv", 6697), timeout=30)
            sock = ssl.create_default_context().wrap_socket(raw, server_hostname="irc.chat.twitch.tv")
            nick = f"justinfan{int(time.time()) % 100000}"
            sock.sendall(f"PASS SCHMOOPIIE\r\nNICK {nick}\r\nJOIN #{channel}\r\n".encode())
            print(f"[OK] Verbunden mit Chat von #{channel}")
            buf = ""
            while True:
                try:
                    data = sock.recv(4096).decode(errors="ignore")
                except socket.timeout:
                    yield None
                    continue
                if not data:
                    raise ConnectionError("Verbindung getrennt")
                buf += data
                *lines, buf = buf.split("\r\n")
                for line in lines:
                    if line.startswith("PING"):
                        sock.sendall(b"PONG :tmi.twitch.tv\r\n")
                    elif " PRIVMSG " in line:
                        yield line.split(" :", 2)[-1] if line.startswith(":") else line
        except Exception as e:
            print(f"[!] Chat-Verbindung verloren ({e}), neuer Versuch in 5 s ...")
            time.sleep(5)


def main():
    cfg = load_config()
    users = helix(cfg, "GET", "users", f"login={cfg['channel']}")
    if not users or not users.get("data"):
        sys.exit("Kanal nicht gefunden oder Token/Client-ID falsch.")
    bid = users["data"][0]["id"]
    print(f"[OK] Kanal {cfg['channel']} (ID {bid}). Warte auf Highlights ... (Strg+C zum Beenden)")

    det = SpikeDetector(cfg)
    last_check = time.time()
    for msg in chat_lines(cfg["channel"]):
        now = time.time()
        if msg is not None:
            if cfg["chat_command"] and msg.strip().lower() == cfg["chat_command"]:
                if now - det.last_clip >= 20:
                    det.last_clip = now
                    make_clip(cfg, bid, "Chat-Befehl")
                continue
            det.add(msg, now)
        if now - last_check >= 1:
            last_check = now
            spike, score = det.check(now)
            if spike:
                det.last_clip = now
                make_clip(cfg, bid, f"Chat-Spitze: {score:.0f} vs. normal {det.baseline:.1f}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nBeendet.")
