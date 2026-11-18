# Twitch Auto-Clipper

Liest den Chat eines **laufenden** Twitch-Streams mit. Wenn der Chat plötzlich
ausrastet (viel mehr Nachrichten als normal, „KEKW“, „LOL“, „💀“ …), erstellt das
Tool **automatisch einen Twitch-Clip**. Zusätzlich kann jeder im Chat `!clip` schreiben.

Alle Clip-Links landen in `clips.csv` und in deinem Twitch-Konto unter
*Creator Dashboard → Inhalte → Clips*. Dort kannst du sie herunterladen und hochladen.

## Einrichtung (einmalig, geht auch am iPad)

1. **App anlegen:** https://dev.twitch.tv/console → „Register Your Application“
   - Name: beliebig, OAuth Redirect URL: `http://localhost`, Kategorie: *Other*
   - Typ: Client „Public“ → **Client-ID** kopieren.
2. **Token holen:** diese Adresse öffnen (CLIENT_ID ersetzen):
   ```
   https://id.twitch.tv/oauth2/authorize?response_type=token&client_id=CLIENT_ID&redirect_uri=http://localhost&scope=clips:edit
   ```
   Erlauben → die Seite lädt nicht, das ist normal. In der Adresszeile steht
   `access_token=XXXX&...` → den Teil **XXXX** kopieren. (Token läuft nach ca. 60 Tagen ab.)
3. `config.example.json` nach `config.json` kopieren und Kanal, Client-ID und Token eintragen.

## Starten

Braucht nur Python 3.9+ (keine Pakete):
```
python3 autoclipper.py
```
Das Tool muss laufen, **während** der Stream läuft – auf einem PC, Laptop,
Raspberry Pi oder kostenlosem Server. Auf dem iPad selbst läuft es nicht dauerhaft
im Hintergrund.

## Einstellungen (config.json)

| Wert | Bedeutung |
|---|---|
| `spike_factor` | Wie viel mehr Chat als normal ein Highlight ist (kleiner = mehr Clips) |
| `min_messages` | Mindestens so viele Nachrichten in `window_seconds` |
| `cooldown_seconds` | Mindestabstand zwischen zwei Clips |
| `keywords` | Wörter/Emojis, die stärker zählen |
| `chat_command` | Manueller Chat-Befehl, `""` zum Abschalten |

Hinweis: Clips gehen nur, wenn der Kanal Clips erlaubt. Bei fremden Kanälen
darfst du Clips auf Twitch teilen – zum Hochladen auf TikTok/YouTube vorher
den Streamer fragen.
