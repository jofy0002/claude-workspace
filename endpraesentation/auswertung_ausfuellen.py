"""Füllt die leeren Präsentationsbögen M15 (Auswertung) mit den Zahlen von Elbharz.
Aufruf: python3 auswertung_ausfuellen.py <p-3.png> <p-4.png> <ausgabe.pdf>"""
import sys
from PIL import Image, ImageDraw, ImageFont

S = 1653 / 827  # Anzeige-Koordinaten (827 px breit) -> 200-dpi-Scan
INK = (20, 45, 160)
RED = (170, 40, 40)
FONT = "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"
def F(sz): return ImageFont.truetype(FONT, int(sz * S))

def text(d, x, y, t, sz=15, anchor="mm", fill=INK):
    d.text((x * S, y * S), t, font=F(sz), fill=fill, anchor=anchor)

def eur(v): return f"{v:,.0f}".replace(",", ".") + ",–"

runden = {
    "Maschinenanschaffungskosten": [30000, 30000, 0, 0, 0, 0],
    "Personalkosten": [4000, 4000, 8000, 8000, 8000, 8000],
    "Miete": [5000] * 6,
    "Sonstige Kosten": [0, 10000, 5000, 2500, 10000, 6500],
    "Produktionskosten": [15000, 13000, 18000, 1000, 1000, 13000],
    "Materialkosten": [15000, 13000, 18000, 1500, 1000, 13000],
}
umsatz = [45000, 55900, 79940, 5250, 2700, 51480]
kosten = [sum(v[i] for v in runden.values()) for i in range(6)]
wirt = [round(u / k, 2) for u, k in zip(umsatz, kosten)]
schluss = [76000, 56900, 82840, 70090, 47790, 53770]
NAME = "Elbharz Getränke GbR"
print("Kosten", kosten, "Wirtschaftlichkeit", wirt)

# ---------- Seite 1: Wirtschaftlichkeit ----------
im = Image.open(sys.argv[1]).convert("RGB"); d = ImageDraw.Draw(im)
text(d, 82, 207, NAME, 20, "lm")
cols = [341.5, 420.5, 499, 577, 655, 732.5]
rows_y = [326.5, 351.5, 376.5, 401.5, 426.5, 451.5, 476.5, 501.5, 527]
keys = list(runden)
for r, k in enumerate(keys):
    for c in range(6): text(d, cols[c], rows_y[r], eur(runden[k][c]), 14)
for c in range(6):
    text(d, cols[c], rows_y[6], eur(kosten[c]), 14)
    text(d, cols[c], rows_y[7], eur(umsatz[c]), 14)
    text(d, cols[c], rows_y[8], f"{wirt[c]:.2f}".replace(".", ","), 15)
# Säulendiagramm: 0 bei y=1010, 2,0 bei y=684 (Anzeige-Koordinaten)
Y0, PER = 1010, 163
yv = lambda v: Y0 - v * PER
for k in range(1, 11):
    v = k * 0.2
    text(d, 146, yv(v), f"{v:.1f}".replace(".", ","), 11, "rm")
xs = [244, 333, 422, 511, 600, 690]
for x, w in zip(xs, wirt):
    d.rectangle([(x - 17) * S, yv(w) * S, (x + 17) * S, Y0 * S], fill=(120, 150, 215), outline=INK, width=3)
    text(d, x, yv(w) - 11, f"{w:.2f}".replace(".", ","), 14)
d.line([(158 * S, yv(1) * S), (750 * S, yv(1) * S)], fill=RED, width=3)
text(d, 165, yv(1) - 9, "1,0 = Kosten gedeckt", 11, "lm", RED)
p1 = im

# ---------- Seite 2: Kassenbestand ----------
im = Image.open(sys.argv[2]).convert("RGB"); d = ImageDraw.Draw(im)
text(d, 62, 207, NAME, 20, "lm")
cx = [245.5, 337.5, 430, 522, 613.5, 705]
for c in range(6): text(d, cx[c], 380, eur(schluss[c]), 14)
Y0b, PERb = 858, 32.55  # 10.000 € je Rasterlinie
yb = lambda v: Y0b - v / 10000 * PERb
for k in range(0, 12):
    text(d, 126, yb(k * 10000), f"{k*10000:,}".replace(",", "."), 10, "rm")
pts = [(134, yb(100000))] + [(x, yb(v)) for x, v in zip([226, 315, 404, 494, 583, 672], schluss)]
d.line([(x * S, y * S) for x, y in pts], fill=INK, width=5)
for i, (x, y) in enumerate(pts):
    d.rectangle([(x - 4) * S, (y - 4) * S, (x + 4) * S, (y + 4) * S], fill=INK)
    v = 100000 if i == 0 else schluss[i - 1]
    dx, dy, an = {0: (10, -12, "lm"), 1: (-10, 12, "rm"), 2: (0, 15, "mm"), 3: (0, -12, "mm"),
                  4: (12, -12, "lm"), 5: (0, 15, "mm"), 6: (0, -12, "mm")}[i]
    text(d, x + dx, y + dy, f"{v:,}".replace(",", "."), 11, an)
p2 = im

p1.save(sys.argv[3], save_all=True, append_images=[p2], resolution=200)
