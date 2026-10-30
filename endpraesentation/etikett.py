"""Druckfertiges Etikett (Vorder- und Rückseite) für eine 250-ml-Glasflasche als Requisit. A4, 300 dpi."""
from PIL import Image, ImageDraw, ImageFont
DPI = 300; MM = DPI / 25.4
PETROL, ORANGE, LIME, CREAM, INK = (15, 76, 79), (224, 123, 32), (155, 191, 69), (246, 243, 237), (27, 42, 47)
F = "/usr/share/fonts/opentype/inter/"
def font(name, pt): return ImageFont.truetype(F + name, int(pt * DPI / 72))
W, H = int(90 * MM), int(62 * MM)   # 9 × 6,2 cm

def center(d, y, text, fnt, fill, w=W, x0=0, tracking=0):
    if tracking:
        widths = [d.textlength(c, font=fnt) for c in text]; total = sum(widths) + tracking * (len(text) - 1)
        x = x0 + (w - total) / 2
        for c, cw in zip(text, widths): d.text((x, y), c, font=fnt, fill=fill); x += cw + tracking
    else:
        d.text((x0 + (w - d.textlength(text, font=fnt)) / 2, y), text, font=fnt, fill=fill)

def front():
    im = Image.new("RGB", (W, H), CREAM); d = ImageDraw.Draw(im)
    f_small = font("Inter-Bold.otf", 8); y = int(5 * MM)
    tw = d.textlength("ELBHARZ", font=f_small) + 6 * 8
    center(d, y, "ELBHARZ", f_small, PETROL, tracking=8)
    ly = y + int(1.8 * MM)
    d.line([(int(8 * MM), ly), ((W - tw) / 2 - 20, ly)], fill=PETROL, width=4); d.line([((W + tw) / 2 + 20, ly), (W - int(8 * MM), ly)], fill=PETROL, width=4)
    big = font("InterDisplay-Black.otf", 26)
    while d.textlength("PRIMATENBRAUSE", font=big) - 13 * 4 > W - 10 * MM: big = font("InterDisplay-Black.otf", big.size * 72 / DPI - 0.5)
    center(d, int(9.5 * MM), "PRIMATENBRAUSE", big, PETROL, tracking=-4)
    center(d, int(21 * MM), "TAMARINDE · LIMETTE", font("Inter-Bold.otf", 11), LIME, tracking=3)
    d.line([(int(8 * MM), int(27.5 * MM)), (W - int(8 * MM), int(27.5 * MM))], fill=LIME, width=4)
    f_z = font("Inter-Medium.otf", 7)
    center(d, int(30 * MM), "Wasser · Kohlensäure · Tamarinde ·", f_z, LIME)
    center(d, int(33.5 * MM), "Limettensaft · Meersalz · Magnesium", f_z, LIME)
    center(d, int(39 * MM), "Was dein Körper schon immer wollte.", font("Inter-SemiBold.otf", 8.5), INK)
    d.rectangle([0, int(46 * MM), W, H], fill=ORANGE)
    center(d, int(49.5 * MM), "250 ml", font("Inter-Bold.otf", 15), (255, 255, 255))
    return im

def back():
    im = Image.new("RGB", (W, H), CREAM); d = ImageDraw.Draw(im)
    x = int(7 * MM)
    d.text((x, int(5 * MM)), "RÜCKSEITE", font=font("Inter-Bold.otf", 7.5), fill=ORANGE)
    d.text((x, int(9 * MM)), "Zutaten", font=font("InterDisplay-ExtraBold.otf", 15), fill=PETROL)
    f = font("Inter-Medium.otf", 9)
    zut = ["Wasser", "Kohlensäure", "Tamarinde", "Limettensaft", "Meersalz", "Magnesium"]
    for i, z in enumerate(zut):
        cx = x + (i // 3) * int(38 * MM); cy = int(17.5 * MM) + (i % 3) * int(5 * MM)
        r = int(1.1 * MM); d.ellipse([cx, cy + int(1.2 * MM), cx + 2 * r, cy + int(1.2 * MM) + 2 * r], fill=LIME)
        d.text((cx + int(3.5 * MM), cy), z, font=f, fill=INK)
    d.text((x, int(34 * MM)), "6 Zutaten. Mehr nicht.", font=font("Inter-Bold.otf", 9.5), fill=ORANGE)
    d.text((x, int(39 * MM)), "Mehrweg-Glas · in Hamburg abgefüllt", font=font("Inter-Medium.otf", 7.5), fill=INK)
    d.rectangle([0, int(46 * MM), W, H], fill=PETROL)
    logo = Image.open("assets/logo.png").convert("RGBA"); bb = logo.getbbox(); logo = logo.crop(bb)
    lh = int(12 * MM); logo = logo.resize((int(logo.width * lh / logo.height), lh), Image.LANCZOS)
    im.paste(logo, (x, int(48 * MM)), logo)
    d.text((x + logo.width + int(4 * MM), int(50.5 * MM)), "Elbharz Getränke GbR · Hamburg", font=font("Inter-SemiBold.otf", 8), fill=(255, 255, 255))
    d.text((x + logo.width + int(4 * MM), int(55 * MM)), "Schulprojekt – fiktives Produkt", font=font("Inter-Medium.otf", 6.5), fill=(200, 215, 216))
    return im

A4 = (int(210 * MM), int(297 * MM)); page = Image.new("RGB", A4, "white"); d = ImageDraw.Draw(page)
d.text((int(15 * MM), int(12 * MM)), "Primatenbrause – Etikett zum Ausdrucken (250-ml-Flasche)", font=font("Inter-Bold.otf", 13), fill=PETROL)
d.text((int(15 * MM), int(19 * MM)), "In Originalgröße drucken (Skalierung 100 %, nicht „an Seite anpassen“). An den Linien ausschneiden, Vorderseite vorne, Rückseite hinten aufkleben.", font=font("Inter-Regular.otf", 8), fill=INK)
fr, bk = front(), back()
for row in range(3):
    y = int(30 * MM) + row * (H + int(18 * MM))
    for col, lab, name in ((0, fr, "Vorderseite"), (1, bk, "Rückseite")):
        x = int(12 * MM) + col * (W + int(6 * MM))
        page.paste(lab, (x, y)); d.rectangle([x - 2, y - 2, x + W + 1, y + H + 1], outline=(160, 160, 160), width=2)
        d.text((x, y + H + int(2 * MM)), name, font=font("Inter-Medium.otf", 7), fill=(120, 120, 120))
page.save("Etikett_Primatenbrause_A4.pdf", resolution=DPI)
fr.save("/tmp/claude-0/-home-user-claude-workspace/0b460da1-1cd1-536d-b4aa-b391322290fe/scratchpad/etikett_front.png"); bk.save("/tmp/claude-0/-home-user-claude-workspace/0b460da1-1cd1-536d-b4aa-b391322290fe/scratchpad/etikett_back.png")
print("ok")
