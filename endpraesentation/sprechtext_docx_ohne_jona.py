"""Erzeugt Sprechtext_Elbharz.docx aus den Notizen der Präsentation."""
from pptx import Presentation
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

PETROL, AMBER, MUTED = RGBColor(0x0F, 0x5C, 0x6B), RGBColor(0xD9, 0x82, 0x2B), RGBColor(0x5F, 0x8F, 0x99)
WER = ["Sönke", "Sönke", "Sönke", "Nio", "Nio", "Nio", "Lennart", "Lennart", "Sönke", "Sönke", "Lennart", "Lennart"]
TITEL = ["Titel", "Wer wir sind", "Nachhaltigkeit", "Produziert, angeboten, verkauft", "Drei Entscheidungen",
         "Zwei Kennzahlen", "Wirtschaftlichkeit & Schlussbestand", "Verkaufswert", "Wirtschaftsdienst",
         "Reflexion", "Bezug zur Gründung", "Schluss + Danke"]
UEBERGABE = {3: "Übergabe an Nio", 6: "Übergabe an Lennart", 8: "Übergabe an Sönke", 10: "Übergabe an Lennart"}
ERSETZEN = {  # Übergabesätze: alle behalten ihre Folien, Jonas Folien 4 + 6 → Nio, 7 → Lennart
    3: [("Wie es im Spiel lief, zeigt euch jetzt Jona.", "Wie es im Spiel lief, zeigt euch jetzt Nio.")],
    4: [("Welche Entscheidungen dahinterstanden, erklärt euch jetzt Nio.", "Welche Entscheidungen dahinterstanden, erkläre ich euch jetzt.")],
    5: [("Ich erkläre euch unsere wichtigsten Entscheidungen.", "Hinter diesem Verlauf stecken unsere Entscheidungen."), ("Wie wir das ausgewertet haben, zeigt euch jetzt Jona.", "Wie wir das ausgewertet haben, zeige ich euch jetzt.")],
    6: [("Was das für unser Geld bedeutet hat, seht ihr jetzt.", "Was das für unser Geld bedeutet hat, zeigt euch jetzt Lennart.")],
    7: [("Was unser Unternehmen am Ende wert war, erklärt euch jetzt Lennart.", "Was unser Unternehmen am Ende wert war, zeige ich euch jetzt.")],
    8: [("Warum es so wenig wurde, erklärt euch Nio.", "Warum es so wenig wurde, erklärt euch Sönke.")],
    9: [("Was wir besser machen würden, erklärt euch jetzt Sönke.", "Was wir besser machen würden, erkläre ich euch jetzt.")],
}

prs = Presentation("Elbharz_Endpraesentation.pptx")
doc = Document()
for sec in doc.sections:
    sec.left_margin = sec.right_margin = Cm(2.2); sec.top_margin = sec.bottom_margin = Cm(2)
st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(12)

h = doc.add_paragraph(); r = h.add_run("Sprechtext – Ersatzaufteilung ohne Jona"); r.bold = True; r.font.size = Pt(22); r.font.color.rgb = PETROL
p = doc.add_paragraph(); r = p.add_run("Elbharz Getränke GbR · Klasse 11b · Aufteilung OHNE Jona (alle behalten ihre Folien): Sönke 1–3, 9–10 · Nio 4–6 · Lennart 7–8, 11–12 · ca. 10 Min."); r.font.color.rgb = MUTED
tips = doc.add_paragraph()
r = tips.add_run("So nutzt ihr den Text: "); r.bold = True
tips.add_run("Das ist ein Vorschlag – sagt es in euren eigenen Worten. Genaue Zahlen stehen auf den Folien, im Sprechen reichen runde Zahlen. "
             "Nicht ablesen, was alle sehen, sondern erzählen, was ihr gedacht habt. "
             "Die Übergabesätze mit Namen stehen schon im Text.")

for i, s in enumerate(list(prs.slides)[:12]):
    n = s.notes_slide.notes_text_frame.text
    zeit = n.splitlines()[0].replace("ZEIT: ", "").replace("ca. ", "").replace(" min", "")
    text = n.split("anpassen):\n")[-1].strip()
    for a, b in ERSETZEN.get(i + 1, []):
        assert a in text, (i + 1, a); text = text.replace(a, b)
    hd = doc.add_paragraph(); hd.paragraph_format.space_before = Pt(16); hd.paragraph_format.keep_with_next = True
    r = hd.add_run(f"Folie {i+1} – {TITEL[i]} · {WER[i]}"); r.bold = True; r.font.size = Pt(15); r.font.color.rgb = PETROL
    meta = doc.add_paragraph(); meta.paragraph_format.keep_with_next = True
    r = meta.add_run(f"{zeit} Min."); r.font.size = Pt(11); r.font.color.rgb = MUTED
    doc.add_paragraph(text)
    if i + 1 in UEBERGABE:
        u = doc.add_paragraph(); r = u.add_run(UEBERGABE[i + 1]); r.italic = True; r.font.color.rgb = AMBER

doc.add_page_break()
h = doc.add_paragraph(); r = h.add_run("Falls die Lehrkraft nachfragt"); r.bold = True; r.font.size = Pt(15); r.font.color.rgb = PETROL
for q, a in [
    ("Wie habt ihr die Wirtschaftlichkeit berechnet?", "Umsatzerlöse geteilt durch alle Kosten der Runde, z. B. R3: 79.940 € ÷ 54.000 € = 1,48. Alle Werte: Anhang-Folie."),
    ("Wie aussagekräftig ist die Kennzahl?", "Maschinenkäufe zählen voll in der Kaufrunde; die Absatzquote sagt bei nur 1 Palette (R4/R5) wenig aus. Rechnet man stattdessen in allen Runden nur den Wertverlust (10 % je Nutzungsrunde), ergibt sich: R1 1,07 · R2 1,16 · R3 1,33 · R4 0,22 · R5 0,09 · R6 ≈ 1,00."),
    ("Wie kommt ihr auf die 31.000 € bei nur einer Maschine?", "Gleiche Preise und gleicher Absatz wie im Spiel: +30.000 € kein Maschinenkauf, +16.000 € ein Mitarbeiter weniger ab R3, ca. +3.000 € weniger Produktion in R3 (am Markt trotzdem 12,6 verkauft), −18.000 € kein Restwert. Wir haben nie mehr als 15 Paletten normal verkauft."),
    ("Warum habt ihr in R2 eine Maschine ohne Mitarbeiter gekauft?", "Wir wollten Kapazität für R3 aufbauen; der Mitarbeiter kam erst in R3. Im Nachhinein hätten wir beides zusammen planen sollen."),
    ("Warum in R4/R5 so ein hoher Preis?", "Wir erwarteten weniger Absatz und wollten mit einer Palette kein Minus machen. Das war ein Denkfehler: Der Preis bestimmt den Absatz."),
    ("Wie kommt ihr auf 18.000 € Restwert für Maschine 2?", "10 % Wertverlust je genutzter Runde; sie lief nur R3–R6, also 4 Runden: 30.000 − 4 × 3.000."),
    ("Wie würde eure Spielerweiterung funktionieren?", "Marketingbudget: je 2.000 € → +1 Palette Absatz. Mehrweg-Option: +1.000 € Kosten je Runde → +1 Palette Absatz."),
]:
    p = doc.add_paragraph(); r = p.add_run(q); r.bold = True
    doc.add_paragraph(a)
doc.save("Sprechtext_Elbharz_ohne_Jona.docx"); print("ok")
