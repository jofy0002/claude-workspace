"""Erzeugt Team_Leitfaden_Elbharz.xlsx – gemeinsame Abhak-Liste für das Team (Excel/OneDrive, ohne Claude nutzbar)."""
from pptx import Presentation
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from datetime import time

PETROL, AMBER, LIME, SAND = "0F5C6B", "D9822B", "8DB33A", "F3F1EA"
F = lambda **k: Font(name="Arial", **k)
HEAD = dict(font=F(bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=PETROL))
INPUT = PatternFill("solid", fgColor="FFF4D6")
GREEN = PatternFill("solid", fgColor="DDEFC6")
thin = Side(style="thin", color="D3D1C8"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

def head(ws, row, values, widths=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row, i, v); c.font = HEAD["font"]; c.fill = HEAD["fill"]; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[chr(64 + i)].width = w

def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = F(bold=True, size=16, color=PETROL)
    ws["A2"] = sub; ws["A2"].font = F(italic=True, size=10, color="5F8F99")

wb = Workbook()
# ---------------------------------------------------------------- Start
st = wb.active; st.title = "Start"
title(st, "Team-Leitfaden Endpräsentation – Elbharz Getränke GbR", "Gemeinsam in OneDrive/SharePoint öffnen – alle können gleichzeitig bearbeiten.")
st["A4"] = "Unsere Namen (hier eintragen – erscheinen dann überall als Spalten und in den Auswahllisten):"; st["A4"].font = F(bold=True)
for i in range(4):
    st.cell(5 + i, 1, f"Person {i+1}").font = F()
    c = st.cell(5 + i, 2); c.fill = INPUT; c.border = BOX; c.font = F(bold=True)
wb.defined_names["Team"] = DefinedName("Team", attr_text="Start!$B$5:$B$8")
legend = [
    ("So funktioniert's", ""),
    ("Gelbe Felder", "Hier tragt ihr etwas ein. Alles andere rechnet die Tabelle selbst."),
    ("Blatt „Aufgaben“, Teil A", "Jede:r muss das selbst machen → in der EIGENEN Spalte „✓“ auswählen (Pfeil am Feld). Beispiel: Person 2 hat geprobt → ✓ in der Spalte von Person 2."),
    ("Blatt „Aufgaben“, Teil B", "Eine:r reicht → bei „Erledigt von“ den eigenen Namen auswählen und Datum eintragen. Dann wird die Zeile grün."),
    ("Blatt „Rollen & Zeitplan“", "Jede Folie einzeln vergeben, auch abwechselnd; optional eine 2. Person pro Folie. Zeit und Anzahl Folien je Person rechnen sich mit."),
    ("Blatt „Sprechtexte“", "Vorschlag je Folie; in der gelben Spalte schreibt jede:r die eigene Version."),
    ("Blatt „Fragen“", "Mögliche Fragen der Lehrkraft; bei „Wer antwortet“ den Namen auswählen."),
]
for k, (a, b) in enumerate(legend):
    r = 11 + k
    st.cell(r, 1, a).font = F(bold=True, color=PETROL if k == 0 else "1B2A2F", size=12 if k == 0 else 10)
    st.cell(r, 2, b).font = F(); st.cell(r, 2).alignment = WRAP
st.column_dimensions["A"].width = 26; st.column_dimensions["B"].width = 90
st["A19"] = "Stand"; st["A19"].font = F(bold=True)
st["B19"] = '=Aufgaben!B3'; st["B19"].font = F(bold=True, color=PETROL)

# ---------------------------------------------------------------- Aufgaben
ws = wb.create_sheet("Aufgaben")
title(ws, "Aufgaben", "Teil A: jede:r einzeln abhaken · Teil B: eine:r reicht")
einzeln = [
    ("Eigenen Sprechtext lesen und in eigene Worte bringen", "Blatt „Sprechtexte“, gelbe Spalte"),
    ("Eigenen Teil laut mit Stoppuhr üben", "Zeiten: Blatt „Rollen & Zeitplan“"),
    ("Zahlen der eigenen Folien ohne Ablesen erklären können", "z. B. Wirtschaftlichkeit 1,48 in R3"),
    ("Fragen zum eigenen Teil beantworten können", "Blatt „Fragen“"),
    ("Übergabesatz an die nächste Person geübt", "Wer übergibt mit welchem Satz?"),
    ("Reflexion (Folie 10) gelesen – einverstanden oder Einwand gesagt", "Muss unsere gemeinsame Meinung sein"),
    ("Komplette Probe mitgemacht", "Ziel: unter 11 Minuten"),
]
gemeinsam = [
    ("Sprecher festlegen und Namen eintragen", "Blätter „Start“ und „Rollen & Zeitplan“"),
    ("Karte Norddeutschland mit roter Elbe besorgen", "Folie 3"),
    ("KI-Teambild erstellen und Tool notieren", "Folie 12 + Quellenfolie"),
    ("Herkunft der Bilder Limette und Tamarinde klären", "Quellenfolie"),
    ("Optional: Fotos Geld, Maschine, Gebäude", "Folie 9 (sonst bleiben die Symbole)"),
    ("Optional: Stichworte zu den Wirtschaftsdienst-Meldungen R2–R6", "Folie 6"),
    ("R3 prüfen: 12,6 + 5,6 = 18,2 Paletten bei 18 produzierten", "Arbeitsblatt Umsatzberechnung"),
    ("R3 klären: Protokoll „keine neuen Mitarbeiter“ vs. Arbeitsblatt 2 Mitarbeiter", "Antwort parat haben, falls gefragt"),
    ("Platzierung nach der Siegerehrung ergänzen", "Optional, Folie 9"),
    ("Präsentation am Schulrechner/Beamer testen", "Schriften, Diagramme, Bilder"),
    ("Diagramme zeigen deutsche Zahlen (10.000 statt 10,000)", "Beim Test prüfen"),
    ("Alle Platzhalter ersetzt („BILD FOLGT“, „[Quelle ergänzen]“)", "Folien 3, 12, 13"),
    ("PDF-Version exportieren", "Falls PowerPoint fehlt"),
    ("Datei auf USB-Stick und in der Cloud", "Backup"),
    ("Präsentation (PPTX + PDF) und Bögen M15 an die Lehrkraft", "Digitale Abgabe"),
]
# Teil A
a0 = 5
ws.cell(a0, 1, "A  Jede:r einzeln – in der eigenen Spalte ✓ setzen").font = F(bold=True, size=12, color=AMBER)
head(ws, a0 + 1, ["Nr", "Aufgabe", "Hinweis", None, None, None, None, "Fortschritt"], [5, 52, 34, 13, 13, 13, 13, 12])
for i in range(4):
    c = ws.cell(a0 + 1, 4 + i, f'=IF(Start!$B${5+i}="","Person {i+1}",Start!$B${5+i})')
    c.font = HEAD["font"]; c.fill = HEAD["fill"]; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
dv_check = DataValidation(type="list", formula1='"✓"', allow_blank=True, showDropDown=False)
dv_check.error = "Bitte nur ✓ auswählen oder Feld leeren."; ws.add_data_validation(dv_check)
ra1, ra2 = a0 + 2, a0 + 1 + len(einzeln)
for k, (t, h) in enumerate(einzeln):
    r = ra1 + k
    ws.cell(r, 1, k + 1); ws.cell(r, 2, t); ws.cell(r, 3, h)
    for col in range(4, 8):
        c = ws.cell(r, col); c.fill = INPUT; c.alignment = Alignment(horizontal="center", vertical="center"); c.font = F(bold=True, size=12, color="3D6610")
        dv_check.add(c.coordinate)
    ws.cell(r, 8, f'=COUNTIF(D{r}:G{r},"✓")&" / 4"')
    for col in range(1, 9):
        c = ws.cell(r, col); c.border = BOX
        if col < 4 or col == 8: c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="center")
ws.conditional_formatting.add(f"D{ra1}:G{ra2}", CellIsRule(operator="equal", formula=['"✓"'], fill=GREEN))
ws.conditional_formatting.add(f"A{ra1}:C{ra2}", FormulaRule(formula=[f'COUNTIF($D{ra1}:$G{ra1},"✓")=4'], fill=GREEN))
ws.conditional_formatting.add(f"H{ra1}:H{ra2}", FormulaRule(formula=[f'COUNTIF($D{ra1}:$G{ra1},"✓")=4'], fill=GREEN))
# Teil B
b0 = ra2 + 3
ws.cell(b0, 1, "B  Eine:r reicht – bei „Erledigt von“ den eigenen Namen auswählen").font = F(bold=True, size=12, color=AMBER)
head(ws, b0 + 1, ["Nr", "Aufgabe", "Hinweis", "Erledigt von", "Datum", "Status", "Notiz", None])
ws.merge_cells(start_row=b0 + 1, start_column=7, end_row=b0 + 1, end_column=8)
dv_name = DataValidation(type="list", formula1="=Team", allow_blank=True); ws.add_data_validation(dv_name)
dv_date = DataValidation(type="date", operator="greaterThan", formula1="45000", allow_blank=True)
dv_date.error = "Bitte ein Datum eingeben, z. B. 20.10.2026"; ws.add_data_validation(dv_date)
rb1, rb2 = b0 + 2, b0 + 1 + len(gemeinsam)
for k, (t, h) in enumerate(gemeinsam):
    r = rb1 + k
    ws.cell(r, 1, k + 1); ws.cell(r, 2, t); ws.cell(r, 3, h)
    d = ws.cell(r, 4); d.fill = INPUT; dv_name.add(d.coordinate)
    e = ws.cell(r, 5); e.fill = INPUT; e.number_format = "DD.MM.YYYY"; dv_date.add(e.coordinate)
    ws.cell(r, 6, f'=IF(D{r}="","offen","✓ erledigt")')
    g = ws.cell(r, 7); g.fill = INPUT
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
    for col in range(1, 9):
        c = ws.cell(r, col); c.border = BOX; c.font = F(bold=(col == 6)); c.alignment = Alignment(wrap_text=True, vertical="center")
ws.conditional_formatting.add(f"A{rb1}:F{rb2}", FormulaRule(formula=[f'$D{rb1}<>""'], fill=GREEN))
ws["B3"] = (f'="Stand – Teil A: "&COUNTIF(D{ra1}:G{ra2},"✓")&" von "&{len(einzeln)*4}&" Häkchen · Teil B: "'
            f'&COUNTIF(F{rb1}:F{rb2},"✓ erledigt")&" von "&{len(gemeinsam)}&" Aufgaben erledigt"')
ws["B3"].font = F(bold=True, color=PETROL)
ws.freeze_panes = "B5"

# ---------------------------------------------------------------- Rollen & Zeitplan
zp = wb.create_sheet("Rollen & Zeitplan")
title(zp, "Rollen & Zeitplan", "Jede Folie einzeln vergeben – gern abwechselnd. Teilen sich zwei eine Folie, zweite Person eintragen: die Zeit wird halbiert.")
head(zp, 4, ["Folie", "Titel", "Pflichtteil", "Sprecher:in", "2. Sprecher:in (optional)", "Dauer", "bis (Stoppuhr)", "Übergabe / Notiz"], [7, 40, 13, 17, 17, 9, 13, 34])
plan = [(1, "Primatenbrause (Titel)", "–", 0, 20), (2, "Wer wir sind – in 30 Sekunden", "Einleitung", 0, 45), (3, "Nachhaltigkeit war Teil der Idee", "Einleitung", 0, 45),
        (4, "Produziert, angeboten, verkauft", "Verlauf", 1, 0), (5, "Warum wir so entschieden haben", "Verlauf", 1, 5), (6, "Wirtschaftsdienst: Kosten von außen", "Verlauf", 0, 50),
        (7, "Unsere Kennzahlen je Runde", "Auswertung", 1, 0), (8, "Wirtschaftlichkeit und Schlussbestand", "Auswertung", 0, 55), (9, "Was unser Unternehmen am Ende wert ist", "Auswertung", 0, 50),
        (10, "Reflexion", "Reflexion", 1, 20), (11, "Bezug zur Gründung", "Gründung", 1, 10), (12, "Vielen Dank", "–", 0, 10)]
dv_n2 = DataValidation(type="list", formula1="=Team", allow_blank=True); zp.add_data_validation(dv_n2)
for k, (n, t, p, m, s_) in enumerate(plan):
    r = 5 + k
    zp.cell(r, 1, n); zp.cell(r, 2, t); zp.cell(r, 3, p)
    for col in (4, 5, 8):
        c = zp.cell(r, col); c.fill = INPUT
        if col < 8: dv_n2.add(c.coordinate)
    e = zp.cell(r, 6, time(0, m, s_)); e.number_format = "[m]:ss"
    f = zp.cell(r, 7, f"=SUM($F$5:F{r})"); f.number_format = "[m]:ss"
    for col in range(1, 9):
        c = zp.cell(r, col); c.border = BOX; c.font = F(); c.alignment = Alignment(vertical="center", wrap_text=True)
zp.cell(18, 2, "Summe Sprechzeit (ohne Wechsel)").font = F(bold=True)
z = zp.cell(18, 6, "=SUM(F5:F16)"); z.number_format = "[m]:ss"; z.font = F(bold=True)
zp.cell(19, 2, "Folien ohne Sprecher:in").font = F(bold=True)
zp.cell(19, 6, "=COUNTBLANK(D5:D16)").font = F(bold=True, color=AMBER)
head_r = 21
for i, h in enumerate(["", "Zeit und Folien je Person", "", "Folien", "", "Zeit"]):
    c = zp.cell(head_r, 1 + i, h if h else None)
zp.cell(head_r, 2).font = F(bold=True, color=PETROL)
zp.cell(head_r, 4).font = F(bold=True, color=PETROL); zp.cell(head_r, 6).font = F(bold=True, color=PETROL)
for i in range(4):
    r = head_r + 1 + i; nm = f"Start!$B${5+i}"
    zp.cell(r, 2, f'=IF({nm}="","Person {i+1}",{nm})').font = F()
    zp.cell(r, 4, f'=IF({nm}="",0,COUNTIF($D$5:$D$16,{nm})+COUNTIF($E$5:$E$16,{nm}))').font = F()
    c = zp.cell(r, 6, f'=IF({nm}="",0,SUMPRODUCT((($D$5:$D$16={nm})+($E$5:$E$16={nm}))*$F$5:$F$16/(($E$5:$E$16<>"")+1)))')
    c.number_format = "[m]:ss"; c.font = F()
zp.cell(27, 2, "Grenze laut Aufgabe: 10–12 Minuten. Pro Sprecherwechsel etwa 5 Sekunden mehr einplanen – viele Wechsel kosten Zeit.").font = F(italic=True, color="5F8F99")
zp.cell(28, 2, "Faustregel: Wechsel an Kapitelgrenzen wirken ruhiger; ein Wechsel mitten in einer Folie braucht einen klaren Übergabesatz (Spalte H).").font = F(italic=True, color="5F8F99")
zp.freeze_panes = "D5"

# ---------------------------------------------------------------- Sprechtexte
sp = wb.create_sheet("Sprechtexte")
title(sp, "Sprechtexte", "Vorschlag je Folie (aus den Notizen der Präsentation). In der gelben Spalte schreibt jede:r die eigene Version.")
head(sp, 4, ["Folie", "Sprecher:in", "Vorschlag", "Eigene Version"], [7, 16, 80, 60])
prs = Presentation("Elbharz_Endpraesentation.pptx")
for k, s in enumerate(list(prs.slides)[:12]):
    r = 5 + k
    txt = s.notes_slide.notes_text_frame.text.split("anpassen):\n")[-1].strip()
    sp.cell(r, 1, k + 1)
    sp.cell(r, 2, f"=IF('Rollen & Zeitplan'!D{5+k}=\"\",\"\",IF('Rollen & Zeitplan'!E{5+k}=\"\",'Rollen & Zeitplan'!D{5+k},'Rollen & Zeitplan'!D{5+k}&\" + \"&'Rollen & Zeitplan'!E{5+k}))")
    sp.cell(r, 3, txt)
    sp.cell(r, 4).fill = INPUT
    for col in range(1, 5):
        c = sp.cell(r, col); c.border = BOX; c.font = F(); c.alignment = WRAP
    sp.row_dimensions[r].height = max(60, 15 * (len(txt) // 95 + 2))
sp.freeze_panes = "C5"

# ---------------------------------------------------------------- Fragen
fr = wb.create_sheet("Fragen")
title(fr, "Mögliche Fragen der Lehrkraft", "Antworten sind Vorschläge – in eigenen Worten sagen. Bei „Wer antwortet“ den Namen auswählen.")
head(fr, 4, ["Frage", "Antwortidee", "Zahl zum Belegen", "Wer antwortet"], [42, 70, 24, 16])
fragen = [
    ("Warum habt ihr in Runde 2 eine Maschine ohne Mitarbeiter gekauft?", "Wir wollten für Runde 3 Kapazität aufbauen; Personal kam erst in R3. Im Nachhinein hätten wir beides zusammen planen sollen.", "30.000 €, in R2 ungenutzt"),
    ("Warum in R4 und R5 nur eine Palette – und dann so teuer?", "Wir erwarteten weniger Absatz (Protokoll: „Beliebtheit nicht dauerhaft“, Marktanalyse). Der hohe Preis war ein Fehler: Der Preis bestimmt den Absatz.", "0,5 bzw. 0,2 Paletten verkauft"),
    ("Warum habt ihr das Personal in R4/R5 nicht reduziert?", "Ehrlich: nicht bedacht. Ein Mitarbeiter weniger hätte je Runde 4.000 € gespart.", "8.000 € für R4 + R5"),
    ("War der Preis in R1 zu niedrig?", "Vermutlich: Alles war sofort weg. In R3 haben wir zu 5.900 € noch 70 % am Markt verkauft.", "3.000 € vs. 5.900 €"),
    ("Wie habt ihr die Wirtschaftlichkeit berechnet?", "Umsatzerlöse geteilt durch alle Kosten der Runde aus dem Kassenbuch.", "R3: 79.940 ÷ 54.000 = 1,48"),
    ("Was sagt die Absatzquote aus?", "Wie viel Prozent der Produktion am normalen Markt verkauft wurde – also ob der Preis zum Markt passte.", "100 / 50 / 70 / 50 / 20 / 40 %"),
    ("Wie kommt ihr auf 18.000 € Restwert für Maschine 2?", "10 % Wertverlust je Runde Nutzungsdauer; sie lief nur R3–R6, also 4 Runden.", "30.000 − 4 × 3.000"),
    ("Was war der Wirtschaftsdienst bei euch?", "Zusatzkosten in R2–R6, im Kassenbuch als sonstige Kosten. [Wortlaut der Meldungen ergänzen]", "34.000 € gesamt"),
    ("Warum liegt der Verkaufswert unter dem Startkapital?", "Zwei Maschinen für 60.000 €, nur zwei wirtschaftliche Runden, hohe Fixkosten in R4/R5 und 34.000 € Wirtschaftsdienst.", "−16.230 €"),
    ("Protokoll sagt „keine neuen Mitarbeiter“ in R3 – das Arbeitsblatt zeigt 2?", "[Gemeinsam klären, welche Angabe stimmt]", "8.000 € Personal in R3"),
    ("Wie würde eure Spielerweiterung funktionieren?", "Marketingbudget: je 2.000 € → +1 Palette Absatz. Mehrweg-Option: +1.000 € Kosten je Runde → +1 Palette Absatz. Nur eine Zeile mehr im Kassenbuch.", "–"),
]
dv_n3 = DataValidation(type="list", formula1="=Team", allow_blank=True); fr.add_data_validation(dv_n3)
for k, (q, a, z_) in enumerate(fragen):
    r = 5 + k
    fr.cell(r, 1, q); fr.cell(r, 2, a); fr.cell(r, 3, z_)
    d = fr.cell(r, 4); d.fill = INPUT; dv_n3.add(d.coordinate)
    for col in range(1, 5):
        c = fr.cell(r, col); c.border = BOX; c.font = F(bold=(col == 1)); c.alignment = WRAP
    fr.row_dimensions[r].height = 46

st.merge_cells("A4:B4")
for sh in wb.worksheets:
    sh.sheet_view.showGridLines = False
    sh.page_setup.orientation = "landscape"; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.sheet_properties.tabColor = {"Start": PETROL, "Aufgaben": AMBER}.get(sh.title, LIME)
wb.active = 1
wb.save("Team_Leitfaden_Elbharz.xlsx")
print("ok")
