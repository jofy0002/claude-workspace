const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/4479ee3f-de1e-4616-9349-269c5bd53f97_e0f7022a-26aa-4f2e-8a7c-986f75ea937c/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Elbharz",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2A2F", lt1: "FFFFFF", dk2: "0F5C6B", lt2: "F3F1EA",
    accent1: "0F5C6B", accent2: "D9822B", accent3: "8DB33A",
    accent4: "5F8F99", accent5: "B8B8B0", accent6: "7A4B1E",
    hlink: "0F5C6B", folHlink: "5F8F99",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Elbharz Getränke GbR – Endpräsentation Planspiel Bertis Brause";
const C = pres.SchemeColor;

const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");
const ICONS = {};
async function loadIcons(names) {
  for (const n of names) {
    const svg = RDS.renderToStaticMarkup(React.createElement(FA[n], { color: "#FFFFFF", size: 256 }));
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    ICONS[n] = "image/png;base64," + buf.toString("base64");
  }
}
function iconCircle(s, name, x, y, d, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: "Icon-Kreis " + name });
  const p = d * 0.28;
  s.addImage({ data: ICONS[name], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: "Symbol", objectName: "Icon " + name });
}

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.4, w: 11.9, h: 0.9, fontSize: 36, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "PRIMATENBRAUSE · ELBHARZ GETRÄNKE GBR", options: { x: 0.7, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.0, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text2 },
  objects: [],
});

let sec = null;
function slide(master, section) {
  if (section && section !== sec) { pres.addSection({ title: section }); sec = section; }
  return pres.addSlide({ masterName: master, sectionTitle: section });
}
function content(title, section, notes) {
  const s = slide("CONTENT", section);
  s.addText(title, { placeholder: "title" });
  s.addNotes(notes);
  return s;
}
// Platzhalter-Kasten für Inhalte, die später folgen
function todo(s, x, y, w, h, label, hint) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { color: C.accent5, width: 1.25, dashType: "dash" }, objectName: "Platzhalter" });
  s.addText([
    { text: label, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
    { text: hint, options: { fontSize: 14, color: C.text1 } },
  ], { x: x + 0.25, y, w: w - 0.5, h, valign: "middle", isTextBox: true, margin: 0 });
}
function card(s, x, y, w, h, head, body, accent, icon) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Karte " + head });
  iconCircle(s, icon, x + 0.3, y + 0.3, 0.7, accent || C.accent1);
  s.addText(head, { x: x + 1.2, y: y + 0.3, w: w - 1.5, h: 0.7, fontSize: 20, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
  s.addText(body, { x: x + 0.3, y: y + 1.3, w: w - 0.6, h: h - 1.5, fontSize: 16, color: C.text1, valign: "top", paraSpaceAfter: 6, isTextBox: true, margin: 0 });
}

function build() {
// 1 Titel
{
  const s = slide("DARK", "Einleitung");
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.2, w: 3.4, h: 3.4, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  s.addShape(pres.ShapeType.ellipse, { x: 8.7, y: 4.1, w: 1.3, h: 1.3, fill: { color: C.accent3, transparency: 25 }, line: { type: "none" }, objectName: "Punkt grün" });
  s.addText("PLANSPIEL BERTIS BRAUSE · ENDPRÄSENTATION", { x: 0.9, y: 1.2, w: 11, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("Primatenbrause", { x: 0.9, y: 2.0, w: 11, h: 1.4, fontSize: 64, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Elbharz Getränke GbR – unser Weg durch sechs Spielrunden", { x: 0.9, y: 3.5, w: 7.6, h: 1.2, fontSize: 24, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Elbharz Getränke GbR · Klasse 11b", { x: 0.9, y: 5.6, w: 11, h: 0.4, fontSize: 16, color: C.accent3, isTextBox: true, margin: 0 });
  s.addNotes("ca. 0:30 min. Gruppe nennen, kurz sagen was kommt: Einleitung, Rundenverlauf, Auswertung, Reflexion, Bezug zur Gründung.");
}

// 2 Einleitung
{
  const s = content("Wer wir sind und warum wir gegründet haben", "Einleitung",
    "ca. 1 min. Hauptaspekte der Firma + Gründungsintention. Quelle: Firmenpräsentation Folie 2–4.");
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 1.6, w: 5.2, h: 4.9, rectRadius: 0.12, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Leitgedanke" });
  s.addText([
    { text: "Ein Getränk, dessen Rückseite man zu Ende lesen kann.", options: { fontSize: 32, bold: true, fontFace: "Cambria", color: C.background1, breakLine: true, paraSpaceAfter: 14 } },
    { text: "Unsere einzige Vorgabe – daran haben wir alles andere gemessen.", options: { fontSize: 16, color: C.background2 } },
  ], { x: 1.1, y: 1.9, w: 4.4, h: 4.3, valign: "middle", isTextBox: true, margin: 0 });
  const rows = [
    ["FaBuilding", "Das Unternehmen", "Elbharz Getränke GbR, Hamburg, 4 Gesellschafter. Produkt: Primatenbrause (Tamarinde-Limette)", C.accent1],
    ["FaBullseye", "Die Zielgruppe", "16–25 Jahre, liest Etiketten aus Misstrauen – Slogan: „Was dein Körper schon immer wollte.“", C.accent2],
    ["FaStore", "Der Vertrieb", "Onlineshop und Social Media, dann regionaler Handel, dazu Messen und Märkte", C.accent3],
  ];
  rows.forEach(([ic, h, t, col], i) => {
    const y = 1.6 + i * 1.7;
    s.addShape(pres.ShapeType.roundRect, { x: 6.2, y, w: 6.4, h: 1.5, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Karte " + h });
    iconCircle(s, ic, 6.5, y + 0.4, 0.7, col);
    s.addText([
      { text: h, options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } },
      { text: t, options: { fontSize: 16, color: C.text1 } },
    ], { x: 7.45, y, w: 4.9, h: 1.5, valign: "middle", isTextBox: true, margin: 0 });
  });
}

// 3 Nachhaltigkeit
{
  const s = content("Nachhaltigkeit von Anfang an", "Einleitung",
    "ca. 1 min. Die Aufgabe fragt ausdrücklich nach Nachhaltigkeit bei der Gründung. Mehrweg lohnt nur bei kurzen Wegen -> regional.");
  card(s, 0.7, 1.6, 3.8, 3.2, "Mehrweg-Glas", "Pfandsystem statt Einweg", C.accent3, "FaRecycle");
  card(s, 4.75, 1.6, 3.8, 3.2, "Regional", "Hier gebraut, hier abgefüllt – kurze Wege, damit Mehrweg ökologisch wirklich besser ist", C.accent3, "FaMapMarkerAlt");
  card(s, 8.8, 1.6, 3.8, 3.2, "Eine Sorte", "Weniger Produktion, Lager und Transport", C.accent3, "FaLeaf");
  todo(s, 0.7, 5.1, 11.9, 1.5, "Verknüpfung zum Spiel", "Wie floss Nachhaltigkeit im Spiel ein? (z. B. Maschinenwahl, Produktionsmenge, Preis) – später ergänzen");
}

// 4 Verlauf der Runden
{
  const s = content("Verlauf der Runden im Überblick", "Verlauf der Runden",
    "ca. 1,5 min. Pro Runde ein Satz: Was haben wir entschieden und was kam dabei raus? Details auf der nächsten Folie.");
  const months = ["Jan./Feb.", "März/Apr.", "Mai/Juni", "Juli/Aug.", "Sep./Okt.", "Nov./Dez."];
  const w = 1.8, gap = 0.22;
  for (let i = 0; i < 6; i++) {
    const x = 0.7 + i * (w + gap);
    s.addShape(pres.ShapeType.ellipse, { x: x + w / 2 - 0.3, y: 1.7, w: 0.6, h: 0.6, fill: { color: i === 5 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "Runde " + (i + 1) });
    s.addText(String(i + 1), { x: x + w / 2 - 0.3, y: 1.7, w: 0.6, h: 0.6, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.background1, isTextBox: true, margin: 0 });
    s.addText("Runde " + (i + 1) + "\n" + months[i], { x, y: 2.4, w, h: 0.7, align: "center", fontSize: 14, bold: true, color: C.text2, isTextBox: true, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.3, w, h: 3.2, rectRadius: 0.1, fill: { color: C.background2 }, line: { color: C.accent5, width: 1, dashType: "dash" }, objectName: "Platzhalter Runde " + (i + 1) });
    s.addText("Entscheidung\n\nErgebnis", { x: x + 0.15, y: 3.3, w: w - 0.3, h: 3.2, fontSize: 14, color: C.accent4, valign: "middle", align: "center", isTextBox: true, margin: 0 });
  }
}

// 5 Entscheidungen (Tabelle)
{
  const s = content("Unsere Entscheidungen je Runde", "Verlauf der Runden",
    "ca. 1,5 min. Nur die 2–3 wichtigsten Entscheidungen erklären, Rest der Tabelle für sich sprechen lassen. Quelle: Entscheidungsprotokoll + Excel.");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0F5C6B" }, align: "center" } });
  const rows = [[hdr("Runde"), hdr("Maschinen"), hdr("Mitarbeiter"), hdr("Paletten produziert"), hdr("Preis / Palette"), hdr("Begründung")]];
  for (let i = 1; i <= 6; i++) rows.push([{ text: String(i), options: { bold: true, align: "center" } }, "–", "–", "–", "–", "–"]);
  s.addTable(rows, { x: 0.7, y: 1.6, w: 11.9, colW: [1.1, 1.9, 1.7, 2.1, 1.8, 3.3], fontSize: 14, color: "1B2A2F", border: { type: "solid", color: "B8B8B0", pt: 0.75 }, rowH: 0.6, valign: "middle", objectName: "Tabelle Entscheidungen" });
  s.addText("Werte aus Entscheidungsprotokoll und Excel eintragen", { x: 0.7, y: 6.3, w: 11.9, h: 0.4, fontSize: 14, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
}

// 6 Wirtschaftsdienst
{
  const s = content("Wirtschaftsdienst: was von außen auf uns wirkte", "Verlauf der Runden",
    "ca. 1 min. Ab Runde 2 gab es Nachrichten mit Auswirkungen auf Kosten und Erfolg. Pro Nachricht: Wie haben wir reagiert, was hat es am Unternehmenswert verändert?");
  todo(s, 0.7, 1.6, 5.8, 4.9, "Nachricht(en) je Runde", "Runde 2 bis 6: Meldung des Wirtschaftsdienstes in einem Satz");
  todo(s, 6.8, 1.6, 5.8, 4.9, "Unsere Reaktion & Wirkung", "Was haben wir geändert, und wie hat sich der Unternehmenswert dadurch verändert?");
}

// 7 Kennzahlen
{
  const s = content("Auswertung: Kennzahlen pro Runde", "Auswertung",
    "ca. 1,5 min. Aufgabe: Welche Kennzahl war in welcher Runde wie hoch, wie berechnet, wie aussagekräftig? Rechenweg: Schlussbestand = Anfangsbestand − Maschinen − Personal − Miete − Produktion − Material + Umsatzerlöse.");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0F5C6B" }, align: "center" } });
  const rows = [[hdr("Kennzahl"), hdr("R1"), hdr("R2"), hdr("R3"), hdr("R4"), hdr("R5"), hdr("R6")]];
  for (const k of ["Umsatzerlöse (€)", "Gesamtkosten (€)", "Gewinn der Runde (€)", "Kassenbestand Ende (€)", "Abgesetzte Paletten", "Auslastung (Absatz / Produktion)"])
    rows.push([{ text: k, options: { bold: true } }, "–", "–", "–", "–", "–", "–"]);
  s.addTable(rows, { x: 0.7, y: 1.6, w: 7.9, colW: [3.1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8], fontSize: 14, color: "1B2A2F", border: { type: "solid", color: "B8B8B0", pt: 0.75 }, rowH: 0.62, valign: "middle", objectName: "Tabelle Kennzahlen" });
  card(s, 8.9, 1.6, 3.7, 4.6, "So haben wir gerechnet", "Umsatz = Menge × Preis (Rest am Schnäppchenmarkt: 1.000 €/Palette)\nKosten = Maschinen, Personal, Miete, Produktion, Material\nSchlussbestand = Anfangsbestand − Kosten + Umsatz", C.accent1, "FaChartLine");
  s.addText("Aussagekraft der Kennzahlen: später ergänzen", { x: 0.7, y: 6.3, w: 7.9, h: 0.4, fontSize: 14, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
}

// 8 Diagramme
{
  const s = content("Auswertung: Entwicklung über die Runden", "Auswertung",
    "ca. 1,5 min. Zwei Diagramme: links Umsatz/Kosten/Gewinn je Runde, rechts Kassenbestand. Je Diagramm eine Aussage, z. B. wo der Knick war und warum.");
  todo(s, 0.7, 1.6, 5.8, 4.2, "Diagramm: Umsatz, Kosten, Gewinn", "Säulen je Runde (R1–R6)");
  todo(s, 6.8, 1.6, 5.8, 4.2, "Diagramm: Kassenbestand", "Linie je Runde, Start 100.000 €");
  s.addText("Kernaussage:", { x: 0.7, y: 6.0, w: 1.6, h: 0.5, fontSize: 16, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
  s.addText("Ein Satz zur Entwicklung – später ergänzen", { x: 2.3, y: 6.0, w: 10.3, h: 0.5, fontSize: 16, color: C.text1, isTextBox: true, margin: 0 });
}

// 9 Endbestand
{
  const s = content("Endbestand und Unternehmensverkaufswert", "Auswertung",
    "ca. 1 min. Verkaufswert = Schlussbestand nach Runde 6 + Restwert der Maschinen (10 % Wertverlust je Runde auf den Anschaffungswert).");
  const stat = (x, num, lbl, col) => {
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.6, w: 3.8, h: 2.6, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Kennzahl " + lbl });
    s.addText(num, { x, y: 1.85, w: 3.8, h: 1.2, align: "center", valign: "middle", fontSize: 48, bold: true, fontFace: "Cambria", color: col, isTextBox: true, margin: 0 });
    s.addText(lbl, { x: x + 0.2, y: 3.15, w: 3.4, h: 0.8, align: "center", fontSize: 16, color: C.text1, isTextBox: true, margin: 0 });
  };
  stat(0.7, "– €", "Schlussbestand nach Runde 6", C.text2);
  stat(4.75, "– €", "Restwert der Maschinen", C.text2);
  stat(8.8, "– €", "Unternehmensverkaufswert", C.accent2);
  todo(s, 0.7, 4.6, 11.9, 1.9, "Einordnung", "Platzierung im Vergleich zu den anderen Gruppen? Was hat am meisten zum Ergebnis beigetragen?");
}

// 10 Reflexion
{
  const s = content("Reflexion: Würden wir wieder so entscheiden?", "Reflexion",
    "ca. 1,5 min. Begründet Stellung nehmen und Handlungsalternativen zeigen (Aufgabentext). Pro Entscheidung: Bewertung + was wir stattdessen tun könnten.");
  card(s, 0.7, 1.6, 3.8, 3.9, "Das lief gut", "– Entscheidung 1: Begründung\n– Entscheidung 2: Begründung", C.accent3, "FaThumbsUp");
  card(s, 4.75, 1.6, 3.8, 3.9, "Das lief schlechter", "– Entscheidung 1: Begründung\n– Entscheidung 2: Begründung", C.accent2, "FaThumbsDown");
  card(s, 8.8, 1.6, 3.8, 3.9, "Alternativen", "– Statt X hätten wir Y tun können\n– Erwartete Wirkung auf den Unternehmenswert", C.accent1, "FaExchangeAlt");
  todo(s, 0.7, 5.75, 11.9, 0.95, "Gesamtfazit", "Würden wir als Gruppe insgesamt wieder so vorgehen? Ein Satz");
}

// 11 Bezug zur Gründung
{
  const s = content("Bezug zur Gründung: was im Spiel fehlte", "Bezug zur Gründung",
    "ca. 1,5 min. Aufgabe: Welche Gründungsideen fanden im Spiel keine Berücksichtigung, welche Auswirkung hätte das gehabt, und wie könnte man das Spiel erweitern, ohne es zu komplex zu machen?");
  card(s, 0.7, 1.6, 3.8, 3.9, "Nicht im Spiel", "– Mehrweg-Glas und Pfand\n– Social-Media-Marketing / Onlineshop\n– Nur eine Sorte, Tamarinde\n– Zielgruppe, Vertriebsweg", C.accent1, "FaBan");
  card(s, 4.75, 1.6, 3.8, 3.9, "Mögliche Auswirkung", "Wie hätte sich das auf Kosten, Absatz und Preis ausgewirkt?\n(später mit euren Zahlen untermauern)", C.accent1, "FaChartLine");
  card(s, 8.8, 1.6, 3.8, 3.9, "Idee zur Spielerweiterung", "Konkreter, einfacher Vorschlag, z. B. eine zusätzliche Entscheidung pro Runde (Marketingbudget), ohne das Rechnen aufzublähen", C.accent2, "FaLightbulb");
  todo(s, 0.7, 5.75, 11.9, 0.95, "Kernbotschaft", "Was hätte die Gründungsidee im Spiel verändert? Ein Satz");
}

// 12 Danke
{
  const s = slide("DARK", "Abschluss");
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.9, w: 3.0, h: 3.0, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  s.addText("Vielen Dank", { x: 0.9, y: 2.4, w: 11, h: 1.3, fontSize: 60, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Wir freuen uns über eure Fragen.", { x: 0.9, y: 3.8, w: 11, h: 0.6, fontSize: 24, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Quellen: Planspiel Bertis Brause (Cornelsen), eigene Berechnungen aus Excel und Entscheidungsprotokoll", { x: 0.9, y: 6.3, w: 11.5, h: 0.4, fontSize: 12, color: C.accent3, isTextBox: true, margin: 0 });
  s.addNotes("Stehen lassen, solange Fragen kommen.");
}

}

(async () => {
  await loadIcons(["FaBuilding","FaLightbulb","FaBullseye","FaRecycle","FaMapMarkerAlt","FaLeaf","FaThumbsUp","FaThumbsDown","FaExchangeAlt","FaBan","FaChartLine","FaBullhorn","FaStore"]);
  build();
  await pres.writeFile({ fileName: "Elbharz_Endpraesentation_Geruest.pptx" });
  await applyTheme("Elbharz_Endpraesentation_Geruest.pptx", THEME);
})();
