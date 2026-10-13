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
const NOTE = (t, txt) => "ZEIT: " + t + "\n\nSPRECHTEXT (Entwurf, bitte persönlich anpassen):\n" + txt;
const HEXP = "0F5C6B", HEXA = "D9822B", HEXT = "1B2A2F";
const chartBase = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", catAxisLabelColor: HEXT, valAxisLabelColor: HEXT, legendColor: HEXT, titleColor: HEXT, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, legendFontSize: 12, titleFontSize: 14, showTitle: true, valGridLine: { color: "D9D9D2", size: 0.5 }, catGridLine: { style: "none" }, valAxisLabelFormatCode: "#,##0" };

// 1 Titel
{
  const s = slide("DARK", "Einleitung");
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.2, w: 3.4, h: 3.4, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  s.addShape(pres.ShapeType.ellipse, { x: 8.7, y: 4.1, w: 1.3, h: 1.3, fill: { color: C.accent3, transparency: 25 }, line: { type: "none" }, objectName: "Punkt grün" });
  s.addText("PLANSPIEL BERTIS BRAUSE · ENDPRÄSENTATION", { x: 0.9, y: 1.2, w: 8, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("Primatenbrause", { x: 0.9, y: 2.0, w: 8, h: 1.4, fontSize: 64, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Elbharz Getränke GbR – unser Weg durch sechs Spielrunden", { x: 0.9, y: 3.5, w: 7.6, h: 1.2, fontSize: 24, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Elbharz Getränke GbR · Klasse 11b", { x: 0.9, y: 5.6, w: 11, h: 0.4, fontSize: 16, color: C.accent3, isTextBox: true, margin: 0 });
  s.addNotes(NOTE("ca. 0:30 min", "Guten Tag, wir sind die Elbharz Getränke GbR aus der 11b. Im Planspiel „Bertis Brause“ haben wir sechs Runden lang ein Brauseunternehmen geführt. Heute zeigen wir euch unseren Weg: von der Gründungsidee über die Entscheidungen in den Runden bis zur Auswertung und zur Reflexion."));
}

// 2 Einleitung
{
  const s = content("Wer wir sind und warum wir gegründet haben", "Einleitung",
    NOTE("ca. 1 min", "Wir sind die Elbharz Getränke GbR aus Hamburg, vier Gesellschafter, und stellen die Primatenbrause her, in der Sorte Tamarinde-Limette. Gegründet haben wir mit einer Beobachtung: Fast jedes Getränk im Regal verspricht irgendetwas, und die Zutatenliste liest niemand zu Ende. Wir wollten das Gegenteil, ein Getränk, dessen Rückseite man zu Ende lesen kann. Unsere Zielgruppe sind 16- bis 25-Jährige, die Etiketten lesen. Verkaufen wollen wir zuerst über Onlineshop und Social Media, dann über den regionalen Handel, dazu Messen und Märkte."));
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
    NOTE("ca. 1 min", "Nachhaltigkeit war uns schon bei der Gründung wichtig. Wir füllen in Mehrweg-Glas im Pfandsystem ab. Das ist aber nur dann wirklich besser als Einweg, wenn die Wege kurz sind, denn Glas ist schwer. Deshalb produzieren und liefern wir regional. Und wir starten mit nur einer einzigen Sorte, das spart Produktion, Lager und Transport. Im Planspiel selbst kam Nachhaltigkeit als Entscheidung nicht vor, dazu kommen wir am Ende noch."));
  card(s, 0.7, 1.6, 3.8, 3.2, "Mehrweg-Glas", "Pfandsystem statt Einweg", C.accent3, "FaRecycle");
  card(s, 4.75, 1.6, 3.8, 3.2, "Regional", "Hier gebraut, hier abgefüllt – kurze Wege, damit Mehrweg ökologisch wirklich besser ist", C.accent3, "FaMapMarkerAlt");
  card(s, 8.8, 1.6, 3.8, 3.2, "Eine Sorte", "Weniger Produktion, Lager und Transport", C.accent3, "FaLeaf");
  todo(s, 0.7, 5.1, 11.9, 1.5, "Verknüpfung zum Spiel", "Wie floss Nachhaltigkeit im Spiel ein? – gemeinsam entscheiden und ergänzen");
}

// 4 Verlauf der Runden
{
  const s = content("Verlauf der Runden im Überblick", "Verlauf der Runden",
    NOTE("ca. 1:30 min", "Runde 1: Eine Economy-Maschine, 15 Paletten zu 3.000 Euro. Alles wurde verkauft, 45.000 Euro Umsatz, der Preis war aber offenbar zu niedrig. Runde 2: Wir haben den Preis drastisch auf 7.600 Euro erhöht und eine zweite Maschine gekauft. Nur die Hälfte ging am Markt weg, der Rest zum Schnäppchenpreis. Runde 3: Wir sind auf 5.900 Euro runtergegangen, das war unsere beste Runde mit 79.940 Euro. In Runde 4 und 5 haben wir weniger Absatz erwartet und nur noch eine Palette zu 9.500 Euro angeboten, der Umsatz war fast null. In Runde 6 haben wir noch einmal 13 Paletten produziert und alles verkauft."));
  const months = ["Jan./Feb.", "März/Apr.", "Mai/Juni", "Juli/Aug.", "Sep./Okt.", "Nov./Dez."];
  const dec = ["1 Maschine\n15 Paletten\nà 3.000 €", "2. Maschine\n13 Paletten\nà 7.600 €", "18 Paletten\nà 5.900 €", "1 Palette\nà 9.500 €", "1 Palette\nà 9.500 €", "13 Paletten\nà 8.400 €"];
  const res = ["45.000 €", "55.900 €", "79.940 €", "5.250 €", "2.700 €", "51.480 €"];
  const w = 1.8, gap = 0.22;
  for (let i = 0; i < 6; i++) {
    const x = 0.7 + i * (w + gap);
    const best = i === 2;
    s.addShape(pres.ShapeType.ellipse, { x: x + w / 2 - 0.3, y: 1.7, w: 0.6, h: 0.6, fill: { color: best ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "Runde " + (i + 1) });
    s.addText(String(i + 1), { x: x + w / 2 - 0.3, y: 1.7, w: 0.6, h: 0.6, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.background1, isTextBox: true, margin: 0 });
    s.addText("Runde " + (i + 1) + "\n" + months[i], { x, y: 2.4, w, h: 0.7, align: "center", fontSize: 14, bold: true, color: C.text2, isTextBox: true, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.3, w, h: 3.2, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Karte Runde " + (i + 1) });
    s.addText("Entscheidung", { x: x + 0.15, y: 3.45, w: w - 0.3, h: 0.3, fontSize: 12, color: C.accent4, align: "center", isTextBox: true, margin: 0 });
    s.addText(dec[i], { x: x + 0.1, y: 3.8, w: w - 0.2, h: 1.2, fontSize: 14, color: C.text1, align: "center", valign: "top", isTextBox: true, margin: 0 });
    s.addText("Umsatz", { x: x + 0.15, y: 5.2, w: w - 0.3, h: 0.3, fontSize: 12, color: C.accent4, align: "center", isTextBox: true, margin: 0 });
    s.addText(res[i], { x: x + 0.1, y: 5.55, w: w - 0.2, h: 0.6, fontSize: 20, bold: true, color: best ? C.accent2 : C.text2, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  }
}

// 5 Entscheidungen (Tabelle)
{
  const s = content("Unsere Entscheidungen je Runde", "Verlauf der Runden",
    NOTE("ca. 1:30 min", "Hier sehen wir alle Entscheidungen auf einen Blick. Beim Maschinenkauf haben wir uns für Economy entschieden, weil sich Standard mit dem zusätzlichen Mitarbeiter nicht gelohnt hätte. Beim Preis haben wir viel ausprobiert: 3.000 Euro in Runde 1, dann 7.600, dann 5.900. Dazwischen haben wir immer aus dem Markt gelernt. In Runde 4 und 5 haben wir die Menge stark zurückgefahren, weil wir mit weniger Absatz gerechnet haben, und in Runde 6 wollten wir vor dem Abschluss noch einmal alles verkaufen."));
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0F5C6B" }, align: "center" } });
  const rows = [[hdr("Runde"), hdr("Maschinen"), hdr("Mitarb."), hdr("Paletten"), hdr("Preis"), hdr("Begründung")]];
  const data = [
    ["1", "1", "1", "15", "3.000 €", "Standard lohnt sich mit Extra-Mitarbeiter nicht; niedriger Preis, um alles loszuwerden"],
    ["2", "2", "1", "13", "7.600 €", "Preis drastisch erhöht; zweite Maschine für mehr Produktion"],
    ["3", "2", "2", "18", "5.900 €", "Letzter Preis war zu hoch; Annahme: Konkurrenz liegt bei ca. 6.000 €"],
    ["4", "2", "2", "1", "9.500 €", "Weniger Absatz erwartet; Minus begrenzen"],
    ["5", "2", "2", "1", "9.500 €", "Marktanalyse sagt weniger Absatz voraus"],
    ["6", "2", "2", "13", "8.400 €", "Alles verkaufen und vor Abschluss Umsatz machen"],
  ];
  data.forEach(r => rows.push(r.map((t, i) => ({ text: t, options: { bold: i === 0, align: i < 5 ? "center" : "left" } }))));
  s.addTable(rows, { x: 0.7, y: 1.6, w: 11.9, colW: [0.9, 1.3, 1.1, 1.2, 1.2, 6.2], fontSize: 14, color: "1B2A2F", border: { type: "solid", color: "B8B8B0", pt: 0.75 }, rowH: [0.5, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7], valign: "middle", objectName: "Tabelle Entscheidungen" });
}

// 6 Wirtschaftsdienst
{
  const s = content("Wirtschaftsdienst: was von außen auf uns wirkte", "Verlauf der Runden",
    NOTE("ca. 1 min", "[Noch offen: Nachrichten des Wirtschaftsdienstes für Runde 2 bis 6 einfügen.] Pro Nachricht: Was wurde gemeldet, wie haben wir reagiert, wie hat sich das auf den Unternehmenswert ausgewirkt? Hinweis: In unserem Kassenbuch stehen sonstige Kosten von 10.000 € (R2), 5.000 € (R3), 2.500 € (R4), 10.000 € (R5) und 6.500 € (R6); vermutlich hängen sie mit dem Wirtschaftsdienst zusammen."));
  todo(s, 0.7, 1.6, 5.8, 4.9, "Nachricht(en) je Runde", "Runde 2 bis 6: Meldung des Wirtschaftsdienstes in einem Satz");
  todo(s, 6.8, 1.6, 5.8, 4.9, "Unsere Reaktion & Wirkung", "Sonstige Kosten laut Kassenbuch: R2 10.000 €, R3 5.000 €, R4 2.500 €, R5 10.000 €, R6 6.500 € – woher kamen sie?");
}

// 7 Kennzahlen
{
  const s = content("Auswertung: Kennzahlen pro Runde", "Auswertung",
    NOTE("ca. 1:30 min", "Wir haben pro Runde Umsatz, Kosten und Gewinn berechnet. Der Umsatz ist Menge mal Preis, dazu kommt, was wir am Schnäppchenmarkt für 1.000 Euro pro Palette losgeworden sind. Die Kosten sind Maschinen, Personal, Miete, sonstige Kosten, Produktion und Material. Gewinn ist Umsatz minus Kosten. Zur Aussagekraft: In Runde 1 und 2 stecken die Maschinenkäufe von zusammen 60.000 Euro, deshalb sind diese Runden negativ, obwohl der laufende Betrieb in Runde 1 und 2 im Plus gewesen wäre. Positiv waren nur Runde 3 und Runde 6."));
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: "0F5C6B" }, align: "center" } });
  const rows = [[hdr("Kennzahl in €"), hdr("R1"), hdr("R2"), hdr("R3"), hdr("R4"), hdr("R5"), hdr("R6")]];
  const data = [
    ["Umsatz", "45.000", "55.900", "79.940", "5.250", "2.700", "51.480"],
    ["Gesamtkosten", "69.000", "75.000", "54.000", "18.000", "25.000", "45.500"],
    ["Gewinn / Verlust", "−24.000", "−19.100", "+25.940", "−12.750", "−22.300", "+5.980"],
    ["Kassenbestand", "76.000", "56.900", "82.840", "70.090", "47.790", "53.770"],
    ["Paletten am Markt", "15", "6,5", "12,6", "0,5", "0,2", "5,2"],
    ["Paletten Schnäppchen", "0", "6,5", "5,6", "0,5", "0,8", "7,8"],
  ];
  data.forEach(r => rows.push(r.map((t, i) => {
    const o = { bold: i === 0, align: i === 0 ? "left" : "right" };
    if (i > 0 && r[0] === "Gewinn / Verlust") o.color = t.startsWith("+") ? "3F6B12" : "9A3B1C";
    return { text: t, options: o };
  })));
  s.addTable(rows, { x: 0.7, y: 1.6, w: 8.5, colW: [2.5, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0], fontSize: 14, color: "1B2A2F", border: { type: "solid", color: "B8B8B0", pt: 0.75 }, rowH: 0.62, valign: "middle", objectName: "Tabelle Kennzahlen" });
  card(s, 9.5, 1.6, 3.1, 4.6, "So haben wir gerechnet", "Umsatz = Menge × Preis\nKosten = Maschinen, Personal, Miete, sonstige, Produktion, Material\nGewinn = Umsatz − Kosten", C.accent1, "FaChartLine");
  s.addText("Gesamtkosten R1 und R2 enthalten je 30.000 € Maschinenkauf.", { x: 0.7, y: 6.1, w: 8.5, h: 0.4, fontSize: 14, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
}

// 8 Diagramme
{
  const s = content("Auswertung: Entwicklung über die Runden", "Auswertung",
    NOTE("ca. 1:30 min", "Links sieht man Umsatz und Kosten je Runde. Nur in Runde 3 und Runde 6 lag der Umsatz über den Kosten. In Runde 4 und 5 sind die Kosten niedrig, weil wir kaum produziert haben, aber Personal und Miete liefen weiter. Rechts sieht man unseren Kassenbestand. Wir sind mit 100.000 Euro gestartet, der Tiefpunkt war Runde 5 mit 47.790 Euro, und am Ende stehen 53.770 Euro."));
  const L = ["R1", "R2", "R3", "R4", "R5", "R6"];
  s.addChart(pres.charts.BAR, [
    { name: "Umsatz", labels: L, values: [45000, 55900, 79940, 5250, 2700, 51480] },
    { name: "Kosten", labels: L, values: [69000, 75000, 54000, 18000, 25000, 45500] },
  ], Object.assign({}, chartBase, { x: 0.7, y: 1.5, w: 5.9, h: 4.5, barDir: "col", barGapWidthPct: 60, chartColors: [HEXP, HEXA], title: "Umsatz und Kosten je Runde (€)", showLegend: true, legendPos: "b" }));
  s.addChart(pres.charts.LINE, [
    { name: "Kassenbestand", labels: ["Start", ...L], values: [100000, 76000, 56900, 82840, 70090, 47790, 53770] },
  ], Object.assign({}, chartBase, { x: 6.8, y: 1.5, w: 5.8, h: 4.5, chartColors: [HEXP], lineSize: 3, lineDataSymbolSize: 9, title: "Kassenbestand am Rundenende (€)", showLegend: false, valAxisMinVal: 0 }));
  s.addText([
    { text: "Kernaussage: ", options: { bold: true, color: C.accent2 } },
    { text: "Nur in Runde 3 und 6 lag der Umsatz über den Kosten (R1/R2 enthalten 60.000 € Maschinenkäufe).", options: { color: C.text1 } },
  ], { x: 0.7, y: 6.15, w: 11.9, h: 0.6, fontSize: 16, isTextBox: true, margin: 0, valign: "middle" });
}

// 9 Endbestand
{
  const s = content("Endbestand und Unternehmensverkaufswert", "Auswertung",
    NOTE("ca. 1 min", "Unser Endbestand nach Runde 6 beträgt 53.770 Euro. Dazu kommt der Restwert unserer beiden Economy-Maschinen mit zusammen 30.000 Euro. Daraus ergibt sich ein Unternehmensverkaufswert von 83.770 Euro. Das sind 16.230 Euro weniger als unser Startkapital von 100.000 Euro. Eine Platzierung im Vergleich zu den anderen Gruppen gibt es noch nicht. [PRÜFEN: Die zweite Maschine wurde in Runde 2 gekauft, nach den Spielregeln zählt die Kaufrunde mit, das wären 5 Nutzungsrunden und 15.000 Euro Restwert statt 18.000 Euro, also 80.770 Euro Gesamtwert. Bitte mit der Lehrkraft klären.]"));
  const stat = (x, num, lbl, col) => {
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.6, w: 3.8, h: 2.6, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Kennzahl " + lbl });
    s.addText(num, { x, y: 1.85, w: 3.8, h: 1.2, align: "center", valign: "middle", fontSize: 48, bold: true, fontFace: "Cambria", color: col, isTextBox: true, margin: 0 });
    s.addText(lbl, { x: x + 0.2, y: 3.15, w: 3.4, h: 0.8, align: "center", fontSize: 16, color: C.text1, isTextBox: true, margin: 0 });
  };
  stat(0.7, "53.770 €", "Schlussbestand nach Runde 6", C.text2);
  stat(4.75, "30.000 €", "Restwert der zwei Maschinen", C.text2);
  stat(8.8, "83.770 €", "Unternehmensverkaufswert", C.accent2);
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 4.6, w: 11.9, h: 1.9, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Einordnung" });
  iconCircle(s, "FaChartLine", 1.0, 5.15, 0.8, C.accent1);
  s.addText([
    { text: "Einordnung", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } },
    { text: "Startkapital 100.000 € – unser Verkaufswert liegt 16.230 € darunter. Platzierung im Vergleich zu den anderen Gruppen: folgt.", options: { fontSize: 16, color: C.text1 } },
  ], { x: 2.1, y: 4.6, w: 10.2, h: 1.9, valign: "middle", isTextBox: true, margin: 0 });
}

// 10 Reflexion
{
  const s = content("Reflexion: Würden wir wieder so entscheiden?", "Reflexion",
    NOTE("ca. 1:30 min", "[Entwurf aus unserem Entscheidungsprotokoll, bitte gemeinsam prüfen und ergänzen.] Gut war, dass wir nach Runde 2 den Preis angepasst haben, das hat in Runde 3 die beste Runde gebracht. Schlechter lief Runde 2 mit einem zu hohen Preis und vor allem Runde 4 und 5: Wir haben fast nichts produziert, aber Personal und Miete liefen weiter, das hat die Kasse stark belastet. Die zweite Maschine für 30.000 Euro war außerdem nur in wenigen Runden richtig ausgelastet. Als Alternative hätten wir den Preis aus Runde 3 beibehalten und weiter normal produziert."));
  card(s, 0.7, 1.6, 3.8, 3.9, "Das lief gut", "Preis nach Runde 2 an den Markt angepasst – beste Runde mit 79.940 €\nRunde 6: alles produziert und verkauft", C.accent3, "FaThumbsUp");
  card(s, 4.75, 1.6, 3.8, 3.9, "Das lief schlechter", "Runde 2: Preis 7.600 € zu hoch, Hälfte zum Schnäppchenpreis\nRunde 4 und 5: nur 1 Palette, Fixkosten liefen weiter\n2. Maschine kaum ausgelastet", C.accent2, "FaThumbsDown");
  card(s, 8.8, 1.6, 3.8, 3.9, "Alternativen", "Preis um 5.900 € halten und weiter produzieren\nNur mit einer Maschine arbeiten (spart 30.000 € und Personal)", C.accent1, "FaExchangeAlt");
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.75, w: 11.9, h: 0.95, rectRadius: 0.12, fill: { color: C.background2 }, line: { color: C.accent5, width: 1.25, dashType: "dash" }, objectName: "Gesamtfazit" });
  s.addText([
    { text: "Gesamtfazit: ", options: { bold: true, color: C.text2 } },
    { text: "Der Preis war unser größter Hebel – zu hoch oder zu niedrig kostete uns jeweils Gewinn. (Entwurf)", options: { color: C.text1 } },
  ], { x: 0.95, y: 5.75, w: 11.4, h: 0.95, fontSize: 16, valign: "middle", isTextBox: true, margin: 0 });
}

// 11 Bezug zur Gründung
{
  const s = content("Bezug zur Gründung: was im Spiel fehlte", "Bezug zur Gründung",
    NOTE("ca. 1:30 min", "[Entwurf, bitte gemeinsam prüfen.] Viele unserer Gründungsideen kamen im Spiel nicht vor: Mehrweg und Pfand, Marketing auf Social Media, unser Onlineshop, die Zielgruppe und die besondere Sorte Tamarinde. Im Spiel entscheidet nur der Preis über den Absatz. Hätten Marketing und Vertrieb eine Rolle gespielt, hätten wir in Runde 4 und 5, als kaum etwas verkauft wurde, vielleicht mit Werbung gegensteuern können. Unser Vorschlag zur Erweiterung: ein Marketingbudget als zusätzliche Entscheidung pro Runde. Es kostet Geld, erhöht aber den Absatz. Das wäre nur eine Zeile mehr im Kassenbuch und das Spiel bliebe einfach."));
  card(s, 0.7, 1.6, 3.8, 3.9, "Nicht im Spiel", "Mehrweg-Glas und Pfand\nMarketing auf Social Media\nOnlineshop als Vertriebsweg\nZielgruppe und besondere Sorte", C.accent1, "FaBan");
  card(s, 4.75, 1.6, 3.8, 3.9, "Mögliche Auswirkung", "Werbung hätte in Runde 4 und 5 den Absatz stützen können\nMehrweg kostet, spart aber Material", C.accent1, "FaChartLine");
  card(s, 8.8, 1.6, 3.8, 3.9, "Idee zur Spielerweiterung", "Marketingbudget als zusätzliche Entscheidung pro Runde: kostet Geld, erhöht den Absatz – nur eine Zeile mehr im Kassenbuch", C.accent2, "FaLightbulb");
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.75, w: 11.9, h: 0.95, rectRadius: 0.12, fill: { color: C.background2 }, line: { color: C.accent5, width: 1.25, dashType: "dash" }, objectName: "Kernbotschaft" });
  s.addText([
    { text: "Kernbotschaft: ", options: { bold: true, color: C.text2 } },
    { text: "Im Spiel zählt nur der Preis – unsere Gründungsidee lebt von Marke, Vertrieb und Nachhaltigkeit. (Entwurf)", options: { color: C.text1 } },
  ], { x: 0.95, y: 5.75, w: 11.4, h: 0.95, fontSize: 16, valign: "middle", isTextBox: true, margin: 0 });
}

// 12 Danke
{
  const s = slide("DARK", "Abschluss");
  s.addShape(pres.ShapeType.ellipse, { x: 9.5, y: 1.9, w: 3.0, h: 3.0, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  s.addText("Vielen Dank", { x: 0.9, y: 2.4, w: 8, h: 1.3, fontSize: 60, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Wir freuen uns über eure Fragen.", { x: 0.9, y: 3.8, w: 8, h: 0.6, fontSize: 24, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Quellen: Planspiel Bertis Brause (Cornelsen), eigene Berechnungen aus Arbeitsblättern und Entscheidungsprotokoll", { x: 0.9, y: 6.3, w: 11.5, h: 0.4, fontSize: 12, color: C.accent3, isTextBox: true, margin: 0 });
  s.addNotes(NOTE("ca. 0:30 min", "Das war unsere Präsentation zur Elbharz Getränke GbR und zur Primatenbrause. Vielen Dank fürs Zuhören, wir beantworten gern eure Fragen."));
}
}

(async () => {
  await loadIcons(["FaBuilding","FaLightbulb","FaBullseye","FaRecycle","FaMapMarkerAlt","FaLeaf","FaThumbsUp","FaThumbsDown","FaExchangeAlt","FaBan","FaChartLine","FaBullhorn","FaStore"]);
  build();
  await pres.writeFile({ fileName: "Elbharz_Endpraesentation.pptx" });
  await applyTheme("Elbharz_Endpraesentation.pptx", THEME);
})();
