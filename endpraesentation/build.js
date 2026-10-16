// Endpräsentation Elbharz Getränke GbR – Version 1
// Erzeugen: node build.js  (Zahlen siehe FAKTEN.md)
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/4479ee3f-de1e-4616-9349-269c5bd53f97_e0f7022a-26aa-4f2e-8a7c-986f75ea937c/pptx/scripts/apply_theme.js");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");
const fs = require("fs");

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
const HEX = { petrol: "0F5C6B", amber: "D9822B", lime: "8DB33A", text: "1B2A2F", grid: "D9D9D2", muted: "5F8F99", sand: "F3F1EA" };
const OUT = "Elbharz_Endpraesentation.pptx";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13,33 × 7,5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Primatenbrause – Endpräsentation Planspiel Bertis Brause";
pres.author = "Elbharz Getränke GbR, Klasse 11b";
const C = pres.SchemeColor;

// ---------- Bilder ----------
const img = (f) => "image/png;base64," + fs.readFileSync("assets/" + f).toString("base64");
const IMG = {
  logo: img("logo.png"), flasche: img("flasche.png"), limette: img("limette.png"), tamarinde: img("tamarinde.png"),
  umriss: img("flasche_umriss_petrol.png"),
};
const FLASCHE_RATIO = 262 / 621;

const ICONS = {};
async function loadIcons(names) {
  for (const n of names) {
    for (const [suffix, color] of [["", "#FFFFFF"], ["_p", "#0F5C6B"]]) {
      const svg = RDS.renderToStaticMarkup(React.createElement(FA[n], { color, size: 256 }));
      ICONS[n + suffix] = "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
    }
  }
}

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.4, w: 11.9, h: 0.9, fontSize: 36, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.0, y: 6.95, w: 0.6, h: 0.3, fontSize: 11, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({ title: "DARK", background: { color: C.text2 }, objects: [] });

// ---------- Bausteine ----------
const CHAPTERS = ["Einleitung", "Verlauf", "Auswertung", "Reflexion", "Gründung"];
let sec = null;
function slide(master, section) {
  if (section !== sec) { pres.addSection({ title: section }); sec = section; }
  return pres.addSlide({ masterName: master, sectionTitle: section });
}
// Kapitel-Leiste: zeigt, welcher Pflichtteil gerade dran ist
function tracker(s, active) {
  CHAPTERS.forEach((c, i) => {
    const x = 0.7 + i * 1.55, on = c === active;
    s.addShape(pres.ShapeType.ellipse, { x, y: 7.04, w: 0.12, h: 0.12, fill: { color: on ? C.accent2 : C.accent5 }, line: { type: "none" }, objectName: "Kapitelpunkt " + c });
    s.addText(c, { x: x + 0.2, y: 6.95, w: 1.3, h: 0.3, fontSize: 11, bold: on, color: on ? C.accent2 : C.accent4, isTextBox: true, margin: 0, valign: "middle" });
  });
}
function content(title, chapter, notes) {
  const s = slide("CONTENT", chapter);
  s.addText(title, { placeholder: "title" });
  if (CHAPTERS.includes(chapter)) tracker(s, chapter);
  s.addNotes(notes);
  return s;
}
function iconCircle(s, name, x, y, d, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: "Icon-Kreis " + name });
  const p = d * 0.27;
  s.addImage({ data: ICONS[name], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name.replace("Fa", "Symbol "), objectName: "Icon " + name });
}
function box(s, x, y, w, h, fill, name, dashed) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: dashed ? { color: C.accent5, width: 1.25, dashType: "dash" } : { type: "none" }, objectName: name });
}
function card(s, x, y, w, h, head, body, accent, icon, bodySize) {
  box(s, x, y, w, h, C.background2, "Karte " + head);
  iconCircle(s, icon, x + 0.3, y + 0.3, 0.7, accent || C.accent1);
  s.addText(head, { x: x + 1.2, y: y + 0.3, w: w - 1.45, h: 0.7, fontSize: 20, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
  s.addText(body, { x: x + 0.3, y: y + 1.2, w: w - 0.6, h: h - 1.4, fontSize: bodySize || 16, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
}
const bullets = (arr, size) => arr.map((t, i) => ({ text: t, options: { bullet: { indent: 16 }, fontSize: size || 16, paraSpaceAfter: 8, breakLine: i < arr.length - 1 } }));
const chartBase = {
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: HEX.text, valAxisLabelColor: HEX.muted, legendColor: HEX.text, titleColor: HEX.text, dataLabelColor: HEX.text,
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, legendFontSize: 12, titleFontSize: 14, dataLabelFontSize: 11,
  showTitle: true, titleAlign: "left", valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" },
  catAxisLineShow: true, valAxisLineShow: false,
};
const NOTE = (t, txt) => `ZEIT: ${t}\n\nSPRECHTEXT (Vorschlag – bitte in eigenen Worten anpassen):\n${txt}`;
const fmt = (n) => n.toLocaleString("de-DE");

// ---------- Daten (siehe FAKTEN.md) ----------
const R = ["R1", "R2", "R3", "R4", "R5", "R6"];
const preis = [3000, 7600, 5900, 9500, 9500, 8400];
const produziert = [15, 13, 18, 1, 1, 13];
const markt = [15, 6.5, 12.6, 0.5, 0.2, 5.2];
const schnaeppchen = [0, 6.5, 5.6, 0.5, 0.8, 7.8];
const kostenZeilen = [
  ["Maschinen", [30000, 30000, 0, 0, 0, 0]],
  ["Personal", [4000, 4000, 8000, 8000, 8000, 8000]],
  ["Miete", [5000, 5000, 5000, 5000, 5000, 5000]],
  ["Wirtschaftsdienst", [0, 10000, 5000, 2500, 10000, 6500]],
  ["Produktion", [15000, 13000, 18000, 1000, 1000, 13000]],
  ["Material", [15000, 13000, 18000, 1500, 1000, 13000]],
];
const kosten = R.map((_, i) => kostenZeilen.reduce((a, [, v]) => a + v[i], 0));
const umsatz = [45000, 55900, 79940, 5250, 2700, 51480];
const wirt = umsatz.map((u, i) => Math.round((u / kosten[i]) * 100) / 100);
const schluss = [76000, 56900, 82840, 70090, 47790, 53770];
const wd = kostenZeilen[3][1];

function build() {
// =========================================================== 1 Titel
{
  const s = slide("DARK", "Einleitung");
  s.addImage({ data: IMG.logo, x: 0.8, y: 0.45, w: 1.5, h: 1.5, altText: "Logo Elbharz", objectName: "Logo" });
  s.addText("PLANSPIEL BERTIS BRAUSE · ENDPRÄSENTATION", { x: 0.9, y: 2.2, w: 6.6, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("Primatenbrause", { x: 0.9, y: 2.7, w: 7.2, h: 1.3, fontSize: 60, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Von der Idee zum Verkaufswert – unser Weg durch sechs Spielrunden", { x: 0.9, y: 4.1, w: 6.2, h: 1.0, fontSize: 22, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Elbharz Getränke GbR · Klasse 11b", { x: 0.9, y: 6.3, w: 6, h: 0.4, fontSize: 16, color: C.accent3, isTextBox: true, margin: 0 });
  // Flasche im orangen Kreis, Zutaten am grünen Kreis (Ideen der Gruppe)
  s.addShape(pres.ShapeType.ellipse, { x: 8.75, y: 1.05, w: 4.0, h: 4.0, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  const fh = 5.9; s.addImage({ data: IMG.flasche, x: 10.75 - (fh * FLASCHE_RATIO) / 2, y: 0.55, w: fh * FLASCHE_RATIO, h: fh, altText: "Flasche Primatenbrause", objectName: "Flasche" });
  s.addShape(pres.ShapeType.ellipse, { x: 7.35, y: 4.55, w: 2.1, h: 2.1, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "Zutatenkreis" });
  s.addImage({ data: IMG.tamarinde, x: 7.65, y: 4.95, w: 1.55, h: 1.55, altText: "Tamarinde", objectName: "Tamarinde" });
  s.addImage({ data: IMG.limette, x: 7.05, y: 4.45, w: 1.5, h: 0.75, altText: "Limette", objectName: "Limette" });
  s.addNotes(NOTE("ca. 0:35 min", "Wer von euch dreht beim Einkaufen eigentlich mal die Flasche um und liest, was drin ist? [kurz warten] Genau für solche Leute haben wir die Primatenbrause erfunden. Wir sind die Elbharz Getränke GbR aus der 11b, und in den nächsten elf Minuten zeigen wir euch, was aus unserer Idee im Planspiel geworden ist. Ehrlich gesagt: Nicht alles lief gut. Aber wir wissen jetzt ziemlich genau, warum. Los geht's mit unserer Idee."));
}

// =========================================================== 2 Wer wir sind
{
  const s = content("Wer wir sind – in 30 Sekunden", "Einleitung",
    NOTE("ca. 0:45 min", "Unsere Idee war ganz einfach: ein Getränk, dessen Rückseite man zu Ende lesen kann. Fast jedes Getränk im Regal verspricht heute irgendwas, mehr Energie, mehr Fokus. Wir wollten bewusst wenig versprechen, nur das, was man auf dem Etikett nachprüfen kann. Das war unsere einzige Vorgabe, und daran haben wir bei der Gründung alles andere gemessen, vom Namen bis zur Flasche. Gedacht ist die Primatenbrause für Leute zwischen 16 und 25, so wie unsere Persona Jasper: neunzehn, Azubi, und er dreht im Laden jede Flasche um. Und Tamarinde haben wir genommen, weil es die in Deutschland kaum in Getränken gibt. Herb statt süß."));
  box(s, 0.7, 1.6, 4.1, 4.9, C.accent1, "Leitgedanke");
  s.addText([
    { text: "Ein Getränk, dessen Rückseite man zu Ende lesen kann.", options: { fontSize: 30, bold: true, fontFace: "Cambria", color: C.background1, breakLine: true, paraSpaceAfter: 14 } },
    { text: "Unsere einzige Vorgabe – daran haben wir bei der Gründung alles andere gemessen.", options: { fontSize: 16, color: C.background2 } },
  ], { x: 1.05, y: 1.9, w: 3.4, h: 4.3, valign: "middle", isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: 5.1, y: 1.6, w: 3.4, h: 4.9, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent1, width: 2 }, shadow: { type: "outer", color: "1B2A2F", opacity: 0.18, blur: 8, offset: 3, angle: 90 }, objectName: "Etikett Rückseite" });
  s.addText("RÜCKSEITE", { x: 5.4, y: 1.85, w: 2.8, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3, isTextBox: true, margin: 0 });
  s.addText("Zutaten", { x: 5.4, y: 2.15, w: 2.8, h: 0.5, fontSize: 22, bold: true, fontFace: "Cambria", color: C.text2, isTextBox: true, margin: 0 });
  ["Wasser", "Kohlensäure", "Tamarinde", "Limettensaft", "Meersalz", "Magnesium"].forEach((z, i) => {
    const y = 2.8 + i * 0.42;
    s.addShape(pres.ShapeType.ellipse, { x: 5.42, y: y + 0.11, w: 0.14, h: 0.14, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "Punkt " + z });
    s.addText(z, { x: 5.7, y, w: 2.5, h: 0.36, fontSize: 16, color: C.text1, isTextBox: true, margin: 0, valign: "middle" });
  });
  s.addText("6 Zutaten. Mehr nicht.", { x: 5.4, y: 5.55, w: 2.8, h: 0.5, fontSize: 16, bold: true, italic: true, color: C.accent2, isTextBox: true, margin: 0 });
  const rows = [
    ["FaBuilding", "Unternehmen", "Elbharz Getränke GbR, Hamburg – 4 Gesellschafter", C.accent1],
    ["FaUserAlt", "Zielgruppe", "16–25 Jahre, Persona Jasper (19): liest Zutatenlisten", C.accent2],
    ["FaStore", "Vertrieb", "Online & Social Media, dann regionaler Handel", C.accent3],
  ];
  rows.forEach(([ic, h, t, col], i) => {
    const y = 1.6 + i * 1.7;
    box(s, 8.8, y, 3.8, 1.5, C.background2, "Karte " + h);
    iconCircle(s, ic, 9.05, y + 0.4, 0.7, col);
    s.addText([{ text: h, options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } }, { text: t, options: { fontSize: 14, color: C.text1 } }],
      { x: 9.95, y, w: 2.5, h: 1.5, valign: "middle", isTextBox: true, margin: 0 });
  });
}

// =========================================================== 3 Nachhaltigkeit
{
  const s = content("Nachhaltigkeit war Teil der Idee", "Einleitung",
    NOTE("ca. 0:40 min", "Nachhaltigkeit war uns von Anfang an wichtig. Wir wollten Mehrweg-Glas statt Einweg. Das lohnt sich aber nur bei kurzen Wegen, denn Glas ist schwer. Deshalb regional: in Hamburg herstellen und im Norden verkaufen. Und nur eine Sorte, weil das bei Produktion, Lager und Transport Aufwand spart. Für uns hieß das: lieber wenige Dinge richtig machen, als überall ein bisschen grün zu wirken. Im Planspiel selbst spielte das allerdings gar keine Rolle. Dort ging es nur um Preis, Menge und Kosten. Dazu sagen wir am Ende noch etwas."));
  const cw = 3.85, gx = 0.175;
  box(s, 0.7, 1.6, cw, 3.75, C.background2, "Karte Mehrweg");
  s.addImage({ data: IMG.umriss, x: 0.7 + cw / 2 - (2.0 * FLASCHE_RATIO) / 2, y: 1.85, w: 2.0 * FLASCHE_RATIO, h: 2.0, altText: "Leere Mehrwegflasche (Umriss)", objectName: "Leere Flasche" });
  s.addText([{ text: "Mehrweg-Glas", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: "Leere Flasche zurück, spülen, neu befüllen – Pfand statt Einweg", options: { fontSize: 15, color: C.text1 } }],
    { x: 0.95, y: 4.0, w: cw - 0.5, h: 1.2, valign: "top", isTextBox: true, margin: 0 });
  const x2 = 0.7 + cw + gx;
  box(s, x2, 1.6, cw, 3.75, C.background2, "Karte Regional");
  iconCircle(s, "FaMapMarkerAlt", x2 + cw / 2 - 0.75, 2.0, 1.5, C.accent3);
  s.addText("Hamburg · Norddeutschland", { x: x2 + 0.25, y: 3.55, w: cw - 0.5, h: 0.35, fontSize: 14, bold: true, color: C.accent4, align: "center", isTextBox: true, margin: 0 });
  s.addText([{ text: "Regional", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: "In Hamburg hergestellt, im Norden verkauft – kurze Wege, damit Mehrweg sich lohnt", options: { fontSize: 15, color: C.text1 } }],
    { x: x2 + 0.25, y: 4.0, w: cw - 0.5, h: 1.2, valign: "top", isTextBox: true, margin: 0 });
  const x3 = x2 + cw + gx;
  box(s, x3, 1.6, cw, 3.75, C.background2, "Karte Eine Sorte");
  [["FaIndustry", "Produktion"], ["FaWarehouse", "Lager"], ["FaTruck", "Transport"]].forEach(([ic, l], i) => {
    const cx = x3 + 0.35 + i * 1.15;
    iconCircle(s, ic, cx, 1.95, 0.8, C.accent3);
    s.addText(l, { x: cx - 0.2, y: 2.85, w: 1.2, h: 0.35, fontSize: 13, color: C.text1, align: "center", isTextBox: true, margin: 0 });
    if (i < 2) s.addText("›", { x: cx + 0.8, y: 2.05, w: 0.35, h: 0.6, fontSize: 28, bold: true, color: C.accent3, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText([{ text: "Eine Sorte", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: "Weniger Aufwand in jedem dieser drei Schritte", options: { fontSize: 15, color: C.text1 } }],
    { x: x3 + 0.25, y: 4.0, w: cw - 0.5, h: 1.2, valign: "top", isTextBox: true, margin: 0 });
  box(s, 0.7, 5.6, 11.9, 1.05, C.accent1, "Hinweis Spiel");
  s.addText([{ text: "Und im Planspiel? ", options: { bold: true, color: C.accent2 } }, { text: "Dort entschied nur der Preis über den Absatz – Nachhaltigkeit kam nicht vor. Mehr dazu am Ende.", options: { color: C.background1 } }],
    { x: 1.0, y: 5.6, w: 11.4, h: 1.05, fontSize: 17, valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 4 Produziert, angeboten, verkauft
{
  const s = content("Produziert, angeboten, verkauft", "Verlauf",
    NOTE("ca. 1:00 min", "Jetzt zum Spiel. Hier seht ihr alle sechs Runden auf einen Blick. Jede Säule ist das, was wir produziert haben, darunter steht unser Preis pro Palette. Der dunkle Teil wurde normal verkauft. Der orange Teil ging für nur tausend Euro an den Schnäppchenmarkt, weil Brause nach einer Runde verdirbt. Man sieht ein Muster: In Runde eins waren wir billig, und alles war weg. In Runde zwei waren wir zu teuer, die Hälfte blieb liegen. In Runde drei lagen wir mit knapp sechstausend Euro richtig, das war unsere beste Runde. Und in Runde vier und fünf haben wir fast nichts mehr produziert. Auffällig ist: Genau in den Runden mit dem höchsten Preis haben wir am wenigsten verkauft. Warum das so war, schauen wir uns jetzt genauer an."));
  const labels = R.map((r, i) => `${r} · ${fmt(preis[i])}\u00A0€`);
  s.addChart(pres.charts.BAR, [
    { name: "Am Markt verkauft", labels, values: markt },
    { name: "Schnäppchenmarkt (1.000\u00A0€)", labels, values: schnaeppchen },
  ], Object.assign({}, chartBase, {
    x: 0.7, y: 1.5, w: 7.9, h: 5.25, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
    chartColors: [HEX.petrol, HEX.amber], title: "Paletten je Runde (unter der Säule: unser Angebotspreis)",
    showLegend: true, legendPos: "t", valAxisMaxVal: 20, valAxisMajorUnit: 5, valAxisLabelFormatCode: "0",
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: '[<1]"";General', dataLabelFontSize: 11,
  }));
  const facts = [
    ["100 %", "verkauft in R1 – zu 3.000\u00A0€ ging alles weg", C.accent1],
    ["70 %", "verkauft in R3 bei 5.900\u00A0€ – unsere beste Runde", C.accent3],
    ["1 Palette", "in R4 und R5 – wir erwarteten kaum Absatz", C.accent2],
  ];
  facts.forEach(([n, t, col], i) => {
    const y = 1.6 + i * 1.75;
    box(s, 8.9, y, 3.7, 1.55, C.background2, "Fakt " + n);
    s.addText(n, { x: 9.15, y: y + 0.12, w: 3.3, h: 0.6, fontSize: 28, bold: true, fontFace: "Cambria", color: col, isTextBox: true, margin: 0 });
    s.addText(t, { x: 9.15, y: y + 0.75, w: 3.3, h: 0.7, fontSize: 15, color: C.text1, isTextBox: true, margin: 0 });
  });
}

// =========================================================== 5 Drei Entscheidungen
{
  const s = content("Drei Entscheidungen, die unser Spiel geprägt haben", "Verlauf",
    NOTE("ca. 1:10 min", "Drei Entscheidungen haben unser Spiel geprägt. Erstens Runde zwei: Weil in Runde eins alles weg war, dachten wir, die anderen sind teurer, und sind von 3.000 auf 7.600 Euro hoch. Gleichzeitig haben wir eine zweite Maschine für 30.000 Euro gekauft, aber keinen zweiten Mitarbeiter. Die Maschine stand also eine ganze Runde lang still. Zweitens Runde drei: Wir haben gemerkt, dass wir zu teuer waren, und sind knapp unter die 6.000 Euro gegangen, die wir bei der Konkurrenz vermutet haben. Dazu kam der zweite Mitarbeiter. Das war unsere einzige Runde mit richtig Gewinn, fast 26.000 Euro. Drittens Runde vier und fünf: In unserem Protokoll steht, wir dachten, die Beliebtheit hält nicht an, und eine Marktanalyse sagte weniger Absatz voraus. Also haben wir nur noch eine Palette produziert, aber für 9.500 Euro. Wir haben mit weniger Nachfrage gerechnet und trotzdem den höchsten Preis verlangt. Im Nachhinein passt das nicht zusammen."));
  const cards = [
    ["Runde 2", "7.600 €", "„Alles war weg – die anderen sind wohl teurer.“", "Preis drastisch rauf. 2. Maschine gekauft – aber ohne Mitarbeiter.", "−19.100 €", false],
    ["Runde 3", "5.900 €", "„Wir waren zu teuer. Die Konkurrenz liegt bei ca. 6.000\u00A0€.“", "Knapp darunter angeboten, 2. Mitarbeiter eingestellt.", "+25.940 €", true],
    ["Runde 4 + 5", "9.500 €", "„Die Beliebtheit hält nicht an.“ – Marktanalyse: weniger Absatz", "Nur 1 Palette – aber zum Höchstpreis.", "−35.050 €", false],
  ];
  cards.forEach(([r, p, gedanke, tat, erg, good], i) => {
    const x = 0.7 + i * 4.025, w = 3.85;
    box(s, x, 1.55, w, 5.15, C.background2, "Entscheidung " + r);
    s.addText(r, { x: x + 0.3, y: 1.75, w: w - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.accent4, isTextBox: true, margin: 0 });
    s.addText(p, { x: x + 0.3, y: 2.15, w: w - 0.6, h: 0.75, fontSize: 36, bold: true, fontFace: "Cambria", color: C.text2, isTextBox: true, margin: 0 });
    s.addText("Was wir dachten", { x: x + 0.3, y: 3.05, w: w - 0.6, h: 0.3, fontSize: 13, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
    s.addText(gedanke, { x: x + 0.3, y: 3.35, w: w - 0.6, h: 0.95, fontSize: 16, italic: true, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
    s.addText("Was wir gemacht haben", { x: x + 0.3, y: 4.35, w: w - 0.6, h: 0.3, fontSize: 13, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
    s.addText(tat, { x: x + 0.3, y: 4.65, w: w - 0.6, h: 0.95, fontSize: 16, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: 5.75, w: w - 0.6, h: 0.65, rectRadius: 0.08, fill: { color: good ? "E3EFCF" : C.background1 }, line: { type: "none" }, objectName: "Ergebnis " + r });
    s.addText([{ text: "Ergebnis  ", options: { fontSize: 14, color: C.text1 } }, { text: erg, options: { fontSize: 22, bold: true, fontFace: "Cambria", color: good ? "3D6610" : "8A4A12" } }],
      { x: x + 0.45, y: 5.75, w: w - 0.9, h: 0.65, valign: "middle", isTextBox: true, margin: 0 });
  });
}

// =========================================================== 7 Kennzahlen
{
  const s = content("So haben wir ausgewertet: zwei Kennzahlen", "Auswertung",
    NOTE("ca. 1:00 min", "Um das auszuwerten, haben wir zwei Kennzahlen berechnet. Die erste ist die Wirtschaftlichkeit: Umsatz geteilt durch Kosten. In Runde drei hatten wir knapp 80.000 Euro Umsatz und 54.000 Euro Kosten. Das ergibt 1,48. Für jeden Euro, den wir ausgegeben haben, kamen also fast anderthalb Euro zurück. Alles über eins ist Gewinn. Die zweite Kennzahl ist die Absatzquote: Wie viel unserer Produktion haben wir zum normalen Preis verkauft? In Runde drei waren das 70 Prozent. Die Zahl zeigt, ob unser Preis zum Markt gepasst hat. Wichtig ist noch: Die Wirtschaftlichkeit rechnet Maschinenkäufe komplett in die Runde, in der wir sie gekauft haben. Deshalb sehen Runde eins und zwei schlechter aus, als der laufende Betrieb eigentlich war. Alle Werte für jede Runde stehen in unserem Auswertungsbogen und im Anhang."));
  const ks = [
    ["FaBalanceScale", "Wirtschaftlichkeit", "Umsatzerlöse ÷ Kosten", "79.940 € ÷ 54.000 € = 1,48", "Über 1 = Gewinn. Zeigt, ob der Umsatz die Kosten einer Runde deckt."],
    ["FaChartLine", "Absatzquote", "am Markt verkauft ÷ produziert", "12,6 ÷ 18 Paletten = 70 %", "Zeigt, ob unser Preis zum Markt gepasst hat."],
  ];
  ks.forEach(([ic, h, f, bsp, auss], i) => {
    const x = 0.7 + i * 6.05, w = 5.85;
    box(s, x, 1.55, w, 4.6, C.background2, "Kennzahl " + h);
    iconCircle(s, ic, x + 0.35, 1.85, 0.8, i ? C.accent2 : C.accent1);
    s.addText(h, { x: x + 1.35, y: 1.85, w: w - 1.6, h: 0.8, fontSize: 26, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(f, { x: x + 0.35, y: 2.95, w: w - 0.7, h: 0.6, fontSize: 22, bold: true, color: C.text1, isTextBox: true, margin: 0 });
    s.addText("Beispiel Runde 3", { x: x + 0.35, y: 3.7, w: w - 0.7, h: 0.35, fontSize: 14, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
    s.addText(bsp, { x: x + 0.35, y: 4.05, w: w - 0.7, h: 0.6, fontSize: 24, fontFace: "Cambria", color: C.text2, isTextBox: true, margin: 0 });
    s.addText(auss, { x: x + 0.35, y: 4.8, w: w - 0.7, h: 1.1, fontSize: 16, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });
  s.addText("Alle Werte je Runde: Auswertungsbogen M15 und Anhang", { x: 0.7, y: 6.3, w: 11.9, h: 0.4, fontSize: 14, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
}

// =========================================================== 8 Entwicklung
{
  const s = content("Wirtschaftlichkeit und Schlussbestand", "Auswertung",
    NOTE("ca. 0:50 min", "Links seht ihr die Wirtschaftlichkeit für jede Runde. Die gestrichelte Linie ist die Eins, also die Grenze zwischen Gewinn und Verlust. Darüber liegen nur Runde drei und Runde sechs. Über alle Runden zusammen kommen wir auf 0,84: rund 240.000 Euro Umsatz bei rund 287.000 Euro Kosten. Rechts seht ihr, was mit unserem Startkapital passiert ist. Wir sind mit 100.000 Euro gestartet. Die Maschinenkäufe haben uns am Anfang runtergezogen, nach Runde drei ging es wieder hoch, und nach Runde fünf waren wir mit unter 48.000 Euro am Tiefpunkt. Man sieht gut: Nach unserer besten Runde ging es durch zwei schwache Runden direkt wieder nach unten. Am Ende stehen knapp 54.000 Euro in der Kasse."));
  s.addChart([
    { type: pres.charts.BAR, data: [{ name: "Wirtschaftlichkeit", labels: R, values: wirt }], options: { chartColors: [HEX.petrol], barGapWidthPct: 60, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00" } },
    { type: pres.charts.LINE, data: [{ name: "Grenze 1,0 (Kosten gedeckt)", labels: R, values: [1, 1, 1, 1, 1, 1] }], options: { chartColors: [HEX.amber], lineSize: 2, lineDash: "dash", lineDataSymbol: "none", showValue: false } },
  ], Object.assign({}, chartBase, { x: 0.7, y: 1.5, w: 6.0, h: 4.55, title: "Wirtschaftlichkeit je Runde", showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 1.6, valAxisMajorUnit: 0.2, valAxisLabelFormatCode: "0.0", catAxisLabelFontSize: 13, valAxisLabelFontSize: 12, dataLabelFontSize: 13 }));
  s.addChart(pres.charts.LINE, [{ name: "Schlussbestand", labels: ["Start", ...R], values: [100000, ...schluss] }], Object.assign({}, chartBase, {
    x: 6.9, y: 1.5, w: 5.7, h: 3.55, chartColors: [HEX.petrol], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8,
    title: "Schlussbestand je Runde (€)", showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 120000, valAxisMajorUnit: 40000, valAxisLabelFormatCode: "#,##0",
    showValue: false, catAxisLabelFontSize: 13, valAxisLabelFontSize: 12,
  }));
  [["Start", "100.000 €", C.text2], ["Höchststand R3", "82.840 €", "3D6610"], ["Tiefpunkt R5", "47.790 €", "8A4A12"], ["Ende R6", "53.770 €", C.accent2]].forEach(([l, v, col], i) => {
    const x = 6.9 + i * 1.45;
    box(s, x, 5.15, 1.35, 0.9, C.background2, "Eckwert " + l);
    s.addText([{ text: v, options: { fontSize: 15, bold: true, color: col, breakLine: true } }, { text: l, options: { fontSize: 11, color: C.text1 } }],
      { x: x + 0.05, y: 5.15, w: 1.25, h: 0.9, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  box(s, 0.7, 6.2, 11.9, 0.6, C.background2, "Kernaussage");
  s.addText([{ text: "Kernaussage: ", options: { bold: true, color: C.accent2 } }, { text: "Nur 2 von 6 Runden über 1 – gesamt 0,84 (240.270 € Umsatz ÷ 286.500 € Kosten).", options: { color: C.text1 } }],
    { x: 0.95, y: 6.2, w: 11.4, h: 0.6, fontSize: 16, valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 9 Verkaufswert
{
  const s = content("Was unser Unternehmen am Ende wert ist", "Auswertung",
    NOTE("ca. 0:50 min", "Am Ende des Spiels wird das Unternehmen verkauft. Der Wert ist das Geld in der Kasse plus das, was die Maschinen noch wert sind. Jede Maschine verliert pro Runde, in der sie läuft, zehn Prozent. Die erste lief sechs Runden und ist noch 12.000 Euro wert. Die zweite lief nur vier Runden, weil sie in Runde zwei stillstand. Sie ist noch 18.000 Euro wert. Zusammen kommen wir auf 83.770 Euro. Das sind gut 16.000 Euro weniger, als wir am Anfang hatten. Allein die zwei Maschinen haben 60.000 Euro gekostet, davon steckt nur noch die Hälfte im Unternehmenswert. Und hier kommt die Frage: warum nur so wenig?"));
  const tiles = [
    ["FaCoins", "53.770\u00A0€", "Schlussbestand nach R6", C.accent1],
    ["FaCogs", "12.000\u00A0€", "Maschine 1 (6 Runden genutzt)", C.accent1],
    ["FaCogs", "18.000\u00A0€", "Maschine 2 (4 Runden genutzt)", C.accent1],
    ["FaBuilding", "83.770\u00A0€", "Unternehmens­verkaufswert", C.accent2],
  ];
  const tw = 2.55, gap = 0.55;
  tiles.forEach(([ic, n, l, col], i) => {
    const x = 0.7 + i * (tw + gap), last = i === 3;
    box(s, x, 1.6, tw, 2.75, last ? C.accent1 : C.background2, "Baustein " + l);
    iconCircle(s, ic, x + tw / 2 - 0.4, 1.85, 0.8, last ? C.accent2 : col);
    s.addText(n, { x, y: 2.75, w: tw, h: 0.7, fontSize: 30, bold: true, fontFace: "Cambria", color: last ? C.background1 : C.text2, align: "center", isTextBox: true, margin: 0 });
    s.addText(l, { x: x + 0.15, y: 3.45, w: tw - 0.3, h: 0.7, fontSize: 14, color: last ? C.background2 : C.text1, align: "center", isTextBox: true, margin: 0 });
    if (i < 3) s.addText(i < 2 ? "+" : "=", { x: x + tw, y: 2.5, w: gap, h: 0.8, fontSize: 36, bold: true, color: C.accent2, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  });
  // Vergleich mit dem Startkapital
  s.addText("Vergleich mit dem Startkapital", { x: 0.7, y: 4.65, w: 6, h: 0.4, fontSize: 18, bold: true, color: C.text2, isTextBox: true, margin: 0 });
  const full = 9.0;
  [["Startkapital", 100000, C.accent5, "100.000\u00A0€"], ["Verkaufswert", 83770, C.accent2, "83.770\u00A0€"]].forEach(([l, v, col, t], i) => {
    const y = 5.15 + i * 0.6;
    s.addText(l, { x: 0.7, y, w: 1.7, h: 0.45, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x: 2.4, y, w: (full * v) / 100000, h: 0.45, rectRadius: 0.06, fill: { color: col }, line: { type: "none" }, objectName: "Balken " + l });
    s.addText(t, { x: 2.5, y, w: 2, h: 0.45, fontSize: 14, bold: true, color: i ? C.background1 : C.text1, valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText([{ text: "−16.230\u00A0€", options: { fontSize: 22, bold: true, color: C.accent2, fontFace: "Cambria", breakLine: true } }, { text: "(−16,2 %)", options: { fontSize: 14, color: C.text1 } }],
    { x: 11.5, y: 5.1, w: 1.3, h: 1.15, align: "center", valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 6 Wirtschaftsdienst
{
  const s = content("Warum nur 83.770\u00A0€? Der Wirtschaftsdienst", "Auswertung",
    NOTE("ca. 0:50 min", "Ein Teil davon lag nicht an uns. Ab Runde zwei kamen die Meldungen des Wirtschaftsdienstes, und jede hat uns Geld gekostet. Insgesamt waren das 34.000 Euro, am meisten in Runde zwei und fünf mit je 10.000. Rein rechnerisch hätten wir ohne diese Kosten am Ende rund 118.000 Euro gehabt, also mehr als unser Startkapital. Das soll keine Ausrede sein. Die Meldungen konnten wir nicht verhindern, aber wir hätten Rücklagen einplanen können, statt in Runde zwei direkt eine zweite Maschine zu kaufen. Für uns war das eine wichtige Erkenntnis: Ein Unternehmen braucht einen Puffer für Dinge, die man nicht planen kann. [Optional: pro Runde ein Stichwort zur Meldung nennen.]"));
  box(s, 0.7, 1.6, 3.6, 2.35, C.accent1, "Kennzahl Zusatzkosten");
  iconCircle(s, "FaNewspaper", 0.95, 1.85, 0.65, C.accent2);
  s.addText("34.000\u00A0€", { x: 0.95, y: 2.55, w: 3.2, h: 0.8, fontSize: 40, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Zusatzkosten in den Runden 2–6", { x: 0.95, y: 3.3, w: 3.2, h: 0.45, fontSize: 15, color: C.background2, isTextBox: true, margin: 0 });
  box(s, 0.7, 4.15, 3.6, 2.5, C.background2, "Einordnung Zusatzkosten");
  s.addText(bullets(["≈ 12 % unserer Gesamtkosten", "Am höchsten in R2 und R5 (je 10.000\u00A0€)", "Ohne sie: Verkaufswert 117.770\u00A0€ – über dem Startkapital"], 15),
    { x: 0.95, y: 4.3, w: 3.2, h: 2.2, color: C.text1, valign: "middle", isTextBox: true, margin: 0 });
  s.addChart(pres.charts.BAR, [{ name: "Zusatzkosten (€)", labels: R, values: wd }], Object.assign({}, chartBase, {
    x: 4.6, y: 1.5, w: 8.0, h: 5.2, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.amber],
    title: "Zusatzkosten durch den Wirtschaftsdienst je Runde (€)", showLegend: false,
    valAxisLabelFormatCode: "#,##0", valAxisMaxVal: 12000, valAxisMajorUnit: 2000,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: 12,
  }));
}

// =========================================================== 10 Reflexion
{
  const s = content("Reflexion: Würden wir wieder so entscheiden?", "Reflexion",
    NOTE("ca. 1:10 min", "Würden wir wieder so entscheiden? Teilweise. Gut war, dass wir vorsichtig mit nur einer Maschine gestartet sind, das empfiehlt sogar die Spielanleitung, und dass wir nach Runde zwei schnell reagiert haben. Schlecht war vor allem unser Preis ab Runde vier. Die Regel kannten wir ja: Der Preis bestimmt den Absatz. Trotzdem haben wir genau dann die höchsten Preise verlangt, als wir mit wenig Nachfrage gerechnet haben. Dazu kamen die Fixkosten: Für eine einzige Palette haben wir jeweils 13.000 Euro für Personal und Miete bezahlt. Und die zweite Maschine haben wir eigentlich nur in Runde drei gebraucht. Beim nächsten Mal würden wir beim Preis aus Runde drei bleiben, Maschinen nur kaufen, wenn wir sie auch auslasten und besetzen können, und das Personal an die Produktion anpassen. Unser wichtigstes Learning: Der Preis war unser stärkster Hebel, und Fixkosten laufen weiter, auch wenn man kaum etwas verkauft."));
  const cols = [
    ["FaThumbsUp", "Das lief gut", C.accent3, ["Vorsichtiger Start mit nur 1 Maschine", "Nach R2 schnell reagiert → beste Runde"]],
    ["FaThumbsDown", "Das lief schlecht", C.accent2, ["Höchstpreis trotz erwarteter Flaute (R4–R6)", "13.000\u00A0€ Fixkosten für 1 Palette", "2. Maschine kaum gebraucht"]],
    ["FaExchangeAlt", "Das machen wir anders", C.accent1, ["Preis nahe 5.900\u00A0€ halten", "Maschinen nur kaufen, wenn ausgelastet und besetzt", "Personal an die Produktion anpassen"]],
  ];
  cols.forEach(([ic, h, col, items], i) => {
    const x = 0.7 + i * 4.025;
    box(s, x, 1.6, 3.85, 3.85, C.background2, "Karte " + h);
    iconCircle(s, ic, x + 0.3, 1.85, 0.7, col);
    s.addText(h, { x: x + 1.2, y: 1.85, w: 2.5, h: 0.7, fontSize: 19, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(bullets(items, 18), { x: x + 0.3, y: 2.75, w: 3.3, h: 2.6, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });
  box(s, 0.7, 5.65, 11.9, 1.0, C.accent1, "Fazit");
  s.addText([{ text: "Unser Fazit: ", options: { bold: true, color: C.accent2 } }, { text: "Der Preis war unser stärkster Hebel – und Fixkosten laufen weiter, auch wenn man kaum verkauft.", options: { color: C.background1 } }],
    { x: 1.0, y: 5.65, w: 11.4, h: 1.0, fontSize: 17, valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 11 Bezug zur Gründung
{
  const s = content("Bezug zur Gründung: Was im Spiel fehlte", "Gründung",
    NOTE("ca. 1:05 min", "Zum Schluss der Blick zurück auf unsere Gründung. Vieles, was uns wichtig war, kam im Spiel gar nicht vor. Mehrweg hätte uns zusätzliche Kosten gebracht, für Rücknahme und Spülen. Unsere Zielgruppe hätte bedeutet, dass nicht nur der Preis zählt, sondern auch, wen man anspricht. Mit Social Media hätten wir in Runde vier und fünf Werbung machen können, statt nur eine Palette anzubieten. Und Tamarinde-Limette ist etwas Besonderes, damit hätten wir vielleicht einen höheren Preis verlangen können. Im Spiel war aber jede Brause gleich. Deshalb schlagen wir zwei einfache Erweiterungen vor: ein Marketingbudget, bei dem es für je 2.000 Euro eine Palette mehr Absatz gibt, und eine Mehrweg-Option, die 1.000 Euro pro Runde kostet und dafür eine Palette mehr Absatz bringt. Beides ist nur eine Zeile mehr im Kassenbuch, und das Spiel käme näher an das, was wir uns bei der Gründung überlegt haben."));
  s.addText("Was im Spiel fehlte – und was es bewirkt hätte", { x: 0.7, y: 1.5, w: 7.6, h: 0.4, fontSize: 16, bold: true, color: C.accent4, isTextBox: true, margin: 0 });
  [["FaRecycle", "Mehrweg & Pfand", "Mehr Kosten für Rücknahme und Spülen"],
   ["FaUserAlt", "Zielgruppe & Persona Jasper", "Nicht nur der Preis zählt, auch wen man anspricht"],
   ["FaBullhorn", "Social Media & Onlineshop", "Werbung statt nur 1 Palette in R4/R5"],
   ["FaLeaf", "Sorte Tamarinde-Limette", "Besonderheit erlaubt höheren Preis"]].forEach(([ic, t, e], i) => {
    const y = 2.0 + i * 1.17;
    box(s, 0.7, y, 7.6, 1.02, C.background2, "Fehlt " + t);
    iconCircle(s, ic, 0.9, y + 0.17, 0.68, C.accent1);
    s.addText([{ text: t, options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: "→ " + e, options: { fontSize: 16, color: C.text1 } }],
      { x: 1.8, y, w: 6.35, h: 1.02, valign: "middle", isTextBox: true, margin: 0 });
  });
  box(s, 8.6, 1.5, 4.0, 5.2, C.accent1, "Spielerweiterung");
  iconCircle(s, "FaLightbulb", 8.85, 1.75, 0.7, C.accent2);
  s.addText("Unsere Vorschläge", { x: 9.7, y: 1.75, w: 2.8, h: 0.7, fontSize: 19, bold: true, color: C.background1, valign: "middle", isTextBox: true, margin: 0 });
  s.addText([
    { text: "1  Marketingbudget", options: { bold: true, fontSize: 20, color: C.accent2, breakLine: true } },
    { text: "Je 2.000\u00A0€ → +1 Palette Absatz", options: { fontSize: 17, color: C.background1, breakLine: true, paraSpaceAfter: 22 } },
    { text: "2  Mehrweg-Option", options: { bold: true, fontSize: 20, color: C.accent2, breakLine: true } },
    { text: "+1.000\u00A0€ Kosten je Runde → +1 Palette Absatz bei umweltbewussten Kunden", options: { fontSize: 17, color: C.background1, breakLine: true, paraSpaceAfter: 22 } },
    { text: "Je nur eine Zeile mehr im Kassenbuch.", options: { fontSize: 16, italic: true, color: C.background2 } },
  ], { x: 8.85, y: 2.7, w: 3.5, h: 3.8, valign: "top", isTextBox: true, margin: 0 });
}

// =========================================================== 12 Schluss
{
  const s = slide("DARK", "Abschluss");
  s.addText("Unsere Rückseite kann man zu Ende lesen.", { x: 0.9, y: 1.5, w: 7.0, h: 1.0, fontSize: 20, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
  s.addText("Unsere Bilanz auch.", { x: 0.9, y: 2.3, w: 7.0, h: 1.1, fontSize: 44, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Sie war nicht gut – aber wir wissen jetzt, warum.", { x: 0.9, y: 3.45, w: 6.6, h: 0.9, fontSize: 22, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Vielen Dank! Wir freuen uns auf eure Fragen.", { x: 0.9, y: 4.75, w: 6.6, h: 0.5, fontSize: 20, bold: true, color: C.accent3, isTextBox: true, margin: 0 });
  s.addImage({ data: IMG.logo, x: 0.8, y: 5.75, w: 1.1, h: 1.1, altText: "Logo Elbharz", objectName: "Logo" });
  s.addShape(pres.ShapeType.ellipse, { x: 8.9, y: 1.2, w: 3.8, h: 3.8, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  const fh = 5.4; s.addImage({ data: IMG.flasche, x: 10.8 - (fh * FLASCHE_RATIO) / 2, y: 0.75, w: fh * FLASCHE_RATIO, h: fh, altText: "Flasche Primatenbrause", objectName: "Flasche" });
  s.addShape(pres.ShapeType.ellipse, { x: 7.3, y: 4.5, w: 2.2, h: 2.2, fill: { color: C.accent3, transparency: 15 }, line: { color: C.background1, width: 1.5, dashType: "dash" }, objectName: "Platz KI-Teambild" });
  s.addText("BILD FOLGT: KI-Teambild", { x: 7.4, y: 4.5, w: 2.0, h: 2.2, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  s.addNotes(NOTE("ca. 0:30 min", "Am Anfang haben wir gesagt: Die Rückseite unserer Flasche kann man zu Ende lesen. Unsere Bilanz auch. Sie war nicht gut, aber wir wissen jetzt genau, warum. Und unser Slogan, „Was dein Körper schon immer wollte“, verspricht bewusst nichts, was nicht auf dem Etikett steht. Genau so ehrlich wollten wir auch diese Präsentation halten. Vielen Dank fürs Zuhören, wir freuen uns auf eure Fragen."));
}

// =========================================================== 13 Quellen
{
  const s = content("Quellen", "Quellen", "Quellenfolie – wird nicht vorgetragen, steht für die Lehrkraft.");
  const blocks = [
    ["Spielmaterial", ["Planspiel „Bertis Brause“, Cornelsen Verlag 2006 (Spielanleitung, Arbeitsblätter, Bögen M14/M15)", "Aufgabenstellung Endpräsentation, Wirtschaftsgymnasium Am Lämmermarkt"]],
    ["Daten", ["Eigene Arbeitsblätter: Maschinen/Personal, Produktion/Material, Kassenbuch, Angebotskarte, Umsatzberechnung, Unternehmensverkauf", "Eigenes Entscheidungsprotokoll; alle Kennzahlen selbst berechnet"]],
    ["Bilder", ["Logo und Flasche: KI-generiert mit Perplexity (aus unserer Gründungspräsentation)", "Limette, Tamarinde: [Quelle ergänzen]", "KI-Teambild: [Tool ergänzen]", "Symbole: Font Awesome (CC BY 4.0)"]],
  ];
  blocks.forEach(([h, items], i) => {
    const x = 0.7 + i * 4.025;
    box(s, x, 1.6, 3.85, 3.7, C.background2, "Quellen " + h);
    s.addText(h, { x: x + 0.3, y: 1.8, w: 3.3, h: 0.5, fontSize: 19, bold: true, color: C.text2, isTextBox: true, margin: 0 });
    s.addText(bullets(items, 14), { x: x + 0.3, y: 2.4, w: 3.3, h: 2.75, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });
}
// =========================================================== 14 Anhang
{
  const s = content("Anhang: Kosten und Kennzahlen je Runde", "Anhang", "Backup-Folie für Rückfragen – wird nicht vorgetragen.");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.petrol }, align: "center" } });
  const rows = [[hdr("in\u00A0€"), ...R.map(hdr)]];
  kostenZeilen.forEach(([n, v]) => rows.push([{ text: "– " + n }, ...v.map((x) => ({ text: fmt(x), options: { align: "right" } }))]));
  rows.push([{ text: "= Kosten gesamt", options: { bold: true } }, ...kosten.map((x) => ({ text: fmt(x), options: { bold: true, align: "right" } }))]);
  rows.push([{ text: "Umsatzerlöse", options: { bold: true } }, ...umsatz.map((x) => ({ text: fmt(x), options: { bold: true, align: "right" } }))]);
  rows.push([{ text: "Wirtschaftlichkeit", options: { bold: true } }, ...wirt.map((w) => ({ text: w.toFixed(2).replace(".", ","), options: { bold: true, align: "center", color: w >= 1 ? "3D6610" : "8A4A12", fill: { color: w >= 1 ? "E3EFCF" : "F7E4CF" } } }))]);
  rows.push([{ text: "Absatzquote", options: { bold: true } }, ...R.map((_, i) => ({ text: Math.round((markt[i] / produziert[i]) * 100) + " %", options: { align: "center" } }))]);
  rows.push([{ text: "Schlussbestand", options: { bold: true } }, ...schluss.map((x) => ({ text: fmt(x), options: { align: "right" } }))]);
  s.addTable(rows, { x: 0.7, y: 1.5, w: 11.9, colW: [3.2, 1.45, 1.45, 1.45, 1.45, 1.45, 1.45], fontSize: 15, color: HEX.text, border: { type: "solid", color: "D3D1C8", pt: 0.75 }, rowH: 0.42, valign: "middle", objectName: "Tabelle Anhang" });
  s.addText("Wirtschaftlichkeit gesamt: 240.270 € ÷ 286.500 € = 0,84 · Wirtschaftsdienst = sonstige Kosten laut Kassenbuch", { x: 0.7, y: 6.5, w: 11.9, h: 0.35, fontSize: 12, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
}

}

(async () => {
  await loadIcons(["FaBuilding", "FaUserAlt", "FaStore", "FaIndustry", "FaWarehouse", "FaTruck", "FaNewspaper", "FaBalanceScale", "FaCoins", "FaCogs",
    "FaThumbsUp", "FaThumbsDown", "FaExchangeAlt", "FaRecycle", "FaBullhorn", "FaLeaf", "FaChartLine", "FaLightbulb", "FaMapMarkerAlt"]);
  build();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Kosten", kosten, "Wirtschaftlichkeit", wirt);
})();
