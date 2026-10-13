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
  s.addNotes(NOTE("ca. 0:20 min", "Hallo zusammen, wir sind die Elbharz Getränke GbR aus der 11b, und das ist unsere Primatenbrause, Tamarinde-Limette. In den nächsten elf Minuten zeigen wir euch, was aus unserer Idee im Planspiel geworden ist: von der Gründung über sechs Spielrunden bis zum Verkaufswert unseres Unternehmens."));
}

// =========================================================== 2 Wer wir sind
{
  const s = content("Wer wir sind – in 30 Sekunden", "Einleitung",
    NOTE("ca. 0:45 min", "Kurz zur Erinnerung, wer wir sind. Unser Ausgangspunkt war ein Getränk, dessen Rückseite man zu Ende lesen kann. Das war unsere einzige Vorgabe, und daran haben wir bei der Gründung alles andere gemessen. Hier seht ihr die Rückseite: sechs Zutaten, mehr nicht – Wasser, Kohlensäure, Tamarinde, Limettensaft, Meersalz und Magnesium. Wir sind vier Gesellschafter aus Hamburg. Unsere Zielgruppe sind 16- bis 25-Jährige wie unsere Persona Jasper, der im Laden die Flasche umdreht und die Zutaten liest. Verkaufen wollten wir zuerst online und über Social Media, danach im regionalen Handel. Und warum Tamarinde? Die gibt es in Deutschland in Getränken kaum – herb statt süß."));
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
    NOTE("ca. 0:45 min", "Nachhaltigkeit hatten wir von Anfang an im Blick. Erstens Mehrweg: Glasflaschen im Pfandsystem statt Einweg – die leere Flasche kommt zurück, wird gespült und neu befüllt. Zweitens regional: Glas ist schwer, Mehrweg lohnt sich deshalb nur bei kurzen Wegen. Darum wollten wir in Norddeutschland herstellen und abfüllen. Drittens nur eine Sorte: Das spart Aufwand bei Produktion, Lager und Transport. Wichtig war uns dabei, nur zu versprechen, was man auf dem Etikett überprüfen kann. Ehrlich gesagt spielte das im Planspiel keine Rolle, dort entschied nur der Preis über den Absatz. Darauf kommen wir am Ende beim Bezug zur Gründung zurück."));
  const cw = 3.85, gx = 0.175;
  box(s, 0.7, 1.6, cw, 3.75, C.background2, "Karte Mehrweg");
  s.addImage({ data: IMG.umriss, x: 0.7 + cw / 2 - (2.0 * FLASCHE_RATIO) / 2, y: 1.85, w: 2.0 * FLASCHE_RATIO, h: 2.0, altText: "Leere Mehrwegflasche (Umriss)", objectName: "Leere Flasche" });
  s.addText([{ text: "Mehrweg-Glas", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: "Leere Flasche zurück, spülen, neu befüllen – Pfand statt Einweg", options: { fontSize: 15, color: C.text1 } }],
    { x: 0.95, y: 4.0, w: cw - 0.5, h: 1.2, valign: "top", isTextBox: true, margin: 0 });
  const x2 = 0.7 + cw + gx;
  box(s, x2, 1.6, cw, 3.75, C.background2, "Karte Regional");
  box(s, x2 + 0.25, 1.85, cw - 0.5, 2.0, C.background1, "Bildplatz Karte", true);
  s.addText("BILD FOLGT: Karte Norddeutschland, Elbe rot markiert", { x: x2 + 0.4, y: 1.85, w: cw - 0.8, h: 2.0, fontSize: 12, color: C.accent4, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  s.addText([{ text: "Regional", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: "Hier hergestellt, hier abgefüllt – kurze Wege, damit Mehrweg sich lohnt", options: { fontSize: 15, color: C.text1 } }],
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
    NOTE("ca. 1:00 min", "Hier seht ihr alle sechs Runden auf einen Blick. Jede Säule zeigt, wie viele Paletten wir produziert haben, darunter steht unser Angebotspreis. Petrol ist, was wir am normalen Markt verkauft haben. Orange mussten wir für nur 1.000 Euro pro Palette am Schnäppchenmarkt loswerden, denn Brause verdirbt nach einer Runde. In Runde 1 haben wir zu 3.000 Euro alles verkauft. In Runde 2 haben wir den Preis auf 7.600 Euro erhöht, da blieb die Hälfte liegen. Mit 5.900 Euro in Runde 3 lief es am besten: 70 Prozent gingen am Markt weg. In Runde 4 und 5 haben wir nur noch eine Palette produziert und davon trotzdem nur die Hälfte bzw. ein Fünftel verkauft. In Runde 6 waren es noch einmal 13 Paletten."));
  const labels = R.map((r, i) => `${r} · ${fmt(preis[i])}\u00A0€`);
  s.addChart(pres.charts.BAR, [
    { name: "Am Markt verkauft", labels, values: markt },
    { name: "Schnäppchenmarkt (1.000\u00A0€)", labels, values: schnaeppchen },
  ], Object.assign({}, chartBase, {
    x: 0.7, y: 1.5, w: 7.9, h: 4.85, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55,
    chartColors: [HEX.petrol, HEX.amber], title: "Paletten je Runde (unter der Säule: unser Angebotspreis)",
    showLegend: true, legendPos: "t", valAxisMaxVal: 20, valAxisMajorUnit: 5, valAxisLabelFormatCode: "0",
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFormatCode: '[<1]"";General', dataLabelFontSize: 11,
  }));
  s.addText("R4: 0,5 am Markt + 0,5 Schnäppchen · R5: 0,2 + 0,8 · R3 laut Arbeitsblatt 12,6 + 5,6 Paletten", { x: 0.7, y: 6.4, w: 7.9, h: 0.35, fontSize: 12, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
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

// =========================================================== 5 Warum wir so entschieden haben
{
  const s = content("Warum wir so entschieden haben", "Verlauf",
    NOTE("ca. 1:05 min", "Warum haben wir so entschieden? Das steht in unserem Entscheidungsprotokoll, und unten auf jeder Karte seht ihr, was die Runde für unseren Kassenstand bedeutet hat. In Runde 1 haben wir eine Economy-Maschine gekauft, weil sich Standard mit dem zusätzlichen Mitarbeiter nicht gelohnt hätte, und bewusst günstig angeboten. Weil alles weg war, dachten wir in Runde 2, die anderen sind teurer, und sind drastisch auf 7.600 Euro hoch. Außerdem haben wir eine zweite Maschine gekauft – aber noch ohne Mitarbeiter, sie stand also still. In Runde 3 kam der zweite Mitarbeiter dazu, und wir haben knapp unter den vermuteten 6.000 Euro der Konkurrenz angeboten. Das war unsere einzige Runde mit deutlichem Plus. In Runde 4 und 5 haben wir weniger Absatz erwartet und nur eine Palette produziert, aber zu 9.500 Euro. In Runde 6 wollten wir zum Schluss noch einmal Umsatz machen."));
  const months = ["Jan./Feb.", "März/Apr.", "Mai/Juni", "Juli/Aug.", "Sep./Okt.", "Nov./Dez."];
  const why = [
    "Economy statt Standard. Günstig starten und alles verkaufen.",
    "Alles war weg → Preis drastisch erhöht. 2. Maschine gekauft, noch ohne Personal.",
    "Knapp unter den vermuteten 6.000\u00A0€ der Konkurrenz. 2. Mitarbeiter → Maschine 2 läuft.",
    "Weniger Absatz erwartet („Beliebtheit nicht dauerhaft“) → nur 1 Palette.",
    "Marktanalyse sagte weniger Absatz voraus → wieder nur 1 Palette.",
    "Letzte Runde: noch einmal 13 Paletten, um Umsatz zu machen.",
  ];
  const cw = 3.85, ch = 2.5;
  R.forEach((r, i) => {
    const x = 0.7 + (i % 3) * (cw + 0.175), y = 1.5 + Math.floor(i / 3) * (ch + 0.15);
    const good = wirt[i] >= 1;
    box(s, x, y, cw, ch, C.background2, "Runde " + (i + 1));
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.25, y: y + 0.22, w: 0.55, h: 0.55, fill: { color: good ? C.accent3 : C.accent1 }, line: { type: "none" }, objectName: "Nr " + r });
    s.addText(String(i + 1), { x: x + 0.25, y: y + 0.22, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(months[i], { x: x + 0.95, y: y + 0.22, w: 1.4, h: 0.55, fontSize: 13, color: C.accent4, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(`${fmt(preis[i])}\u00A0€`, { x: x + 2.0, y: y + 0.18, w: 1.6, h: 0.62, fontSize: 22, bold: true, fontFace: "Cambria", color: C.text2, align: "right", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(why[i], { x: x + 0.25, y: y + 0.88, w: cw - 0.5, h: 0.9, fontSize: 15, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
    const erg = umsatz[i] - kosten[i];
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.25, y: y + 1.85, w: cw - 0.5, h: 0.45, rectRadius: 0.06, fill: { color: good ? "E3EFCF" : C.background1 }, line: { type: "none" }, objectName: "Ergebnis " + r });
    s.addText(`Ergebnis ${erg >= 0 ? "+" : "−"}${fmt(Math.abs(erg))}\u00A0€ → Kasse ${fmt(schluss[i])}\u00A0€`, { x: x + 0.35, y: y + 1.85, w: cw - 0.7, h: 0.45, fontSize: 14, bold: true, color: good ? "3D6610" : C.text2, valign: "middle", isTextBox: true, margin: 0 });
  });
}

// =========================================================== 6 Wirtschaftsdienst
{
  const s = content("Wirtschaftsdienst: Kosten von außen", "Verlauf",
    NOTE("ca. 0:50 min", "Ab Runde 2 kamen die Meldungen des Wirtschaftsdienstes dazu. Bei uns stehen sie im Kassenbuch als sonstige Kosten: insgesamt 34.000 Euro, am meisten in Runde 2 und Runde 5 mit je 10.000 Euro. Das sind rund zwölf Prozent unserer gesamten Kosten. Rein rechnerisch läge unser Verkaufswert ohne diese Kosten bei 117.770 Euro, also über dem Startkapital. Diese Kosten konnten wir nicht beeinflussen, aber wir hätten sie bei Preis und Menge einplanen können. Das soll keine Ausrede sein, zeigt aber, wie stark äußere Einflüsse wirken – und dass wir Rücklagen gebraucht hätten. [Wenn ihr den Wortlaut der Meldungen noch wisst: hier pro Runde ein Stichwort nennen.]"));
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

// =========================================================== 7 Kennzahlen
{
  const s = content("Unsere Kennzahlen je Runde", "Auswertung",
    NOTE("ca. 1:00 min", "Für die Auswertung haben wir zwei Kennzahlen berechnet. Die wichtigste ist die Wirtschaftlichkeit: Umsatzerlöse geteilt durch die Kosten der Runde. In Runde 3 waren das 79.940 Euro Umsatz geteilt durch 54.000 Euro Kosten, also 1,48 – für jeden eingesetzten Euro kamen 1,48 Euro zurück. Über 1 heißt Gewinn, unter 1 Verlust; das schaffen nur Runde 3 und 6. Die zweite Kennzahl ist die Absatzquote: Wie viel Prozent unserer Produktion haben wir am normalen Markt verkauft? Sie zeigt, ob unser Preis zum Markt gepasst hat – von 100 Prozent in Runde 1 bis 20 Prozent in Runde 5. Zur Aussagekraft: Die Wirtschaftlichkeit rechnet Maschinenkäufe voll in die Kaufrunde. Würde man nur den Wertverlust von 10 Prozent ansetzen, wären auch Runde 1 und 2 knapp wirtschaftlich gewesen. Die vollständige Kostentabelle steht in unserem Auswertungsbogen M15."));
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.petrol }, align: "center" } });
  const rows = [[hdr("in\u00A0€"), ...R.map(hdr)]];
  const sum = (idx) => R.map((_, i) => idx.reduce((a, k) => a + kostenZeilen[k][1][i], 0));
  const zeilen = [
    ["Maschinen", sum([0])],
    ["Personal + Miete", sum([1, 2])],
    ["Wirtschaftsdienst", sum([3])],
    ["Produktion + Material", sum([4, 5])],
  ];
  zeilen.forEach(([n, v]) => rows.push([{ text: "– " + n }, ...v.map((x) => ({ text: fmt(x), options: { align: "right" } }))]));
  rows.push([{ text: "= Kosten gesamt", options: { bold: true } }, ...kosten.map((x) => ({ text: fmt(x), options: { bold: true, align: "right" } }))]);
  rows.push([{ text: "Umsatzerlöse", options: { bold: true } }, ...umsatz.map((x) => ({ text: fmt(x), options: { bold: true, align: "right" } }))]);
  rows.push([{ text: "Wirtschaftlichkeit", options: { bold: true, color: "FFFFFF", fill: { color: HEX.petrol } } }, ...wirt.map((w) => ({
    text: w.toFixed(2).replace(".", ","), options: { bold: true, align: "center", color: w >= 1 ? "3D6610" : "8A4A12", fill: { color: w >= 1 ? "E3EFCF" : "F7E4CF" } },
  }))]);
  rows.push([{ text: "Absatzquote", options: { bold: true, color: "FFFFFF", fill: { color: HEX.petrol } } }, ...R.map((_, i) => ({
    text: Math.round((markt[i] / produziert[i]) * 100) + " %", options: { bold: true, align: "center", color: HEX.petrol, fill: { color: "E6EEF0" } },
  }))]);
  s.addTable(rows, { x: 0.7, y: 1.55, w: 8.4, colW: [2.4, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0], fontSize: 15, color: HEX.text, border: { type: "solid", color: "D3D1C8", pt: 0.75 }, rowH: 0.5, valign: "middle", objectName: "Tabelle Kennzahlen" });
  s.addText("Vollständige Kostentabelle: Auswertungsbogen M15", { x: 0.7, y: 6.15, w: 8.4, h: 0.35, fontSize: 12, italic: true, color: C.accent4, isTextBox: true, margin: 0 });
  box(s, 9.35, 1.55, 3.25, 5.15, C.background2, "Formeln");
  iconCircle(s, "FaBalanceScale", 9.6, 1.75, 0.65, C.accent1);
  s.addText("So rechnen wir", { x: 10.4, y: 1.75, w: 2.1, h: 0.65, fontSize: 18, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
  s.addText([
    { text: "Wirtschaftlichkeit", options: { bold: true, breakLine: true } },
    { text: "= Umsatzerlöse ÷ Kosten", options: { breakLine: true } },
    { text: "R3: 79.940 ÷ 54.000 = 1,48", options: { color: C.accent4, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Absatzquote", options: { bold: true, breakLine: true } },
    { text: "= am Markt verkauft ÷ produziert", options: { breakLine: true } },
    { text: "R3: 12,6 ÷ 18 = 70 %", options: { color: C.accent4, breakLine: true, paraSpaceAfter: 10 } },
    { text: "Aussagekraft: ", options: { bold: true } },
    { text: "über 1 = Gewinn. Die Absatzquote zeigt, ob der Preis zum Markt passte.", options: {} },
  ], { x: 9.6, y: 2.55, w: 2.8, h: 4.0, fontSize: 14, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
}

// =========================================================== 8 Entwicklung
{
  const s = content("Wirtschaftlichkeit und Schlussbestand", "Auswertung",
    NOTE("ca. 0:55 min", "Beide Grafiken kennt ihr aus dem Auswertungsbogen. Links die Wirtschaftlichkeit je Runde, die gestrichelte Linie markiert 1, ab dort lohnt sich eine Runde. Nur Runde 3 und 6 liegen darüber, über alle sechs Runden kommen wir auf 0,84. Runde 4 und 5 sind besonders schwach: Wir haben kaum etwas verkauft, aber Personal und Miete liefen weiter. Rechts seht ihr, wie sich unser Startkapital über die Schlussbestände entwickelt hat. Wir sind mit 100.000 Euro gestartet. Die Maschinenkäufe drücken Runde 1 und 2, nach Runde 3 hatten wir mit 82.840 Euro unseren besten Stand, nach Runde 5 mit 47.790 Euro den Tiefpunkt, und am Ende stehen 53.770 Euro in der Kasse. Was das Unternehmen insgesamt wert ist, zeigt die nächste Folie."));
  s.addChart([
    { type: pres.charts.BAR, data: [{ name: "Wirtschaftlichkeit", labels: R, values: wirt }], options: { chartColors: [HEX.petrol], barGapWidthPct: 60, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00" } },
    { type: pres.charts.LINE, data: [{ name: "Grenze 1,0 (Kosten gedeckt)", labels: R, values: [1, 1, 1, 1, 1, 1] }], options: { chartColors: [HEX.amber], lineSize: 2, lineDash: "dash", lineDataSymbol: "none", showValue: false } },
  ], Object.assign({}, chartBase, { x: 0.7, y: 1.5, w: 6.0, h: 4.6, title: "Wirtschaftlichkeit je Runde", showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 1.6, valAxisMajorUnit: 0.2, valAxisLabelFormatCode: "0.0", catAxisLabelFontSize: 13, valAxisLabelFontSize: 12, dataLabelFontSize: 12 }));
  s.addChart(pres.charts.LINE, [{ name: "Schlussbestand", labels: ["Start", ...R], values: [100000, ...schluss] }], Object.assign({}, chartBase, {
    x: 6.9, y: 1.5, w: 5.7, h: 4.6, chartColors: [HEX.petrol], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8,
    title: "Schlussbestand je Runde (€), Start 100.000\u00A0€", showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 120000, valAxisMajorUnit: 20000, valAxisLabelFormatCode: "#,##0",
    showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "#,##0", dataLabelFontSize: 11, catAxisLabelFontSize: 13, valAxisLabelFontSize: 12,
  }));
  box(s, 0.7, 6.2, 11.9, 0.6, C.background2, "Kernaussage");
  s.addText([{ text: "Kernaussage: ", options: { bold: true, color: C.accent2 } }, { text: "Nur 2 von 6 Runden über 1, gesamt 0,84 – Schlussbestand 53.770\u00A0€ (mit Maschinen: nächste Folie).", options: { color: C.text1 } }],
    { x: 0.95, y: 6.2, w: 11.4, h: 0.6, fontSize: 16, valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 9 Verkaufswert
{
  const s = content("Was unser Unternehmen am Ende wert ist", "Auswertung",
    NOTE("ca. 0:50 min", "Am Ende wird unser Unternehmen verkauft. Der Wert ist der Schlussbestand plus der Restwert der Maschinen. Der Schlussbestand nach Runde 6 sind 53.770 Euro. Jede Maschine verliert pro genutzter Runde zehn Prozent ihres Kaufpreises: Die erste Maschine lief sechs Runden und ist noch 12.000 Euro wert, die zweite lief vier Runden, weil sie in Runde 2 ohne Mitarbeiter stillstand, und ist noch 18.000 Euro wert. Zusammen ergibt das einen Verkaufswert von 83.770 Euro. Das sind 16.230 Euro weniger als unser Startkapital. Zum Vergleich: Allein die zwei Maschinen haben 60.000 Euro gekostet, davon stecken noch 30.000 Euro im Unternehmenswert."));
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

// =========================================================== 10 Reflexion
{
  const s = content("Reflexion: Würden wir wieder so entscheiden?", "Reflexion",
    NOTE("ca. 1:20 min", "Im Nachhinein bewerten wir unsere Entscheidungen so. Gut war, dass wir vorsichtig mit nur einer Maschine gestartet sind, wie es die Spielanleitung empfiehlt, und dass wir nach Runde 2 schnell reagiert haben: Mit 5.900 Euro kam unsere beste Runde. Schlecht war vor allem unser Preis ab Runde 4. Wir haben weniger Absatz erwartet und trotzdem die höchsten Preise des ganzen Spiels verlangt, 9.500 und 8.400 Euro – dabei bestimmt der Preis den Absatz. Verkauft haben wir nur 20 bis 50 Prozent. Dazu kamen je 13.000 Euro Fixkosten für eine einzige Palette. Und die zweite Maschine für 30.000 Euro haben wir nur in Runde 3 wirklich gebraucht, sonst hätte eine gereicht. In Runde 1 war der Preis dagegen eher zu niedrig, da ging alles sofort weg. Anders machen würden wir: beim Preis von Runde 3 bleiben, Maschinen nur kaufen, wenn wir sie auslasten und besetzen können, und das Personal an die Produktion anpassen. Unser Fazit: Der Preis war unser stärkster Hebel, und Fixkosten laufen weiter, auch wenn man kaum verkauft."));
  const cols = [
    ["FaThumbsUp", "Das lief gut", C.accent3, ["Vorsichtiger Start mit 1 Maschine – wie in der Spielanleitung empfohlen", "Nach R2 schnell reagiert: 5.900\u00A0€ → beste Runde (1,48)"]],
    ["FaThumbsDown", "Das lief schlecht", C.accent2, ["R4–R6: hohe Preise trotz erwarteter Flaute – nur 20–50 % verkauft", "R4/R5: je 13.000\u00A0€ Fixkosten für 1 Palette", "2. Maschine (30.000\u00A0€) nur in R3 wirklich gebraucht"]],
    ["FaExchangeAlt", "Das machen wir anders", C.accent1, ["Preis nahe 5.900\u00A0€ halten – lief in R3 am besten", "Maschinen nur kaufen, wenn wir sie auslasten und besetzen", "Personal an die Produktion anpassen (4.000\u00A0€ je Person und Runde)"]],
  ];
  cols.forEach(([ic, h, col, items], i) => {
    const x = 0.7 + i * 4.025;
    box(s, x, 1.6, 3.85, 3.85, C.background2, "Karte " + h);
    iconCircle(s, ic, x + 0.3, 1.85, 0.7, col);
    s.addText(h, { x: x + 1.2, y: 1.85, w: 2.5, h: 0.7, fontSize: 19, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(bullets(items, 16), { x: x + 0.3, y: 2.75, w: 3.3, h: 2.6, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });
  box(s, 0.7, 5.65, 11.9, 1.0, C.accent1, "Fazit");
  s.addText([{ text: "Unser Fazit: ", options: { bold: true, color: C.accent2 } }, { text: "Der Preis war unser stärkster Hebel – und Fixkosten laufen weiter, auch wenn man kaum verkauft.", options: { color: C.background1 } }],
    { x: 1.0, y: 5.65, w: 11.4, h: 1.0, fontSize: 17, valign: "middle", isTextBox: true, margin: 0 });
}

// =========================================================== 11 Bezug zur Gründung
{
  const s = content("Bezug zur Gründung: Was im Spiel fehlte", "Gründung",
    NOTE("ca. 1:10 min", "Zum Schluss der Blick zurück auf unsere Gründung. Vieles, was uns wichtig war, kam im Spiel nicht vor, und jedes davon hätte etwas verändert. Mehrweg hätte Kosten gebracht, für Rücknahme, Spülen und Rücktransport. Unsere Zielgruppe mit Persona Jasper hätte bedeutet, dass der Absatz nicht nur vom Preis abhängt, sondern auch davon, wen man anspricht. Mit Social Media und Onlineshop hätten wir in Runde 4 und 5 Werbung machen können, statt nur eine Palette anzubieten. Und unsere besondere Sorte Tamarinde-Limette ist ein Alleinstellungsmerkmal – damit hätten wir vielleicht einen höheren Preis bei gleichem Absatz verlangen können. Im Spiel war aber jede Brause gleich. Wir schlagen deshalb zwei einfache Erweiterungen vor: Erstens ein Marketingbudget – je 2.000 Euro gibt es eine Palette mehr Absatz. Zweitens eine Mehrweg-Option: 1.000 Euro Zusatzkosten je Runde, dafür eine Palette mehr Absatz bei umweltbewussten Kunden. Beides ist nur eine Zeile mehr im Kassenbuch."));
  s.addText("Was im Spiel fehlte – und was es bewirkt hätte", { x: 0.7, y: 1.5, w: 7.6, h: 0.4, fontSize: 16, bold: true, color: C.accent4, isTextBox: true, margin: 0 });
  [["FaRecycle", "Mehrweg & Pfand", "Mehr Kosten: Rücknahme, Spülen, Rücktransport"],
   ["FaUserAlt", "Zielgruppe & Persona Jasper", "Absatz hinge auch davon ab, wen man anspricht"],
   ["FaBullhorn", "Social Media & Onlineshop", "Werbung hätte R4/R5 mehr Absatz bringen können"],
   ["FaLeaf", "Sorte Tamarinde-Limette", "Alleinstellung: höherer Preis bei gleichem Absatz"]].forEach(([ic, t, e], i) => {
    const y = 2.0 + i * 1.17;
    box(s, 0.7, y, 7.6, 1.02, C.background2, "Fehlt " + t);
    iconCircle(s, ic, 0.9, y + 0.17, 0.68, C.accent1);
    s.addText([{ text: t, options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: "→ " + e, options: { fontSize: 15, color: C.text1 } }],
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

// =========================================================== 12 Danke
{
  const s = slide("DARK", "Abschluss");
  s.addText("Vielen Dank!", { x: 0.9, y: 2.2, w: 6.5, h: 1.3, fontSize: 60, bold: true, fontFace: "Cambria", color: C.background1, isTextBox: true, margin: 0 });
  s.addText("Was dein Körper schon immer wollte – und was wir im Planspiel gelernt haben.", { x: 0.9, y: 3.6, w: 6.3, h: 1.0, fontSize: 20, color: C.background2, isTextBox: true, margin: 0 });
  s.addText("Wir freuen uns auf eure Fragen.", { x: 0.9, y: 4.8, w: 6.3, h: 0.5, fontSize: 20, bold: true, color: C.accent2, isTextBox: true, margin: 0 });
  s.addImage({ data: IMG.logo, x: 0.8, y: 5.75, w: 1.1, h: 1.1, altText: "Logo Elbharz", objectName: "Logo" });
  s.addShape(pres.ShapeType.ellipse, { x: 8.9, y: 1.2, w: 3.8, h: 3.8, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Sonne" });
  const fh = 5.4; s.addImage({ data: IMG.flasche, x: 10.8 - (fh * FLASCHE_RATIO) / 2, y: 0.75, w: fh * FLASCHE_RATIO, h: fh, altText: "Flasche Primatenbrause", objectName: "Flasche" });
  s.addShape(pres.ShapeType.ellipse, { x: 7.0, y: 4.3, w: 2.4, h: 2.4, fill: { color: C.accent3, transparency: 15 }, line: { color: C.background1, width: 1.5, dashType: "dash" }, objectName: "Platz KI-Teambild" });
  s.addText("BILD FOLGT: KI-Teambild", { x: 7.1, y: 4.3, w: 2.2, h: 2.4, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  s.addNotes(NOTE("ca. 0:10 min", "Das war unser Weg mit der Primatenbrause, von der Idee bis zum Verkaufswert. Vielen Dank fürs Zuhören, wir freuen uns auf eure Fragen."));
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
}

(async () => {
  await loadIcons(["FaBuilding", "FaUserAlt", "FaStore", "FaIndustry", "FaWarehouse", "FaTruck", "FaNewspaper", "FaBalanceScale", "FaCoins", "FaCogs",
    "FaThumbsUp", "FaThumbsDown", "FaExchangeAlt", "FaRecycle", "FaBullhorn", "FaLeaf", "FaChartLine", "FaLightbulb"]);
  build();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Kosten", kosten, "Wirtschaftlichkeit", wirt);
})();
