# Plan: Endpräsentation Version 1 – Elbharz Getränke GbR

Zahlen und Aussagen: siehe `FAKTEN.md` (einzige Quelle). Erzeugen: `node build.js`.

## Leitidee
**„Von der Idee zum Verkaufswert“** – die Endpräsentation erzählt das Spiel, nicht noch einmal die Gründung.
- Die Firmenpräsentation (Phase I) wird nur kurz aufgegriffen (2 Folien), damit sie wiedererkennbar ist, aber nicht wiederholt wird.
- Schwerpunkt auf dem, was 60 % der Note trägt: Verlauf, Auswertung, Reflexion, Bezug zur Gründung.
- Ehrlicher roter Faden: Wir haben 16.230 € verloren – und können genau erklären, warum.

## Design
- **Farben unverändert:** Petrol `0F5C6B`, Bernstein `D9822B`, Limette `8DB33A`, gebrochenes Weiß `F3F1EA`, Text `1B2A2F`.
- **Wiedererkennung:** Logo, Flasche, Limette und Tamarinde aus der Firmenpräsentation; Sonne/Kreis aus dem Logo als Motiv.
- **Abgrenzung zur Firmenpräsentation:** Serifen-Überschriften (Cambria) statt Großbuchstaben-Sans, abgerundete Karten statt Balken oben/unten, Daten im Mittelpunkt.
- **Kapitel-Leiste unten auf jeder Inhaltsfolie:** Einleitung · Verlauf · Auswertung · Reflexion · Gründung (aktuelles Kapitel orange). Zeigt der Lehrkraft, dass alle Pflichtteile da sind.
- **Diagramme:** native PowerPoint-Diagramme (bearbeitbar), eine Achse, Legende ab 2 Reihen, gezielte Beschriftung, ruhiges Raster.

## Folien (Ziel ca. 11 Minuten)
| # | Folie | Pflichtteil | Visuell | Min. |
|---|---|---|---|---|
| 1 | Titel | – | Flasche im orangen Kreis, Limette + Tamarinde am grünen Kreis (Idee der Gruppe), Logo | 0:15 |
| 2 | Wer wir sind | Einleitung | Leitgedanke + **Rückseite**: Etikett mit den 6 Zutaten (Idee der Gruppe) + 3 Fakten | 0:45 |
| 3 | Nachhaltigkeit | Einleitung | Mehrweg / Regional / Eine Sorte (Produktion → Lager → Transport als Icon-Kette) + ehrlicher Hinweis „im Spiel keine Rolle“ | 0:45 |
| 4 | Produziert, angeboten, verkauft | Verlauf | Gestapelte Säulen je Runde: Markt vs. Schnäppchenmarkt, Preis über jeder Säule | 1:00 |
| 5 | Warum wir so entschieden haben | Verlauf | 6 Rundenkarten (Preis, Kernaussage aus dem Protokoll, Ergebnis-Marke) | 1:15 |
| 6 | Wirtschaftsdienst | Verlauf | Zusatzkosten je Runde (Säulen) + Kennzahl 34.000 € + Wirkung | 0:45 |
| 7 | Kennzahlen je Runde | Auswertung | Kompakte Kostentabelle, Wirtschaftlichkeit + Absatzquote farbig, Formeln und Aussagekraft | 1:00 |
| 8 | Entwicklung | Auswertung | Säulen Wirtschaftlichkeit mit Linie bei 1,0 + Kurve Kassenbestand (beides wie M15) | 1:15 |
| 9 | Verkaufswert | Auswertung | Rechnung 53.770 + 12.000 + 18.000 = 83.770 € mit Icons (Geld, Maschinen, Gebäude) + Vergleich zum Startkapital | 0:45 |
| 10 | Reflexion | Reflexion | gut / schlechter / Alternative je Entscheidung, mit Zahlen belegt | 1:15 |
| 11 | Bezug zur Gründung | Gründung | Was fehlte (Mehrweg, Persona Jasper, Marketing, Sorte) → Auswirkung → Spielerweiterung | 1:15 |
| 12 | Danke | – | Flasche + Platz für KI-Teambild | 0:15 |
| 13 | Quellen | – | Spielmaterial, eigene Berechnungen, KI-Bilder gekennzeichnet | – |

## Bilder
- Im Repo verfügbar und eingebaut: Logo, Flasche, Limette, Tamarinde (aus der Firmenpräsentation).
- Als Icons gelöst: Produktion/Lager/Transport, Geld, Maschinen, Gebäude.
- **Fehlen (Netzwerk blockiert Bilddienste):** echte Fotos für leere Mehrwegflasche, Karte Norddeutschland mit Elbe, optional Geldstapel/Abfüllmaschine/Firmengebäude; KI-Teambild. → Platzhalter sind markiert, Bilder liefert die Gruppe oder der Zugriff wird freigegeben.

## Skills
- Genutzt: `pptx` (Aufbau, Prüfung), `dataviz` (Diagrammregeln), `pdf` (Auswertungsbogen M15).
- Team-Leitfaden: als Dokument zum Teilen (Sprechtexte, Rollen, Zeitplan, Fragen-Vorbereitung).
- Optional, nur nach Rückfrage: `the-council` (5 Agenten prüfen die Präsentation kritisch, kostet viele Tokens), Goodnotes-Verbindung (Leitfaden direkt in Goodnotes ablegen).

## Prüfung
1. Datei-Validierung und Render jeder Folie.
2. Blick als Zuschauer und als Lehrkraft: Pflichtteile abgedeckt? Zahlen nachvollziehbar? Kein Widerspruch zur Firmenpräsentation?
3. Zahlen-Abgleich Folie ↔ `FAKTEN.md`.

## Offen (braucht die Gruppe)
- [ ] Herkunft der Bilder Limette und Tamarinde (für die Quellenfolie)
- [ ] Fotos: leere Flasche, Karte Norddeutschland, optional Geld/Maschine/Gebäude
- [ ] KI-Teambild (mit Tool-Angabe)
- [ ] Optional: Wortlaut der Wirtschaftsdienst-Meldungen; Ursache Materialpreis R4
- [ ] Sprecher festlegen, Sprechtexte persönlich anpassen
- [ ] Reflexion gemeinsam prüfen (ist eure Bewertung, nicht meine)
- [ ] Platzierung nach der Siegerehrung ergänzen (optional)

## Bereits erledigt
- [x] Version 1 gebaut (`Elbharz_Endpraesentation.pptx` + `.pdf`), unabhängig geprüft (Zuschauer- und Lehrkraftsicht), Korrekturen eingearbeitet
- [x] Team-Leitfaden (Rollen, Zeitplan, Sprechtexte, Fragen, To-dos, Checkliste): https://claude.ai/code/artifact/4ef3e79e-1018-4975-a9c0-da88face15e6
- [x] Zahlen aller 6 Runden, Entscheidungsprotokoll, Verkaufswert, Auswertungsbogen M15 (`Auswertung_Elbharz_ausgefuellt.pdf`)
- [x] Wirtschaftsdienst = sonstige Kosten; R2 zweite Maschine ungenutzt; Restwert über Nutzungsdauer geklärt
- [x] Ideen der Gruppe aus den Stift-Notizen übernommen (`vorlagen/`)
