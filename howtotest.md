# Lokaal testen — Grade Importer

Dit document beschrijft hoe je het script / de bookmarklet lokaal kan testen zonder dat je op een live iBaMaFlex-pagina moet werken.

## Wat staat er in de repo?

- [importer.js](importer.js) — de volledige UI + import-logica (leesbare versie, bron van waarheid).
- [bookmarklet_loader.js](bookmarklet_loader.js) — minimale loader die SheetJS van CDN trekt en `importer.js` injecteert (handig tijdens dev).
- [install.html](install.html) — bundelt de inhoud van `<script id="importer-source">` tot één bookmarklet-URL (de productieflow).
- [Demo/AP iBaMaFlex!_ Puntenlijsten.html](Demo/AP%20iBaMaFlex%21_%20Puntenlijsten.html) — opgeslagen HTML-kopie van de echte puntenlijst-pagina. Hierop kan je testen.
- [Demo/SubgroepEmail.xlsx](Demo/SubgroepEmail.xlsx) — voorbeeld-Excel om te uploaden.

---

## Methode 1 — Snel: plakken in DevTools console (aanbevolen voor dev)

Geen webserver nodig. Goed voor iteratief debuggen van `importer.js`.

1. Open [Demo/AP iBaMaFlex!_ Puntenlijsten.html](Demo/AP%20iBaMaFlex%21_%20Puntenlijsten.html) in Chrome/Edge (dubbelklik volstaat — `file://` werkt).
2. Open DevTools (`F12`) → tab **Console**.
3. Laad eerst SheetJS:
   ```js
   var s=document.createElement('script');
   s.src='https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
   document.head.appendChild(s);
   ```
4. Wacht tot `typeof XLSX` `"object"` is, plak dan de **volledige inhoud van [importer.js](importer.js)** in de console en druk Enter.
5. De floating UI verschijnt rechtsboven. Upload [Demo/SubgroepEmail.xlsx](Demo/SubgroepEmail.xlsx) en test de matching.

> Bij elke wijziging in `importer.js`: refresh de demo-pagina (`Ctrl+F5`) en herhaal stap 3–4.

---

## Methode 2 — Lokale webserver + loader-bookmarklet

Dichter bij de productieflow: je test daadwerkelijk de `<script>`-injectie via een bookmarklet. Voordeel: je hoeft niet bij elke wijziging code opnieuw in de console te plakken — gewoon de pagina refreshen + bookmark opnieuw klikken.

1. Start een statische server vanuit de repo-root:
   ```powershell
   python -m http.server 8080
   ```
   (Of `npx serve -p 8080`.)
2. Maak in je browser een nieuwe bookmark met als URL:
   ```
   javascript:(function(){var s=document.createElement('script');s.src='http://127.0.0.1:8080/bookmarklet_loader.js?t='+Date.now();document.head.appendChild(s);})();
   ```
   Deze stub laadt [bookmarklet_loader.js](bookmarklet_loader.js), dat op zijn beurt SheetJS van CDN haalt en daarna `importer.js` van localhost. De `?t=...` cache-bust zorgt dat je altijd de laatste versie krijgt.
3. Open [http://127.0.0.1:8080/Demo/AP%20iBaMaFlex!_%20Puntenlijsten.html](http://127.0.0.1:8080/Demo/AP%20iBaMaFlex!_%20Puntenlijsten.html) (of de echte iBaMaFlex-pagina).
4. Klik de bookmark → SheetJS + importer.js worden geladen en de UI verschijnt rechtsboven.

> Bij elke wijziging in `importer.js`: pagina refreshen (de oude overlay verdwijnt) en bookmark opnieuw klikken. De server hoeft niet te herstarten.

---

## Methode 3 — Volledige productie-bookmarklet testen

Test de exacte versie die eindgebruikers krijgen via GitHub Pages.

1. Serveer de repo via een lokale webserver (zie Methode 2 stap 1) **of** push en gebruik de GitHub Pages URL. `install.html` haalt `importer.js` op via `fetch()`, dus `file://` werkt niet — de knop toont dan een rode foutmelding.
2. Open `install.html` via die URL (bv. [http://127.0.0.1:8080/install.html](http://127.0.0.1:8080/install.html)). De pagina bundelt `importer.js` + SheetJS-loader in één `javascript:`-URL op de knop.
3. Sleep de knop **🔢IbamaflexImporter** naar je bookmarks bar.
4. Open een puntenlijst-pagina (de Demo-HTML óf de echte iBaMaFlex-pagina).
5. Klik de bookmark → SheetJS wordt van CDN geladen, daarna de UI.

> `install.html` heeft sinds kort geen eigen kopie meer van de importer-source — het bestand fetcht `importer.js` runtime. Eén bron van waarheid, geen drift meer.

---

## Wat te verifiëren

Na een import zou je het volgende moeten zien op de Demo-pagina:

- Status-blok: `Done! Matched: X / Not found: Y`.
- Score-inputs in de tabel hebben de waarde uit Excel.
- Bij `Empty grade = Absent` aangevinkt: lege Excel-cellen → score-input lichtblauw (`#cfe2ff`) met een `+`-key-event afgevuurd.
- **Show Log**-knop toont studenten die niet matchten met de Excel.
- Split-modus (achternaam/voornaam): namen worden in beide volgorden geprobeerd én via token-set match.

## Veelvoorkomende issues

- **`XLSX is not defined`** in console → SheetJS niet (volledig) geladen voor je `importer.js` plakte. Herlaad en wacht langer.
- **`Cannot read properties of null (reading 'cells')`** → de Demo-HTML heeft de grid-tabel niet (of de selector `#ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00 tbody tr` matcht niets). Check in DevTools of die ID nog bestaat in de opgeslagen pagina.
- **Bookmarklet werkt niet vanaf `file://`** → sommige browsers blokkeren `javascript:` op lokale bestanden. Gebruik dan Methode 2 (lokale server).
