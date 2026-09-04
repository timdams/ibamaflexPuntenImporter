# CLAUDE.md — iBaMaFlex Punten Importer

Instructies en context voor Claude Code bij het werken in deze repo.

## Wat is dit?

Een **bookmarklet** die op een iBaMaFlex-puntenlijstpagina een floating overlay injecteert,
een Excel (`.xlsx`) inleest en de cijfers automatisch in de score-inputs van de pagina zet.
Doelgroep: AP-docenten. Alles draait client-side; er wordt niets geüpload.

- Live: https://timdams.github.io/ibamaflexPuntenImporter/ (redirect naar `install.html`)
- Repo: https://github.com/timdams/ibamaflexPuntenImporter
- Wishlist/issues staan op GitHub Issues, niet meer in `wishlist.md`.

**Geen build, geen dependencies, geen package.json.** Vanilla JS, statische GitHub Pages.
De enige externe library is SheetJS (`xlsx-0.20.1`), die runtime van CDN wordt geladen.

## Bestanden

| Bestand | Rol |
|---|---|
| `importer.js` | **De hele applicatie** — UI + matching + import-logica. Single source of truth. Bijna al het werk gebeurt hier. |
| `install.html` | Installatiegids voor eindgebruikers. Fetcht `importer.js` runtime, minifiet het naïef en bouwt daar de `javascript:`-bookmarklet-URL van. |
| `index.html` | Redirect naar `install.html` (fixt 404 op de root-URL). |
| `bookmarklet_loader.js` | **Dev-only** loader: haalt SheetJS + `importer.js` van localhost. Eindgebruikers raken dit nooit aan. |
| `version.json` | Wordt bij deploy gegenereerd. In de repo staat enkel de `"dev"`-placeholder. |
| `.github/workflows/deploy.yml` | Stempelt de versie en publiceert de repo-root naar GitHub Pages. |
| `changelog.md` | Handmatig bijgehouden, nieuwste bovenaan, datum als heading. |
| `howtotest.md` | Drie manieren om lokaal te testen. Lees dit voor je iets test. |
| `Demo/` | **Gitignored.** Lokale kopie van een echte puntenlijstpagina + voorbeeld-Excel om op te testen. |
| `README.MD` | Gebruikersdocumentatie (installatie, gebruik, matching, afwezigheden, FAQ). |

## Versionering — dit gebeurt automatisch, niet met de hand

**Kort antwoord: nee, je hoeft bij een commit géén versienummer te verhogen. Dat doet de workflow.**

Bij elke push naar `master` draait `deploy.yml` en:

1. `VERSION=$(git log -1 --format=%h -- importer.js)` — de korte commit-hash van de **laatste
   commit die `importer.js` aanraakte**.
2. `sed -i "s/__GI_VERSION__/$VERSION/" importer.js` — vervangt de placeholder in het
   gepubliceerde artefact (de repo behoudt de placeholder).
3. Schrijft `version.json` met diezelfde hash + datum.

De draaiende bookmarklet vergelijkt zijn ingebakken `GI_VERSION` met de opgehaalde
`version.json` en toont een gele "Nieuwe versie beschikbaar"-balk bij verschil.

Gevolgen om te onthouden:

- De versie verandert **alleen** wanneer `importer.js` mee in de commit zit. Een commit die
  enkel `README.MD` of `install.html` wijzigt, triggert dus terecht géén update-melding bij
  gebruikers — die hebben immers dezelfde importer-code.
- Verwijder of hernoem `__GI_VERSION__` in `importer.js` niet. Lokaal blijft de placeholder
  staan en slaat `checkForUpdate()` zichzelf over (`if (GI_VERSION.indexOf('__') === 0) return;`).
- Verzin geen semver-nummers en bump niets handmatig. Wél handmatig: **een entry in
  `changelog.md`** bij elke betekenisvolle wijziging.
- `versie 12` in de footer van `install.html` (regel ~241) is een los, handmatig label dat
  niets met dit mechanisme te maken heeft. Het drift makkelijk — bump het als je het opmerkt
  of laat het staan, maar verwar het niet met de echte versie.

## Werkafspraken per commit

1. Wijziging in `importer.js` → getest volgens `howtotest.md` (minstens methode 1).
2. Regel bijgeschreven in `changelog.md` onder een datum-heading (nieuwste bovenaan,
   Nederlands, in de stijl van de bestaande entries). Meerdere releases op één dag krijgen
   `## YYYY-MM-DD (2)`.
3. Raakt het de gebruikersflow of de opties? Dan ook `README.MD` en/of de uitlegzinnetjes in
   de overlay bijwerken.
4. Committen/pushen alleen wanneer de gebruiker daarom vraagt. Push naar `master` = deploy naar
   productie voor echte docenten.

## Technische valkuilen

**De naïeve minifier in `install.html`.** De bookmarklet wordt gebouwd met
`code.replace(/(^|[^:])\/\/[^\n]*/g, '$1').replace(/\s+/g, ' ')`. Concreet betekent dat voor
`importer.js`:

- `//` in string literals overleeft **enkel** wanneer er een `:` voor staat (`https://…`).
  Schrijf dus geen protocol-loze `//cdn…`-URLs en geen `//` in andere strings of regexes.
- Alle whitespace wordt tot één spatie gecollapst: **geen ASI**. Elke statement moet met een
  `;` eindigen en er mogen geen `//`-comments blijven die de rest van een regel opeten.
- Newlines binnen template literals worden spaties. Dat is oké voor de CSS/HTML-strings die er
  nu in staan; introduceer geen template literal waar een echte newline betekenis heeft.
- Gebruik `/* */`-comments liever niet — die worden niet gestript en blijven meelopen.

**DOM-koppeling met iBaMaFlex.** Twee harde selectors:

- Rijen: `#ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00 tbody tr`
- Naam in `row.cells[1]`, score-input via `input[name$="txtScore"]`

Verandert iBaMaFlex zijn ASP.NET-markup, dan breekt dit. Er is geen fallback; toon dan een
duidelijke foutmelding in plaats van stil te falen.

**Afwezigheid.** Wordt niet als tekst `"A"` ingevuld maar door een `+`-keyevent op de input te
simuleren — dat is wat iBaMaFlex' eigen handler nodig heeft om "Niet deelgenomen" te zetten.
Niet vervangen door een simpele `value = 'A'`.

**Naam-matching** (`stripNameMarkers` / `normalizeName` / `nameVariants` / `tokenize`):
iBaMaFlex plakt codes achter namen — `{8}`, `<J>`, `[B]` — die worden weggeknipt. Een bijnaam
tussen haakjes wordt zowel mét als zonder geprobeerd. Er is een split-modus voor Excels met
aparte achternaam-/voornaamkolommen, met token-set fallback. Matching blijft bewust **strikt**
(geen fuzzy/levenshtein): een fout gematcht punt is erger dan een niet-gevonden student. Houd
dat zo.

**Kommapunten.** Standaard worden decimale cijfers overgeslagen en gelogd; de optie
"Kommapunten toestaan" neemt ze over en zet `.` om naar `,`.

**Logica van de log.** Er wordt vanuit de studenten in iBaMaFlex geredeneerd, niet vanuit de
Excel. Studenten die in de Excel staan maar niet in iBaMaFlex, worden genegeerd. Dat is
opzettelijk (de Excel is mogelijk outdated) en staat zo ook in de README.

## Taal en toon

Alle gebruikersgerichte tekst — overlay, `install.html`, `README.MD`, `changelog.md` — is
**Nederlands**. Code-comments zijn gemengd NL/EN; nieuwe comments in het Nederlands is prima.
De README gebruikt 🚨 voor waarschuwingen die docenten écht moeten lezen; die toon aanhouden.

## Testen

Zie `howtotest.md`. Snelste lus: `Demo/AP iBaMaFlex!_ Puntenlijsten.html` openen, SheetJS in de
console laden, dan de volledige `importer.js` plakken. Voor de bookmarklet-flow (inclusief de
minifier!) heb je een lokale server nodig — `install.html` fetcht en werkt niet via `file://`:

```powershell
python -m http.server 8080
```

Test na een wijziging aan de minifier of aan string literals **altijd** methode 3, want een
minifier-bug is onzichtbaar in methode 1.
