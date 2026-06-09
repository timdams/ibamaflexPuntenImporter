# Changelog

## 2026-06-09
- Optie "Kommapunten toestaan" toegevoegd: standaard worden decimale cijfers (bv. 12,5) overgeslagen en in de log gemeld; aanvinken neemt ze wél over (punt wordt automatisch omgezet naar komma). Summary toont aantal overgeslagen kommapunten.

## 2026-06-03
- `install.html` haalt `importer.js` nu runtime via `fetch()` — geen gedupliceerde importer-source meer in install.html (single source of truth)
- `bookmarklet_loader.js` opgeschoond tot een werkende dev-loader (laadt SheetJS + importer.js vanaf localhost)
- `debug_bookmarklet.js` verwijderd (verouderde orphan)
- `howtotest.md` toegevoegd met drie testmethodes (console-paste, lokale server, productie-bookmarklet)
- Naam-matching uitgebreid: optie voor namen verdeeld over 2 kolommen (achternaam/voornaam) met token-set fallback
- `index.html` toegevoegd (redirect naar `install.html`) — fixt de 404 op de root-URL
- App-overlay volledig vernederlandst + korte uitleg-zinnetjes bij elke optie (o.a. wat "eerste rij = kolomtitels" betekent)
- `install.html` herschreven tot een stap-voor-stap installatiegids (bladwijzerbalk tonen met Ctrl+Shift+B, knop slepen, gebruik in iBaMaFlex) i.p.v. enkel een versielabel

## 2026-01-07
- readme
- Add GitHub Pages deployment workflow
- Remove Demo folder from repository, keep locally
- betere naam bookmark
- log feature
- empty grades= absent optie
- werkende versie die wél cijfers bijhoudt :)
- Update README with GitHub Pages link
- Final version for import
- prototype 1
- geen flexibele matching
