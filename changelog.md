# Changelog

## 2026-09-04 (2)
- Update-waarschuwing: bij het openen haalt de tool `version.json` van GitHub Pages op en toont een gele balk in het venster wanneer je bladwijzer een oudere versie van `importer.js` bevat, met link naar de installatiepagina. Mislukt de check (offline, geblokkeerd), dan gebeurt er gewoon niets.
- De versie wordt bij het deployen automatisch gestempeld: de workflow vervangt `__GI_VERSION__` in `importer.js` door de korte commit-hash van dat bestand en schrijft dezelfde waarde in `version.json`. Enkel wijzigingen aan `importer.js` triggeren dus een melding, niet bv. een README-aanpassing.
- De actieve versie staat nu in de titelbalk van het venster (bv. `v2026-09-04`); de tooltip toont ook de commit-hash. Een niet-gestempelde build toont "lokale versie".
- GitHub Pages stond op `legacy` (publiceerde de master-branch rechtstreeks), waardoor de workflow-output nooit online kwam en de versie-placeholders bleven staan. Pages staat nu op `build_type: workflow`, zodat het artifact van `deploy.yml` gepubliceerd wordt. Meteen ook het einde van de dubbele deployment (legacy `pages build and deployment` naast de eigen workflow).
- Minifier in `install.html` laat `//` na een `:` staan, zodat URLs in string literals de bookmarklet-generatie overleven.

## 2026-09-04
- Naam-matching houdt nu rekening met de extra codes die iBaMaFlex achter de naam plakt: `{8}`, `<J>` en `[B]` worden weggeknipt (vroeger enkel een `[...]` helemaal op het einde). Dat gebeurt aan beide kanten, dus ook codes die in de Excel achter de naam staan (bv. `Barry Aliou <N>`) worden genegeerd. De bijnaam tussen haakjes wordt zowel mét als zonder meegenomen bij het vergelijken, bv. `Intzidis Alki (Alkiviadis) {8} <J> [B]` matcht met `Intzidis Alki`.

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
