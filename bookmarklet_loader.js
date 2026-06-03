/**
 * Grade Importer — DEV bookmarklet loader
 *
 * Doel: tijdens lokale ontwikkeling laden vanaf een localhost-server, zodat je
 * importer.js kan editen + browser refreshen zonder elke keer een nieuwe
 * bookmarklet te bouwen via install.html.
 *
 * Setup:
 *   1. Serveer de repo-root via een statische server, bv:
 *        python -m http.server 8080
 *   2. Maak een bookmark met deze URL (de inline versie van wat dit bestand doet):
 *
 *      javascript:(function(){var s=document.createElement('script');s.src='http://127.0.0.1:8080/bookmarklet_loader.js?t='+Date.now();document.head.appendChild(s);})();
 *
 *   3. Open de iBaMaFlex puntenlijst (of Demo/AP iBaMaFlex!_ Puntenlijsten.html
 *      via dezelfde server) en klik de bookmark.
 *
 * Dit bestand laadt eerst SheetJS van CDN, dan importer.js van dezelfde
 * localhost-server. De `?t=...` cache-bust zorgt dat je altijd de laatste
 * versie van importer.js krijgt.
 *
 * Voor eindgebruikers: zie install.html, niet dit bestand.
 */
(function () {
    if (window.GradeImporter) {
        alert('Grade Importer is already loaded!');
        return;
    }

    // Resolve the origin we were loaded from, so importer.js is fetched from
    // the same server (typically http://127.0.0.1:8080).
    var loaderScript = document.currentScript ||
        (function () {
            var scripts = document.getElementsByTagName('script');
            for (var i = scripts.length - 1; i >= 0; i--) {
                if (/bookmarklet_loader\.js/.test(scripts[i].src)) return scripts[i];
            }
            return null;
        })();

    var importerUrl = loaderScript
        ? loaderScript.src.replace(/bookmarklet_loader\.js.*$/, 'importer.js?t=' + Date.now())
        : 'http://127.0.0.1:8080/importer.js?t=' + Date.now();

    console.log('[GradeImporter dev loader] Loading SheetJS…');

    var sheetJs = document.createElement('script');
    sheetJs.src = 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
    sheetJs.onload = function () {
        console.log('[GradeImporter dev loader] SheetJS loaded. Loading importer.js from', importerUrl);
        var importer = document.createElement('script');
        importer.src = importerUrl;
        importer.onerror = function () {
            alert('Kon importer.js niet laden van ' + importerUrl + '. Draait je localhost-server?');
        };
        document.head.appendChild(importer);
    };
    sheetJs.onerror = function () {
        alert('Kon SheetJS niet laden van CDN.');
    };
    document.head.appendChild(sheetJs);
})();
