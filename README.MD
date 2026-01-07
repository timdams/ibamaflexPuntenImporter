# AP iBaMaFlex! Punten Importer

Een handige tool om puntenlijsten vanuit Excel automatisch in te laden in iBaMaFlex.

## 🚀 Installatie

1.  Navigeer naar [de installatie pagina](https://timdams.github.io/ibamaflexPuntenImporter/install.html).
2.  Zorg dat je **Bladwijzerbalk** (Bookmarks bar) zichtbaar is.
    *   *Sneltoets*: `Ctrl + Shift + B`
3.  Sleep de grote blauwe knop **"⬇️ Drag Me to Bookmarks Bar ⬇️"** naar je bladwijzerbalk.
4.  Klaar! Je ziet nu een nieuwe knop **"Import Grades"** (of de naam die je koos) in je balk staan.

## 📝 Gebruik

1.  **Navigeer naar iBaMaFlex**
    *   Open de puntenlijst pagina in iBaMaFlex waar je de punten wilt invoeren.

2.  **Start de Tool**
    *   Klik op de **"Import Grades"** knop in je bladwijzerbalk.
    *   Er verschijnt een pop-up venster rechtsboven in je scherm.

3.  **Excel Inladen**
    *   Klik op "Bestand kiezen" en selecteer je Excel-bestand (`.xlsx`).
    
4.  **Kolommen Kiezen**
    *   De tool probeert automatisch te raden welke kolommen de **Student** en het **Cijfer** bevatten.
    *   Controleer de selectie in de dropdown menu's en pas aan indien nodig.

5.  **Importeren**
    *   Klik op **"Import Grades"**.
    *   De tool verwerkt de lijst:
    *   ✅ **Groen**: Punt succesvol ingevuld.
    *   🔵 **Blauw**: Afwezigheid (`A`) ingevuld (simuleert `+` toets).
    *   ⚠️ **Geel**: Student gevonden in Excel, maar geen punt aanwezig.
    *   ⚪ **Wit**: Student niet gevonden in Excel.

6.  **Controleren & Opslaan**
    *   Kijk de ingevulde punten na.
    *   Klik onderaan de pagina op **Opslaan** om te bevestigen.

## ⚙️ Hoe werkt het matchen?

*   **Strikte Naamcontrole**: De naam in Excel moet **exact** overeenkomen met de naam in iBaMaFlex.
    *   *Voorbeeld*: "Tim Dams" matcht **niet** met "Timm Dams" (typfout).
*   **Suffix Negeren**: De tool negeert automatisch toevoegingen tussen vierkante haken in iBaMaFlex.
    *   *Voorbeeld*: "Tim Dams" in Excel matcht met "Tim Dams [B]" in iBaMaFlex.

## 🅰️ Afwezigheden

*   Als je in de Excel een **"A"** (of "a") als cijfer zet, zal de tool dit herkennen als "Afwezig".
*   In plaats van "A" in te vullen, simuleert de tool een druk op de **`+` toets**.
*   Dit zorgt ervoor dat iBaMaFlex correct de status "Niet deelgenomen" activeert.

## ❓ Veelgestelde Vragen

**Worden mijn gegevens ergens opgeslagen?**
Nee. De tool werkt volledig lokaal in je browser. Je Excel-bestand verlaat nooit je computer.

**Werkt dit met elke Excel?**
Ja, zolang er maar een kolom is met namen en een kolom met cijfers.
