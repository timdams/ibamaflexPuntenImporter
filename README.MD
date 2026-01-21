# AP iBaMaFlex! Punten Importer

Een handige tool om puntenlijsten vanuit Excel automatisch in te laden in iBaMaFlex.

## 🚀 Installatie

1.  Navigeer naar [de installatie pagina](https://timdams.github.io/ibamaflexPuntenImporter/install.html).
2.  Zorg dat je **Bladwijzerbalk** (Bookmarks bar) zichtbaar is.
    *   *Sneltoets*: `Ctrl + Shift + B`
3.  Sleep de grote blauwe knop **"🔢IbamaflexImporter"** naar je bladwijzerbalk.
4.  Klaar! Je ziet nu een nieuwe knop **"🔢IbamaflexImporter"** (of de naam die je koos) in je balk staan.

**🚨VOOR JE DE TOOL GEBRUIKT: Lees alsjeblieft goed de werking hierna na EN zeker de teksten met 🚨.**

## 📝 Gebruik

1.  **Navigeer naar iBaMaFlex**
    *   Open de puntenlijst pagina in iBaMaFlex waar je de punten wilt invoeren. **OPGELET: doe dit best groep per groep.**

2.  **Start de Tool**
    *   Klik op de **"🔢IbamaflexImporter"** knop in je bladwijzerbalk.
    *   Er verschijnt een pop-up venster rechtsboven in je scherm.

3.  **Excel Inladen**
    *   Klik op "Bestand kiezen" en selecteer je Excel-bestand (`.xlsx`).
    *  Merk op dat deze puntenlijst ook punten mag bevatten van andere klassen of groeppen. Alleen de studenten die in iBaMaFlex staan, zullen punten krijgen.


4.  **Instellingen**
    *   **Excel contains headers**: Vink dit uit als je eerste rij direct een student bevat, en niet de titels van de kolommen (standaard aan).
    *   **Empty grade in excel = Absent**: Vink dit aan om studenten op afwezig te zetten indien ze in je excel geen waarde ingevuld hebben (standaard uit). (niet gevonden studenten in je excel worden altijd als leeg aangeduid)

5.  **Kolommen Kiezen**
    *   De tool probeert automatisch te matchen welke kolommen de **Student** en het **Cijfer** bevatten.
    *   Controleer de selectie in de dropdown menu's en pas aan indien nodig.

6.  **Importeren & Log**
    *   Klik op **"Import Grades"**.
    *   De tool verwerkt de lijst en toont een korte samenvatting.
    *   **Show Log**: Klik op deze knop (verschijnt na import) om details te zien over studenten die niet gevonden zijn of die op "Afwezig" gezet zijn. **🚨BEKIJK DEZE LOG GOED.🚨 ER WORDT VANUIT DE STUDENTEN OP IBAMAFLEX GEWERKT. NIET VANUIT JE MOGELIJK AL OUTDATED EXCEL.**
    *   ✅ **Groen**: Punt succesvol ingevuld.
    *   🔵 **Blauw**: Afwezigheid (`A`) ingevuld (simuleert `+` toets).
    *   ⚪ **Wit**: Student niet gevonden in Excel.

7.  **Controleren & Opslaan**
    *   Kijk de ingevulde punten na.
    *   Klik onderaan de pagina op **Opslaan** om te bevestigen.
    *   **🚨KLIK ZEKER NIET OP NAAR VOLGENDE PUNTENLIJST🚨**: druk op stoppen (**NIET VERGRENDELEN**) zodat je terug in het hoofdscherm van ibamaflex komt. En heropen dan de puntenlijst voor je dan beslist om al dan niet te vergrendelen.

## ⚙️ Hoe werkt het matchen?

*   **Strikte Naamcontrole**: De naam in Excel moet **exact** overeenkomen met de naam in iBaMaFlex.
    *   *Voorbeeld*: "Tim Dams" matcht **niet** met "Timm Dams" (typfout).
*   **Suffix Negeren**: De tool negeert automatisch toevoegingen tussen vierkante haken in iBaMaFlex.
    *   *Voorbeeld*: "Tim Dams" in Excel matcht met "Tim Dams [B]" in iBaMaFlex.

## 🅰️ Afwezigheden

*   Als je in de Excel een **"A"** (of "a") als cijfer zet, zal de tool dit herkennen als "Afwezig".
*   Je kan ook de optie **"Empty grade in excel = Absent"** aanvinken. Dan worden lege cellen ook als afwezigheid behandeld.
*   In plaats van "A" in te vullen, simuleert de tool een druk op de **`+` toets**.
*   Dit zorgt ervoor dat iBaMaFlex correct de status "Niet deelgenomen" activeert.

## ❓ Veelgestelde Vragen

**Worden mijn gegevens ergens opgeslagen?**
Nee. De tool werkt volledig lokaal in je browser. Je Excel-bestand verlaat nooit je computer.

**Werkt dit met elke Excel?**
Ja, zolang er maar een kolom is met namen en een kolom met cijfers.
