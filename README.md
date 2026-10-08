# CSV & Excel Converter (Ommen & Zwolle) — Technische Documentatie

Dit document beschrijft de **exacte werking van de huidige codebase** (`src/App.tsx`). Het documenteert nauwkeurig hoe bestanden worden ingelezen, hoe de koppeling en samenvoeging plaatsvinden, welke normalisatieregels gelden, hoe duplicaten worden afgehandeld en hoe de CSV-exports zijn gestructureerd.

---

## Inhoudsopgave

1. [Overzicht & Doel](#overzicht--doel)
2. [Ondersteunde Bestandsformaten & Inlezen](#ondersteunde-bestandsformaten--inlezen)
3. [Koppeling & Zoekindex (Klantreferentie, E-mail, Naam)](#koppeling--zoekindex-klantreferentie-e-mail-naam)
4. [Samenvoegen van Klantprofielen (Fase 1: Registratie)](#samenvoegen-van-klantprofielen-fase-1-registratie)
5. [Deduplicatie (Fase 2: Groepering)](#deduplicatie-fase-2-groepering)
6. [Normalisatieregels & Veldformaten](#normalisatieregels--veldformaten)
7. [High Value Logica (Drempelwaarde €400)](#high-value-logica-drempelwaarde-400)
8. [Productcategorisatie & Trefwoorden](#productcategorisatie--trefwoorden)
9. [Filters, Uitzonderingen & Niet-Gekoppelde Producten](#filters-uitzonderingen--niet-gekoppelde-producten)
10. [CSV-Export Specificaties](#csv-export-specificaties)
11. [Interface & Downloadmogelijkheden](#interface--downloadmogelijkheden)
12. [Gegevensverwerking & Privacy (AVG / GDPR Context)](#gegevensverwerking--privacy-avg--gdpr-context)
13. [Lokale Installatie & Ontwikkeling](#lokale-installatie--ontwikkeling)

---

## Overzicht & Doel

De applicatie is gebouwd in React (TypeScript) en Vite om data uit fitness- en clubmanagementsoftware te combineren en voor te bereiden voor exports naar bijvoorbeeld advertentieplatformen (zoals Meta / Facebook) of spreadsheetprogramma's.

Er zijn twee uploadsecties: één voor vestiging **Ommen** en één voor vestiging **Zwolle**. Binnen elke vestiging kunnen één of meerdere bestanden worden geüpload (zoals een klantenbestand en/of een producten-/abonnementenbestand).

---

## Ondersteunde Bestandsformaten & Inlezen

### Bestandsformaten
- **CSV (`.csv`)**: Wordt ingelezen via **PapaParse** met `header: true` en `skipEmptyLines: 'greedy'`. Scheidingstekens (zoals `;`, `,` of tab) worden automatisch door de parser gedetecteerd. Byte Order Marks (`\ufeff`) en aanhalingstekens rond kolomnamen worden verwijderd.
- **Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`)**: Wordt ingelezen via **SheetJS (`xlsx`)**. Uitsluitend het **eerste werkblad** (`workbook.SheetNames[0]`) wordt uitgelezen met datumweergave `yyyy-mm-dd`. Lege rijen worden gefilterd.

### Automatische Rol-detectie (`detectFileRole`)
Bij het uploaden krijgt elk bestand automatisch een rol toegewezen: `'customers'` (Klanten) of `'products'` (Producten):
1. **Bestandsnaam**:
   - Bevat de naam `product`, `abonnement`, `tarief`, `pakket`, `dienst`, `contract`, `omzet` of `factuur` (en géén `klant` of `leden`) $\rightarrow$ `'products'`.
   - Bevat de naam `klant`, `leden`, `member`, `customer` of `relatie` (en géén `product`) $\rightarrow$ `'customers'`.
2. **Inhoudsanalyse (eerste 50 rijen)**:
   - Als $\ge 20\%$ van de steekproefrijen een e-mailadres bevat $\rightarrow$ `'customers'`.
   - Als kolomnamen trefwoorden bevatten zoals `tarief`/`prijs` of `product`/`abonnement`, en er géén e-mailadressen en géén telefoonnummers in de steekproef staan $\rightarrow$ `'products'`.
3. **Disambiguatie bij 2 bestanden**: Als er tegelijk exact 2 bestanden voor een vestiging worden geüpload met dezelfde gedetecteerde rol, en één van de twee bevat productgerelateerde trefwoorden in de naam, worden de rollen automatisch verdeeld over `products` en `customers`.
4. **Handmatige aanpassing**: In de interface kan per geüpload bestand via een dropdown altijd handmatig worden gewisseld tussen `👤 Klanten` en `🏷️ Producten`.

---

## Koppeling & Zoekindex (Klantreferentie, E-mail, Naam)

De koppeling tussen gegevens vindt plaats in de functie `findCustomer`:

### 1. Zoekvolgorde
Wanneer een regel wordt gematcht met een bestaand klantprofiel, zoekt het systeem in deze strikte volgorde:
1. **Klantreferentie (`ref`)**:
   - Zoekt exact op de getrimde waarde (`profilesByRef.get(trimmed)`).
   - Zoekt case-insensitive (`profilesByRef.get(trimmed.toLowerCase())`).
   - **Zonder voorloopnullen**: Indien de referentie volledig uit cijfers bestaat (`/^\d+$/`), wordt ook gezocht op de numerieke string zonder voorloopnullen (`profilesByRef.get(String(Number(trimmed)))`). Bijvoorbeeld: `'00123'` matcht met `'123'`.
2. **E-mailadres (`email`)**:
   - Indien aanwezig en geldig volgens `isLikelyEmail`: zoekt in `profilesByEmail` op `email.toLowerCase().trim()`.
3. **Volledige naam (`fn` en `ln`)**:
   - Indien zowel voornaam als achternaam aanwezig zijn: zoekt in `profilesByName` op `${fn.toLowerCase().trim()}_${ln.toLowerCase().trim()}`.

### 2. Gecombineerde Zoekindex over Beide Vestigingen
- De zoekindex (`profilesByRef`, `profilesByEmail`, `profilesByName`) wordt **gezamenlijk opgebouwd over alle geüploade bestanden**, dus over Ommen én Zwolle samen.
- De index wordt **niet** per vestiging geïsoleerd. Een klantreferentie of e-mailadres uit een bestand van Zwolle kan daardoor matchen met een eerder ingelezen klantprofiel uit een bestand van Ommen.

### 3. Beperking bij Bijwerken van de Zoekindex
- De maps `profilesByRef`, `profilesByEmail` en `profilesByName` worden **uitsluitend gevuld bij het allereerste aanmaken van een nieuw profiel** (in de `else`-tak van `registerOrMergeCustomer`).
- Wanneer een bestaand profiel via een latere regel wordt aangevuld met een voorheen ontbrekend e-mailadres, ontbrekende voor-/achternaam of ontbrekende klantreferentie, worden de zoekindexen **niet met terugwerkende kracht bijgewerkt**.
- **Gevolg:** Latere regels die uitsluitend via die nieuw aangevulde e-mail, naam of referentie zouden kunnen matchen, vinden het profiel niet in de zoekindex en kunnen daardoor een koppeling missen (en eventueel als nieuw los profiel worden aangemaakt).

### 4. Herkende Kolomnamen voor Klantreferentie (`getKlantRef`)
De functie `getKlantRef` verwijdert alle niet-alfanumerieke tekens uit kolomnamen en zoekt naar:
- Exact genormaliseerd: `klantref`, `klantreferentie`, `referentie`, `ref`, `klantnummer`, `klantnr`, `lidnummer`, `lidnr`, `relatienummer`, `relatienr`, `memberid`, `clientid`, `klantid`, `relatieid`.
- Als fallback worden kolomnamen geaccepteerd die `klantref`, `lidnummer` of `klantnummer` bevatten.

---

## Samenvoegen van Klantprofielen (Fase 1: Registratie)

Voordat er gefilterd of geëxporteerd wordt, worden klantbestanden verwerkt via `registerOrMergeCustomer`.

### Volgorde van Verwerking
- De bestandenlijst wordt samengesteld als:
  ```ts
  const allFilesWithSource = [
    ...ommenFiles.map(f => ({ ...f, source: 'Ommen' })),
    ...zwolleFiles.map(f => ({ ...f, source: 'Zwolle' }))
  ];
  ```
- **Ommen-bestanden worden hierdoor altijd vóór Zwolle-bestanden verwerkt.**
- Er is **geen** instelbare vestigingsvoorkeur in de applicatie aanwezig; de volgorde ligt vast in de code.

### Leidend Profiel & Aanvullende Gegevens
- Het **eerst aangetroffen profiel blijft leidend**. De oorspronkelijke vestigingsbron (`source: 'Ommen'`) en de ruwe rij (`raw`) van het eerste record blijven behouden.
- Wanneer een later record matcht via referentie, e-mail of naam, vult dit latere record **uitsluitend ontbrekende velden aan**:
  - `email`: alleen overgenomen als het bestaande profiel nog geen e-mail had.
  - `phone`: alleen overgenomen als het bestaande profiel nog geen telefoon had.
  - `fn`, `ln`, `zip`, `ct`, `dob`, `doby`, `age`, `gen`, `ref`: alleen ingevuld indien leeg in het bestaande profiel.
  - `baseProduct`, `baseValue`, `baseDate`: alleen ingevuld indien het bestaande profiel nog geen basisproduct had.
- **Tariefbehandeling bij dubbele klantregels:** Een later samengevoegde klantregel verhoogt `existing.maxSingleTariff` **niet**. Mocht een tweede klantregel voor dezelfde persoon een hoger tarief bevatten (bijvoorbeeld €500 t.o.v. een eerder basistarief van €50), dan gaat dit hogere tarief in het klantprofiel verloren. Gekoppelde productregels (in Stap 2) verhogen `maxSingleTariff` daarentegen **wél**.
- **Let op:** Deze eerste samenvoegstap (`registerOrMergeCustomer`) blijft **altijd actief**, ongeacht of de deduplicatieschakelaar in de interface aan- of uitstaat.

---

## Deduplicatie (Fase 2: Groepering)

Na het samenstellen van de basisrecords (`allRecords`) volgt een optionele tweede deduplicatiestap (`if (deduplicate)`):

### 1. Groeperingssleutel
De deduplicatiestap groepeert records op basis van de volgende sleutel:
1. `ref_${rec._klantRef.toLowerCase().trim()}` (indien `_klantRef` aanwezig is).
2. Anders: `email_${rec.email.toLowerCase().trim()}` (indien `email` aanwezig is).
3. Anders: `name_${rec.fn.toLowerCase().trim()}_${rec.ln.toLowerCase().trim()}`.

> **Belangrijk:** In deze deduplicatiefase vindt **géén matching op telefoonnummer** plaats. Alleen klantreferentie, e-mailadres en volledige naam worden als groeperingssleutel gebruikt.

### 2. Selectie van het Beste Record binnen een Groep
Als een groep meerdere records bevat, wordt het winnende record gekozen door sortering op:
1. **Compleetheidsscore**:
   $$\text{score} = (\text{heeft email} \times 4) + (\text{heeft phone} \times 3) + (\text{heeft zip} \times 2) + (\text{heeft dob} \times 2)$$
   Het record met de hoogste score komt bovenaan.
2. **Recentste datum (`_actiefSindsDate`)**:
   Bij gelijke score wint het record waarvan de startdatum het meest recent is (`b.date - a.date`).

### 3. Exacte Regels voor Waardebehoud bij Deduplicatie
Binnen de groep wordt de hoogste individuele aankoopwaarde (`maxSingleTariff`) berekend. De toewijzing aan `bestRow.value` gebeurt via deze specifieke condities:
```ts
if (maxSingleTariff >= HIGH_VALUE_THRESHOLD && (bestRow.value === "" || Number(bestRow.value) < HIGH_VALUE_THRESHOLD)) {
  bestRow.value = maxSingleTariff;
} else if ((bestRow.value === "" || bestRow.value === 0 || Number(bestRow.value) === 0) && maxSingleTariff > 0) {
  bestRow.value = maxSingleTariff;
}
bestRow._maxHistoricalValue = maxSingleTariff;
```
Dit betekent in de praktijk:
- Als `maxSingleTariff >= 400`, wordt de waarde van het beste record **alleen overschreven als die waarde leeg is óf strikt lager dan €400** (`< 400`).
- **Belangrijke nuance:** Heeft het geselecteerde beste record zelf al een waarde van $\ge 400$ (bijvoorbeeld **€450**), en had een ander record in dezelfde groep een nog hogere waarde van bijvoorbeeld **€600**, dan evalueert `Number(bestRow.value) < 400` naar `false`. De waarde van **€450 wordt in dat geval niet vervangen door €600** (de waarde blijft €450, al wordt `_maxHistoricalValue` wel op €600 gezet).
- Als de waarde van het beste record leeg of 0 is, en er is een waarde $> 0$ in de groep, dan wordt die overgenomen.

### 4. Uitschakelen van Deduplicatie
Wanneer de schakelaar *"Dedupliceren activeren"* in de interface wordt uitgevinkt:
- Wordt de tweede groeperingsstap (Stap B) overgeslagen.
- Blijft Fase 1 (`registerOrMergeCustomer`) wel actief (klanten die in dezelfde run al gekoppeld zijn, blijven samengevoegd in `profilesList`).

---

## Normalisatieregels & Veldformaten

De velden in elk record worden als volgt verwerkt door `processRawRow`:

| Veld | Normalisatieregel in Code | Wat er WEL gebeurt | Wat er NIET gebeurt |
| :--- | :--- | :--- | :--- |
| **`email`** | `isLikelyEmail(emailRaw) ? emailRaw.trim().toLowerCase() : ""` | - Spaties aan begin/eind getrimd<br>- Omgezet naar kleine letters<br>- Leeggemaakt als het niet voldoet aan `isLikelyEmail` | Geen complexe RFC-regex. De controle is uitsluitend: bevat `@` en `.`, lengte $\ge 5$, geen spaties, begint/eindigt niet met `@`. |
| **`phone`** | `formatPhone(phoneRaw)` | - Spaties, streepjes, haakjes verwijderd: `/[\s\-\(\)]/g`<br>- `0031...` $\rightarrow$ `+31...`<br>- `31...` $\rightarrow$ `+31...`<br>- `06...` $\rightarrow$ `+316...`<br>- Begint met `5` én exact 9 cijfers $\rightarrow$ `0` voorgevoegd (`05...`) | - **Het `+` teken blijft behouden** (`+316...`)<br>- Vaste nummers worden **niet** standaard naar `+31` omgezet<br>- **Geen lengtevalidatie** op 9–15 cijfers; kortere/langere nummers worden niet afgekeurd. |
| **`fn`** | Direct uit kolomnaam | - Waarde getrimd via `getValueByHeader` | **Geen kapitalisatie / Title Case**. Invoer `JAN` blijft `JAN`. |
| **`ln`** | Direct uit kolomnaam | - Als aparte tussenvoegselkolom aanwezig is én ontbreekt in achternaam: samengevoegd als `${tussenvoegsel} ${lnRaw}` | **Geen kapitalisatie / Title Case**. Invoer `DE BOER` blijft `DE BOER`. |
| **`zip`** | Direct uit kolomnaam | - Waarde getrimd | **Geen automatische postcode-formattering**. Spaties worden niet verwijderd (`7731 BC` blijft `7731 BC`). |
| **`ct`** | Direct uit kolomnaam | - Waarde getrimd | **Geen automatische hoofdlettercorrectie**. `ommen` blijft `ommen`. |
| **`country`** | Vaste waarde | - Altijd ingesteld op **`'NL'`** (hoofdletters) | Geen dynamische landdetectie. |
| **`dob`** | `format(dobDate, 'yyyy-MM-dd')` | - Datumnotaties (`DD-MM-YYYY`, `YYYY-MM-DD`, etc.) geparsed naar **`YYYY-MM-DD`** | Wordt **niet** geformatteerd als `YYYYMMDD` zonder streepjes. |
| **`doby`** | `format(dobDate, 'yyyy')` | - 4-cijferig geboortejaar (`YYYY`) | - |
| **`gen`** | Case-insensitive mapping | - `'man'`, `'m'`, `'mannelijk'` $\rightarrow$ **`'M'`**<br>- `'vrouw'`, `'f'`, `'v'`, `'vrouwelijk'` $\rightarrow$ **`'F'`** | **Onbekende waarden blijven behouden** zoals ingevoerd (bijv. `'Dhr'` of `'Onbekend'` blijft staan). |
| **`age`** | `calculateAge(dobDate)` | - Berekend t.o.v. de huidige datum (jaarverschil gecorrigeerd voor verjaardag) | Leeg als er geen geboortedatum is. |
| **`value`** | `parseCurrency(tariffRaw)` | - Valutasymbolen verwijderd, komma's/punten genormaliseerd, afgerond op 2 decimalen (`Math.round(val * 100) / 100`) | Krijgt in de geëxporteerde CSV **niet** verplicht twee decimalen (kan `400` of `299.5` zijn). |
| **`product`** | Direct uit kolomnaam | - Naam van het product/abonnement | **Wordt NIET opgenomen in de CSV-export** (alleen intern en in de UI gebruikt). |

---

## High Value Logica (Drempelwaarde €400)

De selectie voor de High Value lijst (`highValue`) volgt deze specifieke programmalogica:

1. **Vaste Drempelwaarde**:
   - `const HIGH_VALUE_THRESHOLD = 400;`
2. **Minimaal Één Los Tarief (Geen Optelsom)**:
   - Een klant kwalificeert als een individueel productrecord of geregistreerd tarief een numerieke waarde van **$\ge 400$** heeft (`!isNaN(num) && num >= 400`).
   - Losse maandbedragen (zoals 12 keer €45) worden **niet** bij elkaar opgeteld.
3. **Hoe de Hoogste Waarde en Productomschrijving worden Bewaard**:
   - **Belangrijk onderscheid:** `cust.maxSingleTariff` wordt geïnitialiseerd op de waarde van het *eerste* klantrecord (`registerOrMergeCustomer`). Latere klantregels voor dezelfde persoon verhogen deze waarde **niet**.
   - Bij het inlezen van het **productenbestand** wordt `cust.maxSingleTariff` wél actief bijgewerkt wanneer een gekoppeld product een hogere waarde heeft:
     ```ts
     if (!isNaN(valNum) && valNum > cust.maxSingleTariff) {
       cust.maxSingleTariff = valNum;
     }
     if (!isNaN(valNum) && valNum >= HIGH_VALUE_THRESHOLD) {
       cust.highValueProduct = prodName || cust.highValueProduct;
     }
     ```
   - Bij het samenstellen van `allRecords`:
     - Als `cust.maxSingleTariff >= 400`:
       - `displayValue = cust.maxSingleTariff;`
       - `displayProduct = cust.highValueProduct || (latestProd ? latestProd.name : cust.baseProduct);`
     - Als `cust.maxSingleTariff < 400`:
       - Wordt de waarde en productnaam van het **meest recente product** gekozen (`latestProd`), of als dat ontbreekt het basisproduct van de klant (`cust.baseValue` / `cust.baseProduct`).

---

## Productcategorisatie & Trefwoorden

Records worden toegekend aan productcategorieën via de functie `matchesCategory`. 

### Zoekbereik
De matching zoekt in:
1. De veldwaarde `rec.product` (omgezet naar hoofdletters).
2. Indien geen match: **alle gecombineerde tekstwaarden van de oorspronkelijke rij** (`Object.values(rec._raw).filter(Boolean).join(' ').toUpperCase()`).

### Volledige Trefwoordenlijsten per Categorie

#### 1. No More Pain (`no_more_pain`)
- **Trefwoorden:**
  `NO MORE PAIN`, `NO-MORE-PAIN`, `NOMOREPAIN`, `3 REVALIDATIE BEHANDELINGEN`, `REVALIDATIE BEHANDELINGEN`, `REVALIDATIE`, `REVALIDATIETRAJECT`, `REVALIDATIE TRAJECT`, `NMP`, `PIJNVRIJ`, `PIJN VRIJ`, `PIJN`, `PAIN`, `FYSIO`, `FYSIOTHERAPIE`, `HERSTEL`, `REVALIDEREN`, `BEHANDELING`, `BEHANDELINGEN`, `BLESSURE`, `RUGKLACHTEN`, `SCHOUDERKLACHTEN`.
- **Speciale Uitzonderingsregel in Code:**
  Bevat de productnaam of de ruwe rij de term `COACHING` of `PERSOONLIJKE ONTWIKKELING`, én is de tariefwaarde exact `299` (of `299.00`), dan kwalificeert dit record **ook** voor de categorie *No More Pain*.

#### 2. Personal Training (`personal_training`)
- **Trefwoorden:**
  `PERSONAL TRAINING`, `PERSONAL-TRAINING`, `PERSONALTRAINING`, `PERSOONLIJKE ONTWIKKELING`, `PERSONAL COACH`, `PERSONAL COACHING`, `1-OP-1`, `1 OP 1`, `1-ON-1`, `1 ON 1`, `1:1`, `DUO TRAINING`, `DUO-TRAINING`, `DUOTRAINING`, `INDIVIDUEEL`, `TRAJECT`, `COACHING`, `COACH`, `BEGELEIDING`, ` PT`, `PT `, `PT-`, `PT/`, `(PT)`, `[PT]`, `/PT`, `-PT`, `_PT`, ` PT `.

#### 3. Groepslessen (`groepslessen`)
- **Trefwoorden:**
  `GROEPSLESSEN`, `GROEPSLES`, `VIKING MODE`, `VIKINGMODE`, `SOLID STRONG`, `SOLIDSTRONG`, `GOLDEN TIGER`, `GOLDENTIGER`, `SMALL GROUP`, `SMALLGROUP`, `SMALL-GROUP`, `BOOTCAMP`, `CIRCUIT`, `GROEP`, `GROUP`, `FIT & STRONG`, `FIT AND STRONG`, `UNLIMITED`, `ONBEPERKT`, `FITNESS`, `SPORTER`, `ABONNEMENT`, `LIDMAATSCHAP`, `COMMUNITY`, `OPEN GYM`, `VRIJ SPORTEN`, `VRIJ TRAINEN`, `GYM`, `CROSSFIT`, `WORKOUT`, `1X PER WEEK`, `2X PER WEEK`, `3X PER WEEK`, `1 X PER WEEK`, `2 X PER WEEK`, `3 X PER WEEK`, `PER WEEK`, `PER MAAND`, `STRIPPENKAART`, `RITTENKAART`.

#### 4. Get Leaner (`get_leaner`)
- **Trefwoorden:**
  `GET LEANER`, `GET-LEANER`, `GETLEANER`, `LEANER`, `LEAN`, `AFVALLEN`, `VETVERLIES`, `GEWICHTSVERLIES`, `VOEDING`, `VOEDINGSCHEMA`, `VOEDINGSBEGELEIDING`, `NUTRITION`, `DIET`, `DIEET`, `SHRED`, `BODY COMPOSITION`, `TRANSFORMATIE`, `CHALLENGE`, `LIFESTYLE`.

### Overlap tussen Categorieën
De categorieën sluiten elkaar **niet** uit. Een regel met bijvoorbeeld `"Small Group PT"` bevat zowel trefwoorden voor *Personal Training* als *Groepslessen* en zal in **beide** exports verschijnen.

---

## Filters, Uitzonderingen & Niet-Gekoppelde Producten

### Geen Filter op Actieve Klanten of Betaalstatus
- De applicatie bevat **geen filter** op actieve versus inactieve leden, opgezegde abonnementen of betaalstatus. Alle rijen in de bronbestanden worden verwerkt.

### Werking van het Tarieffilter (`tariffFilter`)
- De checkbox *"Verwijder records zonder tarief"* staat **standaard uitgeschakeld** (`tariffFilter = false`).
- Wanneer ingeschakeld, filtert deze uitsluitend rijen met `value === ""` of `Number(value) === 0` uit de totale klantenlijst (`all`) en daarmee de High Value lijst.
- **Productexports worden hierdoor niet beïnvloed:** De exports per productcategorie (`productExports`) worden samengesteld uit `candidateProductRecords` en worden niet gefilterd door de `tariffFilter`-schakelaar.

### Niet-Gekoppelde Producten
- Als in een productenbestand een rij staat waarvan de `"Klant ref."`, het e-mailadres of de naam niet voorkomt in het klantenbestand:
  ```ts
  const standalone = processRawRow(rawRow, f.source);
  candidateProductRecords.push(standalone);
  ```
- Dit record wordt **wél** meegenomen in de beoordeling voor productcategorieën (`productExports`).
- Dit record wordt **niet** toegevoegd aan `profilesList`, en verschijnt daarom **niet** in de totale klantenlijst (`all`) of in de High Value lijst (`highValue`) wanneer er afzonderlijke klantenbestanden zijn ingelezen.

---

## CSV-Export Specificaties

Alle CSV-bestanden worden gegenereerd via `Papa.unparse` met de volgende instellingen:

1. **Scheidingsteken**:
   - `delimiter: ";"` (Europese puntkomma, geoptimaliseerd voor Nederlandse Excel-installaties).
2. **Aanhalingstekens**:
   - Wordt bepaald door de instelling `noQuotes` (standaard ingeschakeld). Bij `noQuotes = true` worden alleen aanhalingstekens geplaatst waar nodig (`quotes: false`).
3. **Exacte Kolomvolgorde**:
   De functie `cleanForExport` levert exact de volgende 12 kolommen op:
   ```csv
   email;phone;fn;ln;zip;ct;country;dob;doby;gen;age;value
   ```
   > **Let op:** Het veld `product` zit **niet** in deze lijst en wordt dus **niet meegeëxporteerd** in de CSV-bestanden!
4. **Bedragnotatie**:
   - Bedragen krijgen in de CSV geen verplichte twee decimalen (bijvoorbeeld `400` of `299.5`). Alleen in de tabellen in de gebruikersinterface wordt `.toFixed(2)` getoond.

---

## Interface & Downloadmogelijkheden

De interface biedt specifieke downloadknoppen:

### 1. Klantenlijsten & High Value
De standaard bestandsnamen bevatten de datum van vandaag (`DD-MM-YYYY`):
- **Volledige Lijst**: `All Customers - DD-MM-YYYY.csv`
- **High Value Totaal**: `High Value Customers - DD-MM-YYYY.csv` (alle records $\ge$ €400)
- **High Value Ommen**: `High Value - Ommen - DD-MM-YYYY.csv`
- **High Value Zwolle**: `High Value - Zwolle - DD-MM-YYYY.csv`

*(De bestandsnamen van deze vier exports kunnen via invoervelden in de interface worden aangepast).*

### 2. Product- & Locatie-exports
Per productcategorie (*No More Pain*, *Personal Training*, *Groepslessen*, *Get Leaner*) zijn er drie downloadknoppen:
- `[Categorienaam] - Ommen - DD-MM-YYYY.csv`
- `[Categorienaam] - Zwolle - DD-MM-YYYY.csv`
- `[Categorienaam] - Totaal - DD-MM-YYYY.csv`

Daarnaast is er een knop **"Download Alle Product Exports"**:
- Deze knop triggert achtereenvolgens (met een tussentijd van 300 ms per bestand) de afzonderlijke downloads van alle niet-lege productbestanden. Er wordt **geen ZIP-bestand** gegenereerd.

### 3. Wat de Interface NIET Biedt
- Er is **geen exportknop** voor een `duplicaten_overzicht.csv`. Het deduplicatielog is uitsluitend in te zien via het tabblad *"Gededupliceerd Log"* in de browser.
- Er is **geen algemene knop** om alle klanten- én productbestanden tegelijk in één archief/ZIP te downloaden.

---

## Gegevensverwerking & Privacy (AVG / GDPR Context)

- **Lokale Client-Side Verwerking**:
  Alle bestandsverwerking (parsen, matchen, transformeren en genereren van downloads) wordt lokaal in de webbrowser van de gebruiker uitgevoerd via JavaScript. Er worden geen persoonsgegevens naar een server van deze applicatie verzonden of daarin opgeslagen.
- **Geen Automatische AVG-Garantie**:
  Hoewel lokale verwerking het risico op datalekken via een tussenliggende applicatieserver uitsluit, vormt dit **geen garantie dat automatisch aan alle verplichtingen van de Algemene Verordening Gegevensbescherming (AVG / GDPR) wordt voldaan**. Organisaties die persoonsgegevens verwerken blijven zelf verantwoordelijk voor:
  - Een geldige wettelijke grondslag voor verwerking (zoals toestemming voor marketing of gerechtvaardigd belang).
  - Voldoen aan doorgiftevereisten bij het uploaden van geëxporteerde lijsten naar advertentieplatformen van derden (zoals Meta Platforms Ireland Ltd.).
  - Het respecteren van bewaartermijnen en rechten van betrokkenen (zoals het recht op inzage of verwijdering).

---

## Lokale Installatie & Ontwikkeling

### Vereisten
- Node.js (versie 18 of hoger)
- npm (of pnpm / bun)

### Installatie & Uitvoering
```bash
# 1. Clone de repository
git clone https://github.com/jouw-organisatie/csv-converter.git
cd csv-converter

# 2. Installeer afhankelijkheden
npm install

# 3. Start ontwikkelserver
npm run dev

# 4. Type-checken (linter)
npm run lint

# 5. Bouwen voor productie
npm run build
```
