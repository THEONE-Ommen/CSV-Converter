# CSV & Excel Converter voor Meta Ads & CRM (Ommen & Zwolle)

Een snelle, veilige en 100% lokale webapplicatie voor het verwerken, koppelen, normaliseren, dedupliceren en exporteren van leden- en productdata uit fitness- en clubmanagementsoftware. De tool is specifiek afgestemd op de vestigingen **Ommen** en **Zwolle** en levert kant-en-klare CSV-bestanden op volgens de strikte standaarden van **Meta Ads (Facebook Custom Audiences)** en moderne CRM-systemen.

---

## Inhoudsopgave

1. [Belangrijkste Functies](#belangrijkste-functies)
2. [Ondersteunde Bestandsformaten & Delimiters](#ondersteunde-bestandsformaten--delimiters)
3. [Workflow: Hoe Werkt Het?](#workflow-hoe-werkt-het)
4. [Koppeling van Bestanden op "Klant ref."](#koppeling-van-bestanden-op-klant-ref)
5. [Alle Ingestelde Formatteringen (Meta Ads Standaarden)](#alle-ingestelde-formatteringen-meta-ads-standaarden)
6. [High Value Klanten (Drempelwaarde ≥ €400)](#high-value-klanten-drempelwaarde--400)
7. [Productcategorisatie](#productcategorisatie)
8. [Deduplicatie & Prioriteringsregels](#deduplicatie--prioriteringsregels)
9. [Overzicht van Beschikbare Exports](#overzicht-van-beschikbare-exports)
10. [Privacy & Gegevensbescherming (AVG / GDPR)](#privacy--gegevensbescherming-avg--gdpr)
11. [Installatie & Lokale Ontwikkeling](#installatie--lokale-ontwikkeling)

---

## Belangrijkste Functies

- **Multi-formaat upload**: Ondersteuning voor zowel `.csv` als Microsoft Excel (`.xlsx` en `.xls`).
- **Vestiging-specifieke verwerking**: Gescheiden upload-dropzones voor **Ommen** en **Zwolle**.
- **Dual-file koppeling**: Upload per vestiging een **Klantenbestand** (met alle contact- en persoonsgegevens) en een **Productenbestand** (met tarieven en afgenomen producten).
- **Slimme rol-detectie**: Herkent automatisch of een geüpload bestand een klantenlijst of een productenlijst is, inclusief handmatige dropdown-wisselaar.
- **Automatische koppeling op relatienummer**: Brengt klantgegevens en productinformatie naadloos samen via de kolom `"Klant ref."`.
- **Meta Ads Custom Audience normalisatie**: Automatische validatie en formattering van e-mailadressen, E.164-telefoonnummers, postcodes, geboortedata, leeftijden en geslachten.
- **Geavanceerde deduplicatie**: Voorkomt dubbele contacten op basis van e-mail, telefoonnummer of volledige naam, met instelbare vestigingsvoorkeur (Ommen vs. Zwolle eerst) en compleetheidsscoring.
- **High Value segmentatie**: Filtert klanten met een losse aankoop of tarief van **minimaal €400** (puur los tarief, geen cumulatief maandtotaal).
- **Automatische productsegmentatie**: Directe exports voor *No More Pain*, *Personal Training*, *Groepslessen* en *Get Leaner*.
- **Interactieve preview**: Inzicht in geüploade bestanden, koppelingsstatus, statistieken en datatabellen vóór het downloaden.
- **100% Client-side privacy**: Alle berekeningen vinden plaats in de webbrowser van de gebruiker. Er worden geen persoonsgegevens naar externe servers verzonden.

---

## Ondersteunde Bestandsformaten & Delimiters

| Formaat | Extensie | Parser | Bijzonderheden |
| :--- | :--- | :--- | :--- |
| **CSV** | `.csv` | PapaParse | Automatische scheidingsteken-detectie (puntkomma `;`, komma `,` of tab `\t`). Ondersteunt quotes en UTF-8 encoding. |
| **Excel Modern** | `.xlsx` | SheetJS (`xlsx`) | Leest automatisch het eerste werkblad uit met behoud van datums en numerieke waarden. |
| **Excel Klassiek** | `.xls` | SheetJS (`xlsx`) | Volledige ondersteuning voor oudere Excel 97-2004 werkmappen. |

---

## Workflow: Hoe Werkt Het?

```
┌───────────────────────────────────┐     ┌───────────────────────────────────┐
│     Vestiging Ommen               │     │     Vestiging Zwolle              │
│  - Klantenbestand (CSV / Excel)   │     │  - Klantenbestand (CSV / Excel)   │
│  - Productenbestand (CSV / Excel) │     │  - Productenbestand (CSV / Excel) │
└─────────────────┬─────────────────┘     └─────────────────┬─────────────────┘
                  │                                         │
                  ▼                                         ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. INLEZEN & ROL-DETECTIE                                                   │
│    - Bepaal per bestand: 'Klanten' (NAW/contact) of 'Producten' (tarieven)  │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 2. KOPPELING OP "Klant ref."                                                │
│    - Match productrecords aan klantrecords via relatienummer                │
│    - Verrijk productrecords met NAW, email, telefoon, geboortedatum         │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 3. NORMALISATIE NAAR META ADS STANDAARDEN                                   │
│    - E-mail lowercase & regex validatie                                     │
│    - Telefoonnummer omzetten naar '316...' (E.164 zonder plus)              │
│    - Postcode naar '1234AB', geboortedatum naar 'YYYYMMDD', leeftijd 'age'  │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 4. DEDUPLICATIE & BRONPRIORITERING                                          │
│    - Unieke match op email, telefoon of naam                                │
│    - Rangschikking op basis van voorkeursvestiging en compleetheidsscore    │
│    - Behoud van hoogste individuele tarief (≥ €400)                         │
└─────────────────────────────────────┬───────────────────────────────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ 5. EXPORTS                                                                  │
│    - Alle Klanten (Totaal)       - High Value (≥ €400)                      │
│    - No More Pain                - Personal Training                        │
│    - Groepslessen                - Get Leaner                               │
│    - Vestiging Ommen             - Vestiging Zwolle                         │
│    - Duplicaten Overzicht        - Download Alles (Bundle)                  │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Koppeling van Bestanden op "Klant ref."

In fitness- en ledenadministratiesoftware worden persoonsgegevens en aankoopgegevens vaak gescheiden geëxporteerd:
1. **Klantenbestand (`customers`)**: Bevat naam, e-mailadres, telefoonnummer, adres, woonplaats, postcode, geslacht, geboortedatum en het klantnummer.
2. **Productenbestand (`products`)**: Bevat de afgenomen lidmaatschappen, strippenkaarten, tarieven, startdata en het bijbehorende klantnummer.

### Herkende Kolomnamen voor het Koppelveld:
De applicatie herkent automatisch de volgende kolomkoppen (ongeacht hoofdletters of leestekens):
- `Klant ref.` / `Klant ref`
- `Klantnummer` / `Klantnr` / `Klant nr`
- `Relatienummer` / `Relatienr`
- `Klant ID` / `CustomerID` / `Customer ID`
- `Debiteurnummer` / `Lidnummer`

### Hoe de Koppeling Verloopt:
- Wanneer zowel een klantenbestand als een productenbestand voor een vestiging aanwezig zijn, zoekt de matcher voor elk record in het productenbestand de bijbehorende klant op via het klantnummer.
- Alle ontbrekende contactvelden (`email`, `phone`, `fn`, `ln`, `zip`, `ct`, `dob`, `doby`, `gen`, `age`) in het productrecord worden overgenomen uit het stam-klantenbestand.
- Het productrecord levert de actuele `product`-naam en het individuele `value`-tarief.
- Als een klant in het klantenbestand staat maar nog geen product heeft, blijft het contact behouden en worden de beschikbare persoonsgegevens meegenomen.
- Als een bestand per ongeluk als verkeerde rol is aangemerkt, kan de gebruiker via het dropdown-menu bij het bestand direct wisselen tussen **"Klantenbestand"** en **"Productenbestand"**.

---

## Alle Ingestelde Formatteringen (Meta Ads Standaarden)

Elk geëxporteerd CSV-bestand voldoet exact aan de kolomnamen en dataspecificaties van Meta Custom Audiences:

| Kolom | Beschrijving | Formattering & Validatieregels | Voorbeeld Invoer | Voorbeeld Uitvoer |
| :--- | :--- | :--- | :--- | :--- |
| **`email`** | E-mailadres | - Volledig in kleine letters (lowercase)<br>- Witruimtes en tabs getrimd<br>- Strikte regex-validatie (`/^[^\s@]+@[^\s@]+\.[^\s@]+$/`)<br>- **Leeg indien ongeldig** (voorkomt dat klantnummers of ongeldige tekst in de e-mailkolom belanden) | ` Jan.Jansen@Gmail.com ` | `jan.jansen@gmail.com` |
| **`phone`** | Telefoonnummer | - Alle leestekens, spaties, haakjes en streepjes verwijderd<br>- Voorloopnullen bij Nederlandse mobiele nummers (`06...` of `0031...`) omgezet naar landcode **`31`**<br>- Internationale `+` wordt verwijderd (`+316...` $\rightarrow$ `316...`)<br>- Geldigheidscontrole: 9 t/m 15 numerieke cijfers<br>- Leeg gelaten indien onvolledig of ongeldig | `+31 (0)6 - 12 34 56 78` | `31612345678` |
| **`fn`** | Voornaam | - Eerste letter hoofdletter, rest kleine letters (Title Case)<br>- Titels en voorvoegsels opgeschoond | `JAN` | `Jan` |
| **`ln`** | Achternaam | - Juiste kapitalisatie met behoud van Nederlandse tussenvoegsels (`de`, `van`, `van der`, `ter`, `in 't`, etc.) | `VAN DER VEEN` | `van der Veen` |
| **`zip`** | Postcode | - Nederlandse postcodes omgezet naar 4 cijfers + 2 hoofdletters **zonder spatie**<br>- Buitenlandse postcodes ontdaan van ongeldige leestekens | `7731 bc` of `7731 BC` | `7731BC` |
| **`ct`** | Woonplaats | - Volledige plaatsnaam met correcte hoofdletters | `ommen` | `Ommen` |
| **`country`** | Landcode | - Vaste 2-letterige ISO landcode in kleine letters | `Nederland` / leeg | `nl` |
| **`dob`** | Geboortedatum | - ISO numerieke notatie: **`YYYYMMDD`**<br>- Herkent Nederlandse datumnotaties (`DD-MM-YYYY`, `DD/MM/YYYY`) en Excel datum-serienummers | `24-03-1992` | `19920324` |
| **`doby`** | Geboortejaar | - 4-cijferig geboortejaar (**`YYYY`**) | `24-03-1992` | `1992` |
| **`gen`** | Geslacht | - `m` voor man / dhr / heer<br>- `f` voor vrouw / mevr / dame<br>- Leeg indien onbekend of niet ingevuld | `Vrouw` of `Mevr.` | `f` |
| **`age`** | Leeftijd | - Dynamisch berekend als geheel getal op basis van geboortedatum t.o.v. de huidige datum | Geboren in 1992 | `34` *(of actuele leeftijd)* |
| **`value`** | Tarief / Waarde | - Numeriek bedrag met 2 decimalen (geen valutasymbolen of punten als scheidingsteken)<br>- Behoudt het **hoogste individuele losse tarief** van de klant | `€ 450,00` | `450.00` |
| **`product`** | Product / Lidmaatschap | - Schone tekstuele benaming van het meest recente of hoogst gewaardeerde product | `"10 Rittenkaart PT"` | `10 Rittenkaart PT` |

---

## High Value Klanten (Drempelwaarde ≥ €400)

De selectie voor de **High Value** export is gebaseerd op het principe van een **los aankoopbedrag**:

1. **Puur Los Tarief (Niet Cumulatief)**:
   - Klanten worden **niet** geselecteerd op basis van een optelsom van vele kleine maandelijkse contributies (zoals 12 × €35 = €420).
   - Een klant kwalificeert uitsluitend als er minimaal één individuele transactie, lidmaatschap of productafname is met een tarief van **minimaal €400** (bijvoorbeeld een Personal Training pakket van €450 of een No More Pain traject).
2. **Drempelwaarde**:
   - De drempelwaarde is vast ingesteld op **€400.00**.
3. **Behoud van Waarde bij Deduplicatie**:
   - Als een klant meerdere records heeft (bijvoorbeeld een regulier maandabonnement van €45 én een PT-pakket van €600), behoudt het gededupliceerde klantrecord automatisch het hoogste losse tarief (€600.00).

---

## Productcategorisatie

De converter analyseert de productomschrijving van elk record en segmenteert automatisch in de volgende marketingcategorieën:

| Categorie | Trefwoorden & Detectieregels | Doelgroep / Toepassing |
| :--- | :--- | :--- |
| **No More Pain** | `no more pain`, `nmp`, `pijn`, `revalidatie`, `fysio`, `rugklachten` | Klanten die specifiek een pijnvrij- of revalidatietraject volgen. |
| **Personal Training** | `personal training`, `pt `, ` pt`, `1-op-1`, `1 op 1`, `coaching`, `traject` | Klanten met individuele coaching en hoogwaardige begeleiding. |
| **Groepslessen** | `groepsles`, `groepslessen`, `small group`, `sgt`, `bootcamp`, `spinning`, `yoga`, `pilates`, `fitness` | Deelnemers aan groepsactiviteiten en algemene fitnessabonnementen. |
| **Get Leaner** | `get leaner`, `leaner`, `afvallen`, `vetverlies`, `lifestyle`, `voeding` | Deelnemers aan afval- en leefstijlprogramma's. |

---

## Deduplicatie & Prioriteringsregels

Wanneer bestanden uit meerdere vestigingen (Ommen en Zwolle) worden samengevoegd, kunnen klanten dubbel voorkomen. De applicatie past een intelligent meerstaps deduplicatie-algoritme toe:

### 1. Detectie van Duplicaten
Twee records worden als dezelfde persoon beschouwd als er een overeenkomst is op:
1. **E-mailadres** (ongevoelig voor hoofdletters)
2. **Telefoonnummer** (genormaliseerd naar E.164 standaard)
3. **Volledige naam** (`fn` + `ln` identiek, bij ontbreken van telefoon en email)

### 2. Bepaling van het Primaire Record ("Beste Record")
Wanneer een duplicaat wordt gevonden, wordt het winnende stamrecord gekozen op basis van:
1. **Vestigingsprioriteit**: Instelbaar in de interface:
   - *Optie A*: **Ommen heeft prioriteit** $\rightarrow$ Ommen-record behoudt voorrang.
   - *Optie B*: **Zwolle heeft prioriteit** $\rightarrow$ Zwolle-record behoudt voorrang.
2. **Compleetheidsscore**: Indien binnen dezelfde vestiging of bij gelijke voorkeur, wint het record met de meeste ingevulde velden (aanwezigheid van telefoon, e-mail, adres, geboortedatum).
3. **Recentheid**: Het meest recent actieve lidmaatschap (`Actief sinds`) geniet voorrang.
4. **Tariefbehoud**: Het winnende record erft altijd het **hoogste individuele tarief** (`value`) van alle samengevoegde records.

---

## Overzicht van Beschikbare Exports

Alle exports worden gegenereerd met UTF-8 encoding en standaard komma-gescheiden waarden, direct klaar voor upload in Meta Ads Manager of import in Excel/CRM:

1. **`klanten_totaal_gededupliceerd.csv`**: Complete gecombineerde database van alle unieke klanten uit beide vestigingen, verrijkt met productdata.
2. **`klanten_high_value_400plus.csv`**: Exclusieve lijst van unieke klanten met een individueel tarief $\ge$ €400 (met volledige contactgegevens).
3. **`product_no_more_pain.csv`**: Alle klanten die het No More Pain programma volgen of hebben gevolgd.
4. **`product_personal_training.csv`**: Alle klanten met Personal Training afnames.
5. **`product_groepslessen.csv`**: Alle klanten met groepsles- of fitnessabonnementen.
6. **`product_get_leaner.csv`**: Alle klanten van het Get Leaner leefstijltraject.
7. **`klanten_ommen.csv`**: Unieke klanten behorend bij vestiging Ommen.
8. **`klanten_zwolle.csv`**: Unieke klanten behorend bij vestiging Zwolle.
9. **`duplicaten_overzicht.csv`**: Gedetailleerd audit-rapport van alle gedetecteerde duplicaten met reden van samenvoeging.
10. **Knop "Download Alle CSV's"**: Downloadt in één klik alle afzonderlijke doelgroepbestanden.

---

## Privacy & Gegevensbescherming (AVG / GDPR)

- **Geen serveropslag**: De gehele verwerking (parsing, koppeling, opschoning, deduplicatie) draait in het geheugen van de client-browser.
- **Geen externe tracking**: Er worden geen persoonsgegevens, e-mailadressen of telefoonnummers gelogd of verstuurd naar externe analytics- of AI-diensten.
- **Veilig voor gevoelige data**: Voldoet aan de vereisten voor veilige verwerking van persoonsgegevens onder de Algemene Verordening Gegevensbescherming (AVG).

---

## Installatie & Lokale Ontwikkeling

### Vereisten
- [Node.js](https://nodejs.org/) versie 18 of hoger
- `npm` (of `bun` / `pnpm`)

### Installatie

1. Clone de repository:
   ```bash
   git clone https://github.com/jouw-gebruikersnaam/csv-converter.git
   cd csv-converter
   ```

2. Installeer de benodigde afhankelijkheden:
   ```bash
   npm install
   ```

3. Start de lokale ontwikkelserver:
   ```bash
   npm run dev
   ```
   Open vervolgens [http://localhost:3000](http://localhost:3000) (of de poort getoond in de terminal) in uw browser.

4. Productie-build genereren:
   ```bash
   npm run build
   ```

---

## Licentie

Dit project is ontwikkeld voor intern gebruik en geoptimaliseerd voor marketing- en leadgeneratiedoeleinden van de vestigingen Ommen en Zwolle.
