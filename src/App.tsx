/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import { 
  Upload, 
  Download, 
  FileText, 
  Trash2, 
  CheckCircle2, 
  AlertCircle, 
  MapPin, 
  SlidersHorizontal,
  FileSpreadsheet,
  Info,
  DollarSign,
  Users,
  Settings,
  Sparkles,
  Package,
  FolderDown,
  Tag,
  Link2
} from 'lucide-react';

interface RawRow {
  [key: string]: any;
}

interface ProcessedRow {
  email: string;
  phone: string;
  fn: string;
  ln: string;
  zip: string;
  ct: string;
  country: string;
  dob: string;
  doby: string;
  gen: string;
  age: number | string;
  value: number | string;
  product: string;
  _actiefSindsDate: Date | null;
  _source: 'Ommen' | 'Zwolle';
  _maxHistoricalValue?: number;
  _totalHistoricalValue?: number;
  _klantRef?: string;
  _raw: RawRow;
}

interface DuplicateLog {
  key: string;
  records: ProcessedRow[];
  selected: ProcessedRow;
}

// Product Categories definition
const HIGH_VALUE_THRESHOLD = 400;

interface ProductCategoryDef {
  id: string;
  name: string;
  description: string;
  keywords: string[];
}

const PRODUCT_CATEGORIES: ProductCategoryDef[] = [
  {
    id: 'no_more_pain',
    name: 'No More Pain',
    description: 'Producten met "No More Pain", "3 Revalidatie behandelingen", "Pijnvrij", Fysio, Rug/Schouder/Nek herstel of Coaching (€299)',
    keywords: [
      'NO MORE PAIN', 'NO-MORE-PAIN', 'NOMOREPAIN', '3 REVALIDATIE BEHANDELINGEN', 
      'REVALIDATIE BEHANDELINGEN', 'REVALIDATIE', 'REVALIDATIETRAJECT', 'REVALIDATIE TRAJECT',
      'NMP', 'PIJNVRIJ', 'PIJN VRIJ', 'PIJN', 'PAIN', 'FYSIO', 'FYSIOTHERAPIE', 'HERSTEL',
      'REVALIDEREN', 'BEHANDELING', 'BEHANDELINGEN', 'BLESSURE', 'RUGKLACHTEN', 'SCHOUDERKLACHTEN'
    ]
  },
  {
    id: 'personal_training',
    name: 'Personal Training',
    description: 'Personal Training, 1-op-1 coaching, Duo Training & Persoonlijke ontwikkeling',
    keywords: [
      'PERSONAL TRAINING', 'PERSONAL-TRAINING', 'PERSONALTRAINING', 'PERSOONLIJKE ONTWIKKELING', 
      'PERSONAL COACH', 'PERSONAL COACHING', '1-OP-1', '1 OP 1', '1-ON-1', '1 ON 1', '1:1',
      'DUO TRAINING', 'DUO-TRAINING', 'DUOTRAINING', 'INDIVIDUEEL', 'TRAJECT',
      'COACHING', 'COACH', 'BEGELEIDING',
      ' PT', 'PT ', 'PT-', 'PT/', '(PT)', '[PT]', '/PT', '-PT', '_PT', ' PT '
    ]
  },
  {
    id: 'groepslessen',
    name: 'Groepslessen',
    description: 'Groepslessen, Viking Mode, Solid Strong, Golden Tiger, Small Group, Fitness & Vrij Trainen',
    keywords: [
      'GROEPSLESSEN', 'GROEPSLES', 'VIKING MODE', 'VIKINGMODE', 'SOLID STRONG', 
      'SOLIDSTRONG', 'GOLDEN TIGER', 'GOLDENTIGER', 'SMALL GROUP', 'SMALLGROUP', 'SMALL-GROUP',
      'BOOTCAMP', 'CIRCUIT', 'GROEP', 'GROUP', 'FIT & STRONG', 'FIT AND STRONG',
      'UNLIMITED', 'ONBEPERKT', 'FITNESS', 'SPORTER', 'ABONNEMENT', 'LIDMAATSCHAP', 
      'COMMUNITY', 'OPEN GYM', 'VRIJ SPORTEN', 'VRIJ TRAINEN', 'GYM', 'CROSSFIT', 'WORKOUT',
      '1X PER WEEK', '2X PER WEEK', '3X PER WEEK', '1 X PER WEEK', '2 X PER WEEK', '3 X PER WEEK',
      'PER WEEK', 'PER MAAND', 'STRIPPENKAART', 'RITTENKAART'
    ]
  },
  {
    id: 'get_leaner',
    name: 'Get Leaner',
    description: 'Producten met "Get Leaner", vetverlies, voedingsbegeleiding & lifestyle',
    keywords: [
      'GET LEANER', 'GET-LEANER', 'GETLEANER', 'LEANER', 'LEAN', 'AFVALLEN', 
      'VETVERLIES', 'GEWICHTSVERLIES', 'VOEDING', 'VOEDINGSCHEMA', 'VOEDINGSBEGELEIDING',
      'NUTRITION', 'DIET', 'DIEET', 'SHRED', 'BODY COMPOSITION', 'TRANSFORMATIE', 'CHALLENGE', 'LIFESTYLE'
    ]
  }
];

function matchesCategory(rec: ProcessedRow, cat: ProductCategoryDef): boolean {
  // 1. Direct check on product field
  const productStr = (rec.product || '').toUpperCase().trim();
  if (productStr) {
    if (cat.keywords.some(kw => productStr.includes(kw))) {
      return true;
    }
  }

  // 2. Check all string/text values in the raw imported row as fallback
  if (rec._raw) {
    const rawValuesStr = Object.values(rec._raw).filter(Boolean).join(' ').toUpperCase();
    if (cat.keywords.some(kw => rawValuesStr.includes(kw))) {
      return true;
    }

    // Special rule: No More Pain also includes coaching or personal development with tariff 299
    if (cat.id === 'no_more_pain') {
      const isCoaching = productStr.includes('COACHING') || productStr.includes('PERSOONLIJKE ONTWIKKELING') ||
                         rawValuesStr.includes('COACHING') || rawValuesStr.includes('PERSOONLIJKE ONTWIKKELING');
      const valNum = typeof rec.value === 'number' ? rec.value : parseFloat(String(rec.value));
      if (isCoaching && (valNum === 299 || valNum === 299.00)) {
        return true;
      }
    }
  }

  return false;
}

interface UploadedFileItem {
  id: string;
  file: File;
  rows: RawRow[];
  role: 'customers' | 'products';
}

// Extraction helper for Klant ref. across any variation of casing or punctuation
function getKlantRef(row: RawRow): string {
  if (!row) return "";
  for (const [key, value] of Object.entries(row)) {
    const kNorm = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (
      kNorm === 'klantref' || 
      kNorm === 'klantreferentie' || 
      kNorm === 'referentie' || 
      kNorm === 'ref' || 
      kNorm === 'klantnummer' || 
      kNorm === 'klantnr' || 
      kNorm === 'lidnummer' || 
      kNorm === 'lidnr' || 
      kNorm === 'relatienummer' || 
      kNorm === 'relatienr' || 
      kNorm === 'memberid' || 
      kNorm === 'clientid' ||
      kNorm === 'klantid' ||
      kNorm === 'relatieid'
    ) {
      if (value !== undefined && value !== null && String(value).trim() !== "") {
        return String(value).trim();
      }
    }
  }
  // Secondary check: keys that contain "klantref" or "lidnummer" or "klantnummer"
  for (const [key, value] of Object.entries(row)) {
    const kNorm = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if ((kNorm.includes('klantref') || kNorm.includes('lidnummer') || kNorm.includes('klantnummer')) && value !== undefined && value !== null && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
}

// Check if a string looks like a legitimate email address
function isLikelyEmail(val: any): boolean {
  if (!val || typeof val !== 'string') return false;
  const s = val.trim();
  return s.includes('@') && s.includes('.') && s.length >= 5 && !s.includes(' ') && !s.startsWith('@') && !s.endsWith('@');
}

// Auto-detect whether an uploaded file is primarily Customers (NAW, email, phone) or Products (prices, packages, Klant ref.)
function detectFileRole(rows: RawRow[], fileName: string): 'customers' | 'products' {
  const lowerName = fileName.toLowerCase();

  const isProductByName = (
    lowerName.includes('product') || 
    lowerName.includes('abonnement') || 
    lowerName.includes('tarief') || 
    lowerName.includes('pakket') || 
    lowerName.includes('dienst') || 
    lowerName.includes('contract') ||
    lowerName.includes('omzet') ||
    lowerName.includes('factuur')
  ) && !lowerName.includes('klant') && !lowerName.includes('leden');

  const isCustomerByName = (
    lowerName.includes('klant') || 
    lowerName.includes('leden') || 
    lowerName.includes('member') || 
    lowerName.includes('customer') || 
    lowerName.includes('relatie')
  ) && !lowerName.includes('product');

  if (isProductByName) return 'products';
  if (isCustomerByName) return 'customers';

  if (rows && rows.length > 0) {
    const sample = rows.slice(0, Math.min(rows.length, 50));
    let emailCount = 0;
    let phoneCount = 0;
    let productHeaderCount = 0;
    let tariffHeaderCount = 0;

    const firstRowKeys = Object.keys(rows[0]).map(k => k.toLowerCase().replace(/[^a-z0-9]/g, ''));
    if (firstRowKeys.some(k => k.includes('product') || k.includes('abonnement') || k.includes('dienst'))) {
      productHeaderCount += 2;
    }
    if (firstRowKeys.some(k => k.includes('tarief') || k.includes('prijs') || k.includes('bedrag') || k.includes('tariff'))) {
      tariffHeaderCount += 2;
    }

    sample.forEach(r => {
      for (const v of Object.values(r)) {
        if (isLikelyEmail(v)) {
          emailCount++;
          break;
        }
      }
      for (const [k, v] of Object.entries(r)) {
        const kNorm = k.toLowerCase().replace(/[^a-z0-9]/g, '');
        if ((kNorm.includes('telefoon') || kNorm.includes('mobiel') || kNorm.includes('phone')) && v && String(v).replace(/\D/g, '').length >= 7) {
          phoneCount++;
          break;
        }
      }
    });

    if (emailCount >= Math.max(1, Math.floor(sample.length * 0.2))) {
      return 'customers';
    }

    if ((productHeaderCount > 0 || tariffHeaderCount > 0) && emailCount === 0 && phoneCount === 0) {
      return 'products';
    }
  }

  return 'customers';
}

export default function App() {
  // Uploaded files states (multiple files per location)
  const [ommenFiles, setOmmenFiles] = useState<UploadedFileItem[]>([]);
  const [zwolleFiles, setZwolleFiles] = useState<UploadedFileItem[]>([]);

  // Derived rows from all uploaded files
  const ommenRows = useMemo(() => ommenFiles.flatMap(item => item.rows), [ommenFiles]);
  const zwolleRows = useMemo(() => zwolleFiles.flatMap(item => item.rows), [zwolleFiles]);

  // Settings
  const [deduplicate, setDeduplicate] = useState<boolean>(true);
  const [tariffFilter, setTariffFilter] = useState<boolean>(false); // Disabled by default for full history uploads!
  const [noQuotes, setNoQuotes] = useState<boolean>(true); // Request 6: prevent quotes around values
  const [activeTab, setActiveTab] = useState<'all' | 'high_value' | 'duplicates' | string>('all');

  // Drag and drop states
  const [dragOverOmmen, setDragOverOmmen] = useState<boolean>(false);
  const [dragOverZwolle, setDragOverZwolle] = useState<boolean>(false);

  // Status message for parsing
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'err' | 'info'; text: string } | null>(null);

  // Get formatted date string for filenames
  const getTodayString = () => {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  };

  const todayStr = getTodayString();
  const [allCustomersFilename, setAllCustomersFilename] = useState<string>(`All Customers - ${todayStr}`);
  const [highValueCustomersFilename, setHighValueCustomersFilename] = useState<string>(`High Value Customers - ${todayStr}`);
  const [highValueOmmenFilename, setHighValueOmmenFilename] = useState<string>(`High Value - Ommen - ${todayStr}`);
  const [highValueZwolleFilename, setHighValueZwolleFilename] = useState<string>(`High Value - Zwolle - ${todayStr}`);

  // Refactored Phone Formatting Logic (Request 1-3, 9 & 29)
  const formatPhone = useCallback((phone: string) => {
    if (!phone) return "";
    
    // Remove all spaces, dashes, parentheses and non-numeric characters first (but keep eventual '+' at start temporarily)
    let cleaned = phone.trim().replace(/[\s\-\(\)]/g, '');

    // 1. If starts with 0031 -> change to +31
    if (cleaned.startsWith('0031')) {
      cleaned = '+31' + cleaned.substring(4);
    }
    // 2. If starts with 31 (but not +31) -> change to +31
    else if (cleaned.startsWith('31') && !cleaned.startsWith('+31')) {
      cleaned = '+31' + cleaned.substring(2);
    }
    // 3. If starts with 06 -> transform to +316
    else if (cleaned.startsWith('06')) {
      cleaned = '+316' + cleaned.substring(2);
    }
    // 4. Request 3 & 9: If a local landline number starts with 5 and is 9 digits, prepend 0
    else if (cleaned.startsWith('5') && cleaned.length === 9) {
      cleaned = '0' + cleaned;
    }
    
    // Remove any remaining spaces or non-digit/non-plus signs
    cleaned = cleaned.replace(/[^\d+]/g, '');

    return cleaned;
  }, []);

  // Safe Date parsing helper for DD-MM-YYYY, YYYY-MM-DD, ISO, etc.
  const parseDate = useCallback((dateStr: string): Date | null => {
    if (!dateStr) return null;
    const s = dateStr.trim();
    
    // Match YYYY-MM-DD or YYYY/MM/DD (with optional time)
    let match = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) {
      const d = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
      if (!isNaN(d.getTime())) return d;
    }
    
    // Match DD-MM-YYYY or DD/MM/YYYY (with optional time)
    match = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (match) {
      const d = new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
      if (!isNaN(d.getTime())) return d;
    }

    // Attempt native JS Date parse
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;

    return null;
  }, []);

  // Robust Dutch / European & International currency parser
  const parseCurrency = useCallback((raw: any): number | "" => {
    if (raw === null || raw === undefined) return "";
    if (typeof raw === 'number') {
      return isNaN(raw) ? "" : Math.round(raw * 100) / 100;
    }
    let s = String(raw).trim();
    if (!s) return "";

    // Replace Dutch zero cent notations e.g. ",-" or ".-" -> ",00"
    s = s.replace(/[,-]-$/, ',00').replace(/\.-$/, '.00');

    // Strip currency symbols and letters, keeping digits, commas, dots, and minus
    s = s.replace(/[^0-9,\.\-]/g, '').trim();
    if (!s || s === '-' || s === '.' || s === ',') return "";

    // Handle combinations of dots and commas
    if (s.includes('.') && s.includes(',')) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
        // Dutch/EU format: 1.250,50 -> 1250.50
        s = s.replace(/\./g, '').replace(',', '.');
      } else {
        // US format: 1,250.50 -> 1250.50
        s = s.replace(/,/g, '');
      }
    } else if (s.includes(',')) {
      // Single comma, e.g. "500,00" or "299,00"
      s = s.replace(',', '.');
    } else if (s.includes('.')) {
      // Single dot: e.g. "1.250" (thousands in NL) vs "500.00" / "12.50" (decimal)
      if (/^\d{1,3}\.\d{3}$/.test(s)) {
        s = s.replace('.', '');
      }
    }

    const num = parseFloat(s);
    return isNaN(num) ? "" : Math.round(num * 100) / 100;
  }, []);

  // Age calculation
  const calculateAge = useCallback((dobDate: Date | null): number => {
    if (!dobDate) return 0;
    const today = new Date();
    let age = today.getFullYear() - dobDate.getFullYear();
    const m = today.getMonth() - dobDate.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < dobDate.getDate())) {
      age--;
    }
    return age;
  }, []);

  // Helper to extract properties cleanly from different header capitalizations and formats
  const getValueByHeader = useCallback((row: RawRow, headerNames: string[]): string => {
    if (!row) return "";
    const rowKeys = Object.keys(row);

    // 1. Direct exact match
    for (const hn of headerNames) {
      if (row[hn] !== undefined && row[hn] !== null && String(row[hn]).trim() !== "") {
        return String(row[hn]).trim();
      }
    }

    // 2. Case-insensitive trim match
    for (const hn of headerNames) {
      const lowerHn = hn.trim().toLowerCase();
      for (const key of rowKeys) {
        if (key.trim().toLowerCase() === lowerHn) {
          const val = row[key];
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            return String(val).trim();
          }
        }
      }
    }

    // 3. Normalized exact match (ignoring spaces, underscores, dots, hyphens, BOM)
    for (const hn of headerNames) {
      const normalizedHn = hn.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (!normalizedHn) continue;
      for (const key of rowKeys) {
        const normalizedKey = key.toLowerCase().replace(/^\ufeff/, '').replace(/[^a-z0-9]/g, '');
        if (normalizedKey === normalizedHn) {
          const val = row[key];
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            return String(val).trim();
          }
        }
      }
    }

    // 4. Prefix/Suffix match: The CSV column header contains the search keyword (normalizedHn)
    // E.g. CSV column "klant_achternaam" contains "achternaam", or "totale_omzet" contains "omzet"
    // CRITICAL: normalizedKey must be strictly equal or longer than normalizedHn to prevent false partial matches on generic words
    for (const hn of headerNames) {
      const normalizedHn = hn.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (normalizedHn.length >= 4) {
        for (const key of rowKeys) {
          const normalizedKey = key.toLowerCase().replace(/^\ufeff/, '').replace(/[^a-z0-9]/g, '');
          if (normalizedKey.length >= normalizedHn.length && normalizedKey.includes(normalizedHn)) {
            const val = row[key];
            if (val !== undefined && val !== null && String(val).trim() !== "") {
              return String(val).trim();
            }
          }
        }
      }
    }

    return "";
  }, []);

  // Convert raw row to the target layout
  const processRawRow = useCallback((row: RawRow, source: 'Ommen' | 'Zwolle'): ProcessedRow => {
    // 1. Email: Only accept actual valid email addresses (never set email to numeric ref!)
    const emailRaw = getValueByHeader(row, [
      "Klant e-mailadres", "E-mailadres", "Emailadres", "E-mail", "Email", "e-mailadres", "emailadres", "e-mail", "email",
      "Klant email", "Klant emailadres", "E-mailadres klant", "Email klant", "Klant e-mail", "E-mail contactpersoon", "Mail", "mail"
    ]);
    const refRaw = getKlantRef(row);
    const email = isLikelyEmail(emailRaw) ? emailRaw.trim().toLowerCase() : "";

    // 2. Phone formatting (mobile with +316, landline starting with 5 to 05...)
    const phoneRaw = getValueByHeader(row, [
      "Klant telefoonnummer", "Telefoonnummer", "Mobiel", "Telefoon", "Mobiel nummer", "Telefoon nummer", "Mobielnr", "Tel", "Phone", "Mobile", "GSM", "GSM nummer"
    ]);
    const phone = formatPhone(phoneRaw);

    // 3. Names (combining tussenvoegsel if separately provided)
    const fn = getValueByHeader(row, ["Klant voornaam", "Voornaam", "First name", "Firstname", "Roepnaam"]);
    const tussenvoegsel = getValueByHeader(row, ["Klant tussenvoegsel", "Tussenvoegsel", "Voorvoegsel", "Prefix"]);
    const lnRaw = getValueByHeader(row, ["Klant achternaam", "Achternaam", "Last name", "Lastname", "Tussenvoegsel + achternaam", "Achternaam klant"]);
    const ln = tussenvoegsel && !lnRaw.startsWith(tussenvoegsel) ? `${tussenvoegsel} ${lnRaw}`.trim() : lnRaw;

    // 4. Zip code & City
    const zip = getValueByHeader(row, ["Klant postcode", "Postcode", "Zip", "Zipcode", "Postal code"]);
    const ct = getValueByHeader(row, ["Klant woonplaats", "Woonplaats", "Plaats", "City", "Stad", "Dorp"]);
    const country = "NL";

    // 5. Birthdate & Age
    const dobRaw = getValueByHeader(row, ["Klant geboortedatum", "Geboortedatum", "Geboorte datum", "Birthdate", "Date of birth", "DOB"]);
    const dobDate = parseDate(dobRaw);
    const dob = dobDate ? format(dobDate, 'yyyy-MM-dd') : "";
    const doby = dobDate ? format(dobDate, 'yyyy') : "";
    const age = dobDate ? calculateAge(dobDate) : "";

    // 6. Gender mapping (Man -> M, Vrouw -> F)
    const genRaw = getValueByHeader(row, ["Klant geslacht", "Geslacht", "Gender", "Geslacht / aanhef", "Aanhef", "Sex"]);
    let gen = "";
    if (genRaw) {
      const gLower = genRaw.toLowerCase().trim();
      if (gLower === "man" || gLower === "m" || gLower === "mannelijk") {
        gen = "M";
      } else if (gLower === "vrouw" || gLower === "f" || gLower === "v" || gLower === "vrouwelijk") {
        gen = "F";
      } else {
        gen = genRaw;
      }
    }

    // 7. Tariff / Value Parsing with comprehensive field name support & scanning
    let tariffRaw = getValueByHeader(row, [
      "Tarief", "tarief", "Prijs", "prijs", "Bedrag", "bedrag", "Tariff", "tariff", "Value", "value", 
      "Abonnementsprijs", "Abonnementsbedrag", "Contributie", "Maandbedrag", "Kosten", "Omzet", 
      "Totale omzet", "Totaal omzet", "Factuurbedrag", "Totaal besteed", "LTV", "Lifetime value", 
      "Klantwaarde", "Totaalbedrag", "Totaal bedrag", "Totaal", "Total", "Investering", "Aankoopbedrag", 
      "Prijs incl. btw", "Prijs incl btw", "Prijs excl. btw", "Prijs excl btw", "Prijs per maand", "Tarief per maand", "Price", "Amount"
    ]);

    // Fallback: Scan row keys for specific price headers if not found yet
    if (!tariffRaw && row) {
      const priceKeywords = ['tarief', 'prijs', 'bedrag', 'kosten', 'contributie', 'waarde', 'omzet', 'factuurbedrag', 'totaalbedrag', 'maandbedrag'];
      for (const [k, v] of Object.entries(row)) {
        const kNorm = k.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (priceKeywords.some(pk => kNorm.includes(pk)) && v !== undefined && v !== null && String(v).trim() !== "") {
          tariffRaw = String(v).trim();
          break;
        }
      }
    }

    let value = parseCurrency(tariffRaw);

    const actiefSindsRaw = getValueByHeader(row, [
      "Actief sinds", "Startdatum", "Begindatum", "Datum actief", "Actief vanaf", "Vanaf datum", "Ingangsdatum", "Datum", "Lid sinds", "Datum aanmaak", "Aanmaakdatum", "Datum inschrijving", "Inschrijfdatum"
    ]);
    const actiefSindsDate = parseDate(actiefSindsRaw);

    // 8. Product name extraction from all possible product / subscription columns
    let product = getValueByHeader(row, [
      "Naam van product", 
      "Productnaam", 
      "Product", 
      "Abonnement", 
      "Dienst", 
      "Naam abonnement", 
      "Naam lidmaatschap", 
      "Lidmaatschap", 
      "Contract", 
      "Omschrijving", 
      "Productomschrijving", 
      "Artikel", 
      "Artikelnaam", 
      "Membership", 
      "Plan", 
      "Service", 
      "Soort lidmaatschap", 
      "Soort abonnement", 
      "Pakket", 
      "Pakketnaam", 
      "Package", 
      "Programma", 
      "Traject", 
      "Cursus", 
      "Item", 
      "Type"
    ]);

    // Fallback: Scan row keys for product/subscription columns if not found
    if (!product && row) {
      const prodKeywords = ['product', 'abonnement', 'lidmaatschap', 'dienst', 'contract', 'pakket', 'membership', 'omschrijving', 'artikel', 'traject', 'programma'];
      for (const [k, v] of Object.entries(row)) {
        const kNorm = k.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (prodKeywords.some(pk => kNorm.includes(pk)) && v !== undefined && v !== null && String(v).trim() !== "") {
          product = String(v).trim();
          break;
        }
      }
    }

    // Extra fallback: if value is still empty, see if product text contains price (e.g. €299 or €500 or €650)
    if (value === "" && product) {
      const matchPrice = product.match(/(?:€|eur|euro)\s*([0-9]+(?:[.,][0-9]{2})?)/i);
      if (matchPrice) {
        const extracted = parseCurrency(matchPrice[1]);
        if (typeof extracted === 'number') {
          value = extracted;
        }
      }
    }

    return {
      email,
      phone,
      fn,
      ln,
      zip,
      ct,
      country,
      dob,
      doby,
      gen,
      age: age !== "" ? Number(age) : "",
      value,
      product,
      _actiefSindsDate: actiefSindsDate,
      _source: source,
      _klantRef: refRaw,
      _raw: row
    };
  }, [formatPhone, parseDate, calculateAge, getValueByHeader, parseCurrency]);

  // Helper to parse either an Excel (.xlsx, .xls) or CSV (.csv) file into raw row objects
  const parseFileToRows = async (file: File): Promise<RawRow[]> => {
    const fileName = file.name.toLowerCase();
    const isExcel = fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.xlsm') || fileName.endsWith('.xlsb');

    if (isExcel) {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        return [];
      }
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawJson = XLSX.utils.sheet_to_json<Record<string, any>>(firstSheet, {
        defval: '',
        raw: false,
        dateNF: 'yyyy-mm-dd'
      });

      const cleanedRows: RawRow[] = rawJson.map(row => {
        const cleanRow: RawRow = {};
        for (const [k, v] of Object.entries(row)) {
          const cleanKey = k.trim().replace(/^\ufeff/, '').replace(/^["']|["']$/g, '');
          if (cleanKey) {
            cleanRow[cleanKey] = v !== null && v !== undefined ? String(v).trim() : '';
          }
        }
        return cleanRow;
      }).filter(r => Object.values(r).some(v => v !== ''));

      return cleanedRows;
    } else {
      // CSV parse using PapaParse wrapped in Promise
      return new Promise<RawRow[]>((resolve, reject) => {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: 'greedy',
          dynamicTyping: false,
          transformHeader: (h: string) => h.trim().replace(/^\ufeff/, '').replace(/^["']|["']$/g, ''),
          complete: (results) => {
            const rows = (results.data || []) as RawRow[];
            const cleaned = rows.filter(r => r && Object.values(r).some(v => v !== null && v !== undefined && String(v).trim() !== ''));
            resolve(cleaned);
          },
          error: (err) => {
            reject(err);
          }
        });
      });
    }
  };

  // Handle parsing multiple CSV or Excel files with robust separation configuration
  const handleMultipleUploads = async (files: FileList | File[], source: 'Ommen' | 'Zwolle') => {
    setStatusMessage(null);
    const fileArray = Array.from(files);
    const validFiles = fileArray.filter(f => {
      const name = f.name.toLowerCase();
      return name.endsWith('.csv') || name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.xlsm') || name.endsWith('.xlsb') ||
             f.type === 'text/csv' || f.type === 'application/vnd.ms-excel' || f.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    });

    if (validFiles.length === 0) {
      setStatusMessage({
        type: 'err',
        text: 'Geen geldige .csv of Excel (.xlsx, .xls) bestanden gevonden.'
      });
      return;
    }

    let successCount = 0;
    let totalRowsAdded = 0;
    const newItems: UploadedFileItem[] = [];
    const errors: string[] = [];

    for (const file of validFiles) {
      try {
        const rows = await parseFileToRows(file);
        if (rows.length > 0) {
          successCount++;
          totalRowsAdded += rows.length;
          const detectedRole = detectFileRole(rows, file.name);
          newItems.push({
            id: `${file.name}-${file.size}-${Date.now()}-${Math.random()}`,
            file,
            rows,
            role: detectedRole
          });
        } else {
          errors.push(`${file.name} (geen data rijen gevonden)`);
        }
      } catch (err) {
        errors.push(`${file.name}: ${err instanceof Error ? err.message : 'Onbekende fout'}`);
      }
    }

    // Auto-disambiguate if exactly 2 files uploaded together and both have same role
    if (newItems.length === 2 && newItems[0].role === newItems[1].role) {
      const lower0 = newItems[0].file.name.toLowerCase();
      const lower1 = newItems[1].file.name.toLowerCase();
      const p0 = lower0.includes('product') || lower0.includes('abonnement') || lower0.includes('tarief');
      const p1 = lower1.includes('product') || lower1.includes('abonnement') || lower1.includes('tarief');
      if (p0 && !p1) {
        newItems[0].role = 'products';
        newItems[1].role = 'customers';
      } else if (p1 && !p0) {
        newItems[1].role = 'products';
        newItems[0].role = 'customers';
      }
    }

    if (newItems.length > 0) {
      if (source === 'Ommen') {
        setOmmenFiles(prev => [...prev, ...newItems]);
      } else {
        setZwolleFiles(prev => [...prev, ...newItems]);
      }
    }

    if (successCount > 0) {
      setStatusMessage({
        type: 'success',
        text: `${successCount} bestand(en) voor ${source} geïmporteerd (${totalRowsAdded} rijen toegevoegd)!${errors.length > 0 ? ` (Waarschuwing: ${errors.join(', ')})` : ''}`
      });
    } else {
      setStatusMessage({
        type: 'err',
        text: `Fout bij importeren van bestanden voor ${source}: ${errors.join(', ')}`
      });
    }
  };

  // Helper to change file role (Klantenbestand vs Productenbestand)
  const updateFileRole = (source: 'Ommen' | 'Zwolle', id: string, newRole: 'customers' | 'products') => {
    if (source === 'Ommen') {
      setOmmenFiles(prev => prev.map(f => f.id === id ? { ...f, role: newRole } : f));
    } else {
      setZwolleFiles(prev => prev.map(f => f.id === id ? { ...f, role: newRole } : f));
    }
    setStatusMessage({
      type: 'info',
      text: `Bestandsrol aangepast naar ${newRole === 'customers' ? 'Klantenbestand' : 'Productenbestand'}. Data wordt opnieuw gekoppeld op "Klant ref.".`
    });
  };

  // Drag over handlers
  const handleDragOver = (e: React.DragEvent, source: 'Ommen' | 'Zwolle') => {
    e.preventDefault();
    if (source === 'Ommen') {
      setDragOverOmmen(true);
    } else {
      setDragOverZwolle(true);
    }
  };

  const handleDragLeave = (source: 'Ommen' | 'Zwolle') => {
    if (source === 'Ommen') {
      setDragOverOmmen(false);
    } else {
      setDragOverZwolle(false);
    }
  };

  const handleDrop = (e: React.DragEvent, source: 'Ommen' | 'Zwolle') => {
    e.preventDefault();
    if (source === 'Ommen') {
      setDragOverOmmen(false);
    } else {
      setDragOverZwolle(false);
    }
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleMultipleUploads(e.dataTransfer.files, source);
    }
  };

  const removeSingleFile = (source: 'Ommen' | 'Zwolle', id: string) => {
    if (source === 'Ommen') {
      setOmmenFiles(prev => prev.filter(item => item.id !== id));
    } else {
      setZwolleFiles(prev => prev.filter(item => item.id !== id));
    }
    setStatusMessage({
      type: 'info',
      text: `Bestand verwijderd uit ${source}.`
    });
  };

  const clearAllFiles = (source: 'Ommen' | 'Zwolle') => {
    if (source === 'Ommen') {
      setOmmenFiles([]);
    } else {
      setZwolleFiles([]);
    }
    setStatusMessage({
      type: 'info',
      text: `Alle bestanden voor ${source} verwijderd.`
    });
  };

  // Compile all data and execute business logic:
  // Match Customers and Products on "Klant ref.", compute single high-value tariffs (>= €400),
  // enrich exports with full customer details (email, phone, address, dob, gen, age).
  const processedData = useMemo(() => {
    const allFilesWithSource = [
      ...ommenFiles.map(f => ({ ...f, source: 'Ommen' as const })),
      ...zwolleFiles.map(f => ({ ...f, source: 'Zwolle' as const }))
    ];

    if (allFilesWithSource.length === 0) {
      const emptyProductExports: Record<string, { ommen: ProcessedRow[]; zwolle: ProcessedRow[]; all: ProcessedRow[] }> = {};
      PRODUCT_CATEGORIES.forEach(cat => {
        emptyProductExports[cat.id] = { ommen: [], zwolle: [], all: [] };
      });
      return {
        all: [],
        highValue: [],
        highValueOmmen: [],
        highValueZwolle: [],
        duplicates: [],
        productExports: emptyProductExports,
        matchedProductsCount: 0,
        totalCustomersCount: 0
      };
    }

    interface CustProfile {
      ref: string;
      email: string;
      phone: string;
      fn: string;
      ln: string;
      zip: string;
      ct: string;
      country: string;
      dob: string;
      doby: string;
      gen: string;
      age: number | string;
      baseValue: number | string;
      baseProduct: string;
      baseDate: Date | null;
      source: 'Ommen' | 'Zwolle';
      raw: RawRow;
      maxSingleTariff: number;
      highValueProduct: string;
      products: Array<{
        name: string;
        value: number | string;
        date: Date | null;
        raw: RawRow;
        source: 'Ommen' | 'Zwolle';
      }>;
    }

    const customerFiles = allFilesWithSource.filter(f => f.role === 'customers');
    const productFiles = allFilesWithSource.filter(f => f.role === 'products');

    const profilesByRef = new Map<string, CustProfile>();
    const profilesByEmail = new Map<string, CustProfile>();
    const profilesByName = new Map<string, CustProfile>();
    const profilesList: CustProfile[] = [];

    const candidateProductRecords: ProcessedRow[] = [];
    let matchedProductsCount = 0;

    // Helper to find existing customer profile
    const findCustomer = (ref: string, email: string, fn: string, ln: string): CustProfile | undefined => {
      if (ref) {
        const trimmed = ref.trim();
        let match = profilesByRef.get(trimmed);
        if (match) return match;
        match = profilesByRef.get(trimmed.toLowerCase());
        if (match) return match;
        if (/^\d+$/.test(trimmed)) {
          match = profilesByRef.get(String(Number(trimmed)));
          if (match) return match;
        }
      }
      if (email && isLikelyEmail(email)) {
        const match = profilesByEmail.get(email.toLowerCase().trim());
        if (match) return match;
      }
      if (fn && ln) {
        const match = profilesByName.get(`${fn.toLowerCase().trim()}_${ln.toLowerCase().trim()}`);
        if (match) return match;
      }
      return undefined;
    };

    // Helper to register or merge customer profile
    const registerOrMergeCustomer = (rawRow: RawRow, source: 'Ommen' | 'Zwolle') => {
      const rowProcessed = processRawRow(rawRow, source);
      const ref = getKlantRef(rawRow);
      const existing = findCustomer(ref, rowProcessed.email, rowProcessed.fn, rowProcessed.ln);

      if (existing) {
        if (!existing.email && rowProcessed.email) existing.email = rowProcessed.email;
        if (!existing.phone && rowProcessed.phone) existing.phone = rowProcessed.phone;
        if (!existing.fn && rowProcessed.fn) existing.fn = rowProcessed.fn;
        if (!existing.ln && rowProcessed.ln) existing.ln = rowProcessed.ln;
        if (!existing.zip && rowProcessed.zip) existing.zip = rowProcessed.zip;
        if (!existing.ct && rowProcessed.ct) existing.ct = rowProcessed.ct;
        if (!existing.dob && rowProcessed.dob) {
          existing.dob = rowProcessed.dob;
          existing.doby = rowProcessed.doby;
          existing.age = rowProcessed.age;
        }
        if (!existing.gen && rowProcessed.gen) existing.gen = rowProcessed.gen;
        if (!existing.ref && ref) existing.ref = ref;
        if (rowProcessed.product && !existing.baseProduct) {
          existing.baseProduct = rowProcessed.product;
          existing.baseValue = rowProcessed.value;
          existing.baseDate = rowProcessed._actiefSindsDate;
        }
      } else {
        const valNum = typeof rowProcessed.value === 'number' ? rowProcessed.value : (rowProcessed.value ? Number(rowProcessed.value) : 0);
        const prof: CustProfile = {
          ref,
          email: rowProcessed.email,
          phone: rowProcessed.phone,
          fn: rowProcessed.fn,
          ln: rowProcessed.ln,
          zip: rowProcessed.zip,
          ct: rowProcessed.ct,
          country: 'NL',
          dob: rowProcessed.dob,
          doby: rowProcessed.doby,
          gen: rowProcessed.gen,
          age: rowProcessed.age,
          baseValue: rowProcessed.value,
          baseProduct: rowProcessed.product,
          baseDate: rowProcessed._actiefSindsDate,
          source,
          raw: rawRow,
          maxSingleTariff: !isNaN(valNum) ? valNum : 0,
          highValueProduct: (!isNaN(valNum) && valNum >= HIGH_VALUE_THRESHOLD) ? rowProcessed.product : '',
          products: []
        };

        profilesList.push(prof);
        if (ref) {
          profilesByRef.set(ref.trim(), prof);
          profilesByRef.set(ref.trim().toLowerCase(), prof);
          if (/^\d+$/.test(ref.trim())) {
            profilesByRef.set(String(Number(ref.trim())), prof);
          }
        }
        if (rowProcessed.email) {
          profilesByEmail.set(rowProcessed.email.toLowerCase().trim(), prof);
        }
        if (rowProcessed.fn && rowProcessed.ln) {
          profilesByName.set(`${rowProcessed.fn.toLowerCase().trim()}_${rowProcessed.ln.toLowerCase().trim()}`, prof);
        }
      }
    };

    // 1. Process Customer files (or all files if user only uploaded 1 combined file)
    const filesForCustomers = customerFiles.length > 0 ? customerFiles : allFilesWithSource;
    filesForCustomers.forEach(f => {
      f.rows.forEach(r => registerOrMergeCustomer(r, f.source));
    });

    // 2. Process Product files and link to Customers on "Klant ref."
    if (productFiles.length > 0) {
      productFiles.forEach(f => {
        f.rows.forEach(rawRow => {
          const ref = getKlantRef(rawRow);
          const rawEmail = getValueByHeader(rawRow, ["Klant e-mailadres", "E-mailadres", "Email"]);
          const rawFn = getValueByHeader(rawRow, ["Klant voornaam", "Voornaam"]);
          const rawLn = getValueByHeader(rawRow, ["Klant achternaam", "Achternaam"]);

          const cust = findCustomer(ref, rawEmail, rawFn, rawLn);

          const prodName = getValueByHeader(rawRow, [
            "Naam van product", "Productnaam", "Product", "Abonnement", "Dienst", "Naam abonnement", 
            "Naam lidmaatschap", "Lidmaatschap", "Contract", "Omschrijving", "Artikel", "Pakket", "Pakketnaam", "Programma", "Traject"
          ]);
          const tariffRaw = getValueByHeader(rawRow, [
            "Tarief", "tarief", "Prijs", "prijs", "Bedrag", "bedrag", "Tariff", "tariff", "Value", "value", 
            "Abonnementsprijs", "Contributie", "Maandbedrag", "Factuurbedrag", "Totaalbedrag", "Investering", "Aankoopbedrag"
          ]);
          const prodValue = parseCurrency(tariffRaw);
          const valNum = typeof prodValue === 'number' ? prodValue : (prodValue ? Number(prodValue) : 0);

          const prodDateRaw = getValueByHeader(rawRow, [
            "Actief sinds", "Startdatum", "Begindatum", "Datum actief", "Actief vanaf", "Ingangsdatum", "Datum"
          ]);
          const prodDate = parseDate(prodDateRaw);

          if (cust) {
            matchedProductsCount++;
            cust.products.push({
              name: prodName,
              value: prodValue,
              date: prodDate,
              raw: rawRow,
              source: f.source
            });

            if (!isNaN(valNum) && valNum > cust.maxSingleTariff) {
              cust.maxSingleTariff = valNum;
            }
            if (!isNaN(valNum) && valNum >= HIGH_VALUE_THRESHOLD) {
              cust.highValueProduct = prodName || cust.highValueProduct;
            }

            // Create enriched candidate product record with full customer info!
            candidateProductRecords.push({
              email: cust.email,
              phone: cust.phone,
              fn: cust.fn,
              ln: cust.ln,
              zip: cust.zip,
              ct: cust.ct,
              country: 'NL',
              dob: cust.dob,
              doby: cust.doby,
              gen: cust.gen,
              age: cust.age,
              value: prodValue !== "" ? prodValue : cust.baseValue,
              product: prodName || cust.baseProduct,
              _actiefSindsDate: prodDate || cust.baseDate,
              _source: f.source,
              _klantRef: cust.ref || ref,
              _raw: { ...cust.raw, ...rawRow }
            });
          } else {
            // Unmatched product record
            const standalone = processRawRow(rawRow, f.source);
            candidateProductRecords.push(standalone);
          }
        });
      });
    }

    // Step A: Assemble all records
    const allRecords: ProcessedRow[] = [];

    profilesList.forEach(cust => {
      // Sort products by date: newest first
      if (cust.products.length > 1) {
        cust.products.sort((a, b) => {
          if (!a.date && !b.date) return 0;
          if (!a.date) return 1;
          if (!b.date) return -1;
          return b.date.getTime() - a.date.getTime();
        });
      }

      const latestProd = cust.products[0];

      let displayValue: number | string = "";
      let displayProduct = "";
      let displayDate: Date | null = cust.baseDate;

      // If customer has a single purchase >= HIGH_VALUE_THRESHOLD (400), reflect that value
      if (cust.maxSingleTariff >= HIGH_VALUE_THRESHOLD) {
        displayValue = cust.maxSingleTariff;
        displayProduct = cust.highValueProduct || (latestProd ? latestProd.name : cust.baseProduct);
        displayDate = latestProd ? (latestProd.date || cust.baseDate) : cust.baseDate;
      } else if (latestProd) {
        displayValue = latestProd.value !== "" ? latestProd.value : cust.baseValue;
        displayProduct = latestProd.name || cust.baseProduct;
        displayDate = latestProd.date || cust.baseDate;
      } else {
        displayValue = cust.baseValue;
        displayProduct = cust.baseProduct;
        displayDate = cust.baseDate;
      }

      const row: ProcessedRow = {
        email: cust.email,
        phone: cust.phone,
        fn: cust.fn,
        ln: cust.ln,
        zip: cust.zip,
        ct: cust.ct,
        country: 'NL',
        dob: cust.dob,
        doby: cust.doby,
        gen: cust.gen,
        age: cust.age,
        value: displayValue,
        product: displayProduct,
        _actiefSindsDate: displayDate,
        _source: cust.source,
        _maxHistoricalValue: cust.maxSingleTariff,
        _klantRef: cust.ref,
        _raw: cust.raw
      };

      allRecords.push(row);
    });

    // If no separate product files were uploaded, candidateProductRecords is allRecords
    if (productFiles.length === 0) {
      candidateProductRecords.push(...allRecords);
    }

    // Step B: Deduplication
    let finalRecords = [...allRecords];
    const duplicatesList: DuplicateLog[] = [];

    if (deduplicate) {
      const groups: { [key: string]: ProcessedRow[] } = {};
      finalRecords.forEach(rec => {
        const key = rec._klantRef
          ? `ref_${rec._klantRef.toLowerCase().trim()}`
          : (rec.email ? `email_${rec.email.toLowerCase().trim()}` : `name_${rec.fn.toLowerCase().trim()}_${rec.ln.toLowerCase().trim()}`);
        if (!groups[key]) groups[key] = [];
        groups[key].push(rec);
      });

      const uniqueList: ProcessedRow[] = [];
      Object.keys(groups).forEach(key => {
        const group = groups[key];

        let maxSingleTariff = 0;
        group.forEach(r => {
          const valNum = typeof r.value === 'number' ? r.value : (r.value !== "" ? Number(r.value) : 0);
          const maxVal = r._maxHistoricalValue || 0;
          const best = Math.max(!isNaN(valNum) ? valNum : 0, maxVal);
          if (best > maxSingleTariff) {
            maxSingleTariff = best;
          }
        });

        if (group.length > 1) {
          // Sort by completeness of customer information (email, phone, address, dob)
          group.sort((a, b) => {
            const scoreA = (a.email ? 4 : 0) + (a.phone ? 3 : 0) + (a.zip ? 2 : 0) + (a.dob ? 2 : 0);
            const scoreB = (b.email ? 4 : 0) + (b.phone ? 3 : 0) + (b.zip ? 2 : 0) + (b.dob ? 2 : 0);
            if (scoreB !== scoreA) return scoreB - scoreA;
            if (a._actiefSindsDate && b._actiefSindsDate) return b._actiefSindsDate.getTime() - a._actiefSindsDate.getTime();
            return 0;
          });

          const bestRow: ProcessedRow = { ...group[0] };
          if (maxSingleTariff >= HIGH_VALUE_THRESHOLD && (bestRow.value === "" || Number(bestRow.value) < HIGH_VALUE_THRESHOLD)) {
            bestRow.value = maxSingleTariff;
          } else if ((bestRow.value === "" || bestRow.value === 0 || Number(bestRow.value) === 0) && maxSingleTariff > 0) {
            bestRow.value = maxSingleTariff;
          }
          bestRow._maxHistoricalValue = maxSingleTariff;

          uniqueList.push(bestRow);
          duplicatesList.push({
            key,
            records: group,
            selected: bestRow
          });
        } else {
          const singleRow: ProcessedRow = { ...group[0] };
          singleRow._maxHistoricalValue = maxSingleTariff;
          uniqueList.push(singleRow);
        }
      });
      finalRecords = uniqueList;
    } else {
      finalRecords = finalRecords.map(r => ({
        ...r,
        _maxHistoricalValue: typeof r.value === 'number' ? r.value : (r.value ? Number(r.value) : 0)
      }));
    }

    // Step C: Optional tariff filter
    if (tariffFilter) {
      finalRecords = finalRecords.filter(item => {
        if (item.value === "") return false;
        const num = Number(item.value);
        return !isNaN(num) && num !== 0;
      });
    }

    // High Value Category definition: individual single tariff/value >= 400
    const highValueRecords = finalRecords.filter(item => {
      const num = typeof item.value === 'number' ? item.value : (item.value !== "" ? Number(item.value) : NaN);
      return !isNaN(num) && num >= HIGH_VALUE_THRESHOLD;
    });

    const highValueOmmen = highValueRecords.filter(r => r._source === 'Ommen');
    const highValueZwolle = highValueRecords.filter(r => r._source === 'Zwolle');

    // Step D: Product Categories grouping
    const productExports: Record<string, { ommen: ProcessedRow[]; zwolle: ProcessedRow[]; all: ProcessedRow[] }> = {};
    PRODUCT_CATEGORIES.forEach(cat => {
      productExports[cat.id] = { ommen: [], zwolle: [], all: [] };
    });

    PRODUCT_CATEGORIES.forEach(cat => {
      const matchingRecords = candidateProductRecords.filter(rec => matchesCategory(rec, cat));

      let categoryRecords = matchingRecords;
      if (deduplicate) {
        const groups: { [key: string]: ProcessedRow[] } = {};
        matchingRecords.forEach(rec => {
          const key = rec._klantRef
            ? `ref_${rec._klantRef.toLowerCase().trim()}`
            : (rec.email ? `email_${rec.email.toLowerCase().trim()}` : `name_${rec.fn.toLowerCase().trim()}_${rec.ln.toLowerCase().trim()}`);
          if (!groups[key]) groups[key] = [];
          groups[key].push(rec);
        });

        const uniqueCategoryRecords: ProcessedRow[] = [];
        Object.keys(groups).forEach(key => {
          const group = groups[key];
          if (group.length > 1) {
            group.sort((a, b) => {
              const scoreA = (a.email ? 4 : 0) + (a.phone ? 3 : 0) + (a.zip ? 2 : 0);
              const scoreB = (b.email ? 4 : 0) + (b.phone ? 3 : 0) + (b.zip ? 2 : 0);
              if (scoreB !== scoreA) return scoreB - scoreA;
              if (a._actiefSindsDate && b._actiefSindsDate) return b._actiefSindsDate.getTime() - a._actiefSindsDate.getTime();
              return 0;
            });
            uniqueCategoryRecords.push(group[0]);
          } else {
            uniqueCategoryRecords.push(group[0]);
          }
        });
        categoryRecords = uniqueCategoryRecords;
      }

      categoryRecords.forEach(rec => {
        productExports[cat.id].all.push(rec);
        if (rec._source === 'Ommen') {
          productExports[cat.id].ommen.push(rec);
        } else {
          productExports[cat.id].zwolle.push(rec);
        }
      });
    });

    return {
      all: finalRecords,
      highValue: highValueRecords,
      highValueOmmen,
      highValueZwolle,
      duplicates: duplicatesList,
      productExports,
      matchedProductsCount,
      totalCustomersCount: profilesList.length
    };
  }, [ommenFiles, zwolleFiles, processRawRow, deduplicate, tariffFilter]);

  // Exclude internally used keys (_raw, _source, _actiefSindsDate, _klantRef) when exporting to CSV
  const cleanForExport = (records: ProcessedRow[]) => {
    return records.map(item => ({
      email: item.email || '',
      phone: item.phone || '',
      fn: item.fn || '',
      ln: item.ln || '',
      zip: item.zip || '',
      ct: item.ct || '',
      country: item.country || 'NL',
      dob: item.dob || '',
      doby: item.doby || '',
      gen: item.gen || '',
      age: item.age !== undefined && item.age !== null ? item.age : '',
      value: item.value !== undefined && item.value !== null ? item.value : ''
    }));
  };

  // Perform CSV Unparse and Trigger Download to meet direct integer values constraint (Request 6)
  const downloadCSV = (records: ProcessedRow[], filename: string) => {
    if (records.length === 0) {
      setStatusMessage({
        type: 'err',
        text: `Geen records aanwezig voor "${filename}".`
      });
      return;
    }

    try {
      const exportData = cleanForExport(records);
      
      // We parse values defensively. High value integers and age floats will be written pure without surrounding quotations.
      const csvContent = Papa.unparse(exportData, {
        quotes: !noQuotes, // If noQuotes is true, PapaParse will not insert quotes around numbers or strings that don't need them!
        quoteChar: '"',
        escapeChar: '"',
        delimiter: ";" // Standard European semi-colon CSV format works best in standard Dutch Excel
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `${filename.trim()}.csv`);
      document.body.appendChild(link);
      link.click();
      setTimeout(() => {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }, 100);
      
      setStatusMessage({
        type: 'success',
        text: `Bestand "${filename}.csv" met ${records.length} records succesvol gedownload!`
      });
    } catch (e) {
      setStatusMessage({
        type: 'err',
        text: `Fout bij downloaden van ${filename}: ${e instanceof Error ? e.message : 'Onbekende fout'}`
      });
    }
  };

  // Helper to batch download all active non-empty product CSVs sequentially
  const downloadAllProductCSVs = () => {
    let downloadedCount = 0;
    PRODUCT_CATEGORIES.forEach(cat => {
      const ommenData = processedData.productExports[cat.id]?.ommen || [];
      const zwolleData = processedData.productExports[cat.id]?.zwolle || [];

      if (ommenData.length > 0) {
        setTimeout(() => {
          downloadCSV(ommenData, `${cat.name} - Ommen - ${todayStr}`);
        }, downloadedCount * 300);
        downloadedCount++;
      }

      if (zwolleData.length > 0) {
        setTimeout(() => {
          downloadCSV(zwolleData, `${cat.name} - Zwolle - ${todayStr}`);
        }, downloadedCount * 300);
        downloadedCount++;
      }
    });

    if (downloadedCount === 0) {
      setStatusMessage({
        type: 'info',
        text: 'Geen matchende productrecords gevonden om te exporteren.'
      });
    } else {
      setStatusMessage({
        type: 'success',
        text: `Bezig met downloaden van ${downloadedCount} productbestanden...`
      });
    }
  };

  // Calculate total active product records across all categories
  const totalProductExportsCount = useMemo(() => {
    return PRODUCT_CATEGORIES.reduce((acc, cat) => {
      const ommen = processedData.productExports[cat.id]?.ommen?.length || 0;
      const zwolle = processedData.productExports[cat.id]?.zwolle?.length || 0;
      return acc + ommen + zwolle;
    }, 0);
  }, [processedData.productExports]);

  return (
    <div id="app-root" className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans selection:bg-indigo-150 selection:text-indigo-900">
      {/* Header Bar */}
      <header id="app-header" className="border-b border-slate-200 bg-white sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 md:px-8 py-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-gradient-to-tr from-indigo-505 to-indigo-600 rounded-xl shadow-md shadow-indigo-500/10 text-white">
              <FileSpreadsheet className="w-6 h-6 stroke-[2]" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-950 flex items-center gap-2">
                Klantenbestand CSV & Excel Converter <span className="text-xs bg-indigo-100 text-indigo-800 border border-indigo-200/50 font-mono px-2 py-0.5 rounded-full">v2.2</span>
              </h1>
              <p className="text-xs text-slate-500 mt-0.5">Professionele verwerking van Excel (.xlsx, .xls) & CSV bestanden, duplicatenopschoning en +31-formattering</p>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs bg-slate-100 px-3 py-1.5 rounded-lg border border-slate-200 font-mono text-slate-650">
            <Sparkles className="w-3.5 h-3.5 text-indigo-500" />
            <span>Datum: {todayStr}</span>
          </div>
        </div>
      </header>

      <main className="flex-grow max-w-7xl w-full mx-auto px-4 md:px-8 py-8 flex flex-col gap-8">
        {/* Status Message Display */}
        {statusMessage && (
          <div 
            id="status-bar"
            className={`p-4 rounded-xl border flex items-start gap-3 animate-fade-in transition-all ${
              statusMessage.type === 'success' 
                ? 'bg-emerald-50 border-emerald-200 text-emerald-800' 
                : statusMessage.type === 'err' 
                ? 'bg-rose-50 border-rose-200 text-rose-800' 
                : 'bg-indigo-50 border-indigo-200 text-indigo-800'
            }`}
          >
            {statusMessage.type === 'success' && <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />}
            {statusMessage.type === 'err' && <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />}
            {statusMessage.type === 'info' && <Info className="w-5 h-5 text-indigo-600 shrink-0 mt-0.5" />}
            <span className="text-sm font-medium">{statusMessage.text}</span>
          </div>
        )}

        {/* Section 1: Twin Dropboxes (Ommen & Zwolle) */}
        <section id="upload-slots" className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Ommen Dropbox Slot */}
          <div 
            id="ommen-drop-zone"
            className={`border-2 border-dashed rounded-2xl p-6 transition-all duration-300 flex flex-col items-center justify-center text-center gap-4 cursor-pointer relative ${
              ommenFiles.length > 0 
                ? 'border-emerald-300 bg-emerald-50/20' 
                : dragOverOmmen 
                ? 'border-indigo-400 bg-indigo-50 scale-[1.01]' 
                : 'border-slate-300 bg-white hover:border-slate-400 hover:bg-slate-50/50'
            }`}
            onDragOver={(e) => handleDragOver(e, 'Ommen')}
            onDragLeave={() => handleDragLeave('Ommen')}
            onDrop={(e) => handleDrop(e, 'Ommen')}
          >
            <input 
              type="file" 
              accept=".csv, .xlsx, .xls, .xlsm, .xlsb, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv"
              multiple
              className="absolute inset-0 opacity-0 cursor-pointer z-0"
              onChange={(e) => {
                if (e.target.files && e.target.files.length > 0) {
                  handleMultipleUploads(e.target.files, 'Ommen');
                  e.target.value = '';
                }
              }}
            />
            
            <div className={`p-4 rounded-2xl ${ommenFiles.length > 0 ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-500'}`}>
              <MapPin className="w-8 h-8" />
            </div>

            <div>
              <h3 className="text-lg font-semibold text-slate-900">1. Vestiging OMMEN</h3>
              <p className="text-xs text-slate-500 mt-1 max-w-xs mx-auto">
                Sleep hier Ommen Excel- (.xlsx, .xls) of CSV-bestanden in, of klik om bestanden te kiezen.
              </p>
            </div>

            {ommenFiles.length > 0 ? (
              <div className="z-10 flex flex-col gap-2 w-full max-w-sm">
                <div className="flex items-center justify-between text-[11px] font-semibold text-slate-500 px-1 font-mono">
                  <span>{ommenFiles.length} bestand(en) • {ommenRows.length} rijen</span>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      clearAllFiles('Ommen');
                    }}
                    className="text-rose-600 hover:underline cursor-pointer"
                  >
                    Alles verwijderen
                  </button>
                </div>
                <div className="max-h-48 overflow-y-auto flex flex-col gap-1.5 pr-1">
                  {ommenFiles.map((item) => {
                    const isXls = item.file.name.toLowerCase().endsWith('.xlsx') || item.file.name.toLowerCase().endsWith('.xls') || item.file.name.toLowerCase().endsWith('.xlsm');
                    return (
                      <div key={item.id} className="bg-slate-50 border border-slate-200 rounded-xl p-2.5 flex items-center justify-between gap-3 text-left">
                        <div className="flex items-center gap-2 min-w-0">
                          {isXls ? (
                            <div className="p-1 bg-emerald-100 text-emerald-700 rounded-lg shrink-0">
                              <FileSpreadsheet className="w-4 h-4" />
                            </div>
                          ) : (
                            <div className="p-1 bg-blue-100 text-blue-700 rounded-lg shrink-0">
                              <FileText className="w-4 h-4" />
                            </div>
                          )}
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <p className="text-xs font-semibold text-slate-800 truncate max-w-[140px] sm:max-w-[180px]">{item.file.name}</p>
                              <select
                                value={item.role}
                                onClick={(e) => e.stopPropagation()}
                                onChange={(e) => updateFileRole('Ommen', item.id, e.target.value as 'customers' | 'products')}
                                className={`text-[9px] font-mono uppercase px-1.5 py-0.5 rounded font-bold border cursor-pointer focus:outline-none ${
                                  item.role === 'customers' 
                                    ? 'bg-blue-50 text-blue-700 border-blue-200' 
                                    : 'bg-purple-50 text-purple-700 border-purple-200'
                                }`}
                                title="Klik om te wisselen tussen Klantenbestand en Productenbestand"
                              >
                                <option value="customers">👤 Klanten</option>
                                <option value="products">🏷️ Producten</option>
                              </select>
                              <span className={`text-[9px] font-mono uppercase px-1.5 py-0.2 rounded font-bold ${
                                isXls ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-700'
                              }`}>
                                {isXls ? 'Excel' : 'CSV'}
                              </span>
                            </div>
                            <p className="text-[10px] text-slate-500 font-mono mt-0.5">
                              {(item.file.size / 1024).toFixed(1)} KB • {item.rows.length} rijen
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            removeSingleFile('Ommen', item.id);
                          }}
                          className="p-1 hover:bg-slate-200 text-slate-400 hover:text-rose-600 rounded-lg transition-colors shrink-0"
                          title="Verwijder bestand"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    );
                  })}
                </div>
                {ommenFiles.some(f => f.role === 'customers') && ommenFiles.some(f => f.role === 'products') ? (
                  <div className="mt-1 p-2 bg-emerald-50 border border-emerald-200 rounded-xl text-left flex items-center gap-1.5 text-emerald-800 text-[10px] font-medium">
                    <Link2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span><strong>Koppeling actief:</strong> Klanten & Producten gematcht op <em>"Klant ref."</em></span>
                  </div>
                ) : (
                  <div className="mt-1 p-2 bg-indigo-50/70 border border-indigo-100 rounded-xl text-left flex items-center gap-1.5 text-indigo-700 text-[10px]">
                    <Info className="w-3.5 h-3.5 text-indigo-500 shrink-0" />
                    <span>Upload 1 klantenbestand + 1 productenbestand om te koppelen op <em>"Klant ref."</em></span>
                  </div>
                )}
              </div>
            ) : (
              <div className="text-[10px] bg-slate-100 px-3 py-1 rounded-full text-slate-400 font-mono border border-slate-200/60 flex items-center gap-1.5">
                <span>CSV / XLSX / XLS (Klanten + Producten)</span>
              </div>
            )}
          </div>

          {/* Zwolle Dropbox Slot */}
          <div 
            id="zwolle-drop-zone"
            className={`border-2 border-dashed rounded-2xl p-6 transition-all duration-300 flex flex-col items-center justify-center text-center gap-4 cursor-pointer relative ${
              zwolleFiles.length > 0 
                ? 'border-emerald-300 bg-emerald-50/20' 
                : dragOverZwolle 
                ? 'border-indigo-400 bg-indigo-50 scale-[1.01]' 
                : 'border-slate-300 bg-white hover:border-slate-400 hover:bg-slate-50/50'
            }`}
            onDragOver={(e) => handleDragOver(e, 'Zwolle')}
            onDragLeave={() => handleDragLeave('Zwolle')}
            onDrop={(e) => handleDrop(e, 'Zwolle')}
          >
            <input 
              type="file" 
              accept=".csv, .xlsx, .xls, .xlsm, .xlsb, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv"
              multiple
              className="absolute inset-0 opacity-0 cursor-pointer z-0"
              onChange={(e) => {
                if (e.target.files && e.target.files.length > 0) {
                  handleMultipleUploads(e.target.files, 'Zwolle');
                  e.target.value = '';
                }
              }}
            />
            
            <div className={`p-4 rounded-2xl ${zwolleFiles.length > 0 ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-500'}`}>
              <MapPin className="w-8 h-8" />
            </div>

            <div>
              <h3 className="text-lg font-semibold text-slate-900">2. Vestiging ZWOLLE</h3>
              <p className="text-xs text-slate-500 mt-1 max-w-xs mx-auto">
                Sleep hier Zwolle Excel- of CSV-bestanden in (klanten- én productenbestand).
              </p>
            </div>

            {zwolleFiles.length > 0 ? (
              <div className="z-10 flex flex-col gap-2 w-full max-w-sm">
                <div className="flex items-center justify-between text-[11px] font-semibold text-slate-500 px-1 font-mono">
                  <span>{zwolleFiles.length} bestand(en) • {zwolleRows.length} rijen</span>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      clearAllFiles('Zwolle');
                    }}
                    className="text-rose-600 hover:underline cursor-pointer"
                  >
                    Alles verwijderen
                  </button>
                </div>
                <div className="max-h-48 overflow-y-auto flex flex-col gap-1.5 pr-1">
                  {zwolleFiles.map((item) => {
                    const isXls = item.file.name.toLowerCase().endsWith('.xlsx') || item.file.name.toLowerCase().endsWith('.xls') || item.file.name.toLowerCase().endsWith('.xlsm');
                    return (
                      <div key={item.id} className="bg-slate-50 border border-slate-200 rounded-xl p-2.5 flex items-center justify-between gap-3 text-left">
                        <div className="flex items-center gap-2 min-w-0">
                          {isXls ? (
                            <div className="p-1 bg-emerald-100 text-emerald-700 rounded-lg shrink-0">
                              <FileSpreadsheet className="w-4 h-4" />
                            </div>
                          ) : (
                            <div className="p-1 bg-blue-100 text-blue-700 rounded-lg shrink-0">
                              <FileText className="w-4 h-4" />
                            </div>
                          )}
                          <div className="min-w-0">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <p className="text-xs font-semibold text-slate-800 truncate max-w-[140px] sm:max-w-[180px]">{item.file.name}</p>
                              <select
                                value={item.role}
                                onClick={(e) => e.stopPropagation()}
                                onChange={(e) => updateFileRole('Zwolle', item.id, e.target.value as 'customers' | 'products')}
                                className={`text-[9px] font-mono uppercase px-1.5 py-0.5 rounded font-bold border cursor-pointer focus:outline-none ${
                                  item.role === 'customers' 
                                    ? 'bg-blue-50 text-blue-700 border-blue-200' 
                                    : 'bg-purple-50 text-purple-700 border-purple-200'
                                }`}
                                title="Klik om te wisselen tussen Klantenbestand en Productenbestand"
                              >
                                <option value="customers">👤 Klanten</option>
                                <option value="products">🏷️ Producten</option>
                              </select>
                              <span className={`text-[9px] font-mono uppercase px-1.5 py-0.2 rounded font-bold ${
                                isXls ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-700'
                              }`}>
                                {isXls ? 'Excel' : 'CSV'}
                              </span>
                            </div>
                            <p className="text-[10px] text-slate-500 font-mono mt-0.5">
                              {(item.file.size / 1024).toFixed(1)} KB • {item.rows.length} rijen
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            removeSingleFile('Zwolle', item.id);
                          }}
                          className="p-1 hover:bg-slate-200 text-slate-400 hover:text-rose-600 rounded-lg transition-colors shrink-0"
                          title="Verwijder bestand"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    );
                  })}
                </div>
                {zwolleFiles.some(f => f.role === 'customers') && zwolleFiles.some(f => f.role === 'products') ? (
                  <div className="mt-1 p-2 bg-emerald-50 border border-emerald-200 rounded-xl text-left flex items-center gap-1.5 text-emerald-800 text-[10px] font-medium">
                    <Link2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span><strong>Koppeling actief:</strong> Klanten & Producten gematcht op <em>"Klant ref."</em></span>
                  </div>
                ) : (
                  <div className="mt-1 p-2 bg-indigo-50/70 border border-indigo-100 rounded-xl text-left flex items-center gap-1.5 text-indigo-700 text-[10px]">
                    <Info className="w-3.5 h-3.5 text-indigo-500 shrink-0" />
                    <span>Upload 1 klantenbestand + 1 productenbestand om te koppelen op <em>"Klant ref."</em></span>
                  </div>
                )}
              </div>
            ) : (
              <div className="text-[10px] bg-slate-100 px-3 py-1 rounded-full text-slate-400 font-mono border border-slate-200/60 flex items-center gap-1.5">
                <span>CSV / XLSX / XLS (Klanten + Producten)</span>
              </div>
            )}
          </div>
        </section>

        {/* Dashboard Status / Metric Row */}
        <section id="stats-summary" className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 flex items-center gap-4 shadow-sm hover:shadow-md transition-shadow">
            <div className="p-3 bg-slate-50 text-slate-700 rounded-xl">
              <Users className="w-5 h-5 text-indigo-600" />
            </div>
            <div>
              <span className="text-[11px] text-slate-400 block uppercase tracking-wider font-semibold font-mono">Totaal Geüpload</span>
              <strong className="text-2xl font-bold font-mono text-slate-900">{ommenRows.length + zwolleRows.length}</strong>
              <span className="text-[10px] text-slate-500 block font-mono mt-0.5">
                Ommen: {ommenRows.length} • Zwolle: {zwolleRows.length}
              </span>
            </div>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-slate-200 flex items-center gap-4 shadow-sm hover:shadow-md transition-shadow">
            <div className="p-3 bg-slate-50 text-slate-700 rounded-xl">
              <CheckCircle2 className="w-5 h-5 text-emerald-600" />
            </div>
            <div>
              <span className="text-[11px] text-slate-400 block uppercase tracking-wider font-semibold font-mono">Uniek (Gefilterd)</span>
              <strong className="text-2xl font-bold font-mono text-slate-900">{processedData.all.length}</strong>
              <span className="text-[10px] text-slate-500 block font-mono mt-0.5">
                Klaar voor export
              </span>
            </div>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-amber-200/70 bg-gradient-to-br from-white to-amber-50/20 flex items-center gap-4 shadow-sm hover:shadow-md transition-shadow">
            <div className="p-3 bg-amber-100 text-amber-700 rounded-xl">
              <DollarSign className="w-5 h-5" />
            </div>
            <div>
              <span className="text-[11px] text-amber-700 block uppercase tracking-wider font-semibold font-mono">
                High Value (≥ €400)
              </span>
              <strong className="text-2xl font-bold font-mono text-amber-900">{processedData.highValue.length}</strong>
              <span className="text-[10px] text-amber-700/80 block font-mono mt-0.5">
                Ommen: {processedData.highValueOmmen.length} • Zwolle: {processedData.highValueZwolle.length}
              </span>
            </div>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-slate-200 flex items-center gap-4 shadow-sm hover:shadow-md transition-shadow">
            <div className="p-3 bg-slate-50 text-slate-700 rounded-xl">
              <SlidersHorizontal className="w-5 h-5 text-indigo-500" />
            </div>
            <div>
              <span className="text-[11px] text-slate-400 block uppercase tracking-wider font-semibold font-mono">Gededupliceerd</span>
              <strong className="text-2xl font-bold font-mono text-slate-900">{processedData.duplicates.length}</strong>
              <span className="text-[10px] text-slate-500 block font-mono mt-0.5">
                Dubbele historie samengevoegd
              </span>
            </div>
          </div>
        </section>

        {/* Section 2: Interactive Filter Controls & Configuration */}
        <section id="processing-configurations" className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-5 flex flex-col gap-4">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Settings className="w-4 h-4 text-indigo-650" />
              Verwerkings- & Waardefilters
            </h3>
            <p className="text-xs text-slate-500 leading-relaxed">
              Verfijn hoe de records uit Ommen en Zwolle worden gecombineerd en opgeschoond. Eventuele wijzigingen zijn direct zichtbaar in de onderstaande data-preview.
            </p>

            <div className="flex flex-col gap-3 mt-0.5">
              <label className="flex items-start gap-3 cursor-pointer select-none text-slate-650 hover:text-slate-900 transition-colors">
                <input 
                  type="checkbox" 
                  checked={deduplicate}
                  onChange={(e) => setDeduplicate(e.target.checked)}
                  className="w-4 h-4 mt-0.5 accent-indigo-600 rounded border-slate-350 focus:ring-indigo-505 bg-white"
                />
                <div className="text-xs">
                  <span className="font-semibold block text-slate-900">Dedupliceren activeren</span>
                  Klantgroepen samenvoegen, nieuwste contactgegevens gebruiken en hoogste geregistreerde waarde behouden.
                </div>
              </label>

              <label className="flex items-start gap-3 cursor-pointer select-none text-slate-650 hover:text-slate-900 transition-colors">
                <input 
                  type="checkbox" 
                  checked={tariffFilter}
                  onChange={(e) => setTariffFilter(e.target.checked)}
                  className="w-4 h-4 mt-0.5 accent-indigo-600 rounded border-slate-355 focus:ring-indigo-505 bg-white"
                />
                <div className="text-xs">
                  <span className="font-semibold block text-slate-900">Verwijder records zonder tarief</span>
                  Optionele filter om records zonder tarief of met waarde 0 te negeren in de volledige lijst (standaard uitgeschakeld voor volledige klanthistorie).
                </div>
              </label>

              <label className="flex items-start gap-3 cursor-pointer select-none text-slate-650 hover:text-slate-900 transition-colors">
                <input 
                  type="checkbox" 
                  checked={noQuotes}
                  onChange={(e) => setNoQuotes(e.target.checked)}
                  className="w-4 h-4 mt-0.5 accent-indigo-605 rounded border-slate-355 focus:ring-indigo-505 bg-white"
                />
                <div className="text-xs">
                  <span className="font-semibold block text-slate-900">Geen aanhalingstekens (Integers compatibel)</span>
                  Schrijf getallen en telefoonnummers zonder quotes (") in de CSV, zodat ze als integers/schone tekst registreren.
                </div>
              </label>
            </div>
          </div>

          <div className="lg:col-span-7 flex flex-col justify-between border-t lg:border-t-0 lg:border-l border-slate-205 pt-6 lg:pt-0 lg:pl-6 gap-5">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Download className="w-4 h-4 text-emerald-600" />
                Export- en Bestandsnamen
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Pas de namen van de gegenereerde outputbestanden aan naar wens. Vandaag's datum wordt standaard gebruikt.
              </p>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5 mt-3">
                <div>
                  <label className="block text-[11px] uppercase text-slate-450 font-semibold mb-1 font-mono">Volledige Lijst</label>
                  <div className="relative">
                    <input 
                      type="text"
                      value={allCustomersFilename}
                      onChange={(e) => setAllCustomersFilename(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500" 
                    />
                    <span className="absolute right-3 top-2 text-[10px] text-slate-400 font-mono">.csv</span>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] uppercase text-slate-455 font-semibold mb-1 font-mono">High Value (Totaal)</label>
                  <div className="relative">
                    <input 
                      type="text"
                      value={highValueCustomersFilename}
                      onChange={(e) => setHighValueCustomersFilename(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-505" 
                    />
                    <span className="absolute right-3 top-2 text-[10px] text-slate-400 font-mono">.csv</span>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] uppercase text-emerald-700 font-semibold mb-1 font-mono">High Value Ommen</label>
                  <div className="relative">
                    <input 
                      type="text"
                      value={highValueOmmenFilename}
                      onChange={(e) => setHighValueOmmenFilename(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-505" 
                    />
                    <span className="absolute right-3 top-2 text-[10px] text-slate-400 font-mono">.csv</span>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] uppercase text-blue-700 font-semibold mb-1 font-mono">High Value Zwolle</label>
                  <div className="relative">
                    <input 
                      type="text"
                      value={highValueZwolleFilename}
                      onChange={(e) => setHighValueZwolleFilename(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-505" 
                    />
                    <span className="absolute right-3 top-2 text-[10px] text-slate-400 font-mono">.csv</span>
                  </div>
                </div>
              </div>
            </div>

            {/* Direct Multi-Download Controls */}
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-col gap-3">
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-indigo-600 shrink-0" />
                  <span className="text-xs font-bold text-slate-900">Directe Exportknoppen</span>
                </div>
                <span className="text-[11px] text-slate-500 font-mono">
                  {processedData.all.length} uniek • {processedData.highValue.length} high value (≥ €400)
                </span>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {/* Volledige lijst */}
                <button
                  type="button"
                  disabled={processedData.all.length === 0}
                  onClick={() => downloadCSV(processedData.all, allCustomersFilename)}
                  className="flex items-center justify-center gap-1.5 bg-white hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed text-slate-700 hover:text-slate-950 font-semibold text-xs px-3.5 py-2.5 rounded-xl transition-all border border-slate-250 shadow-xs cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5 text-slate-500" />
                  Volledige Lijst ({processedData.all.length})
                </button>

                {/* High value combined */}
                <button
                  type="button"
                  disabled={processedData.highValue.length === 0}
                  onClick={() => downloadCSV(processedData.highValue, highValueCustomersFilename)}
                  className="flex items-center justify-center gap-1.5 bg-amber-600 hover:bg-amber-500 disabled:bg-slate-200 disabled:text-slate-400 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold text-xs px-3.5 py-2.5 rounded-xl transition-all shadow-sm shadow-amber-200 cursor-pointer"
                >
                  <Sparkles className="w-3.5 h-3.5 text-white shrink-0" />
                  High Value Totaal ({processedData.highValue.length})
                </button>

                {/* High value Ommen */}
                <button
                  type="button"
                  disabled={processedData.highValueOmmen.length === 0}
                  onClick={() => downloadCSV(processedData.highValueOmmen, highValueOmmenFilename)}
                  className="flex items-center justify-center gap-1.5 bg-white hover:bg-emerald-50 disabled:opacity-40 disabled:cursor-not-allowed text-emerald-800 font-semibold text-xs px-3 py-2.5 rounded-xl transition-all border border-emerald-300 cursor-pointer"
                >
                  <MapPin className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                  High Value Ommen ({processedData.highValueOmmen.length})
                </button>

                {/* High value Zwolle */}
                <button
                  type="button"
                  disabled={processedData.highValueZwolle.length === 0}
                  onClick={() => downloadCSV(processedData.highValueZwolle, highValueZwolleFilename)}
                  className="flex items-center justify-center gap-1.5 bg-white hover:bg-blue-50 disabled:opacity-40 disabled:cursor-not-allowed text-blue-800 font-semibold text-xs px-3 py-2.5 rounded-xl transition-all border border-blue-300 cursor-pointer"
                >
                  <MapPin className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                  High Value Zwolle ({processedData.highValueZwolle.length})
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Section 2B: Product & Location Export Section */}
        <section id="product-exports-section" className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col gap-5">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Package className="w-5 h-5 text-indigo-600" />
                Product & Locatie Exports (Ommen / Zwolle)
              </h3>
              <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                Records worden automatisch gegroepeerd per productcategorie op basis van het geüploade bestand (Ommen of Zwolle).
              </p>
            </div>

            <button
              type="button"
              disabled={totalProductExportsCount === 0}
              onClick={downloadAllProductCSVs}
              className="flex items-center justify-center gap-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs px-4 py-2.5 rounded-xl border border-indigo-200/80 transition-all self-start sm:self-auto shrink-0 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm cursor-pointer"
            >
              <FolderDown className="w-4 h-4 text-indigo-600" />
              Download Alle Product Exports ({totalProductExportsCount})
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {PRODUCT_CATEGORIES.map(cat => {
              const ommenRows = processedData.productExports[cat.id]?.ommen || [];
              const zwolleRows = processedData.productExports[cat.id]?.zwolle || [];
              const allCatRows = processedData.productExports[cat.id]?.all || [];
              const ommenFilename = `${cat.name} - Ommen - ${todayStr}`;
              const zwolleFilename = `${cat.name} - Zwolle - ${todayStr}`;
              const allCatFilename = `${cat.name} - Totaal - ${todayStr}`;
              const totalCatCount = allCatRows.length;

              return (
                <div key={cat.id} className="bg-slate-50/70 border border-slate-200 rounded-xl p-4 flex flex-col justify-between gap-4 hover:border-indigo-200 transition-all">
                  <div>
                    <div className="flex items-center justify-between">
                      <h4 className="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                        <Tag className="w-3.5 h-3.5 text-indigo-600" />
                        {cat.name}
                      </h4>
                      <span className="text-[10px] bg-indigo-50 text-indigo-700 font-mono font-bold px-2 py-0.5 rounded-full border border-indigo-100">
                        {totalCatCount} totaal
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-500 mt-1.5 leading-snug">
                      {cat.description}
                    </p>
                  </div>

                  <div className="flex flex-col gap-2 pt-3 border-t border-slate-200/60">
                    {/* Ommen Export Button */}
                    <button
                      type="button"
                      disabled={ommenRows.length === 0}
                      onClick={() => downloadCSV(ommenRows, ommenFilename)}
                      className="w-full flex items-center justify-between bg-white hover:bg-emerald-50 border border-slate-200 hover:border-emerald-300 text-slate-800 disabled:opacity-40 disabled:bg-slate-100 disabled:cursor-not-allowed px-3 py-2 rounded-lg text-xs font-semibold transition-all group cursor-pointer"
                      title={ommenFilename}
                    >
                      <span className="flex items-center gap-1.5 truncate">
                        <MapPin className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                        <span className="truncate">{cat.name} - Ommen</span>
                      </span>
                      <span className="text-[10px] font-mono bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded-full ml-1 shrink-0">
                        {ommenRows.length}
                      </span>
                    </button>

                    {/* Zwolle Export Button */}
                    <button
                      type="button"
                      disabled={zwolleRows.length === 0}
                      onClick={() => downloadCSV(zwolleRows, zwolleFilename)}
                      className="w-full flex items-center justify-between bg-white hover:bg-blue-50 border border-slate-200 hover:border-blue-300 text-slate-800 disabled:opacity-40 disabled:bg-slate-100 disabled:cursor-not-allowed px-3 py-2 rounded-lg text-xs font-semibold transition-all group cursor-pointer"
                      title={zwolleFilename}
                    >
                      <span className="flex items-center gap-1.5 truncate">
                        <MapPin className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                        <span className="truncate">{cat.name} - Zwolle</span>
                      </span>
                      <span className="text-[10px] font-mono bg-blue-100 text-blue-800 font-bold px-2 py-0.5 rounded-full ml-1 shrink-0">
                        {zwolleRows.length}
                      </span>
                    </button>

                    {/* Totaal Export Button */}
                    {allCatRows.length > 0 && (
                      <button
                        type="button"
                        onClick={() => downloadCSV(allCatRows, allCatFilename)}
                        className="w-full flex items-center justify-between bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 px-3 py-1.5 rounded-lg text-[11px] font-semibold transition-all cursor-pointer"
                        title={allCatFilename}
                      >
                        <span className="flex items-center gap-1.5 truncate text-slate-600">
                          <Download className="w-3 h-3 text-slate-500 shrink-0" />
                          <span className="truncate">{cat.name} - Beide Locaties</span>
                        </span>
                        <span className="text-[10px] font-mono bg-slate-200 text-slate-800 font-bold px-1.5 py-0.2 rounded-full ml-1 shrink-0">
                          {allCatRows.length}
                        </span>
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </section>

        {/* Section 3: Detailed Interactive Result View & Logs */}
        { (ommenRows.length > 0 || zwolleRows.length > 0) ? (
          <section id="result-table" className="bg-white rounded-2xl border border-slate-200 shadow-sm flex flex-col overflow-hidden">
            {/* Tabs Selector */}
            <div className="border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/50">
              <div className="flex flex-wrap bg-slate-100 p-1 rounded-xl border border-slate-200 gap-1 self-start">
                <button
                  onClick={() => setActiveTab('all')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    activeTab === 'all' 
                      ? 'bg-white text-indigo-600 shadow-sm' 
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  Alle Klanten ({processedData.all.length})
                </button>
                <button
                  onClick={() => setActiveTab('high_value')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    activeTab === 'high_value' 
                      ? 'bg-white text-indigo-600 shadow-sm' 
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  High Value (≥ €400) ({processedData.highValue.length})
                </button>
                {PRODUCT_CATEGORIES.map(cat => {
                  const count = processedData.productExports[cat.id]?.all?.length || 0;
                  return (
                    <button
                      key={cat.id}
                      onClick={() => setActiveTab(`prod_${cat.id}`)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                        activeTab === `prod_${cat.id}` 
                          ? 'bg-white text-indigo-600 shadow-sm' 
                          : 'text-slate-500 hover:text-slate-900'
                      }`}
                    >
                      {cat.name} ({count})
                    </button>
                  );
                })}
                <button
                  onClick={() => setActiveTab('duplicates')}
                  className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    activeTab === 'duplicates' 
                      ? 'bg-white text-indigo-600 shadow-sm' 
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  Gededupliceerd Log ({processedData.duplicates.length})
                </button>
              </div>

              <div className="flex items-center gap-2">
                <span className="text-[10px] text-slate-500 bg-slate-100 px-2.5 py-1 rounded-md font-mono border border-slate-200">
                  Formaat: email;phone;fn;ln;zip;ct;country;dob;doby;gen;age;value
                </span>
              </div>
            </div>

            {/* TAB PREVIEWS */}
            <div className="overflow-x-auto">
              {(activeTab === 'all' || activeTab.startsWith('prod_')) && (() => {
                const isCat = activeTab.startsWith('prod_');
                const catId = isCat ? activeTab.replace('prod_', '') : '';
                const activeCatDef = PRODUCT_CATEGORIES.find(c => c.id === catId);
                const displayRows = isCat 
                  ? (processedData.productExports[catId]?.all || [])
                  : processedData.all;

                return (
                  <table className="w-full text-left text-xs text-slate-600 border-collapse">
                    <thead className="bg-slate-50 border-b border-slate-200 text-[10px] uppercase font-mono text-slate-500 font-bold">
                      <tr>
                        <th className="py-3 px-4">E-mailadres</th>
                        <th className="py-3 px-4">Klant ref.</th>
                        <th className="py-3 px-4">Telefoonnummer</th>
                        <th className="py-3 px-4">Naam</th>
                        <th className="py-3 px-4">Postcode / Stad</th>
                        <th className="py-3 px-4">Geslacht</th>
                        <th className="py-3 px-4">Geboortedatum (Leeftijd)</th>
                        {isCat && <th className="py-3 px-4">Product / Lidmaatschap</th>}
                        <th className="py-3 px-4 text-right">Tarief / Waarde</th>
                        <th className="py-3 px-4">Bron</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono">
                      {displayRows.slice(0, 100).map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                          <td className="py-3 px-4 text-slate-900 font-semibold truncate max-w-xs">
                            {row.email || <span className="text-slate-400 italic font-normal">Geen e-mail</span>}
                          </td>
                          <td className="py-3 px-4 text-slate-500 text-[11px] font-mono">
                            {row._klantRef ? <span className="bg-slate-100 px-1.5 py-0.5 rounded text-slate-700 font-bold">{row._klantRef}</span> : '-'}
                          </td>
                          <td className="py-3 px-4 text-indigo-600 font-bold">{row.phone || '-'}</td>
                          <td className="py-3 px-4">{row.fn} {row.ln}</td>
                          <td className="py-3 px-4">{row.zip} {row.ct}</td>
                          <td className="py-3 px-4">
                            <span className={`px-2 py-0.5 rounded-full text-[10px] ${
                              row.gen === 'M' ? 'bg-indigo-50 text-indigo-600' : 'bg-pink-50 text-pink-650'
                            }`}>
                              {row.gen || '-'}
                            </span>
                          </td>
                          <td className="py-3 px-4">{row.dob ? `${row.dob} (${row.age})` : '-'}</td>
                          {isCat && (
                            <td className="py-3 px-4 text-slate-700 font-medium truncate max-w-xs">
                              {row.product || 'Onbekend'}
                            </td>
                          )}
                          <td className="py-3 px-4 text-right font-bold text-amber-600">
                            {row.value !== "" ? `€${Number(row.value).toFixed(2)}` : 'Geen'}
                          </td>
                          <td className="py-3 px-4 text-[10px]">
                            <span className={`px-2 py-0.5 rounded-full font-bold ${
                              row._source === 'Ommen' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-blue-50 text-blue-700 border border-blue-105'
                            }`}>
                              {row._source}
                            </span>
                          </td>
                        </tr>
                      ))}
                      {displayRows.length === 0 && (
                        <tr>
                          <td colSpan={isCat ? 10 : 9} className="py-8 text-center text-slate-400">
                            Geen records gevonden {isCat ? `voor de categorie "${activeCatDef?.name}"` : 'na filtering'}.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                );
              })()}

              {activeTab === 'high_value' && (
                <table className="w-full text-left text-xs text-slate-600 border-collapse">
                  <thead className="bg-slate-50 border-b border-slate-200 text-[10px] uppercase font-mono text-slate-500 font-bold">
                    <tr>
                      <th className="py-3 px-4">E-mailadres</th>
                      <th className="py-3 px-4">Klant ref.</th>
                      <th className="py-3 px-4">Telefoonnummer</th>
                      <th className="py-3 px-4">Naam</th>
                      <th className="py-3 px-4">Postcode / Stad</th>
                      <th className="py-3 px-4">Geslacht</th>
                      <th className="py-3 px-4">Geboortedatum (Leeftijd)</th>
                      <th className="py-3 px-4 text-right">Waarde (≥ €400)</th>
                      <th className="py-3 px-4">Bron</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {processedData.highValue.slice(0, 100).map((row, idx) => (
                      <tr key={idx} className="hover:bg-amber-50/20 transition-colors">
                        <td className="py-3 px-4 text-slate-900 font-semibold truncate max-w-xs">
                          {row.email || <span className="text-slate-400 italic font-normal">Geen e-mail</span>}
                        </td>
                        <td className="py-3 px-4 text-slate-500 text-[11px] font-mono">
                          {row._klantRef ? <span className="bg-slate-100 px-1.5 py-0.5 rounded text-slate-700 font-bold">{row._klantRef}</span> : '-'}
                        </td>
                        <td className="py-3 px-4 text-indigo-600 font-bold">{row.phone || '-'}</td>
                        <td className="py-3 px-4">{row.fn} {row.ln}</td>
                        <td className="py-3 px-4">{row.zip} {row.ct}</td>
                        <td className="py-3 px-4">
                          <span className={`px-2 py-0.5 rounded-full text-[10px] ${
                            row.gen === 'M' ? 'bg-indigo-50 text-indigo-600' : 'bg-pink-50 text-pink-650'
                          }`}>
                            {row.gen || '-'}
                          </span>
                        </td>
                        <td className="py-3 px-4">{row.dob ? `${row.dob} (${row.age})` : '-'}</td>
                        <td className="py-3 px-4 text-right font-black text-amber-700 bg-amber-50/40">
                          <div>€{Number(row.value).toFixed(2)}</div>
                          {row._maxHistoricalValue && row._maxHistoricalValue !== Number(row.value) ? (
                            <div className="text-[9px] text-amber-600 font-normal">
                              Max: €{row._maxHistoricalValue.toFixed(2)}
                            </div>
                          ) : null}
                        </td>
                        <td className="py-3 px-4 text-[10px]">
                          <span className={`px-2 py-0.5 rounded-full font-bold ${
                            row._source === 'Ommen' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-blue-50 text-blue-700 border border-blue-105'
                          }`}>
                            {row._source}
                          </span>
                        </td>
                      </tr>
                    ))}
                    {processedData.highValue.length === 0 && (
                      <tr>
                        <td colSpan={9} className="py-8 text-center text-slate-400">
                          Geen High Value klanten met een tarief/waarde ≥ €400.00 gevonden.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              )}

              {activeTab === 'duplicates' && (
                <div className="p-6 flex flex-col gap-4">
                  <div className="flex items-center gap-2 p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-500">
                    <Info className="w-4 h-4 text-indigo-600 shrink-0" />
                    <span>
                      Deduplicatie houdt rekening met identieke e-mails (of namen) en behoudt automatisch de rij met de meest recente <strong>"Actief sinds"</strong> datum.
                    </span>
                  </div>

                  <div className="flex flex-col gap-4">
                    {processedData.duplicates.slice(0, 20).map((log, lIdx) => (
                      <div key={lIdx} className="border border-slate-200 bg-slate-50/30 rounded-xl p-4 flex flex-col gap-3">
                        <div className="flex justify-between items-center">
                          <span className="text-slate-800 font-bold text-xs font-mono">{log.key}</span>
                          <span className="text-[10px] text-indigo-600 bg-indigo-50 font-mono px-2 py-0.5 rounded border border-indigo-100">
                            {log.records.length} combinaties gevonden
                          </span>
                        </div>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-1 text-[11px] font-mono">
                          {log.records.map((rec, rIdx) => {
                            const isSelected = rec === log.selected;
                            return (
                              <div 
                                key={rIdx} 
                                className={`p-2.5 rounded-lg border ${
                                  isSelected 
                                    ? 'bg-emerald-50/50 border-emerald-300 text-slate-800' 
                                    : 'bg-white border-slate-200 text-slate-500'
                                }`}
                              >
                                <div className="flex justify-between font-bold text-[10px] mb-1">
                                  <span>Bron: {rec._source}</span>
                                  {isSelected && <span className="text-emerald-750 font-bold">✓ Geselecteerd (meest recent)</span>}
                                </div>
                                <div>Naam: {rec.fn} {rec.ln}</div>
                                <div>Actief Sinds: {rec._raw["Actief sinds"] || 'Niet ingevuld'}</div>
                                <div>Tarief: {rec.value !== "" ? `€${rec.value}` : 'Nul/Leeg'}</div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                    {processedData.duplicates.length === 0 && (
                      <div className="py-8 text-center text-slate-400 text-xs">
                        Geen dubbele records gedetecteerd.
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

            { (activeTab === 'all' && processedData.all.length > 100) && (
              <div className="border-t border-slate-200 px-6 py-3 text-center text-xs text-slate-400 bg-slate-50/50">
                Toont de eerste 100 rijen. Het gedownloade bestand zal alle {processedData.all.length} records bevatten.
              </div>
            )}
          </section>
        ) : (
          /* Empty State Guidelines */
          <section id="empty-state" className="bg-white rounded-2xl border-2 border-dashed border-slate-300 p-12 text-center flex flex-col items-center justify-center gap-4 shadow-sm">
            <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 text-slate-400 animate-pulse">
              <Upload className="w-10 h-10 stroke-[1.5]" />
            </div>
            <div className="max-w-md">
              <h2 className="text-lg font-bold text-slate-900">Geen bestanden geüpload</h2>
              <p className="text-sm text-slate-500 mt-1 leading-relaxed">
                Upload ten minste één CSV-bestand van vestiging Ommen of Zwolle. De tool combineert ze direct, filtert op unieke klanten en formatteert de telefoonnummers perfect naar internationale standaard (+31).
              </p>
            </div>
          </section>
        )}
      </main>

      {/* Footer */}
      <footer id="app-footer" className="border-t border-slate-200 bg-white py-6 text-center text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© 2026 CSV Converter. Alle data wordt uitsluitend in de browser verwerkt; er worden geen gegevens opgeslagen.</p>
          <div className="flex gap-4">
            <span className="hover:text-slate-600 font-mono text-[10px]">Deduplicatie: Actief sinds</span>
            <span className="hover:text-slate-600 font-mono text-[10px]">Landcodes: +31 & 05</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
