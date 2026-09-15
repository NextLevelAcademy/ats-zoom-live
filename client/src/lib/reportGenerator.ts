import type {
  CountryGroup,
  OptInRow,
  OldStudentExclusionRow,
  ReportData,
  SessionDetails,
  ShowUpMergeRow,
  ShowUpRegRow,
  SignUpRow,
} from "../../../shared/schema";
import { parseCsv, parseCsvSkipRows, parseCsvAutoHeader, getVal } from "./csvParse";
import { deriveMetrics } from "./deriveMetrics";
import * as XLSX from "xlsx";

// ============ Helpers ============

function digits(s: string | null | undefined): string {
  return (s || "").replace(/\D/g, "");
}

/**
 * Parse a Keap / Zoom-style phone like:
 *   `'+65 87661449   `   → cc=65, phone=87661449
 *   `'+65 8915 9826   `  → cc=65, phone=89159826
 *   `81273350   `        → cc="", phone=81273350
 *   `60122974700   `     → cc="60", phone=22974700  (if starts with 60 and >=10 digits)
 *   `'+656594831160   `  → cc="65", phone=6594831160 (no space; cc inferred from prefix)
 */
function parseApostrophePhone(raw: string): { cc: string; local: string } {
  if (!raw) return { cc: "", local: "" };
  // Strip leading apostrophe (Excel text marker) and trailing whitespace/notes.
  let s = raw.replace(/^'+/, "").trim();
  // Remove trailing "(Work)" or similar annotations
  s = s.replace(/\s*\(.*\)\s*$/, "").trim();

  if (s.startsWith("+")) {
    // Format: "+65 87661449" or "+656594831160" (no space)
    const afterPlus = s.slice(1).trim();
    const firstSpace = afterPlus.indexOf(" ");
    if (firstSpace > 0) {
      const cc = digits(afterPlus.slice(0, firstSpace));
      const local = digits(afterPlus.slice(firstSpace + 1));
      return { cc, local };
    }
    // No space — try to split a known country code prefix
    const d = digits(afterPlus);
    if (d.startsWith("65") && d.length >= 10) return { cc: "65", local: d.slice(2) };
    if (d.startsWith("60") && d.length >= 11) return { cc: "60", local: d.slice(2) };
    if (d.startsWith("852") && d.length >= 11) return { cc: "852", local: d.slice(3) };
    if (d.startsWith("1") && d.length >= 11) return { cc: "1", local: d.slice(1) };
    // Fallback: treat as 1- or 2-digit cc (best-effort)
    if (d.length >= 10) return { cc: d.slice(0, 2), local: d.slice(2) };
    return { cc: "", local: d };
  }
  // No leading +: could be a raw 8-digit SG local, or country-code-included intl
  const d = digits(s);
  if (!d) return { cc: "", local: "" };
  // 8-digit SG local (starts with 6,8,9)
  if (d.length === 8 && /^[689]/.test(d)) return { cc: "65", local: d };
  // Starts with 65, length 10 → SG
  if (d.startsWith("65") && d.length === 10) return { cc: "65", local: d.slice(2) };
  // Starts with 60, length 11-12 → MY
  if (d.startsWith("60") && (d.length === 11 || d.length === 12))
    return { cc: "60", local: d.slice(2) };
  // Starts with 1, length 11 → USA
  if (d.startsWith("1") && d.length === 11) return { cc: "1", local: d.slice(1) };
  // Starts with 852, length 11 → HK
  if (d.startsWith("852") && d.length === 11) return { cc: "852", local: d.slice(3) };
  // Fallback: keep as local, no cc
  return { cc: "", local: d };
}

function buildFullPhone(cc: string, local: string): string {
  if (!cc && !local) return "";
  if (!cc) return local;
  if (!local) return cc;
  return cc + local;
}

function detectCountryFromCc(cc: string): CountryGroup {
  const c = digits(cc);
  if (c === "65") return "SG";
  if (c === "60") return "MY";
  // Any other recognizable country code (USA, HK, etc.) → OTHERS
  return "OTHERS";
}

/**
 * Detect country for a row that has only a raw phone (no cc field).
 * No usable country code → INVALID.
 */
function detectCountry(cc: string, local: string): CountryGroup {
  const c = digits(cc);
  if (!c) return "INVALID";
  return detectCountryFromCc(c);
}

// ============ Parsers ============

/**
 * Keap CRM export (First Name, Phone 1, Email). Used only for the optional
 * Tag 4 List (ATS4) — this app has no Keap-sourced Opt-In at all.
 */
interface KeapRow {
  first: string;
  email: string;
  cc: string;
  local: string;
  fullPhone: string;
  country: CountryGroup;
}

function parseKeapRows(rows: Record<string, any>[]): KeapRow[] {
  return rows
    .map((r) => {
      const first = getVal(r, ["First Name", "FirstName"]);
      const phoneRaw = getVal(r, ["Phone 1", "Phone", "Mobile"]);
      const email = getVal(r, ["Email", "Email Address"]).toLowerCase();
      const { cc, local } = parseApostrophePhone(phoneRaw);
      const country = detectCountry(cc, local);
      return {
        first,
        email,
        cc,
        local,
        fullPhone: buildFullPhone(cc, local),
        country,
      };
    })
    .filter((r) => r.email);
}

/**
 * Flags obviously fake/gibberish opt-in contacts as INVALID, in place.
 * Two signals seen in real spam clusters, both about the *phone number*
 * rather than the name (name-only checks like "name equals email" throw
 * off too many real contacts who genuinely use their own name as their
 * email address):
 *  - The same phone number is reused across 3+ otherwise-unrelated
 *    contacts (a shared placeholder number, not a real one per person).
 *  - The number itself is a placeholder pattern — 6 or more of the same
 *    digit in a row (e.g. "95111111") — which real mobile numbers don't
 *    produce.
 * A row's country stays whatever was detected (SG/MY/etc.) if neither
 * signal fires, so real contacts keep their normal classification. Rows
 * caught here can still be manually reclassified on the report page if
 * this over-flags a genuine contact.
 */
function applyGibberishHeuristic(rows: { fullPhone: string; country: CountryGroup }[]): void {
  const phoneCounts = new Map<string, number>();
  for (const r of rows) {
    if (r.fullPhone) {
      phoneCounts.set(r.fullPhone, (phoneCounts.get(r.fullPhone) || 0) + 1);
    }
  }
  const repeatedDigitRun = /(\d)\1{5,}/; // same digit 6+ times in a row
  for (const r of rows) {
    const sharedPhone = !!r.fullPhone && (phoneCounts.get(r.fullPhone) || 0) >= 3;
    const placeholderPhone = !!r.fullPhone && repeatedDigitRun.test(r.fullPhone);
    if (sharedPhone || placeholderPhone) {
      r.country = "INVALID";
    }
  }
}

interface RegRow {
  first: string;
  last: string;
  email: string;
  cc: string;
  local: string;
  fullPhone: string;
  country: CountryGroup;
}

function parseRegRows(rows: Record<string, any>[]): RegRow[] {
  return rows
    .map((r) => {
      const first = getVal(r, ["First Name", "FirstName"]);
      const last = getVal(r, ["Last Name", "LastName"]);
      const email = getVal(r, ["Email", "Email Address"]).toLowerCase();
      const phoneRaw = getVal(r, ["Phone", "Phone Number"]);
      const { cc, local } = parseApostrophePhone(phoneRaw);
      const country = detectCountry(cc, local);
      return {
        first,
        last,
        email,
        cc,
        local,
        fullPhone: buildFullPhone(cc, local),
        country,
      };
    })
    .filter((r) => r.email);
}

interface PartRow {
  name: string;
  email: string;
  durationMinutes: number;
}

function parsePartRows(rows: Record<string, any>[]): PartRow[] {
  return rows
    .map((r) => {
      const name = getVal(r, [
        "Name (original name)",
        "Name (Original Name)",
        "Name",
        "Original Name",
      ]);
      const email = getVal(r, [
        "Email",
        "User Email",
        "Email Address",
        "User Email Address",
      ]).toLowerCase();
      const dur =
        parseInt(
          getVal(r, ["Total duration (minutes)", "Duration (minutes)", "Total Duration"]),
          10
        ) || 0;
      return { name, email, durationMinutes: dur };
    })
    .filter((r) => r.email);
}

interface TCRow {
  first: string;
  last: string;
  email: string;
  phone: string;
  total: number;
  pricingOption: string;
  packageName: string;
  orderDate: string;
  paymentMethod: string;
}

function parseTCRows(rows: Record<string, any>[]): TCRow[] {
  return rows
    .map((r) => {
      const totalStr = getVal(r, ["total", "Total", "amount"]);
      const total = parseFloat(totalStr.replace(/[^0-9.\-]/g, "")) || 0;
      return {
        first: getVal(r, [
          "customer_first_name",
          "first_name",
          "Customer First Name",
          "First Name",
        ]),
        last: getVal(r, [
          "customer_last_name",
          "last_name",
          "Customer Last Name",
          "Last Name",
        ]),
        email: getVal(r, [
          "customer_email",
          "email",
          "Customer Email",
          "Email",
        ]).toLowerCase(),
        phone: getVal(r, [
          "customer_phone",
          "phone",
          "Phone",
          "telephone",
          "Telephone",
          "customer_telephone",
        ]),
        total,
        pricingOption: getVal(r, [
          "relevant_item_pricing_option",
          "pricing_option",
          "Pricing Option",
        ]),
        packageName: getVal(r, [
          "relevant_item_name",
          "Product",
          "Item Name",
        ]),
        orderDate: getVal(r, ["order_date", "Order Date", "date"]),
        paymentMethod: getVal(r, [
          "payment processor",
          "payment_processor",
          "payment method",
          "payment_method",
          "payment gateway",
          "payment_gateway",
          "gateway",
          "processor",
        ]),
      };
    })
    .filter((r) => r.email);
}

interface BTRow {
  fullName: string;
  email: string;
  phone: string;
  intake: string; // "May", "June" — from the Date column
  price: number; // optional override; falls back to session.programPrice if 0
}

function parseBTRows(rows: Record<string, any>[]): BTRow[] {
  return rows
    .map((r) => {
      const first = getVal(r, ["First Name", "FirstName", "first_name"]);
      const last = getVal(r, ["Last Name", "LastName", "last_name"]);
      const fullNameDirect = getVal(r, [
        "Name",
        "Full Name",
        "FullName",
        "name",
      ]);
      const dateRaw = getVal(r, [
        "Date",
        "date",
        "Intake",
        "intake",
        "Intake Date",
        "Month",
        "month",
        "Date of VW",
        "date of vw",
        "VW Date",
        "VW",
      ]);
      const priceRaw = getVal(r, [
        "Price",
        "price",
        "Amount",
        "amount",
        "Total",
        "total",
        "Amount Paid",
        "amount paid",
        "Paid",
        "paid",
      ]);
      const priceNum = Number(
        String(priceRaw).replace(/[^0-9.\-]/g, "")
      );
      return {
        fullName: fullNameDirect || `${first} ${last}`.trim(),
        email: getVal(r, ["Email", "email", "Email Address"]).toLowerCase(),
        phone: getVal(r, [
          "Phone Number",
          "phone number",
          "Phone",
          "phone",
          "Mobile",
          "telephone",
          "Telephone",
        ]),
        intake: extractIntake(dateRaw),
        price: Number.isFinite(priceNum) ? priceNum : 0,
      };
    })
    .filter((r) => r.email || r.phone);
}

// ============ Main entry ============

export interface UploadedFiles {
  registrationFile: File; // Zoom Registration export — the Opt-In source (no Keap involved)
  participantsFile: File; // Zoom Participants export — the Show Up source
  thriveCartFile: File;
  oldStudentsFile?: File | null; // Optional: Old ATS Students (.xlsx, single column of emails)
  bankTransferFile?: File | null;
  // Optional Tag 4 List (ATS4) export from Keap CRM (First Name, Phone 1,
  // Email). Contacts excluded from the No Show Up broadcast. Unrelated to
  // Opt-In sourcing — this app never uses Keap for Opt-In.
  nlow4File?: File | null;
}

/**
 * Parse the Old ATS Students .xlsx — a single column of emails,
 * no header row. Returns a lowercased Set for O(1) exclusion lookups.
 */
async function parseOldStudentsXlsx(
  file: File | null | undefined
): Promise<Set<string>> {
  if (!file) return new Set<string>();
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const out = new Set<string>();
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, blankrows: false, defval: "" });
    for (const row of rows) {
      for (const cell of row) {
        const s = String(cell ?? "").trim().toLowerCase();
        if (s && s.includes("@")) out.add(s);
      }
    }
  }
  return out;
}

/**
 * Parse Bank Transfer Sales — accepts both .csv and .xlsx. The xlsx branch
 * reads the first sheet using the first row as headers, producing the same
 * record shape as parseCsv so parseBTRows works unchanged.
 */
async function parseBankTransfer(file: File): Promise<Record<string, any>[]> {
  const isXlsx = /\.xlsx$/i.test(file.name);
  if (!isXlsx) {
    return parseCsv<Record<string, any>>(file);
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sn = wb.SheetNames[0];
  if (!sn) return [];
  const ws = wb.Sheets[sn];
  const rows = XLSX.utils.sheet_to_json<Record<string, any>>(ws, {
    defval: "",
    raw: false, // stringify numbers/dates like the CSV parser does
  });
  return rows;
}

const INTAKE_MONTHS = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

export function extractIntake(s: string): string {
  if (!s) return "";
  const lower = s.toLowerCase();
  for (const m of INTAKE_MONTHS) {
    if (lower.includes(m)) {
      return m.charAt(0).toUpperCase() + m.slice(1);
    }
  }
  return "";
}

export async function generateReport(
  files: UploadedFiles,
  session: SessionDetails
): Promise<ReportData> {
  const [regRaw, partRaw, tcRaw, btRaw, oldStudentsEmails, nlow4Raw] = await Promise.all([
    parseCsvAutoHeader<Record<string, any>>(
      files.registrationFile,
      ["first name", "email"],
      5
    ),
    parseCsvAutoHeader<Record<string, any>>(
      files.participantsFile,
      ["name", "email"],
      3
    ),
    parseCsv<Record<string, any>>(files.thriveCartFile),
    files.bankTransferFile
      ? parseBankTransfer(files.bankTransferFile)
      : Promise.resolve([]),
    parseOldStudentsXlsx(files.oldStudentsFile),
    files.nlow4File
      ? parseCsv<Record<string, any>>(files.nlow4File)
      : Promise.resolve([]),
  ]);

  // Helper: is this email an Old ATS Student? (case-insensitive)
  const isOldStudent = (email: string | null | undefined): boolean => {
    if (!email) return false;
    return oldStudentsEmails.has(email.toLowerCase());
  };

  const reg = parseRegRows(regRaw);
  applyGibberishHeuristic(reg);
  const part = parsePartRows(partRaw);
  const tc = parseTCRows(tcRaw);
  const bt = parseBTRows(btRaw);
  // Tag 4 List (ATS4) — Keap CRM tag-list export, unrelated to Opt-In sourcing.
  const nlow4Rows = parseKeapRows(nlow4Raw);
  const nlow4ExcludedPhones: string[] = [];
  const nlow4ExcludedEmails: string[] = [];
  for (const r of nlow4Rows) {
    if (r.fullPhone) nlow4ExcludedPhones.push(r.fullPhone.replace(/\D/g, ""));
    if (r.email) nlow4ExcludedEmails.push(r.email.toLowerCase());
  }

  const regByEmail = new Map<string, RegRow>();
  for (const r of reg) regByEmail.set(r.email, r);
  // Phone-based lookup: handles people who use a different email on
  // ThriveCart/Bank Transfer than the one they registered with.
  const regByPhone = new Map<string, RegRow>();
  for (const r of reg) {
    if (r.fullPhone) regByPhone.set(r.fullPhone, r);
  }

  // Sum each participant's duration across all their join segments (Zoom
  // logs a separate row per rejoin), then require MORE than 10 minutes
  // total in the room to count as a show-up — someone who joined and left
  // within 10 minutes is treated as a no-show, matching the reference.
  const partDedupMap = new Map<string, PartRow & { totalDuration: number }>();
  for (const p of part) {
    if (!p.email) continue;
    const ex = partDedupMap.get(p.email);
    if (ex) {
      ex.totalDuration = (ex.totalDuration || 0) + (p.durationMinutes || 0);
    } else {
      partDedupMap.set(p.email, { ...p, totalDuration: p.durationMinutes || 0 });
    }
  }
  const MIN_SHOWUP_MINUTES = 10;
  const partDedup = Array.from(partDedupMap.values()).filter(
    (p) => p.totalDuration > MIN_SHOWUP_MINUTES
  );
  const partByEmail = new Map<string, PartRow>();
  for (const p of partDedup) partByEmail.set(p.email, p);

  // ===== Sign-ups (TC + BT) =====
  const signUpRows: SignUpRow[] = [];
  const seenSignupEmails = new Map<string, SignUpRow>();

  function resolvePhoneAndCountry(
    email: string,
    fallbackPhone: string
  ): { cc: string; local: string; fullPhone: string; country: CountryGroup } {
    const r = regByEmail.get(email);
    if (r && (r.cc || r.local)) {
      return {
        cc: r.cc,
        local: r.local,
        fullPhone: r.fullPhone,
        country: r.country,
      };
    }
    if (fallbackPhone) {
      const { cc, local } = parseApostrophePhone(fallbackPhone);
      return {
        cc,
        local,
        fullPhone: buildFullPhone(cc, local),
        country: detectCountry(cc, local),
      };
    }
    return { cc: "", local: "", fullPhone: "", country: "INVALID" };
  }

  for (const r of tc) {
    const fullName = `${r.first} ${r.last}`.trim() || r.email;
    const resolved = resolvePhoneAndCountry(r.email, r.phone);
    const showedUp = partByEmail.has(r.email);
    // Match Opt-In by email first, then by phone (same person, different
    // email between Opt-In and ThriveCart).
    const inOptIn =
      regByEmail.has(r.email) ||
      (!!resolved.fullPhone && regByPhone.has(resolved.fullPhone));
    const signUp: SignUpRow = {
      fullName,
      email: r.email,
      countryCode: resolved.cc,
      phoneNumber: resolved.local,
      fullPhone: resolved.fullPhone,
      country: resolved.country,
      source: "ThriveCart",
      paymentMethod: r.paymentMethod,
      pricingOption: r.packageName || r.pricingOption,
      intake: extractIntake(`${r.packageName || ""} ${r.pricingOption || ""} ${r.orderDate || ""}`),
      total: r.total,
      orderDate: r.orderDate,
      showedUp,
      inOptIn,
    };
    seenSignupEmails.set(r.email, signUp);
    signUpRows.push(signUp);
  }

  for (const r of bt) {
    const existing = r.email ? seenSignupEmails.get(r.email) : undefined;
    if (existing) {
      existing.source = "ThriveCart+BT";
      // Preserve BT row's intake / price if the TC row didn't have them
      if (!existing.intake && r.intake) existing.intake = r.intake;
      if ((!existing.total || existing.total <= 0) && r.price > 0) {
        existing.total = r.price;
      }
      // Prefer the BT row's name when the TC row's name was missing or
      // just an email fallback. The PayNow file always has an explicit Name
      // column, so its value is usually more authoritative.
      const tcNameMissing =
        !existing.fullName || existing.fullName === existing.email;
      if (tcNameMissing && r.fullName) {
        existing.fullName = r.fullName;
      }
      continue;
    }
    const resolved = resolvePhoneAndCountry(r.email, r.phone);
    const showedUp = !!(r.email && partByEmail.has(r.email));
    const inOptIn =
      (!!r.email && regByEmail.has(r.email)) ||
      (!!resolved.fullPhone && regByPhone.has(resolved.fullPhone));
    const signUp: SignUpRow = {
      fullName: r.fullName || r.email,
      email: r.email,
      countryCode: resolved.cc,
      phoneNumber: resolved.local,
      fullPhone: resolved.fullPhone,
      country: resolved.country,
      source: "BT",
      paymentMethod: "PayNow",
      pricingOption: "",
      intake: r.intake || "",
      total: r.price > 0 ? r.price : session.programPrice,
      orderDate: "",
      showedUp,
      inOptIn,
    };
    if (r.email) seenSignupEmails.set(r.email, signUp);
    signUpRows.push(signUp);
  }

  const signUpEmails = new Set(
    signUpRows.map((s) => s.email).filter(Boolean)
  );
  // Bank Transfer sign-ups often carry no email (or a different one than the
  // person opted in / showed up with) — match those by phone too, so a
  // BT-paying attendee is still recognized as signed up.
  const signUpPhones = new Set(
    signUpRows.map((s) => (s.fullPhone || "").replace(/\D/g, "")).filter(Boolean)
  );
  const isSignedUp = (email: string, fullPhone: string): boolean => {
    if (email && signUpEmails.has(email)) return true;
    const p = (fullPhone || "").replace(/\D/g, "");
    return !!p && signUpPhones.has(p);
  };
  const partEmails = new Set(partDedup.map((p) => p.email));

  // ===== Show up Merge (one row per UNIQUE participant email; sum durations) =====
  const showUpMerge: ShowUpMergeRow[] = partDedup.map((p) => {
    // Attendee data (name/phone/country) comes straight from the Zoom
    // Registration row itself — Opt-In IS the registration list, so there's
    // no separate CRM to cross-reference.
    const reg = regByEmail.get(p.email);
    let name = p.name;
    let cc = "", local = "", fullPhone = "", country: CountryGroup = "INVALID";
    let source: "Registration" | "Unknown" = "Unknown";
    if (reg) {
      source = "Registration";
      name = `${reg.first} ${reg.last}`.trim() || p.name;
      cc = reg.cc;
      local = reg.local;
      fullPhone = reg.fullPhone;
      country = reg.country;
    }
    const signed = isSignedUp(p.email, fullPhone);
    return {
      fullName: name,
      email: p.email,
      countryCode: cc,
      phoneNumber: local,
      fullPhone,
      country,
      durationMinutes: p.totalDuration,
      source,
      signedUp: signed,
      signedUpEmail: signed ? p.email : "",
    };
  });

  // ===== Show up REG (one row per registrant) =====
  const showUpReg: ShowUpRegRow[] = reg.map((r) => ({
    firstName: r.first,
    lastName: r.last,
    fullName: `${r.first} ${r.last}`.trim() || r.email,
    email: r.email,
    countryCode: r.cc,
    phoneNumber: r.local,
    fullPhone: r.fullPhone,
    country: r.country,
    showedUp: partEmails.has(r.email),
    signedUp: isSignedUp(r.email, r.fullPhone),
  }));

  // ===== Opt-Ins (every Zoom registrant) =====
  const optInRows: OptInRow[] = reg.map((r) => ({
    firstName: r.first,
    fullName: `${r.first} ${r.last}`.trim() || r.email,
    email: r.email,
    countryCode: r.cc,
    phoneNumber: r.local,
    fullPhone: r.fullPhone,
    country: r.country,
    showedUp: partEmails.has(r.email),
    signedUp: isSignedUp(r.email, r.fullPhone),
  }));
  // Show-up attendees who never matched a registrant (by email) are auto-added
  // to the Opt-In list, since they clearly attended live — e.g. joined via a
  // forwarded link without registering themselves.
  let showUpAddedToOptInCount = 0;
  for (const s of showUpMerge) {
    if (s.source === "Registration") continue; // already in optInRows by email
    showUpAddedToOptInCount++;
    optInRows.push({
      firstName: s.fullName,
      fullName: s.fullName,
      email: s.email,
      countryCode: s.countryCode,
      phoneNumber: s.phoneNumber,
      fullPhone: s.fullPhone,
      country: s.country,
      showedUp: true,
      signedUp: s.signedUp,
    });
  }

  // ===== Old ATS Students exclusion =====
  // Match by email only (case-insensitive). Excluded from Opt-In + Show Up
  // counts. Sign-ups and revenue are untouched.
  // "Show Up & Opt-In" card = # of Old Students who BOTH showed up AND were
  // present in the Opt-In list (pre-exclusion).
  const optInEmailsPreExclusion = new Set(
    optInRows.map((r) => (r.email || "").toLowerCase()).filter(Boolean)
  );
  const showUpEmailsPreExclusion = new Set(
    showUpMerge.map((r) => (r.email || "").toLowerCase()).filter(Boolean)
  );
  let oldStudentsShowUpOptInCount = 0;
  let oldStudentsShowUpCount = 0;
  for (const e of oldStudentsEmails) {
    if (showUpEmailsPreExclusion.has(e)) {
      oldStudentsShowUpCount++;
      if (optInEmailsPreExclusion.has(e)) {
        oldStudentsShowUpOptInCount++;
      }
    }
  }
  const optInRowsBeforeExclusion = optInRows.length;
  const showUpMergeBeforeExclusion = showUpMerge.length;
  // Capture matched rows by email so we can render them in the "Old Students" tab.
  const excludedFromOptInByEmail = new Map<string, OptInRow>();
  for (const r of optInRows) {
    if (isOldStudent(r.email)) excludedFromOptInByEmail.set(r.email.toLowerCase(), r);
  }
  const excludedFromShowUpByEmail = new Map<string, ShowUpMergeRow>();
  for (const r of showUpMerge) {
    if (isOldStudent(r.email)) excludedFromShowUpByEmail.set(r.email.toLowerCase(), r);
  }
  const oldStudentsExcludedList: OldStudentExclusionRow[] = [];
  const seenExcludedEmails = new Set<string>();
  for (const [email, r] of excludedFromOptInByEmail) {
    const inShow = excludedFromShowUpByEmail.has(email);
    oldStudentsExcludedList.push({
      email: r.email,
      name: r.fullName || r.firstName || "",
      country: r.country,
      foundIn: inShow ? "Opt-In + Show Up" : "Opt-In",
    });
    seenExcludedEmails.add(email);
  }
  for (const [email, r] of excludedFromShowUpByEmail) {
    if (seenExcludedEmails.has(email)) continue;
    oldStudentsExcludedList.push({
      email: r.email,
      name: r.fullName || "",
      country: r.country,
      foundIn: "Show Up",
    });
  }
  oldStudentsExcludedList.sort((a, b) =>
    (a.email || "").localeCompare(b.email || "")
  );
  // Old students who showed up, captured before the exclusion filter below —
  // still needed for the Keap Working export's show-up tag.
  const oldStudentsShowUpRows = showUpMerge.filter((r) => isOldStudent(r.email));
  const optInRowsFiltered = optInRows.filter((r) => !isOldStudent(r.email));
  const showUpMergeFiltered = showUpMerge.filter((r) => !isOldStudent(r.email));
  const oldStudentsExcludedCount =
    (optInRowsBeforeExclusion - optInRowsFiltered.length) +
    (showUpMergeBeforeExclusion - showUpMergeFiltered.length);
  // Replace mutable lists with filtered versions for all downstream metrics.
  optInRows.length = 0;
  optInRows.push(...optInRowsFiltered);
  showUpMerge.length = 0;
  showUpMerge.push(...showUpMergeFiltered);

  // ===== Metrics, country breakdowns, and Student List =====
  const { metrics, optInByCountry, showUpByCountry, signUpByCountry, studentList } =
    deriveMetrics(session, optInRows, showUpMerge, signUpRows, {
      oldStudentsExcludedCount,
      oldStudentsShowUpOptInCount,
      oldStudentsShowUpCount,
      showUpAddedToOptInCount,
    });

  return {
    sessionDetails: session,
    metrics,
    optInByCountry,
    showUpByCountry,
    signUpByCountry,
    optIns: optInRows,
    showUpMerge,
    showUpReg,
    signUps: signUpRows,
    studentList,
    oldStudentsExcluded: oldStudentsExcludedList,
    oldStudentsShowUpRows,
    generatedAt: new Date().toISOString(),
    nlow4ExcludedPhones,
    nlow4ExcludedEmails,
  };
}
