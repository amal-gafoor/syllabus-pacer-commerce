/**
 * Excel → curriculum data generator.
 *
 *   npm run curriculum
 *
 * Reads `textbooks/CHAPTER AND TOPIC HUMANITIES & COMMERCE.xlsx` and writes
 * `lib/curriculum/plus-one-com.js` and `lib/curriculum/plus-two-com.js`.
 *
 * Re-run this whenever the teachers update the workbook — the generated files
 * are build artefacts, so edit the spreadsheet, not the output.
 *
 * ── Sheet layout it expects ──────────────────────────────────────────────
 * Each sheet lays its subjects out side by side in column PAIRS. One header row
 * of the left column holds the subject name; the pair is then (TOPICS,
 * SUBTOPICS). The commerce sheets put that header on row 3, the humanities ones
 * on row 2 — hence `headRow` per target below:
 *
 *      A            B            D                  E          …
 *   3  ECONOMICS                 POLITICAL SCIENCE
 *   4  CHAPTER 1.1 - … (NAME)    CHAPTER 1.1 - … (NAME)
 *   5  TOPICS       SUBTOPICS    TOPICS             SUBTOPICS
 *   6  1. Foo       1. Foo       1. Bar             1. Bar
 *   7                2. Foo b                                    ← extra subtopic
 *
 * Chapter blocks stack vertically down the same pair, separated by blanks. A
 * topic spanning several subtopics leaves the TOPICS cell empty on the
 * continuation rows — we schedule at TOPIC grain, so those rows are skipped.
 *
 * ── Chapters still awaiting a topic breakdown ────────────────────────────
 * Some subjects list their remaining chapters as bare numbered rows in the
 * TOPICS column, with nothing beside them in SUBTOPICS. See PLACEHOLDER_RUN
 * below for how those are told apart from real topics.
 *
 * No page counts exist anywhere in this workbook, which is why the planner is
 * topic-weighted rather than page-weighted.
 *
 * Written against Node's standard library only (its own ZIP + XML readers) so
 * that regenerating the curriculum never depends on an install step.
 */

import { existsSync, readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const WORKBOOK = join(ROOT, "textbooks", "CHAPTER AND TOPIC HUMANITIES & COMMERCE.xlsx");
const OUT_DIR = join(ROOT, "lib", "curriculum");

/**
 * Sheet tab → generated file + exported constant. `headRow` is the row carrying
 * the subject names; the workbook is not consistent about it between streams.
 *
 * The workbook also holds the humanities tabs (P1H / P2H, headRow 2). They are
 * not generated — this app is the commerce pacer — but the parser stays
 * layout-agnostic so adding them back is a one-line change.
 */
const TARGETS = [
  { sheet: "P1C", headRow: 3, file: "plus-one-com.js", constName: "PLUS_ONE_COM", label: "Plus One" },
  { sheet: "P2C", headRow: 3, file: "plus-two-com.js", constName: "PLUS_TWO_COM", label: "Plus Two" },
];

/**
 * Display order for the subject accordion. The workbook's column order differs
 * between the two sheets (P1C runs Accountancy → Business Studies → Computer
 * Application → Statistics, P2C puts Statistics third), so we impose one order
 * and students see the same list in both classes.
 * Anything not listed keeps its sheet order and lands at the end.
 */
const SUBJECT_ORDER = [
  "Economics",
  "Accountancy",
  "Business Studies",
  "Political Science",
  "Statistics",
  "Computer Application",
  "Maths",
];

/**
 * Revision / question-practice rows the teachers keep alongside the syllabus.
 * These are dropped: the pacer schedules syllabus content, and board revision is
 * planned separately.
 *
 * The first pattern matches "7. BOARD", "9. BOARD PYQ 2026", "4. PYQ". The
 * second catches the commerce sheets' practice rows, which are named for the
 * exercise rather than the content — "Advanced Questions", "New Pattern Based
 * Questions", "Previous Year Questions", "Trial Balance Questions". Anchoring on
 * a trailing "Question(s)" keeps real topics safe: no chapter in either sheet
 * ends its title that way.
 */
const REVISION_ROW = [/^(?:board|pyq)\b/i, /\bquestions?$/i];

/** Spelling slips in the source sheet, fixed on the way through. */
const TYPOS = [
  [/\bLIV\s+ING\b/gi, "Living"],
  [/\bAPRAISAL\b/gi, "Appraisal"],
  [/\bHEGEMONEY\b/gi, "Hegemony"],
  [/\bCONTEMPARARY\b/gi, "Contemporary"],
  [/\bELEMENTRY\b/gi, "Elementary"],
  [/\bDEVELOPEMENT\b/gi, "Development"],
  [/\bINDRODUCTION\b/gi, "Introduction"],
  [/\bMethodes\b/gi, "Methods"],
  [/\bstudiying\b/gi, "studying"],
  [/\bDifferenciation\b/gi, "Differentiation"],
  // Commerce sheets. ACCOUTANCY is the P1C subject heading — P2C spells it
  // correctly, and without this the two classes show different subject names.
  [/\bACCOUTANCY\b/gi, "Accountancy"],
  [/\bPropritorship\b/gi, "Proprietorship"],
  [/\bAccountng\b/gi, "Accounting"],
  [/\bStanderds\b/gi, "Standards"],
  // "Compliment of a set" — the set operation is spelt Complement.
  [/\bCompliment\b/gi, "Complement"],
];

/**
 * Chapter titles are tagged with the name of the teacher who owns the chapter,
 * in parentheses — internal bookkeeping that students should never see. They are
 * always a single ALL-CAPS word, so that shape is what we strip, rather than
 * listing the staff names here (this file is public; the workbook is not).
 *
 * A couple of genuine title fragments have the same shape and must survive —
 * "Freedom (LIBERTY)" is a real chapter title, not an author tag.
 */
const KEEP_PARENS = new Set(["LIBERTY"]);
const AUTHOR_TAG = /^[A-Z][A-Z.'-]*$/;

/** Kept upper-case by the title-caser. */
const ACRONYMS = new Set([
  // Shared with the Political Science / Economics columns.
  "US", "UK", "UN", "USA", "USSR", "LPG", "WTO", "NHRC", "GDP", "NEP",
  "FPTP", "PR", "IT", "SAARC", "ASEAN", "NATO", "IMF", "WHO", "NGO",
  // Commerce.
  "GST", "MSME", "BRS", "FS", "AI", "HTML", "CSS", "SQL", "DBMS", "ICT",
  "GNU", "NGO", "ROI", "EPS", "NPV", "IPO", "SEBI", "RBI", "ATM", "OPC",
  "HUF", "NCERT", "PAN", "TDS", "CRM", "ERP", "URL", "LAN", "WAN",
]);

/** Lower-cased by the title-caser unless they lead the title. */
const SMALL_WORDS = new Set([
  "a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into",
  "nor", "of", "on", "or", "the", "to", "up", "vs", "with", "its", "their",
]);

/**
 * Chapters the workbook lists twice under different subjects, keyed by
 * `sheet | subject | chapter number`. Neither commerce sheet has any — the
 * hook is kept because the humanities tabs did, and the teachers copy columns
 * between sheets when a subject is shared.
 */
const DUPLICATES = new Set();

// ── ZIP reader ────────────────────────────────────────────────────────────
// An .xlsx is a ZIP of XML parts. We read the central directory rather than
// scanning local headers, because local headers may carry zeroed sizes when the
// writer used a trailing data descriptor.

function openZip(path) {
  const buf = readFileSync(path);

  let eocd = -1;
  const floor = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= floor; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error(`${path} is not a readable .xlsx (no ZIP end record)`);

  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = new Map();

  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error("Corrupt ZIP central directory");
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    entries.set(buf.toString("utf8", off + 46, off + 46 + nameLen), { method, compSize, localOff });
    off += 46 + nameLen + extraLen + commentLen;
  }

  return (name) => {
    const e = entries.get(name);
    if (!e) throw new Error(`${name} is missing from the workbook`);
    if (buf.readUInt32LE(e.localOff) !== 0x04034b50) throw new Error(`Corrupt ZIP entry: ${name}`);
    const start =
      e.localOff + 30 + buf.readUInt16LE(e.localOff + 26) + buf.readUInt16LE(e.localOff + 28);
    const raw = buf.subarray(start, start + e.compSize);
    return (e.method === 0 ? raw : inflateRawSync(raw)).toString("utf8");
  };
}

// ── XML helpers ───────────────────────────────────────────────────────────
// The parts are machine-generated by Excel, so targeted regexes are sufficient
// and keep this script dependency-free.

const decodeXml = (s) =>
  s
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

/** Shared string table: one entry per <si>, joining its <t> runs. */
function readSharedStrings(read) {
  let xml;
  try { xml = read("xl/sharedStrings.xml"); } catch { return []; }
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(([, body]) =>
    decodeXml(
      [...body.replace(/<rPh[\s\S]*?<\/rPh>/g, "").matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
        .map((m) => m[1])
        .join("")
    )
  );
}

/** Tab name → part path, resolved through the workbook relationships. */
function readSheetIndex(read) {
  const rels = new Map(
    [...read("xl/_rels/workbook.xml.rels").matchAll(/<Relationship\b[^>]*>/g)].map((m) => [
      /Id="([^"]+)"/.exec(m[0])?.[1],
      /Target="([^"]+)"/.exec(m[0])?.[1],
    ])
  );
  const sheets = new Map();
  for (const m of read("xl/workbook.xml").matchAll(/<sheet\b[^>]*\/?>/g)) {
    const name = /name="([^"]+)"/.exec(m[0])?.[1];
    const target = rels.get(/r:id="([^"]+)"/.exec(m[0])?.[1]);
    if (!name || !target) continue;
    sheets.set(decodeXml(name), target.startsWith("/") ? target.slice(1) : `xl/${target}`);
  }
  return sheets;
}

/** Sheet → { cells: Map<"A5", string>, maxRow }, blanks omitted. */
function readSheet(read, path, shared) {
  const xml = read(path);
  const cells = new Map();
  let maxRow = 0;

  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const [, attrs, body = ""] = m;
    const ref = /\br="([A-Z]+\d+)"/.exec(attrs)?.[1];
    if (!ref) continue;
    const type = /\bt="([^"]+)"/.exec(attrs)?.[1];

    let value;
    if (type === "inlineStr") {
      value = [...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join("");
    } else {
      value = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
    }
    if (value == null) continue;

    let text = decodeXml(value);
    if (type === "s") text = shared[Number(text)] ?? "";
    if (!text.trim()) continue;

    cells.set(ref, text);
    maxRow = Math.max(maxRow, Number(/\d+$/.exec(ref)[0]));
  }
  return { cells, maxRow };
}

/** "A" → 1, "Z" → 26, "AA" → 27. */
const colToNum = (col) =>
  [...col].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);

/** 1 → "A", 27 → "AA". */
function numToCol(n) {
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

// ── text cleaning ─────────────────────────────────────────────────────────

const fixTypos = (s) =>
  TYPOS.reduce(
    (acc, [re, fix]) =>
      acc.replace(re, (match) => (match === match.toUpperCase() ? fix.toUpperCase() : fix)),
    s
  );

/** Curly quotes and dashes → ASCII; collapse runs of whitespace (incl. tabs). */
const normalise = (s) =>
  s
    .replace(/[‘’‚‛]/g, "'")
    .replace(/[“”„‟]/g, '"')
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\(\s+/g, "(")
    .replace(/\s+\)/g, ")")
    .replace(/\s+([,;:])/g, "$1")
    .trim();

function titleCase(s) {
  let wordIndex = 0;
  let afterColon = false;
  return s
    .split(/(\s+)/)
    .map((token) => {
      if (!token.trim()) return token;
      // A small word still capitalises when it opens the title or a subtitle.
      const opensClause = wordIndex++ === 0 || afterColon;
      afterColon = /[:;.–-]$/.test(token);
      const letters = token.replace(/[^A-Za-z]/g, "");
      // Preserve an acronym only where the sheet already wrote it in caps.
      if (letters && token === token.toUpperCase() && ACRONYMS.has(letters.toUpperCase())) {
        return token.toUpperCase();
      }
      const lower = token.toLowerCase();
      if (!opensClause && SMALL_WORDS.has(letters.toLowerCase())) return lower;
      return lower.replace(/[a-z]/, (c) => c.toUpperCase());
    })
    .join("");
}

/** Sentence-case only when the source is shouting, so mixed-case stays as written. */
const deShout = (s) => (/[a-z]/.test(s) ? s : titleCase(s));

const upperFirst = (s) => s.replace(/^([a-z])/, (c) => c.toUpperCase());

/**
 * Strip the leading "1." / "2)" ordinal the sheet numbers every row with.
 *
 * Two typing slips in the commerce sheets need the looser shapes: a doubled
 * separator ("4.. Applications Of Statistics") and a bare number with no dot at
 * all ("11  Factors Affects the Supply Curve"). The bare form demands two or
 * more spaces so a genuine title opening with a number — "73rd Constitutional
 * Amendment", "1's Complement Method" — is left alone.
 */
const ORDINAL = /^\s*\d+\s*(?:[.)]\s*)+|^\s*\d+\s{2,}/;
const stripOrdinal = (s) => s.replace(ORDINAL, "").trim();

/** The leading ordinal itself — the chapter number on a placeholder row. */
const leadingOrdinal = (s) => /^\s*(\d+)/.exec(s)?.[1] ?? "";

/** Drop single-word ALL-CAPS parentheticals (author tags, "(BOARD)"). */
const stripNoiseParens = (s) =>
  s.replace(/\s*\(([^()]*)\)/g, (match, inner) => {
    const word = inner.trim();
    return AUTHOR_TAG.test(word) && !KEEP_PARENS.has(word.toUpperCase()) ? " " : match;
  });

/** Tidy the dangling "-" / ":" left behind once a note is removed. */
const trimPunctuation = (s) => s.replace(/^[\s\-–—:.,]+|[\s\-–—:,]+$/g, "");

const CHAPTER_HEAD =
  /^\s*CHAPTER\s*[-–—]?\s*(\d+)\s*(?:[.:]\s*(\d+))?\s*[-–—:.)]*\s*/i;

/** `CHAPTER - 1: 3. UNDERSTANDING …` → { num: "1.3", name: "Understanding …" }. */
function parseChapter(raw) {
  const clean = (s) => deShout(fixTypos(normalise(trimPunctuation(stripNoiseParens(s)))));
  const text = normalise(raw);
  const m = CHAPTER_HEAD.exec(text);
  if (!m) return { num: "", name: clean(text) };
  const num = m[2] ? `${m[1]}.${m[2]}` : m[1];
  return { num, name: clean(text.slice(m[0].length)) || `Chapter ${num}` };
}

// ── workbook → curriculum ─────────────────────────────────────────────────

/**
 * Where a subject's trailing run of "chapter listed, breakdown still to come"
 * rows begins, or 0 if it has none.
 *
 * Some subjects are only part-way through being broken down. The teachers list
 * the remaining chapters as bare numbered rows in the TOPICS column, below the
 * last real chapter block:
 *
 *      P            Q
 *   9  4. Applications Of …      4. Applications Of …   ← last real topic
 *  10                                                   ← blank separator
 *  11  2.Methods of Data Collection and Sampling        ← chapter, no breakdown
 *  12  3.Presentation of Data
 *
 * They look exactly like topics, so the tell is the empty SUBTOPICS cell beside
 * them: every genuine topic row in this workbook is paired with one. Requiring
 * the run to reach the bottom of the column keeps the rule narrow — a single
 * mid-column row where someone merely forgot the subtopic is still read as a
 * topic, which is the safer reading of an ambiguous row.
 *
 * The humanities tabs write the same thing as "CHAPTER 2.Data Collection", which
 * the normal chapter path already handles; this is only for the unprefixed form.
 */
function placeholderRunStart(cells, maxRow, topicCol, subCol, headRow) {
  let start = 0;

  for (let row = maxRow; row > headRow; row--) {
    const cell = cells.get(`${topicCol}${row}`);
    if (!cell) continue;
    // A prefixed chapter or a paired topic ends the run.
    if (/^\s*CHAPTER\b/i.test(cell)) break;
    if (/^(TOPICS|SUBTOPICS)$/i.test(cell.trim())) break;
    if (cells.get(`${subCol}${row}`)) break;
    // Revision rows are dropped anyway — don't let one anchor a false run.
    const topic = normalise(stripOrdinal(cell));
    if (REVISION_ROW.some((re) => re.test(topic))) break;
    start = row;
  }

  if (!start) return 0;

  // One lone row is more likely a missing subtopic than a chapter list.
  let rows = 0;
  for (let row = start; row <= maxRow; row++) if (cells.get(`${topicCol}${row}`)) rows++;
  return rows > 1 ? start : 0;
}

/**
 * Number any topic name a chapter lists more than once.
 *
 * P2C Accountancy 3.2 carries "Redemption Problem" at both row 208 and row 210.
 * Those are two separate practice sessions rather than a copy-paste slip — the
 * sheet numbers them 10 and 12, either side of "Redemption by Payment in
 * Instalments" — but identical names collide in the planner's dedupe key and
 * hand the student two plan rows they cannot tell apart. Numbering keeps both
 * sessions schedulable without inventing a description for either.
 *
 * Reported by the caller so the teachers can name them properly in the sheet,
 * at which point this quietly stops firing.
 */
function numberRepeatedTopics(subjectName, chapters, stats) {
  for (const chapter of chapters) {
    const taken = new Set();
    const count = new Map();

    for (const topic of chapter.topics) {
      if (!taken.has(topic.name)) {
        taken.add(topic.name);
        continue;
      }
      const base = topic.name;
      let n = (count.get(base) ?? 1) + 1;
      while (taken.has(`${base} (${n})`)) n++;
      count.set(base, n);
      stats.repeatsNumbered.push(`${subjectName} ${chapter.num} "${base}" → "${base} (${n})"`);
      topic.name = `${base} (${n})`;
      taken.add(topic.name);
    }
  }
}

function parseSubjects(sheet, sheetName, headRow, stats) {
  const { cells, maxRow } = sheet;

  // Subject headings own their column and the one to its right (its SUBTOPICS).
  const headRef = new RegExp(`^[A-Z]+${headRow}$`);
  const heads = [...cells.keys()]
    .filter((ref) => headRef.test(ref))
    .map((ref) => /^([A-Z]+)/.exec(ref)[1])
    .sort((a, b) => colToNum(a) - colToNum(b));

  const subjects = [];

  for (const topicCol of heads) {
    const name = fixTypos(deShout(normalise(cells.get(`${topicCol}${headRow}`))));
    const subCol = numToCol(colToNum(topicCol) + 1);
    const placeholderFrom = placeholderRunStart(cells, maxRow, topicCol, subCol, headRow);
    const chapters = [];
    let current = null;

    for (let row = headRow + 1; row <= maxRow; row++) {
      const cell = cells.get(`${topicCol}${row}`);
      if (!cell) continue;

      // Trailing chapters listed without a "CHAPTER" prefix or a breakdown.
      if (placeholderFrom && row >= placeholderFrom) {
        const num = leadingOrdinal(cell);
        const chapterName =
          deShout(fixTypos(normalise(trimPunctuation(stripNoiseParens(stripOrdinal(cell)))))) ||
          `Chapter ${num}`;
        chapters.push({ num, name: chapterName, topics: [] });
        stats.placeholders.push(`${name} ${num} ${chapterName}`);
        current = null;
        continue;
      }

      if (CHAPTER_HEAD.test(cell) || /^\s*CHAPTER\b/i.test(cell)) {
        const { num, name: chapterName } = parseChapter(cell);
        if (DUPLICATES.has(`${sheetName}|${name}|${num}`)) {
          stats.duplicatesDropped.push(`${sheetName} ${name} ${num} ${chapterName}`);
          current = null; // skip this block's topic rows too
          continue;
        }
        current = { num, name: chapterName, topics: [] };
        chapters.push(current);
        continue;
      }

      if (/^(TOPICS|SUBTOPICS)$/i.test(cell.trim())) continue;
      if (!current) continue;

      const topic = normalise(stripOrdinal(cell));
      if (!topic) continue;
      if (REVISION_ROW.some((re) => re.test(topic))) { stats.revisionDropped++; continue; }

      current.topics.push({ name: upperFirst(fixTypos(topic)) });
    }

    numberRepeatedTopics(name, chapters, stats);
    if (chapters.length) subjects.push({ name, chapters });
  }

  // Impose a stable cross-class display order.
  subjects.sort((a, b) => {
    const ai = SUBJECT_ORDER.indexOf(a.name);
    const bi = SUBJECT_ORDER.indexOf(b.name);
    return (ai < 0 ? SUBJECT_ORDER.length : ai) - (bi < 0 ? SUBJECT_ORDER.length : bi);
  });

  return subjects;
}

// ── emit ──────────────────────────────────────────────────────────────────

const quote = (s) => JSON.stringify(s);

function render(subjects, constName, label, sheetName) {
  const lines = [
    `// GENERATED FILE — do not edit by hand.`,
    `// Source: textbooks/CHAPTER AND TOPIC HUMANITIES & COMMERCE.xlsx (sheet "${sheetName}")`,
    `// Regenerate with: npm run curriculum`,
    `//`,
    `// ${label} Commerce. There are no page counts in the source workbook, so the`,
    `// planner weights days by topic count instead — see lib/planner/allocateDays.js.`,
    `// Chapters with an empty \`topics\` array are listed in the workbook but have no`,
    `// topic breakdown yet; the picker shows them disabled.`,
    ``,
    `const ${constName} = [`,
  ];

  for (const subject of subjects) {
    lines.push(`  {`);
    lines.push(`    name: ${quote(subject.name)},`);
    lines.push(`    chapters: [`);
    for (const chapter of subject.chapters) {
      const head = `      { num: ${quote(chapter.num)}, name: ${quote(chapter.name)},`;
      if (!chapter.topics.length) {
        lines.push(`${head} topics: [] },`);
        continue;
      }
      lines.push(`${head} topics: [`);
      for (const topic of chapter.topics) lines.push(`        { name: ${quote(topic.name)} },`);
      lines.push(`      ] },`);
    }
    lines.push(`    ],`);
    lines.push(`  },`);
  }

  lines.push(`];`, ``, `export default ${constName};`, ``);
  return lines.join("\n");
}

// ── main ──────────────────────────────────────────────────────────────────

function main() {
  // textbooks/ is gitignored (it carries internal teacher names), so a fresh
  // clone won't have the workbook. The committed curriculum still builds and
  // runs — only regeneration needs the file.
  if (!existsSync(WORKBOOK)) {
    console.error(
      `Workbook not found:\n  ${WORKBOOK}\n\n` +
      `It is deliberately kept out of the repository. Ask the team for\n` +
      `"CHAPTER AND TOPIC HUMANITIES & COMMERCE.xlsx" and drop it into textbooks/.\n` +
      `The generated curriculum in lib/curriculum/ is committed, so you only\n` +
      `need this to pick up syllabus changes.`
    );
    process.exit(1);
  }

  const read = openZip(WORKBOOK);
  const shared = readSharedStrings(read);
  const index = readSheetIndex(read);

  mkdirSync(OUT_DIR, { recursive: true });

  for (const target of TARGETS) {
    const path = index.get(target.sheet);
    if (!path) throw new Error(`Sheet "${target.sheet}" not found. Found: ${[...index.keys()].join(", ")}`);

    const stats = {
      revisionDropped: 0, duplicatesDropped: [], placeholders: [], repeatsNumbered: [],
    };
    const subjects = parseSubjects(
      readSheet(read, path, shared), target.sheet, target.headRow, stats
    );

    writeFileSync(
      join(OUT_DIR, target.file),
      render(subjects, target.constName, target.label, target.sheet),
      "utf8"
    );

    const chapters = subjects.reduce((n, s) => n + s.chapters.length, 0);
    const topics = subjects.reduce(
      (n, s) => n + s.chapters.reduce((m, c) => m + c.topics.length, 0), 0
    );
    const empty = subjects.flatMap((s) =>
      s.chapters.filter((c) => !c.topics.length).map((c) => `${s.name} ${c.num} ${c.name}`)
    );

    console.log(
      `${target.sheet} → lib/curriculum/${target.file}\n` +
      `  ${subjects.length} subjects · ${chapters} chapters · ${topics} topics ` +
      `(dropped ${stats.revisionDropped} board/PYQ/practice rows)`
    );
    for (const d of stats.duplicatesDropped) console.log(`  duplicate removed: ${d}`);
    if (stats.placeholders.length) {
      console.log(`  ${stats.placeholders.length} chapter(s) read from an unprefixed trailing run`);
    }
    if (stats.repeatsNumbered.length) {
      console.log(`  ${stats.repeatsNumbered.length} repeated topic name(s) numbered — fix these in the sheet:`);
      for (const r of stats.repeatsNumbered) console.log(`    - ${r}`);
    }
    if (empty.length) {
      console.log(`  ${empty.length} chapter(s) with no topics yet — the picker disables these:`);
      for (const e of empty) console.log(`    - ${e}`);
    }
    for (const s of subjects) {
      const t = s.chapters.reduce((m, c) => m + c.topics.length, 0);
      console.log(`    ${s.name.padEnd(21)} ${String(s.chapters.length).padStart(2)} ch · ${String(t).padStart(3)} topics`);
    }
  }
}

main();
