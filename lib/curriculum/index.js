import { PALETTE } from "@/lib/constants";

import PLUS_ONE_COM from "./plus-one-com";
import PLUS_TWO_COM from "./plus-two-com";

// class → medium → subject list. The commerce syllabus is English medium only.
// Both files are generated from the P1C / P2C tabs of
// textbooks/CHAPTER AND TOPIC HUMANITIES & COMMERCE.xlsx — run
// `npm run curriculum` after the teachers update the workbook.
export const CURRICULA = {
  "Plus One": { English: PLUS_ONE_COM },
  "Plus Two": { English: PLUS_TWO_COM },
};

let _id = 0;
const nextId = () => ++_id;

/**
 * Look up the subject list for a class + medium. Commerce is English-only, so
 * any medium resolves to the English list; an unknown class falls back to empty.
 */
export const curriculumFor = (medium = "English", cls = "Plus Two") => {
  const byMedium = CURRICULA[cls] || {};
  return byMedium.English || [];
};

/**
 * Turn the static data into the editable shape the UI drives: ids, colours, and
 * open/selected flags.
 *
 * A handful of chapters are listed in the workbook with no topic breakdown yet
 * (Plus One Statistics 2–9 and Computer Application 3–8). They're kept in the
 * list so students can see the chapter exists, but flagged `ready: false` so the
 * picker disables them — the
 * planner can't schedule a chapter it has no topics for, and silently dropping a
 * ticked chapter from the finished plan would be worse than saying so up front.
 */
export const buildCurriculum = (medium = "English", cls = "Plus Two") =>
  curriculumFor(medium, cls).map((sub, i) => ({
    id: nextId(),
    name: sub.name,
    color: PALETTE[i % PALETTE.length],
    open: i === 0,
    chapters: sub.chapters.map((c) => ({
      id: nextId(),
      num: c.num || "",
      name: c.name,
      topics: c.topics || [],
      ready: (c.topics?.length ?? 0) > 0,
      selected: false,
    })),
  }));
