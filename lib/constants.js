export const MEDIUMS = ["English", "Malayalam"];
export const CLASSES = ["Plus One", "Plus Two"];
// Classes whose syllabus we only have in English. The medium picker is locked to
// English for these (see mediumsFor / the Medium field in StudentDetailsForm).
// The commerce syllabus is English medium only.
export const ENGLISH_ONLY_CLASSES = ["Plus One", "Plus Two"];
export const mediumsFor = (cls) =>
  ENGLISH_ONLY_CLASSES.includes(cls) ? ["English"] : MEDIUMS;
export const EPSID_RE = /^EPS\d{6}$/;
export const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
// One colour per subject in the accordion. Commerce runs seven subjects, so the
// list needs seven entries — with six, Maths wrapped around onto Economics.
export const PALETTE = [
  "#4f46e5", "#059669", "#d97706", "#e11d48", "#0284c7", "#7c3aed", "#0d9488",
];

// A day's load is "heavy" past this many topics — used to warn the student.
// The full Plus Two commerce syllabus is ~574 topics: comfortable over a
// six-month run (3-4 a day), but past 6 a day the plan has stopped being
// realistic.
export const HEAVY_TOPICS_PER_DAY = 6;
