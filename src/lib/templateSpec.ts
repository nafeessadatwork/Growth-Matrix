/**
 * Layout tokens derived from Flat OPC reference:
 * docs/MindX_Performance_Appraisal_Form_Final.xml
 * (w:pgSz, w:pgMar, w:tblW, w:tcW, w:tcMar, w:spacing samples, theme fills, wp:extent for logo)
 *
 * When the Word template changes, re-measure from that export and update this file.
 *
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * Word run font size in half-points (w:sz w:val).
 *
 * Verified against XML:
 *   headerTitle  → w:sz=44  (title line in navy cell)
 *   headerAddress→ w:sz=13  (address/contact combined line, subdued)
 *   tagline      → no explicit sz in XML → default 12 pt = 24 half-pts
 *   redBand      → w:sz=24  (all section-banner rows: PERFORMANCE, GOALS, COMMENTS, SIGNATURES)
 *   tableHeader  → no explicit sz on column-header rows → default 12 pt = 24 half-pts
 *   body         → w:sz=17  (info-grid labels & values, summary rows, goal/sig cells)
 *   bodySmall    → w:sz=15  (factor description paragraphs, footer notes)
 *   factorTitle  → w:sz=17  (performance factor numbered title line)
 *   factorDesc   → w:sz=15  (factor description line)
 *   footerNote   → w:sz=15  (goal footnote lines)
 */
export const TYPO = {
  headerTitle:   44,
  headerAddress: 13,
  headerContact: 13,
  tagline:       24,
  redBand:       24,
  tableHeader:   24,
  body:          17,
  bodySmall:     15,
  factorTitle:   17,
  factorDesc:    15,
  footerNote:    15,
} as const;

/**
 * Paragraph spacing (twips) and line spacing.
 *
 * sectionBreak  → w:spacing w:before="120" w:after="120"  (every inter-block <w:p> in XML)
 * tagline       → w:spacing w:before="80"                 (before tagline, no after)
 * address       → w:spacing w:before="80"                 (before address line)
 */
export const SPACING = {
  logoBlock:        { before: 120, after: 120 } as const,
  navyTitleFirst:   { before: 40,  after: 0   } as const,
  navyTitleMid:     { before: 0,   after: 0   } as const,
  navyTitleLast:    { before: 0,   after: 80  } as const,
  tagline:          { before: 80,  after: 0   } as const,
  address:          { before: 80,  after: 0   } as const,
  /** Inter-block gap: w:before="120" w:after="120" */
  sectionBreak:     { before: 120, after: 120 } as const,
  /** Multiline comment cell */
  commentCell:      { before: 120, after: 120 } as const,
  goalNote1:        { before: 120, after:  60 } as const,
  goalNote2:        { before: 0,   after: 240 } as const,
  /** Signature content rows: w:spacing w:before="800" (push lines down in tall cell) */
  signaturePara1:   { before: 800, after: 0   } as const,
  signaturePara2:   { before: 0,   after: 100 } as const,
  signaturePara3:   { before: 0,   after: 200 } as const,
  /** Table body line height (twentieths of a point) — pair with LineRuleType.AT_LEAST */
  tableBodyLineTwips: 276,
  emptyLine:        { after: 400 } as const,
} as const;

/**
 * Table cell padding (twips) — from template w:tcMar samples.
 *
 * logo     → uniform 120 (logo cell)
 * navy     → top/bottom 160, left/right 300 (navy header cell)
 * standard → top/bottom 100, left/right 150 (data / factor / goal cells)
 * factor   → top/bottom 80,  left/right 150 (performance factor rows: w:tcMar top=80)
 * redBand  → top/bottom 110, left/right 200 (section-banner spanning cell)
 * summaryLabel → top/bottom 100, left 150, right 100 (summary label column)
 */
export const CELL = {
  logo:         { top: 120, bottom: 120, left: 120, right: 120 } as const,
  navy:         { top: 160, bottom: 160, left: 300, right: 300 } as const,
  standard:     { top: 100, bottom: 100, left: 150, right: 150 } as const,
  factor:       { top:  80, bottom:  80, left: 150, right: 150 } as const,
  redBand:      { top: 110, bottom: 110, left: 200, right: 200 } as const,
  summaryLabel: { top: 100, bottom: 100, left: 150, right: 100 } as const,
} as const;

/**
 * Brand colors — all verified against XML w:shd/@w:fill and w:color/@w:val.
 *
 * PRIMARY    "E8351A"  red  — section banners, factor titles, accent borders
 * ORANGE     "FF6B2B"  warm orange — tagline, goal-number prefix
 * NAVY       "0C1A2E"  deep navy — header cell, total-score label bg
 * NAVY_ALT   "1A2F4A"  mid navy — summary header bg, comments/sig header bg
 * WHITE      "FFFFFF"
 * LIGHT_GRAY "F5F5F5"  — info label bg, score display cells
 * TINT_SOFT  "FFF0EC"  light salmon — summary label rows (Employee/Reviewer Score)
 * TINT_WARM  "FFF7F5"  very light warm — alternating factor/goal rows, CEO body cell
 * BORDER     "CCCCCC"  — all table cell borders in the template
 * TEXT_MAIN  "1A1A2E"  — weightage values, data cell text
 * TEXT_MUTED "666666"  — factor description paragraphs
 * ADDRESS    "8BAABF"  — header address / contact line text
 * SIG_TEXT   "999999"  — signature & date underline text
 */
export const COLORS = {
  PRIMARY:    "E8351A",
  ORANGE:     "FF6B2B",
  NAVY:       "0C1A2E",
  NAVY_ALT:   "1A2F4A",
  WHITE:      "FFFFFF",
  LIGHT_GRAY: "F5F5F5",
  TINT_SOFT:  "FFF0EC",
  TINT_WARM:  "FFF7F5",
  BORDER:     "CCCCCC",
  TEXT_MAIN:  "1A1A2E",
  TEXT_MUTED: "666666",
  ADDRESS:    "8BAABF",
  SIG_TEXT:   "999999",
} as const;

/** A4 page — w:pgSz w:w="11906" w:h="16838" w:code="9" */
export const PAGE = {
  size: { width: 11906, height: 16838, code: 9 as const },
  margin: {
    top:    1440,
    right:  1440,
    bottom: 1440,
    left:   1440,
    header:  706,
    footer:  706,
    gutter:    0,
  },
} as const;

/** Full text area width (w:tblW dxa on full-bleed tables) */
export const CONTENT_WIDTH_DXA = 9360;

/**
 * Column widths (DXA) — from template w:gridCol/@w:w.
 *
 * headerLogo   2800 + headerNavy 6560 = 9360 ✓
 * info         1500 + 3180 + 1500 + 3180 = 9360 ✓
 * performance  2200 + 1450 + 980 + 1450 + 980 + 2300 = 9360 ✓
 * summaryNarrow 2340 + 1170 + 1170 = 4680 (narrow left-aligned table)
 * goals        4660 + 1000 + 900 + 2800 = 9360 ✓
 * finalCalc    2500 + 1250 + 1250 = 5000 (narrow centered table)
 * triple       3120 + 3120 + 3120 = 9360 ✓ (comments & signatures)
 */
export const COLS = {
  headerLogo:    2800,
  headerNavy:    6560,
  info:          [1500, 3180, 1500, 3180] as const,
  performance:   [2200, 1450,  980, 1450,  980, 2300] as const,
  summaryNarrow: [2340, 1170, 1170] as const,
  goals:         [4660, 1000,  900, 2800] as const,
  finalCalc:     [2500, 1250, 1250] as const,
  triple:        [3120, 3120, 3120] as const,
} as const;

export const TABLE = {
  full:             { size: CONTENT_WIDTH_DXA, type: "dxa" as const },
  summaryNarrow:    { size: 4680,              type: "dxa" as const },
  finalCalc:        { size: 5000,              type: "dxa" as const },
  /** Center narrow tables: (9360 − width) / 2 */
  indentNarrow4680: { size: 2340,              type: "dxa" as const },
  indentNarrow5000: { size: 2180,              type: "dxa" as const },
} as const;

/** Row heights (twips) from template w:trHeight/@w:val */
export const ROW = {
  headerTable: 1400,
  /** Signature / comment body rows: w:trHeight w:val="1200" */
  contentTall: 1200,
  /** Single-line red section headers (atLeast) */
  redBanner:    520,
} as const;

/** Logo drawing extents from template wp:extent (EMU) → ImageRun px @ 96 dpi */
const LOGO_EMU = { cx: 1476375, cy: 723900 } as const;
const EMU_PER_INCH = 914400;
const PX_PER_INCH  = 96;
export const LOGO_PX = {
  width:  Math.round((LOGO_EMU.cx * PX_PER_INCH) / EMU_PER_INCH),
  height: Math.round((LOGO_EMU.cy * PX_PER_INCH) / EMU_PER_INCH),
} as const;