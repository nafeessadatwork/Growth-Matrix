/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  VerticalAlign,
  ShadingType,
  TableLayoutType,
  HeightRule,
  LineRuleType,
} from "docx";
import { AppraisalData } from "../types";
import { calculateAppraisal, formatScore } from "./utils";
import {
  CELL,
  COLORS,
  COLS,
  CONTENT_WIDTH_DXA,
  LOGO_PX,
  PAGE,
  ROW,
  SPACING,
  TABLE,
  TYPO,
} from "./templateSpec";

// ─── Border presets ────────────────────────────────────────────────────────────
// Template uses CCCCCC for all internal cell borders; outer table wrapper uses "auto".
const B = {
  grid: {
    top:             { style: BorderStyle.SINGLE, size: 4, color: "auto" },
    bottom:          { style: BorderStyle.SINGLE, size: 4, color: "auto" },
    left:            { style: BorderStyle.SINGLE, size: 4, color: "auto" },
    right:           { style: BorderStyle.SINGLE, size: 4, color: "auto" },
    insideHorizontal:{ style: BorderStyle.SINGLE, size: 4, color: "auto" },
    insideVertical:  { style: BorderStyle.SINGLE, size: 4, color: "auto" },
  },
  cell: {
    top:   { style: BorderStyle.SINGLE, size: 4, color: COLORS.BORDER },
    bottom:{ style: BorderStyle.SINGLE, size: 4, color: COLORS.BORDER },
    left:  { style: BorderStyle.SINGLE, size: 4, color: COLORS.BORDER },
    right: { style: BorderStyle.SINGLE, size: 4, color: COLORS.BORDER },
  },
  /** CEO cell red accent border (sz=6 in template) */
  ceoBorder: {
    top:   { style: BorderStyle.SINGLE, size: 6, color: COLORS.PRIMARY },
    bottom:{ style: BorderStyle.SINGLE, size: 6, color: COLORS.PRIMARY },
    left:  { style: BorderStyle.SINGLE, size: 6, color: COLORS.PRIMARY },
    right: { style: BorderStyle.SINGLE, size: 6, color: COLORS.PRIMARY },
  },
  none: {
    top:             { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom:          { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left:            { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right:           { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideHorizontal:{ style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideVertical:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  },
} as const;

const bodySpacing = { line: SPACING.tableBodyLineTwips, lineRule: LineRuleType.AT_LEAST };

// ─── Main generator ────────────────────────────────────────────────────────────
export async function generateAppraisalDoc(data: AppraisalData) {
  const stats = calculateAppraisal(data);
  const logoUrl = new URL("../assets/mindx-logo.png", import.meta.url).href;
  const logoBytes = new Uint8Array(await (await fetch(logoUrl)).arrayBuffer());

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size:   PAGE.size,
            margin: PAGE.margin,
          },
        },
        children: [
          // ── HEADER TABLE (Logo + Navy branding cell) ──────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.none,
            rows: [
              new TableRow({
                height: { value: ROW.headerTable, rule: HeightRule.ATLEAST },
                children: [
                  // Logo cell — white bg, uniform 120 twip padding
                  new TableCell({
                    width:  { size: COLS.headerLogo, type: WidthType.DXA },
                    borders: B.none,
                    shading: { fill: COLORS.WHITE, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...SPACING.logoBlock },
                        children: [
                          new ImageRun({
                            data:           logoBytes,
                            type:           "png",
                            transformation: { width: LOGO_PX.width, height: LOGO_PX.height },
                          }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins:       CELL.logo,
                  }),

                  // Navy branding cell — RIGHT-aligned (matches XML w:jc val="right")
                  new TableCell({
                    width:  { size: COLS.headerNavy, type: WidthType.DXA },
                    borders: B.none,
                    shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                    children: [
                      // Title line — "MINDX Growth Matrix Form" (sz=44, bold, white)
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        spacing:   { ...SPACING.navyTitleFirst },
                        children: [
                          new TextRun({
                            text:  "MINDX Growth Matrix Form",
                            bold:  true,
                            size:  TYPO.headerTitle,
                            color: COLORS.WHITE,
                          }),
                        ],
                      }),
                      // Tagline — italic orange (no explicit sz → default; before=80)
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        spacing:   { before: 80, ...bodySpacing },
                        children: [
                          new TextRun({
                            text:    "Technology with a Human Touch • mindxtech.ai",
                            size:    TYPO.tagline,
                            italics: true,
                            color:   COLORS.ORANGE,
                          }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins:       CELL.navy,
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.sectionBreak } }),

          // ── INFO GRID (Employee & Reviewer details) ───────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              createInfoRow("Employee Name",     data.employee.name,          "Reviewer Name",     data.reviewer.name),
              createInfoRow("Employee Position", data.employee.position,      "Reviewer Position", data.reviewer.position),
              createInfoRow("Department",        data.employee.department,    "Department",        data.reviewer.department),
              createInfoRow("Projects Managed",  data.employee.projectsManaged, "Evaluation Due",   data.reviewer.appraisalDue),
              createInfoRow("Type of Evaluation", data.employee.appraisalType,   "",               ""),
              createInfoRow("Evaluation Period",  data.employee.appraisalPeriod, "",               ""),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.sectionBreak } }),

          // ── GROWTH EVALUATION SECTION ─────────────────────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              createBannerRow("GROWTH EVALUATION", COLS.performance.length, COLORS.PRIMARY),
              // Column header row
              new TableRow({
                children: [
                  createHeaderCell("Growth Factor",    COLS.performance[0], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Employee Score",        COLS.performance[1], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Weightage",             COLS.performance[2], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Reviewer Score",        COLS.performance[3], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Weightage",             COLS.performance[4], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Reviewer's Comments",   COLS.performance[5], COLORS.PRIMARY, { noTopBorder: true }),
                ],
              }),

              // Factor data rows — alternating FFFFFF / TINT_WARM fill
              ...data.factors.map((f, idx) =>
                new TableRow({
                  children: [
                    // Factor name + description cell
                    new TableCell({
                      width:   { size: COLS.performance[0], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: { ...bodySpacing },
                          children: [
                            new TextRun({
                              text:  `${String(f.id).padStart(2, "0")}. ${f.title}`,
                              bold:  true,
                              color: COLORS.PRIMARY,
                              size:  TYPO.factorTitle,
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { ...bodySpacing },
                          children: [
                            new TextRun({
                              text:  f.description,
                              size:  TYPO.factorDesc,
                              color: COLORS.TEXT_MUTED,
                            }),
                          ],
                        }),
                      ],
                      margins: CELL.factor,
                    }),

                    // Employee score
                    new TableCell({
                      width:   { size: COLS.performance[1], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [
                            new TextRun({ text: f.employeeScore.toString(), size: TYPO.body, color: COLORS.TEXT_MAIN }),
                          ],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),

                    // Weightage (employee)
                    new TableCell({
                      width:   { size: COLS.performance[2], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [
                            new TextRun({ text: f.weightage.toString(), size: TYPO.bodySmall, color: COLORS.TEXT_MAIN }),
                          ],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),

                    // Reviewer score
                    new TableCell({
                      width:   { size: COLS.performance[3], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [
                            new TextRun({ text: f.reviewerScore.toString(), size: TYPO.body, color: COLORS.TEXT_MAIN }),
                          ],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),

                    // Weightage (reviewer)
                    new TableCell({
                      width:   { size: COLS.performance[4], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [
                            new TextRun({ text: f.weightage.toString(), size: TYPO.bodySmall, color: COLORS.TEXT_MAIN }),
                          ],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),

                    // Reviewer comments
                    new TableCell({
                      width:   { size: COLS.performance[5], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: { ...bodySpacing },
                          children: [new TextRun({ text: f.comments || "", size: TYPO.body })],
                        }),
                      ],
                      verticalAlign: VerticalAlign.TOP,
                      margins: CELL.standard,
                    }),
                  ],
                }),
              ),

              // TOTAL SCORE row — navy bg spans cols 1-2, gray score cells, tint competency note
              new TableRow({
                children: [
                  // Cols 0-1 merged: TOTAL SCORE label (right-aligned, navy bg)
                  new TableCell({
                    columnSpan: 2,
                    width: { size: COLS.performance[0] + COLS.performance[1], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        spacing:   { ...bodySpacing },
                        children: [
                          new TextRun({ text: "TOTAL SCORE", bold: true, color: COLORS.WHITE, size: TYPO.body }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  // Employee total (gray bg)
                  new TableCell({
                    width:   { size: COLS.performance[2], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [new TextRun({ text: formatScore(stats.employeeFactorScore), bold: true, size: TYPO.body, color: COLORS.TEXT_MAIN })],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  // Cols 3: NAVY bg (empty separator)
                  new TableCell({
                    width:   { size: COLS.performance[3], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                    children: [new Paragraph({ spacing: { ...bodySpacing } })],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  // Reviewer total (gray bg)
                  new TableCell({
                    width:   { size: COLS.performance[4], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [new TextRun({ text: formatScore(stats.reviewerFactorScore), bold: true, size: TYPO.body, color: COLORS.TEXT_MAIN })],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  // Competency note (tint-soft bg)
                  new TableCell({
                    width:   { size: COLS.performance[5], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.TINT_SOFT, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [
                          new TextRun({
                            text:    "Competency Level: 1 = Poor | 5 = Excellent",
                            size:    TYPO.factorDesc,
                            italics: true,
                            color:   COLORS.PRIMARY,
                          }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.factor,
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.sectionBreak } }),

          // ── GROWTH SCORE SUMMARY (narrow, left-indented) ─────────────────
          new Table({
            width:  { size: TABLE.summaryNarrow.size, type: WidthType.DXA },
            alignment: AlignmentType.CENTER,
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("Score Component", COLS.summaryNarrow[0], COLORS.NAVY_ALT),
                  createHeaderCell("Score",           COLS.summaryNarrow[1], COLORS.NAVY_ALT),
                  createHeaderCell("Weightage",       COLS.summaryNarrow[2], COLORS.NAVY_ALT),
                ],
              }),
              createSummaryRow("Employee Score", formatScore(stats.employeeFactorScore), "0.3", COLS.summaryNarrow),
              createSummaryRow("Reviewer Score", formatScore(stats.reviewerFactorScore), "0.7", COLS.summaryNarrow),
              // Growth Score (highlighted row)
              new TableRow({
                children: [
                  new TableCell({
                    width:   { size: COLS.summaryNarrow[0], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.TINT_SOFT, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        spacing: { ...bodySpacing },
                        children: [
                          new TextRun({ text: "Growth Score", bold: true, size: TYPO.body, color: COLORS.NAVY }),
                        ],
                      }),
                    ],
                    margins: CELL.summaryLabel,
                  }),
                  createDataCell(formatScore(stats.performanceScore), true, COLS.summaryNarrow[1]),
                  createDataCell("1.0", false, COLS.summaryNarrow[2]),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.sectionBreak } }),

          // ── EMPLOYEE KEY ACHIEVEMENT GOALS ────────────────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              createBannerRow("EMPLOYEE KEY ACHIEVEMENT GOALS", COLS.goals.length, COLORS.PRIMARY),
              new TableRow({
                children: [
                  createHeaderCell("Goal Description", COLS.goals[0], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Score",            COLS.goals[1], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Weightage",        COLS.goals[2], COLORS.PRIMARY, { noTopBorder: true }),
                  createHeaderCell("Comments",         COLS.goals[3], COLORS.PRIMARY, { noTopBorder: true }),
                ],
              }),
              // Goal rows — alternating fill
              ...data.goals.map((g, idx) =>
                new TableRow({
                  children: [
                    new TableCell({
                      width:   { size: COLS.goals[0], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: { ...bodySpacing },
                          children: [
                            new TextRun({
                              text:  `${String(g.id).padStart(2, "0")}. `,
                              bold:  true,
                              color: COLORS.ORANGE,
                              size:  TYPO.body,
                            }),
                            new TextRun({
                              text:  g.description || "",
                              size:  TYPO.body,
                              color: COLORS.TEXT_MAIN,
                            }),
                          ],
                        }),
                      ],
                      margins: CELL.factor,
                    }),
                    new TableCell({
                      width:   { size: COLS.goals[1], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [new TextRun({ text: g.score.toString(), size: TYPO.body, color: COLORS.TEXT_MAIN })],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),
                    new TableCell({
                      width:   { size: COLS.goals[2], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing:   { ...bodySpacing },
                          children: [new TextRun({ text: g.weightage.toString(), size: TYPO.bodySmall, color: COLORS.TEXT_MAIN })],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: CELL.standard,
                    }),
                    new TableCell({
                      width:   { size: COLS.goals[3], type: WidthType.DXA },
                      borders: B.cell,
                      shading: { fill: idx % 2 === 0 ? COLORS.WHITE : COLORS.TINT_WARM, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: { ...bodySpacing },
                          children: [new TextRun({ text: g.comments || "", size: TYPO.body })],
                        }),
                      ],
                      verticalAlign: VerticalAlign.TOP,
                      margins: CELL.standard,
                    }),
                  ],
                }),
              ),
              // Goal total row
              new TableRow({
                children: [
                  new TableCell({
                    columnSpan: 2,
                    width: { size: COLS.goals[0] + COLS.goals[1], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        spacing:   { ...bodySpacing },
                          children: [new TextRun({ text: "Score", bold: true, color: COLORS.TEXT_MAIN, size: TYPO.body })],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  new TableCell({
                    width:   { size: COLS.goals[2], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.NAVY_ALT, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [new TextRun({ text: formatScore(stats.goalAchievementScore), bold: true, color: COLORS.WHITE, size: TYPO.body })],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                  createDataCell("", false, COLS.goals[3]),
                ],
              }),
            ],
          }),

          // Goal footnote paragraphs
          new Paragraph({
            spacing: { ...SPACING.goalNote1, ...bodySpacing },
            children: [
              new TextRun({
                text:    "* Goal descriptions filled jointly by employee and reviewer; scored by reviewer only.",
                italics: true,
                size:    TYPO.footerNote,
                color:   COLORS.TEXT_MUTED,
              }),
              new TextRun({
                text:    "  ",
                size:    TYPO.footerNote,
              }),
              new TextRun({
                text:    "Competency Level: 1 = Poor | 5 = Excellent",
                italics: true,
                size:    TYPO.footerNote,
                color:   COLORS.PRIMARY,
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.sectionBreak } }),

          // ── FINAL SCORE CALCULATION TABLE ─────────────────────────────────
          new Table({
            width:  { size: TABLE.finalCalc.size, type: WidthType.DXA },
            indent: { size: TABLE.indentNarrow5000.size, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              new TableRow({
                children: [
                  createHeaderCell("Score Component",       COLS.finalCalc[0], COLORS.NAVY_ALT),
                  createHeaderCell("Score",                 COLS.finalCalc[1], COLORS.NAVY_ALT),
                  createHeaderCell("Weightage",             COLS.finalCalc[2], COLORS.NAVY_ALT),
                ],
              }),
              createSummaryRow("Growth Score",          formatScore(stats.performanceScore),      "0.8", COLS.finalCalc),
              createSummaryRow("Goal(s) Achievement Score",  formatScore(stats.goalAchievementScore),  "0.2", COLS.finalCalc),
              // FINAL SCORE row: NAVY label + PRIMARY value cells
              new TableRow({
                children: [
                  new TableCell({
                    width:   { size: COLS.finalCalc[0], type: WidthType.DXA },
                    borders: B.none,
                    shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        spacing: { ...bodySpacing },
                        children: [
                          new TextRun({ text: "FINAL SCORE", bold: true, color: COLORS.WHITE, size: TYPO.body }),
                        ],
                      }),
                    ],
                    margins: CELL.summaryLabel,
                  }),
                  new TableCell({
                    columnSpan: 2,
                    width: { size: COLS.finalCalc[1] + COLS.finalCalc[2], type: WidthType.DXA },
                    borders: B.none,
                    shading: { fill: COLORS.PRIMARY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [
                          new TextRun({ text: formatScore(stats.finalScore), bold: true, color: COLORS.WHITE, size: TYPO.body }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                ],
              }),
              // Growth Category row
              new TableRow({
                children: [
                  new TableCell({
                    width:   { size: COLS.finalCalc[0], type: WidthType.DXA },
                    borders: B.cell,
                    shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
                    children: [
                      new Paragraph({
                        spacing: { ...bodySpacing },
                        children: [new TextRun({ text: "Growth Category", bold: true, size: TYPO.body, color: COLORS.TEXT_MAIN })],
                      }),
                    ],
                    margins: CELL.summaryLabel,
                  }),
                  new TableCell({
                    columnSpan: 2,
                    width: { size: COLS.finalCalc[1] + COLS.finalCalc[2], type: WidthType.DXA },
                    borders: B.cell,
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { ...bodySpacing },
                        children: [
                          new TextRun({
                            text:  stats.ratingCategory,
                            bold:  true,
                            color: COLORS.PRIMARY,
                            size:  TYPO.factorTitle,
                          }),
                        ],
                      }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    margins: CELL.standard,
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.emptyLine } }),

          // ── COMMENTS SECTION ──────────────────────────────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              createBannerRow("COMMENTS", COLS.triple.length, COLORS.PRIMARY),
              // Header row: Employee + Reviewer = NAVY, CEO = PRIMARY
              new TableRow({
                children: [
                  createHeaderCell("Comments from Employee", COLS.triple[0], COLORS.NAVY_ALT, { noTopBorder: true }),
                  createHeaderCell("Comments from Reviewer", COLS.triple[1], COLORS.NAVY_ALT, { noTopBorder: true }),
                  createHeaderCell("Comments from CEO",      COLS.triple[2], COLORS.PRIMARY, { noTopBorder: true }),
                ],
              }),
              // Content row: CEO cell gets tint fill + red borders
              new TableRow({
                height: { value: ROW.contentTall, rule: HeightRule.ATLEAST },
                children: [
                  createMultiLineCell(data.comments.employee, COLS.triple[0], COLORS.WHITE,     B.cell),
                  createMultiLineCell(data.comments.reviewer, COLS.triple[1], COLORS.WHITE,     B.cell),
                  createMultiLineCell(data.comments.ceo,      COLS.triple[2], COLORS.TINT_WARM, B.ceoBorder),
                ],
              }),
            ],
          }),

          new Paragraph({ text: "", spacing: { ...SPACING.emptyLine } }),

          // ── SIGNATURES SECTION ────────────────────────────────────────────
          new Table({
            width:  { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
            layout: TableLayoutType.FIXED,
            borders: B.grid,
            rows: [
              createBannerRow("SIGNATURES", COLS.triple.length, COLORS.PRIMARY),
              // Header row: Employee + Reviewer = NAVY_ALT, CEO = PRIMARY
              new TableRow({
                children: [
                  createHeaderCell("Employee Signature",             COLS.triple[0], COLORS.NAVY_ALT, { noTopBorder: true }),
                  createHeaderCell("Reviewer / Team Lead Signature", COLS.triple[1], COLORS.NAVY_ALT, { noTopBorder: true }),
                  createHeaderCell("CEO Signature",                  COLS.triple[2], COLORS.PRIMARY, { noTopBorder: true }),
                ],
              }),
              // Signature content row
              new TableRow({
                height: { value: ROW.contentTall, rule: HeightRule.ATLEAST },
                children: [
                  createSignatureCell(data.signatures.employee.name, data.signatures.employee.date, COLS.triple[0], COLORS.WHITE,     B.cell),
                  createSignatureCell(data.signatures.reviewer.name, data.signatures.reviewer.date, COLS.triple[1], COLORS.WHITE,     B.cell),
                  createSignatureCell(data.signatures.ceo.name,      data.signatures.ceo.date,      COLS.triple[2], COLORS.TINT_WARM, B.ceoBorder),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  return Packer.toBlob(doc);
}

// ─── Helper functions ──────────────────────────────────────────────────────────

/** Info grid row — label cells (LIGHT_GRAY bg, NAVY_ALT text) + value cells (white bg) */
function createInfoRow(label1: string, val1: string, label2: string, val2: string) {
  const [w1, v1, w2, v2] = COLS.info;

  const labelCell = (text: string, width: number) =>
    new TableCell({
      width:   { size: width, type: WidthType.DXA },
      borders: B.cell,
      shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
      children: [
        new Paragraph({
          spacing: { ...bodySpacing },
          children: [
            new TextRun({ text, bold: true, size: TYPO.body, color: COLORS.NAVY_ALT }),
          ],
        }),
      ],
      margins:       CELL.standard,
      verticalAlign: VerticalAlign.CENTER,
    });

  const valueCell = (text: string, width: number) =>
    new TableCell({
      width:   { size: width, type: WidthType.DXA },
      borders: B.cell,
      shading: { fill: COLORS.WHITE, type: ShadingType.CLEAR },
      children: [
        new Paragraph({
          spacing: { ...bodySpacing },
          children: [new TextRun({ text: text || "", size: TYPO.body, color: COLORS.TEXT_MAIN })],
        }),
      ],
      margins:       CELL.standard,
      verticalAlign: VerticalAlign.CENTER,
    });

  return new TableRow({
    children: [
      labelCell(label1, w1),
      valueCell(val1,   v1),
      labelCell(label2, w2),
      valueCell(val2,   v2),
    ],
  });
}

/**
 * Template banner row.
 * In the reference XML this is a `w:tr` inside the same table, using `w:gridSpan` to span all columns.
 */
function createBannerRow(text: string, colSpan: number, fill: string = COLORS.PRIMARY) {
  return new TableRow({
    height: { value: ROW.redBanner, rule: HeightRule.ATLEAST },
    children: [
      new TableCell({
        columnSpan: colSpan,
        // Add an explicit bottom rule so the fill can't visually bleed below the border in Word.
        borders: {
          ...B.none,
          bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.BORDER },
        },
        shading: { fill, type: ShadingType.CLEAR },
        margins: CELL.redBand,
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { ...bodySpacing },
            children: [new TextRun({ text, bold: true, color: COLORS.WHITE, size: TYPO.redBand })],
          }),
        ],
      }),
    ],
  });
}

/** Colored header cell (used for table column headings) */
function createHeaderCell(text: string, widthDxa: number, fill?: string, opts?: { noTopBorder?: boolean }) {
  const noTopBorder = opts?.noTopBorder === true;
  return new TableCell({
    width:   { size: widthDxa, type: WidthType.DXA },
    borders: noTopBorder
      ? { ...B.cell, top: { style: BorderStyle.NONE, size: 0, color: COLORS.BORDER } }
      : B.cell,
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing:   { ...bodySpacing },
        children: [
          new TextRun({ text, bold: true, color: COLORS.WHITE, size: TYPO.tableHeader }),
        ],
      }),
    ],
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL.standard,
  });
}

/** Plain centered data cell */
function createDataCell(text: string, isBold: boolean = false, widthDxa?: number) {
  return new TableCell({
    ...(widthDxa !== undefined ? { width: { size: widthDxa, type: WidthType.DXA } } : {}),
    borders: B.cell,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing:   { ...bodySpacing },
        children: [
          new TextRun({ text: text || " ", bold: isBold, size: TYPO.body, color: COLORS.TEXT_MAIN }),
        ],
      }),
    ],
    verticalAlign: VerticalAlign.CENTER,
    margins: CELL.standard,
  });
}

/** Summary table row — label cell (tint-soft bg) + two centered value cells */
function createSummaryRow(
  label: string,
  score: string,
  weight: string,
  cols: readonly [number, number, number],
) {
  return new TableRow({
    children: [
      new TableCell({
        width:   { size: cols[0], type: WidthType.DXA },
        borders: B.cell,
        shading: { fill: COLORS.TINT_SOFT, type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            spacing: { ...bodySpacing },
            children: [new TextRun({ text: label, bold: true, size: TYPO.body, color: COLORS.NAVY })],
          }),
        ],
        margins: CELL.summaryLabel,
        verticalAlign: VerticalAlign.CENTER,
      }),
      createDataCell(score,  false, cols[1]),
      createDataCell(weight, false, cols[2]),
    ],
  });
}

/** Multi-line text cell (comments sections) — accepts custom fill and border set */
type BorderSet = typeof B.cell | typeof B.ceoBorder | typeof B.none;
function createMultiLineCell(
  text: string,
  widthDxa: number,
  fill: string = COLORS.WHITE,
  borders: BorderSet = B.cell,
) {
  return new TableCell({
    width:   { size: widthDxa, type: WidthType.DXA },
    borders,
    shading: { fill, type: ShadingType.CLEAR },
    children: [
      new Paragraph({
        spacing: { ...SPACING.commentCell, ...bodySpacing },
        children: [new TextRun({ text: text || "", size: TYPO.body, color: COLORS.TEXT_MAIN })],
      }),
    ],
    verticalAlign: VerticalAlign.TOP,
    margins: CELL.standard,
  });
}

/**
 * Signature cell.
 * Template structure: blank top paragraph, then signature line (before=800),
 * then date line — all in 999999 grey, sz=16 (8pt).
 */
function createSignatureCell(
  name: string,
  date: string,
  widthDxa: number,
  fill: string = COLORS.WHITE,
  borders: BorderSet = B.cell,
) {
  return new TableCell({
    width:   { size: widthDxa, type: WidthType.DXA },
    borders,
    shading: { fill, type: ShadingType.CLEAR },
    children: [
      // Blank space at top
      new Paragraph({ spacing: { ...bodySpacing } }),
      // Name line (if provided)
      ...(name
        ? [new Paragraph({
            spacing: { ...bodySpacing },
            children: [new TextRun({ text: name, size: TYPO.body, color: COLORS.TEXT_MAIN })],
          })]
        : []),
      // Signature underline — pushed down with before=800
      new Paragraph({
        spacing: { before: SPACING.signaturePara1.before, ...bodySpacing },
        children: [
          new TextRun({ text: "Signature: ___________________", size: 16, color: COLORS.SIG_TEXT }),
        ],
      }),
      // Date line
      new Paragraph({
        spacing: { ...SPACING.signaturePara2, ...bodySpacing },
        children: [
          new TextRun({ text: `Date: ${date || "________________________"}`, size: 16, color: COLORS.SIG_TEXT }),
        ],
      }),
    ],
    verticalAlign: VerticalAlign.TOP,
    margins: CELL.standard,
  });
}