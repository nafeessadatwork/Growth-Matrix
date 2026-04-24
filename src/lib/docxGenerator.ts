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
  } from "docx";
  import { AppraisalData } from "../types";
  import { calculateAppraisal, formatScore } from "./utils";
  
  const COLORS = {
    PRIMARY: "E8351A", // Red (template)
    ORANGE: "FF6B2B", // Orange (template)
    NAVY: "0C1A2E",   // Navy (template)
    NAVY_ALT: "1A2F4A", // Secondary navy (template)
    WHITE: "FFFFFF",
    LIGHT_GRAY: "F5F5F5",
    BORDER: "999999",
    TEXT_MAIN: "1A1A2E",
    TEXT_MUTED: "777777",
  };
  
  export async function generateAppraisalDoc(data: AppraisalData) {
    const stats = calculateAppraisal(data);
    const logoUrl = new URL("../assets/mindx-logo.png", import.meta.url).href;
    const logoBytes = new Uint8Array(await (await fetch(logoUrl)).arrayBuffer());
  
    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1440,
                bottom: 1440,
                left: 1440,
                right: 1440,
              },
            },
          },
          children: [
            // Header Section (Logo Left, Dark Box Right) - match template
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideHorizontal: { style: BorderStyle.NONE },
                insideVertical: { style: BorderStyle.NONE },
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 30, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new ImageRun({
                              data: logoBytes,
                              type: "png",
                              transformation: { width: 140, height: 90 },
                            }),
                          ],
                          spacing: { before: 120, after: 120 },
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                      width: { size: 70, type: WidthType.PERCENTAGE },
                      shading: {
                        fill: COLORS.NAVY,
                        type: ShadingType.CLEAR,
                      },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "STAFF PERFORMANCE",
                              bold: true,
                              size: 42,
                              color: COLORS.WHITE,
                            }),
                          ],
                          spacing: { before: 40, after: 0 },
                        }),
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "FEEDBACK",
                              bold: true,
                              size: 42,
                              color: COLORS.WHITE,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "FORM",
                              bold: true,
                              size: 42,
                              color: COLORS.WHITE,
                            }),
                          ],
                          spacing: { after: 80 },
                        }),
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "Technology with a Human Touch • mindxtech.ai",
                              size: 18,
                              italics: true,
                              color: COLORS.ORANGE,
                            }),
                          ],
                          spacing: { after: 120 },
                        }),
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "MIND X SDN BHD • B-16-1, Tower B, Vertical Business Suite, Avenue 3, Bangsar South, 59200 Kuala Lumpur",
                              size: 16,
                              color: COLORS.WHITE,
                            }),
                          ],
                          spacing: { after: 60 },
                        }),
                        new Paragraph({
                          alignment: AlignmentType.LEFT,
                          children: [
                            new TextRun({
                              text: "xplorers@mindxtech.ai • +60 3 6043 1700",
                              size: 16,
                              color: COLORS.WHITE,
                            }),
                          ],
                        }),
                      ],
                      verticalAlign: VerticalAlign.CENTER,
                      margins: { top: 180, bottom: 180, left: 220, right: 220 },
                    }),
                  ],
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 300 } }),
  
            // Info Grid
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                createInfoRow("Employee Name", data.employee.name, "Reviewer Name", data.reviewer.name),
                createInfoRow("Employee Position", data.employee.position, "Reviewer Position", data.reviewer.position),
                createInfoRow("Department", data.employee.department, "Department", data.reviewer.department),
                createInfoRow("Type of Appraisal", data.employee.appraisalType, "Projects Managed", data.reviewer.projectsManaged),
                createInfoRow("Appraisal Period", data.employee.appraisalPeriod, "Appraisal Due", data.reviewer.appraisalDue),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 300 } }),
  
            // Performance Appraisal Header
            createRedHeader("PERFORMANCE APPRAISAL"),
  
            // Performance Factors Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Performance Factor", 20, COLORS.PRIMARY),
                    createHeaderCell("Employee Score", 14, COLORS.PRIMARY),
                    createHeaderCell("Weightage", 13, COLORS.PRIMARY),
                    createHeaderCell("Reviewer Score", 14, COLORS.PRIMARY),
                    createHeaderCell("Weightage", 13, COLORS.PRIMARY),
                    createHeaderCell("Reviewer's Comments", 26, COLORS.PRIMARY),
                  ],
                }),
                ...data.factors.map((f) => 
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({ text: `${f.id}. ${f.title}`, bold: true, color: COLORS.PRIMARY, size: 18 }),
                            ],
                          }),
                          new Paragraph({
                            children: [new TextRun({ text: f.description, size: 14, color: COLORS.TEXT_MUTED })],
                          }),
                        ],
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                      }),
                      createDataCell(f.employeeScore.toString()),
                      createDataCell(f.weightage.toString()),
                      createDataCell(f.reviewerScore.toString()),
                      createDataCell(f.weightage.toString()),
                      createDataCell(f.comments),
                    ],
                  })
                ),
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [new TextRun({ text: "TOTAL SCORE", bold: true, color: COLORS.WHITE, size: 18 })],
                        }),
                      ],
                    }),
                    createDataCell(formatScore(stats.employeeFactorScore), true),
                    createDataCell("1.00"),
                    createDataCell(formatScore(stats.reviewerFactorScore), true),
                    createDataCell("1.00"),
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({ text: "Competency Level: 1 = Poor | 5 = Excellent", size: 14, italics: true, color: COLORS.PRIMARY }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 300 } }),
  
            // Scoring Summary
            new Table({
              width: { size: 40, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Score Component", 50, COLORS.NAVY),
                    createHeaderCell("Score", 25, COLORS.NAVY),
                    createHeaderCell("Weightage", 25, COLORS.NAVY),
                  ],
                }),
                createSummaryRow("Employee Score", formatScore(stats.employeeFactorScore), "0.3"),
                createSummaryRow("Reviewer Score", formatScore(stats.reviewerFactorScore), "0.7"),
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: COLORS.LIGHT_GRAY, type: ShadingType.CLEAR },
                      children: [new Paragraph({ children: [new TextRun({ text: "Performance Score", bold: true })] })],
                    }),
                    createDataCell(formatScore(stats.performanceScore), true),
                    createDataCell("1.0"),
                  ],
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 300 } }),
  
            // Goal Achievement Header
            createRedHeader("EMPLOYEE KEY ACHIEVEMENT GOALS"),
  
            // Goals Table
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Goal Description", 53, COLORS.PRIMARY),
                    createHeaderCell("Score", 14, COLORS.PRIMARY),
                    createHeaderCell("Weightage", 13, COLORS.PRIMARY),
                    createHeaderCell("Comments", 20, COLORS.PRIMARY),
                  ],
                }),
                ...data.goals.map((g) => 
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({ children: [new TextRun({ text: `${g.id}. ${g.description || ""}`, size: 16 })] })
                        ],
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                      }),
                      createDataCell(g.score.toString()),
                      createDataCell(g.weightage.toString()),
                      createDataCell(g.comments),
                    ],
                  })
                ),
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: COLORS.NAVY, type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [new TextRun({ text: "Score", bold: true, color: COLORS.WHITE, size: 18 })],
                        }),
                      ],
                    }),
                    createDataCell(formatScore(stats.goalAchievementScore), true),
                    createDataCell("1.00"),
                    createDataCell(""),
                  ],
                }),
              ],
            }),
  
            new Paragraph({
              spacing: { before: 120, after: 60 },
              children: [
                new TextRun({
                  text: "* Goal descriptions filled jointly by employee and reviewer; scored by reviewer only.",
                  italics: true,
                  size: 16,
                  color: COLORS.TEXT_MUTED,
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 0, after: 240 },
              children: [
                new TextRun({
                  text: "Competency Level: 1 = Poor | 5 = Excellent",
                  italics: true,
                  size: 16,
                  color: COLORS.TEXT_MUTED,
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 300 } }),
  
            // Final Calculation Table
            new Table({
              width: { size: 40, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Score Component", 50, COLORS.NAVY),
                    createHeaderCell("Score", 25, COLORS.NAVY),
                    createHeaderCell("Weightage", 25, COLORS.NAVY),
                  ],
                }),
                createSummaryRow("Performance Score", formatScore(stats.performanceScore), "0.8"),
                createSummaryRow("Goal(s) Achievement Score", formatScore(stats.goalAchievementScore), "0.2"),
                new TableRow({
                  children: [
                    new TableCell({
                      shading: { fill: COLORS.PRIMARY, type: ShadingType.CLEAR },
                      children: [new Paragraph({ children: [new TextRun({ text: "FINAL SCORE", bold: true, color: COLORS.WHITE })] })],
                    }),
                    createHeaderCell(formatScore(stats.finalScore), 25, COLORS.PRIMARY),
                    createHeaderCell("1.0", 25, COLORS.PRIMARY),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ shading: { fill: COLORS.LIGHT_GRAY }, children: [new Paragraph({ children: [new TextRun({ text: "Performance Category", bold: true })] })] }),
                    new TableCell({
                      columnSpan: 2,
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [new TextRun({ text: stats.ratingCategory, bold: true, color: COLORS.PRIMARY, size: 18 })],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 400 } }),
  
            // Comments
            createRedHeader("COMMENTS"),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Comments from Employee", 34, COLORS.NAVY),
                    createHeaderCell("Comments from Reviewer", 33, COLORS.NAVY),
                    createHeaderCell("Comments from CEO", 28, COLORS.NAVY),
                  ],
                }),
                new TableRow({
                  children: [
                    createMultiLineCell(data.comments.employee),
                    createMultiLineCell(data.comments.reviewer),
                    createMultiLineCell(data.comments.ceo),
                  ],
                }),
              ],
            }),
  
            new Paragraph({ text: "", spacing: { after: 400 } }),
  
            // Signatures
            createRedHeader("SIGNATURES"),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              layout: TableLayoutType.FIXED,
              rows: [
                new TableRow({
                  children: [
                    createHeaderCell("Employee Signature", 33, COLORS.NAVY),
                    createHeaderCell("Reviewer / Team Lead Signature", 33, COLORS.NAVY),
                    createHeaderCell("CEO Signature", 33, COLORS.NAVY),
                  ],
                }),
                new TableRow({
                  children: [
                    createSignatureCell(data.signatures.employee.name, data.signatures.employee.date),
                    createSignatureCell(data.signatures.reviewer.name, data.signatures.reviewer.date),
                    createSignatureCell(data.signatures.ceo.name, data.signatures.ceo.date),
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
  
  function createInfoRow(label1: string, val1: string, label2: string, val2: string) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 16, type: WidthType.PERCENTAGE },
          shading: { fill: COLORS.LIGHT_GRAY },
          children: [new Paragraph({ children: [new TextRun({ text: label1, bold: true, size: 16 })] })],
          margins: { left: 100, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({ 
          width: { size: 34, type: WidthType.PERCENTAGE },
          children: [new Paragraph({ children: [new TextRun({ text: val1, size: 16 })] })],
          margins: { left: 100, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({
          width: { size: 16, type: WidthType.PERCENTAGE },
          shading: { fill: COLORS.LIGHT_GRAY },
          children: [new Paragraph({ children: [new TextRun({ text: label2, bold: true, size: 16 })] })],
          margins: { left: 100, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
        }),
        new TableCell({ 
          width: { size: 34, type: WidthType.PERCENTAGE },
          children: [new Paragraph({ children: [new TextRun({ text: val2, size: 16 })] })],
          margins: { left: 100, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
        }),
      ],
    });
  }
  
  function createRedHeader(text: string) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLORS.PRIMARY, type: ShadingType.CLEAR },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text, bold: true, color: COLORS.WHITE, size: 22 })],
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
            }),
          ],
          height: { value: 450, rule: "atLeast" },
        }),
      ],
    });
  }
  
  function createHeaderCell(text: string, widthPct: number, fill?: string) {
    return new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      shading: fill
        ? { fill, type: ShadingType.CLEAR }
        : undefined,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: true, color: COLORS.WHITE, size: 16 })],
        }),
      ],
      verticalAlign: VerticalAlign.CENTER,
    });
  }
  
  function createDataCell(text: string, isBold: boolean = false) {
    return new TableCell({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: text || " ", bold: isBold, size: 16 })],
        }),
      ],
      verticalAlign: VerticalAlign.CENTER,
    });
  }
  
  function createSummaryRow(label: string, score: string, weight: string) {
    return new TableRow({
      children: [
        new TableCell({ 
          children: [new Paragraph({ children: [new TextRun({ text: label, size: 16 })] })],
          margins: { left: 100 },
        }),
        createDataCell(score),
        createDataCell(weight),
      ],
    });
  }
  
  function createMultiLineCell(text: string) {
    return new TableCell({
      children: [
        new Paragraph({
          spacing: { before: 120, after: 120 },
          children: [new TextRun({ text: text || "N/A", size: 16 })],
        }),
      ],
      verticalAlign: VerticalAlign.TOP,
      margins: { left: 100, right: 100 },
    });
  }
  
  function createSignatureCell(name: string, date: string) {
    return new TableCell({
      children: [
        new Paragraph({ spacing: { before: 200 } }),
        new Paragraph({ children: [new TextRun({ text: "Signature: ___________________", size: 16 })] }),
        new Paragraph({ spacing: { after: 100 } }),
        new Paragraph({ children: [new TextRun({ text: `Date: ${date}`, size: 16 })] }),
        new Paragraph({ spacing: { after: 200 } }),
      ],
      verticalAlign: VerticalAlign.TOP,
    });
  }
  