import { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';
import { SchoolData, TestPhase, ReportConfig } from '../types';
import { expandSchoolName } from './formatters';

export const generateWordReport = async (
  school: SchoolData,
  phase: TestPhase,
  config: ReportConfig,
  chartImages: { grade3: string; grade4: string; grade5: string }
) => {
  let logoBuffer: ArrayBuffer | null = null;
  try {
    const response = await fetch('/pdf-logo.png');
    logoBuffer = await response.arrayBuffer();
  } catch (e) {
    console.warn("Could not fetch logo for Word export");
  }

  // Helpers
  const getSchoolTitle = () => {
    const fullSchoolName = expandSchoolName(school.School_Name);
    return `${fullSchoolName}`;
  };

  const getFigText = (grade: 3 | 4 | 5) => {
    const classData = school.classes[grade];
    if (!classData) return '';

    const currentPhaseData = phase === 'Baseline' ? classData.baseline : phase === 'Midline' ? classData.midline : classData.endline;
    const avgTotal = currentPhaseData.Total;

    if (phase === 'Endline') {
      const baseline = classData.baseline.Total;
      const percentage = baseline ? (((avgTotal - baseline) / baseline) * 100).toFixed(0) : 0;
      return `Fig ${grade - 2}: A total of ${classData.assessedCount} students were assessed, the English Language Literacy level is increased from baseline to endline ${percentage}%`;
    }
    return `Fig ${grade - 2}: A total of ${classData.assessedCount} students were assessed, the average English Language Literacy level is ${avgTotal.toFixed(2)}`;
  };

  const aboutText = config.phases[phase].aboutText;

  // Convert base64 data URIs to ArrayBuffer
  const base64ToArrayBuffer = (base64: string) => {
    if (!base64 || !base64.includes(',')) return null;
    const base64Data = base64.split(',')[1];
    const binaryString = window.atob(base64Data);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  };

  const grade3Buffer = base64ToArrayBuffer(chartImages.grade3);
  const grade4Buffer = base64ToArrayBuffer(chartImages.grade4);
  const grade5Buffer = base64ToArrayBuffer(chartImages.grade5);

  // Chart rendering helper
  const shortName = school.School_Name.replace(/\s+/g, '_').toUpperCase();

  const noBorders = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  const createHeaderCell = (title: string, color: string) => {
    return new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      shading: { fill: color },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: title, bold: true, size: 28 })],
          spacing: { before: 80, after: 80 }
        })
      ],
      borders: noBorders
    });
  };

  const createChartCell = (grade: number, buffer: ArrayBuffer | null, figText: string) => {
    const cellContents = [];
    
    // Image
    if (buffer) {
      cellContents.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: buffer,
            transformation: {
              width: 290,
              height: 175,
            },
            type: 'png'
          } as any),
        ],
      }));
    } else {
      cellContents.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "No data available", color: "888888" })],
      }));
    }

    // Fig text
    cellContents.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: figText, italics: true, size: 20 })], // 10pt
      spacing: { before: 100, after: 200 }
    }));

    return new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      children: cellContents,
      borders: noBorders
    });
  };

  const getNextStepsCell = () => {
    const contents = [];

    if (phase === 'Baseline') {
      for (let i = 0; i < 4; i++) {
        contents.push(new Paragraph({
          children: [new TextRun({ text: "• __________________________________", size: 24 })],
          spacing: { before: 200, after: 200 }
        }));
      }
    } else {
      config.phases[phase].nextStepsBullets.forEach(step => {
        if (!step) return;
        contents.push(new Paragraph({
          alignment: phase === 'Endline' ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
          children: [
            new TextRun({ 
              text: phase === 'Endline' ? step : "• " + step, 
              size: 24,
              bold: phase === 'Endline'
            })
          ],
          spacing: { before: 150, after: 150 }
        }));
      });
    }

    return new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      children: contents,
      borders: noBorders
    });
  };

  // Build document
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
        }
      },
      children: [
        // Header Table for Title and Logo
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 85, type: WidthType.PERCENTAGE },
                  borders: noBorders,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun({
                          text: getSchoolTitle(),
                          bold: true,
                          italics: true,
                          size: 40, // 20pt
                        }),
                      ],
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun({
                          text: `${school.Mandal ? `${school.Mandal} Mandal, ` : ''}${school.District} District.`,
                          bold: true,
                          italics: true,
                          size: 32, // 16pt
                        }),
                      ],
                      spacing: { after: 400 },
                    }),
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun({
                          text: `${phase} Report`,
                          bold: true,
                          size: 28, // 14pt
                        }),
                      ],
                      spacing: { after: 200 }
                    }),
                  ]
                }),
                new TableCell({
                  width: { size: 15, type: WidthType.PERCENTAGE },
                  borders: noBorders,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      children: logoBuffer ? [
                        new ImageRun({
                          data: logoBuffer,
                          transformation: { width: 80, height: 80 },
                          type: 'png'
                        } as any)
                      ] : [],
                    })
                  ]
                })
              ]
            })
          ]
        }),
        
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: "- By ", size: 24 }),
            new TextRun({ text: "Teach For Change Trust", bold: true, size: 24 }),
          ],
          spacing: { after: 400 }
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [
            new TextRun({ text: "About: ", bold: true, italics: true, size: 24 }),
            new TextRun({ text: aboutText, italics: true, size: 24 }),
          ],
          spacing: { after: 400 }
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "Results: The grade wise details:", bold: true, size: 24 }),
          ],
          spacing: { after: 200 }
        }),
        
        // Table for Grades 3 and 4
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [
            new TableRow({
              children: [
                createHeaderCell("Grade 3", "A9D18E"),
                createHeaderCell("Grade 4", "A9D18E")
              ]
            }),
            new TableRow({
              children: [
                createChartCell(3, grade3Buffer, getFigText(3)),
                createChartCell(4, grade4Buffer, getFigText(4))
              ]
            })
          ]
        }),
        
        // Spacer
        new Paragraph({ spacing: { before: 200, after: 200 } }),

        // Table for Grade 5 and Next Steps
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: noBorders,
          rows: [
            new TableRow({
              children: [
                createHeaderCell("Grade 5", "A9D18E"),
                createHeaderCell(config.phases[phase].nextStepsTitle, "FFD966")
              ]
            }),
            new TableRow({
              children: [
                createChartCell(5, grade5Buffer, getFigText(5)),
                getNextStepsCell()
              ]
            })
          ]
        })
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  return blob;
};
