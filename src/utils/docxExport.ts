/**
 * Evala Edu - Microsoft Word (.docx) Document Generator Engine
 * Generates standards-compliant OpenXML .docx documents for school & accounting workflows.
 * @license Apache-2.0
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  HeadingLevel,
  ShadingType,
} from 'docx';
import { calculateAmounts, formatCZK, calculateTeacherGrade } from './accountingLogic';

/**
 * Triggers browser download for a Blob
 */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

const BORDER_NONE = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

const BORDER_LIGHT = {
  top: { style: BorderStyle.SINGLE, size: 1, color: 'D1D5DB' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'D1D5DB' },
  left: { style: BorderStyle.SINGLE, size: 1, color: 'D1D5DB' },
  right: { style: BorderStyle.SINGLE, size: 1, color: 'D1D5DB' },
};

/**
 * Generates a full Microsoft Word (.docx) document for any accounting document in the workflow
 */
export async function generateDocumentDocx(docType: string, state: any): Promise<Blob> {
  const calc = calculateAmounts(state);
  const isSupp = state.role === 'supplier';
  const c = state.company || {
    name: 'EDU-Trade s.r.o.',
    street: 'Školní 1420/8',
    city: '110 00 Praha 1',
    ico: '87654321',
    dic: 'CZ87654321',
    bankAccount: '2100123456/2010',
    bankName: 'Fio banka a.s.',
  };
  const p = state.partner || {
    name: 'Stavby & Design s.r.o.',
    street: 'Průmyslová 450',
    city: '102 00 Praha 10',
    ico: '27082440',
    dic: 'CZ27082440',
    bankAccount: '123456789/0800',
  };

  const seller = isSupp ? c : p;
  const buyer = isSupp ? p : c;

  const typeTitles: Record<string, string> = {
    CN: 'CENOVÁ NABÍDKA',
    PCN: 'PŘIJATÁ CENOVÁ NABÍDKA',
    OB: 'PŘIJATÁ OBJEDNÁVKA',
    VO: 'VYDANÁ OBJEDNÁVKA',
    ZF: 'ZÁLOHOVÁ FAKTURA',
    PZF: 'PŘIJATÁ ZÁLOHOVÁ FAKTURA',
    DDPP: 'DAŇOVÝ DOKLAD K PŘIJATÉ PLATBĚ',
    PDPP: 'PŘIJATÝ DAŇOVÝ DOKLAD K PLATBĚ',
    DP: 'DAŇOVÝ DOKLAD K PŘIJATÉ PLATBĚ',
    VF: 'FAKTURA – DAŇOVÝ DOKLAD',
    PF: 'PŘIJATÁ FAKTURA – DAŇOVÝ DOKLAD',
    PPD: 'PŘÍJMOVÝ POKLADNÍ DOKLAD',
    VPD: 'VÝDAJOVÝ POKLADNÍ DOKLAD',
    BV: 'BANKOVNÍ VÝPIS Z ÚČTU',
  };

  const docTitle = typeTitles[docType] || 'ÚČETNÍ DOKLAD';
  const today = new Date().toLocaleDateString('cs-CZ');

  const children: any[] = [];

  // 1. Title Banner
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: docTitle,
          bold: true,
          size: 32,
          color: '0F172A',
        }),
      ],
      spacing: { after: 120 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Evidenční číslo dokladu: `,
          bold: true,
          size: 22,
          color: '0D9488',
        }),
        new TextRun({
          text: `${docType}20260001`,
          bold: true,
          size: 22,
          color: '0F172A',
        }),
        new TextRun({
          text: `   |   Datum vystavení: ${today}   |   Variabilní symbol: 20260001`,
          size: 20,
          color: '64748B',
        }),
      ],
      spacing: { after: 240 },
    })
  );

  // 2. Party Cards (Dodavatel & Odběratel) in a 2-column table
  const partyTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_LIGHT,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: { fill: 'F8FAFC', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: 'DODAVATEL (Poskytovatel plnění):',
                    bold: true,
                    size: 18,
                    color: '0D9488',
                  }),
                ],
                spacing: { after: 60 },
              }),
              new Paragraph({
                children: [new TextRun({ text: seller.name, bold: true, size: 22 })],
                spacing: { after: 40 },
              }),
              new Paragraph({
                children: [new TextRun({ text: `${seller.street}, ${seller.city}`, size: 18 })],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `IČO: ${seller.ico}   |   DIČ: ${seller.dic}`,
                    bold: true,
                    size: 18,
                  }),
                ],
                spacing: { before: 60 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Bankovní účet: ${seller.bankAccount || '2100123456/2010'}`,
                    size: 18,
                    color: '0F172A',
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            shading: { fill: 'FFFFFF', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: 'ODBĚRATEL (Zákazník / Příjemce):',
                    bold: true,
                    size: 18,
                    color: '475569',
                  }),
                ],
                spacing: { after: 60 },
              }),
              new Paragraph({
                children: [new TextRun({ text: buyer.name, bold: true, size: 22 })],
                spacing: { after: 40 },
              }),
              new Paragraph({
                children: [new TextRun({ text: `${buyer.street}, ${buyer.city}`, size: 18 })],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `IČO: ${buyer.ico}   |   DIČ: ${buyer.dic}`,
                    bold: true,
                    size: 18,
                  }),
                ],
                spacing: { before: 60 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Způsob úhrady: ${state.paymentMethod === 'cash' ? 'Hotově v pokladně (211)' : 'Běžný účet (221)'}`,
                    size: 18,
                    color: '0F172A',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(partyTable);

  children.push(
    new Paragraph({
      text: '',
      spacing: { after: 200 },
    })
  );

  // 3. Line Items Table
  const itemsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_LIGHT,
    rows: [
      // Header Row
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 45, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Popis položky / Předmět plnění', bold: true, color: 'FFFFFF', size: 18 }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: 'Základ', bold: true, color: 'FFFFFF', size: 18 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'Sazba DPH', bold: true, color: 'FFFFFF', size: 18 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: 'Celkem s DPH', bold: true, color: 'FFFFFF', size: 18 })],
              }),
            ],
          }),
        ],
      }),
      // Item Row
      new TableRow({
        children: [
          new TableCell({
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: state.caseTitle, bold: true, size: 19 }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: calc.isEur
                      ? `Cizí měna: ${calc.foreignBase.toLocaleString('cs-CZ')} EUR (Kurz: ${calc.invoiceRate} CZK/EUR)`
                      : 'Fakturováno v tuzemské měně CZK',
                    size: 16,
                    color: '64748B',
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: formatCZK(calc.base), bold: true, size: 18 })],
              }),
            ],
          }),
          new TableCell({
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: `${calc.vatRate} %`, size: 18 })],
              }),
            ],
          }),
          new TableCell({
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: formatCZK(calc.grandTotal), bold: true, size: 19, color: '0F172A' })],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(itemsTable);

  // 4. Financial Summary Breakdown Table
  const summaryTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_NONE,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            margins: { top: 120, bottom: 120, left: 60, right: 60 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: 'Zákonné náležitosti dokladu:',
                    bold: true,
                    size: 18,
                    color: '0F172A',
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: '• Splňuje veškeré požadavky § 11 zákona č. 563/1991 Sb., o účetnictví.\n• Plátce DPH eviduje plnění dle § 28 zákona č. 235/2004 Sb., o DPH.',
                    size: 16,
                    color: '64748B',
                  }),
                ],
              }),
              ...(calc.isReverseCharge
                ? [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `• REVERSE CHARGE (§ 92a ZDPH): Daň odvede zákazník. Kód předmětu: ${calc.reverseChargeCode}.`,
                          bold: true,
                          size: 16,
                          color: '4338CA',
                        }),
                      ],
                    }),
                  ]
                : []),
            ],
          }),
          new TableCell({
            width: { size: 45, type: WidthType.PERCENTAGE },
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 120, bottom: 120, left: 140, right: 140 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `Základ daně: `, size: 18 }),
                  new TextRun({ text: formatCZK(calc.base), bold: true, size: 18 }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: `DPH (${calc.vatRate} %): `, size: 18 }),
                  new TextRun({ text: formatCZK(calc.vatTotal), bold: true, size: 18 }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: `Záloha s DPH: `, size: 18 }),
                  new TextRun({ text: formatCZK(calc.advanceGross), bold: true, size: 18 }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `DOPLATEK K ÚHRADĚ: `,
                    bold: true,
                    size: 22,
                    color: '0D9488',
                  }),
                  new TextRun({
                    text: formatCZK(calc.remainingPayableGross),
                    bold: true,
                    size: 24,
                    color: '0F172A',
                  }),
                ],
                spacing: { before: 80 },
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(summaryTable);

  children.push(
    new Paragraph({
      text: '',
      spacing: { after: 300 },
    })
  );

  // 5. Signature Section
  const signatureTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_NONE,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Vystavil / Zpracoval:', bold: true, size: 18 }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `${state.studentName || seller.name}\n\n_________________________________________\nRazítko a podpis vystavitele`,
                    size: 16,
                    color: '475569',
                  }),
                ],
                spacing: { before: 140 },
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({ text: 'Schválil / Přijal:', bold: true, size: 18 }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: `${buyer.name}\n\n_________________________________________\nPodpis odpovědné osoby`,
                    size: 16,
                    color: '475569',
                  }),
                ],
                spacing: { before: 140 },
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(signatureTable);

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  return await Packer.toBlob(doc);
}

/**
 * Generates an official Teacher Evaluation Protocol in Microsoft Word (.docx)
 */
export async function generateTeacherProtocolDocx(state: any): Promise<Blob> {
  const calc = calculateAmounts(state);
  const quizScore = 4; // default or from state
  const evalResult = calculateTeacherGrade(state.examErrors || 0, quizScore);
  const isSupp = state.role === 'supplier';
  const flags = isSupp ? state.stepFlags?.supplier : state.stepFlags?.customer;
  const today = new Date().toLocaleDateString('cs-CZ');

  const children: any[] = [];

  // Header
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: 'VÝSLEDKOVÝ PROTOKOL ZKOUŠKY A ÚČTOVÁNÍ (OA)',
          bold: true,
          size: 28,
          color: '0F172A',
        }),
      ],
      spacing: { after: 100 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: 'Evala Edu – Oficiální certifikovaný výkaz práce studenta Obchodní akademie',
          size: 20,
          color: '0D9488',
        }),
      ],
      spacing: { after: 200 },
    })
  );

  // Student & Case Meta Table
  const metaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_LIGHT,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            shading: { fill: 'F8FAFC', type: ShadingType.CLEAR, color: 'auto' },
            children: [
              new Paragraph({
                children: [new TextRun({ text: 'STUDENT:', bold: true, size: 18, color: '0D9488' })],
              }),
              new Paragraph({
                children: [new TextRun({ text: state.studentName || 'Jméno neuvedeno', bold: true, size: 22 })],
              }),
              new Paragraph({
                children: [new TextRun({ text: `Třída: ${state.studentClass || 'Nespecifikována'}   |   Datum: ${today}`, size: 18 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            margins: { top: 100, bottom: 100, left: 140, right: 140 },
            shading: { fill: 'F8FAFC', type: ShadingType.CLEAR, color: 'auto' },
            children: [
              new Paragraph({
                children: [new TextRun({ text: 'ZADANÝ PŘÍPAD:', bold: true, size: 18, color: '0D9488' })],
              }),
              new Paragraph({
                children: [new TextRun({ text: state.caseTitle, bold: true, size: 20 })],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Role: ${isSupp ? 'Dodavatel' : 'Odběratel'}   |   Hodnota: ${formatCZK(calc.grandTotal)}`,
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  children.push(metaTable);

  // Grade Banner
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: `VÝSLEDNÉ HODNOCENÍ: ${evalResult.gradeText} (${evalResult.percentage} %)`,
          bold: true,
          size: 26,
          color: evalResult.grade <= 2 ? '047857' : 'B91C1C',
        }),
      ],
      spacing: { before: 200, after: 200 },
    })
  );

  // 8 Steps Summary Table
  const stepsData = [
    { num: 1, name: 'Registr partnerů (ARES & DPH)', status: flags?.aresDone ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 2, name: 'Číselníky a řady dokladů', status: flags?.seriesDone ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 3, name: 'Cenová nabídka (CN/PCN)', status: (isSupp ? flags?.cnSent : flags?.cnRec) ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 4, name: 'Objednávka (OB/VO)', status: (isSupp ? flags?.obApproved : flags?.voSent) ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 5, name: 'Zálohová faktura & Pokladna/BÚ', status: flags?.bankMatched ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 6, name: 'Daňový doklad k platbě (DDPP)', status: (isSupp ? flags?.ddppDone : flags?.ddppRec) ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 7, name: 'Konečná vyúčtovací faktura (VF/PF)', status: (isSupp ? flags?.vfDone : flags?.pfDone) ? 'SPLNĚNO ✓' : 'Nesplněno' },
    { num: 8, name: 'Závěrečná úhrada faktury & Saldokonto', status: flags?.finalPaid ? 'SPLNĚNO ✓' : 'Nesplněno' },
  ];

  const stepsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_LIGHT,
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [new Paragraph({ children: [new TextRun({ text: 'Krok', bold: true, color: 'FFFFFF', size: 18 })] })],
          }),
          new TableCell({
            width: { size: 60, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [new Paragraph({ children: [new TextRun({ text: 'Metodický krok oběhu dokladů (OA standard)', bold: true, color: 'FFFFFF', size: 18 })] })],
          }),
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Výsledek studenta', bold: true, color: 'FFFFFF', size: 18 })] })],
          }),
        ],
      }),
      ...stepsData.map(
        (s) =>
          new TableRow({
            children: [
              new TableCell({
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ children: [new TextRun({ text: `Krok ${s.num}`, bold: true, size: 18 })] })],
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [new Paragraph({ children: [new TextRun({ text: s.name, size: 18 })] })],
              }),
              new TableCell({
                margins: { top: 80, bottom: 80, left: 100, right: 100 },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: s.status,
                        bold: true,
                        size: 18,
                        color: s.status.includes('SPLNĚNO') ? '047857' : '94A3B8',
                      }),
                    ],
                  }),
                ],
              }),
            ],
          })
      ),
    ],
  });

  children.push(stepsTable);

  // Signatures
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: '\n\nPodpis vyučujícího: ___________________________          Podpis studenta: ___________________________',
          size: 18,
          color: '334155',
        }),
      ],
      spacing: { before: 240 },
    })
  );

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });

  return await Packer.toBlob(doc);
}

/**
 * Generates an editable Student Assignment Worksheet (.docx)
 */
export async function generateWorksheetDocx(state: any): Promise<Blob> {
  const calc = calculateAmounts(state);
  const children: any[] = [];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: 'PRACOVNÍ LIST: OBĚH ÚČETNÍCH DOKLADŮ V ERP EVALA',
          bold: true,
          size: 28,
          color: '0F172A',
        }),
      ],
      spacing: { after: 100 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Zadání pro studenta | Předmět: Účetnictví (OA) | Zakázka: ${state.caseTitle}`,
          size: 20,
          color: '0D9488',
        }),
      ],
      spacing: { after: 200 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Jméno a příjmení studenta: ___________________________     Třída: _________     Datum: _________`,
          bold: true,
          size: 20,
        }),
      ],
      spacing: { after: 200 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: 'I. ZADÁNÍ ÚČETNÍHO PŘÍPADU A PARAMETRY ZAKÁZKY',
          bold: true,
          size: 22,
          color: '0F172A',
        }),
      ],
      spacing: { before: 100, after: 80 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Naše firma EDU-Trade s.r.o. (IČO: 87654321, DIČ: CZ87654321, BÚ: 2100123456/2010) sjednala obchodní případ se zákazníkem ${state.partner?.name || 'Stavby & Design s.r.o.'}.\n` +
            `• Základní cena plnění bez DPH: ${formatCZK(calc.base)}\n` +
            `• Sazba DPH: ${calc.vatRate} %\n` +
            `• Celková cena plnění s DPH: ${formatCZK(calc.grandTotal)}\n` +
            `• Sjednaná záloha: ${state.advancePercent} % (${formatCZK(calc.advanceGross)})\n` +
            `• Způsob platby: ${state.paymentMethod === 'cash' ? 'Hotovostně přes pokladnu (211)' : 'Bezhotovostně převodem na běžný účet (221)'}\n`,
          size: 19,
          color: '334155',
        }),
      ],
      spacing: { after: 180 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: 'II. PRACOVNÍ ÚKOLY K VYPLNĚNÍ (PŘEDKONTACE):',
          bold: true,
          size: 22,
          color: '0F172A',
        }),
      ],
      spacing: { before: 100, after: 80 },
    })
  );

  const tasks = [
    '1. Ověřte IČO v registru ARES a spolehlivost plátce DPH.',
    '2. Vystavte cenovou nabídku (CN) a přijatou objednávku (OB).',
    `3. Zaúčtujte příjem zálohy (${formatCZK(calc.advanceGross)}): Účet MD: ________  /  Účet D: ________`,
    `4. Zaúčtujte přiznání DPH z přijaté zálohy na DDPP (${formatCZK(calc.advanceVat)}): MD: ________  /  D: ________`,
    `5. Zaúčtujte vyúčtovací fakturu – tržba za služby (${formatCZK(calc.base)}): MD: ________  /  D: ________`,
    `6. Zaúčtujte DPH z celkového zdanitelného plnění: MD: ________  /  D: ________`,
    `7. Zaúčtujte zúčtování zaplacené zálohy a odpočet DPH: MD: ________  /  D: ________`,
    `8. Zaúčtujte příjem doplatku konečné faktury (${formatCZK(calc.remainingPayableGross)}): MD: ________  /  D: ________`,
  ];

  tasks.forEach((t) => {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: t, size: 19 })],
        spacing: { after: 80 },
      })
    );
  });

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });

  return await Packer.toBlob(doc);
}

/**
 * Generates an editable Quiz / Exam Paper (.docx)
 */
export async function generateQuizDocx(): Promise<Blob> {
  const children: any[] = [];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: 'PÍSEMNÝ TEST Z ÚČETNICTVÍ A DPH: OBĚH DOKLADŮ',
          bold: true,
          size: 28,
          color: '0F172A',
        }),
      ],
      spacing: { after: 100 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: 'Obchodní akademie | Školní rok 2025/2026 | Časový limit: 20 minut',
          size: 20,
          color: '0D9488',
        }),
      ],
      spacing: { after: 200 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: 'Jméno studenta: ___________________________     Třída: _________     Hodnocení: _________',
          bold: true,
          size: 20,
        }),
      ],
      spacing: { after: 200 },
    })
  );

  const questions = [
    {
      q: '1. Které z následujících údajů NESMÍ být uvedeny na Objednávce ani Cenové nabídce?',
      opts: [
        'A) Datum splatnosti a DUZP (Datum uskutečnění zdanitelného plnění)',
        'B) Datum vystavení a označení dodavatele',
        'C) Základ daně a celková cena zakázky',
        'D) IČO a DIČ partnera',
      ],
    },
    {
      q: '2. Do kolika dnů od přijetí platby je plátce DPH povinen vystavit Daňový doklad k platbě (DDPP)?',
      opts: ['A) Do 3 dnů', 'B) Do 15 dnů (§ 28 ZDPH)', 'C) Do 30 dnů', 'D) Až na konci roku'],
    },
    {
      q: '3. Jak se postupuje v praxi, pokud je u zakázky sjednána 0% záloha?',
      opts: [
        'A) Musí se vystavit zálohová faktura na 0 Kč',
        'B) Zálohová faktura ani DDPP se nevystavují a přechází se rovnou ke konečné faktuře',
        'C) Zálohová faktura se posílá bez DPH',
        'D) Žádá se o výjimku na FÚ',
      ],
    },
    {
      q: '4. Jaký účet se v české účtové osnově používá pro úhrady zálohy v hotovosti?',
      opts: ['A) 221 - Běžný účet', 'B) 211 - Pokladna', 'C) 311 - Odběratelé', 'D) 343 - DPH'],
    },
    {
      q: '5. Jaká předkontace vznikne u dodavatele při přijetí doplatku vyúčtovací faktury na Běžný účet?',
      opts: ['A) MD 221 / D 311', 'B) MD 311 / D 221', 'C) MD 321 / D 221', 'D) MD 211 / D 602'],
    },
  ];

  questions.forEach((item) => {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: item.q, bold: true, size: 20, color: '0F172A' })],
        spacing: { before: 100, after: 40 },
      })
    );
    item.opts.forEach((opt) => {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `[  ]  ${opt}`, size: 18 })],
          spacing: { after: 30 },
        })
      );
    });
  });

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });

  return await Packer.toBlob(doc);
}

/**
 * Generates an accounting journal in Microsoft Word (.docx) table
 */
export async function generateJournalDocx(journal: any[], title: string = 'Účetní deník'): Promise<Blob> {
  const children: any[] = [];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: `ÚČETNÍ DENÍK – ${title.toUpperCase()}`,
          bold: true,
          size: 28,
          color: '0F172A',
        }),
      ],
      spacing: { after: 100 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `Evala Edu | Vygenerováno: ${new Date().toLocaleDateString('cs-CZ')} | Počet zápisů: ${journal.length}`,
          size: 20,
          color: '0D9488',
        }),
      ],
      spacing: { after: 200 },
    })
  );

  const rows: TableRow[] = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 8, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: '#', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 14, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: 'Datum', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 16, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: 'Doklad', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 38, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: 'Popis účetní operace', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 8, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'MD', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 8, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'D', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
        new TableCell({
          width: { size: 16, type: WidthType.PERCENTAGE },
          shading: { fill: '0F172A', type: ShadingType.CLEAR, color: 'auto' },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Částka', bold: true, color: 'FFFFFF', size: 16 })] })],
        }),
      ],
    }),
  ];

  if (journal.length === 0) {
    rows.push(
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 7,
            margins: { top: 120, bottom: 120, left: 100, right: 100 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'Účetní deník zatím neobsahuje žádné zápisy.', italics: true, size: 18 })],
              }),
            ],
          }),
        ],
      })
    );
  } else {
    journal.forEach((entry, idx) => {
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ children: [new TextRun({ text: String(entry.id || idx + 1), size: 16 })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ children: [new TextRun({ text: String(entry.date || ''), size: 16 })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ children: [new TextRun({ text: String(entry.docNum || ''), bold: true, size: 16, color: '0D9488' })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ children: [new TextRun({ text: String(entry.description || ''), size: 16 })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(entry.md || ''), bold: true, size: 16 })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(entry.d || ''), bold: true, size: 16 })] })],
            }),
            new TableCell({
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: formatCZK(entry.amount || 0), bold: true, size: 16 })] })],
            }),
          ],
        })
      );
    });
  }

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDER_LIGHT,
    rows,
  });

  children.push(table);

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });

  return await Packer.toBlob(doc);
}
