/**
 * Evala Edu - Microsoft Word (.docx) Parser Engine
 * Extracts text and accounting metadata from uploaded Word documents.
 * @license Apache-2.0
 */

import mammoth from 'mammoth';

export interface ParsedWordDocument {
  rawText: string;
  foundIco?: string;
  foundDic?: string;
  foundAmount?: number;
  foundInvoiceNumber?: string;
  detectedTitle?: string;
  lines: string[];
}

/**
 * Parses an uploaded .docx file and extracts raw text and key accounting entities
 */
export async function parseDocxFile(file: File): Promise<ParsedWordDocument> {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  const rawText = result.value || '';
  const lines = rawText
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);

  // Heuristic extraction of IČO (8 consecutive digits)
  const icoMatch = rawText.match(/\b\d{8}\b/);
  const foundIco = icoMatch ? icoMatch[0] : undefined;

  // Heuristic extraction of DIČ (CZ followed by 8-10 digits)
  const dicMatch = rawText.match(/\bCZ\d{8,10}\b/i);
  const foundDic = dicMatch ? dicMatch[0].toUpperCase() : undefined;

  // Heuristic extraction of amounts (e.g. 100 000 Kč or 50000)
  const amountMatch = rawText.match(/(\d[\d\s]{2,})\s*(?:Kč|CZK|EUR)/i);
  let foundAmount: number | undefined = undefined;
  if (amountMatch) {
    const cleanNum = Number(amountMatch[1].replace(/\s+/g, ''));
    if (!isNaN(cleanNum) && cleanNum > 0) {
      foundAmount = cleanNum;
    }
  }

  // Invoice number detection (VF2026..., 2026..., etc.)
  const invMatch = rawText.match(/(?:faktur[ay]|doklad[u]?|číslo|č\.)\s*[:#]?\s*([A-Z0-9\-\/]{4,15})/i);
  const foundInvoiceNumber = invMatch ? invMatch[1] : undefined;

  // Detected title
  const detectedTitle = lines[0] || file.name.replace(/\.[^/.]+$/, '');

  return {
    rawText,
    foundIco,
    foundDic,
    foundAmount,
    foundInvoiceNumber,
    detectedTitle,
    lines,
  };
}
