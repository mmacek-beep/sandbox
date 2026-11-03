/**
 * Evala Edu - Pure Accounting & Valuation Logic
 * @license Apache-2.0
 */

export interface CompanyInfo {
  name: string;
  street: string;
  city: string;
  ico: string;
  dic: string;
  bankAccount: string;
  bankName: string;
  iban: string;
  vatPayer: boolean;
}

export interface PartnerInfo {
  name: string;
  street: string;
  city: string;
  ico: string;
  dic: string;
  bankAccount: string;
  email?: string;
  reliable?: boolean;
}

export interface AccountingAmounts {
  isEur: boolean;
  currency: 'CZK' | 'EUR';
  foreignBase: number;
  invoiceRate: number;
  paymentRate: number;
  foreignGrandTotal: number;
  foreignVatTotal: number;
  foreignAdvanceGross: number;
  foreignRemainingPayableGross: number;
  base: number;
  vatRate: number;
  vatTotal: number;
  grandTotal: number;
  advanceGross: number;
  advanceBase: number;
  advanceVat: number;
  remainingPayableGross: number;
  remainingBase: number;
  remainingVat: number;
  actualPaymentCzk: number;
  fxDifference: number;
  isReverseCharge: boolean;
  reverseChargeCode: string;
  reverseChargeVatRate: number;
  reverseChargeVatAmount: number;
}

export function calculateAmounts(params: {
  currency?: 'CZK' | 'EUR';
  isReverseCharge?: boolean;
  reverseChargeCode?: string;
  reverseChargeVatRate?: number;
  foreignBaseAmount?: number;
  customBaseAmount?: number;
  invoiceRate?: number;
  paymentRate?: number;
  vatRate?: number;
  advanceType?: 'pct' | 'czk';
  advancePercent?: number;
  customAdvanceCzk?: number;
}): AccountingAmounts {
  const isEur = params.currency === 'EUR';
  const isRC = Boolean(params.isReverseCharge);
  const rcCode = params.reverseChargeCode || '4';
  const rcVatRate = Number(params.reverseChargeVatRate) || 21;

  const foreignBase = isEur
    ? Number(params.foreignBaseAmount) || 0
    : Number(params.customBaseAmount) || 0;
  const invoiceRate = isEur ? Number(params.invoiceRate) || 25.0 : 1;
  const paymentRate = isEur ? Number(params.paymentRate) || 25.0 : 1;

  const base = isEur
    ? Math.round(foreignBase * invoiceRate)
    : Number(params.customBaseAmount) || 0;

  // In reverse charge mode (§ 92a ZDPH), the invoice has 0 % VAT (tax shifted to buyer)
  const vatRate = isRC ? 0 : Number(params.vatRate) || 0;
  const vatTotal = isRC ? 0 : Math.round(base * (vatRate / 100));
  const grandTotal = isRC ? base : base + vatTotal;
  const rcVatAmount = isRC ? Math.round(base * (rcVatRate / 100)) : 0;

  const foreignVatTotal = isEur ? Math.round(foreignBase * (vatRate / 100)) : 0;
  const foreignGrandTotal = isEur ? foreignBase + foreignVatTotal : 0;

  let advanceGross = 0;
  let foreignAdvanceGross = 0;
  if (params.advanceType === 'pct' || !params.advanceType) {
    const pct = Number(params.advancePercent) ?? 100;
    advanceGross = Math.round(grandTotal * (pct / 100));
    foreignAdvanceGross = isEur ? Math.round(foreignGrandTotal * (pct / 100)) : 0;
  } else {
    advanceGross = Number(params.customAdvanceCzk) || 0;
    if (advanceGross > grandTotal) advanceGross = grandTotal;
  }

  const advanceVat =
    vatRate > 0 ? Math.round(advanceGross * (vatRate / (100 + vatRate))) : 0;
  const advanceBase = advanceGross - advanceVat;

  const remainingPayableGross = grandTotal - advanceGross;
  const remainingBase = base - advanceBase;
  const remainingVat = vatTotal - advanceVat;
  const foreignRemainingPayableGross = isEur
    ? foreignGrandTotal - foreignAdvanceGross
    : 0;

  const actualPaymentCzk = isEur
    ? Math.round(foreignRemainingPayableGross * paymentRate)
    : remainingPayableGross;
  const fxDifference = isEur ? actualPaymentCzk - remainingPayableGross : 0;

  return {
    isEur,
    currency: params.currency || 'CZK',
    foreignBase,
    invoiceRate,
    paymentRate,
    foreignGrandTotal,
    foreignVatTotal,
    foreignAdvanceGross,
    foreignRemainingPayableGross,
    base: Math.round(base),
    vatRate,
    vatTotal,
    grandTotal,
    advanceGross,
    advanceBase,
    advanceVat,
    remainingPayableGross,
    remainingBase,
    remainingVat,
    actualPaymentCzk,
    fxDifference,
    isReverseCharge: isRC,
    reverseChargeCode: rcCode,
    reverseChargeVatRate: rcVatRate,
    reverseChargeVatAmount: rcVatAmount,
  };
}

export function calcPayroll(
  grossWage: number,
  hasTaxpayerRelief: boolean = true,
  childrenCount: number = 0,
  otherDeductions: number = 0
) {
  const gross = Math.max(0, Math.round(grossWage));

  // Employee statutory insurance
  const zpEmployee = Math.ceil(gross * 0.045); // 4.5 %
  const spEmployee = Math.ceil(gross * 0.071); // 7.1 %
  const totalInsEmployee = zpEmployee + spEmployee;

  // Personal income tax (15%)
  const taxBase = Math.ceil(gross / 100) * 100;
  const taxBefore = Math.round(taxBase * 0.15);
  const taxpayerRelief = hasTaxpayerRelief ? 2570 : 0;
  const appliedTaxpayerRelief = Math.min(taxBefore, taxpayerRelief);
  const taxAfterTaxpayer = Math.max(0, taxBefore - appliedTaxpayerRelief);

  let childRelief = 0;
  if (hasTaxpayerRelief && childrenCount > 0) {
    if (childrenCount >= 1) childRelief += 1267;
    if (childrenCount >= 2) childRelief += 1860;
    if (childrenCount >= 3) childRelief += (childrenCount - 2) * 2320;
  }

  let taxAfterRelief = 0;
  let taxBonus = 0;
  if (childRelief <= taxAfterTaxpayer) {
    taxAfterRelief = taxAfterTaxpayer - childRelief;
  } else {
    taxAfterRelief = 0;
    taxBonus = childRelief - taxAfterTaxpayer;
  }

  const netWage = gross - zpEmployee - spEmployee - taxAfterRelief + taxBonus;
  const otherDed = Math.max(0, Math.round(otherDeductions));
  const wageToPay = Math.max(0, netWage - otherDed);

  // Employer statutory social costs
  const zpEmployer = Math.round(gross * 0.09); // 9 %
  const spEmployer = Math.round(gross * 0.248); // 24.8 %
  const totalInsEmployer = zpEmployer + spEmployer;
  const totalCompanyCost = gross + totalInsEmployer;

  return {
    gross,
    zpEmployee,
    spEmployee,
    totalInsEmployee,
    taxBase,
    taxBefore,
    taxpayerRelief,
    appliedTaxpayerRelief,
    childRelief,
    taxAfterRelief,
    taxBonus,
    netWage,
    otherDed,
    wageToPay,
    zpEmployer,
    spEmployer,
    totalInsEmployer,
    totalCompanyCost,
  };
}

export function validateOcrInvoice(inv: {
  baseAmount: number;
  vatRate: number;
  vatAmount: number;
  totalAmount: number;
  supplierIco: string;
  variableSymbol: string;
}) {
  const errors: string[] = [];
  const warnings: string[] = [];

  const calcVat = Math.round(inv.baseAmount * (inv.vatRate / 100));
  const diffVat = Math.abs(calcVat - inv.vatAmount);
  if (diffVat > 2 && inv.vatRate > 0) {
    errors.push(
      `Matematická neshoda v DPH: Základ ${inv.baseAmount.toLocaleString('cs-CZ')} Kč × ${inv.vatRate} % = ${calcVat.toLocaleString('cs-CZ')} Kč, ale v dokladu je ${inv.vatAmount.toLocaleString('cs-CZ')} Kč.`
    );
  }

  const calcTotal = inv.baseAmount + inv.vatAmount;
  if (Math.abs(calcTotal - inv.totalAmount) > 2) {
    errors.push(
      `Součet základu a DPH (${calcTotal.toLocaleString('cs-CZ')} Kč) neodpovídá celkové částce (${inv.totalAmount.toLocaleString('cs-CZ')} Kč).`
    );
  }

  const ico = String(inv.supplierIco || '').replace(/\s+/g, '');
  if (!/^\d{8}$/.test(ico)) {
    errors.push(`IČO "${inv.supplierIco}" není platné (musí mít 8 číslic).`);
  }

  const vs = String(inv.variableSymbol || '').trim();
  if (!vs) {
    warnings.push('Chybí variabilní symbol pro spárování platby.');
  }

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
  };
}

export function calculateTeacherGrade(examErrors: number, quizScore: number): {
  grade: number;
  gradeText: string;
  percentage: number;
} {
  let grade = 1;
  if (examErrors === 0 && quizScore >= 4) grade = 1;
  else if (examErrors <= 2 && quizScore >= 3) grade = 2;
  else if (examErrors <= 4 && quizScore >= 2) grade = 3;
  else if (examErrors <= 6) grade = 4;
  else grade = 5;

  const names: Record<number, string> = {
    1: 'Výborný (1)',
    2: 'Chvalitebný (2)',
    3: 'Dobrý (3)',
    4: 'Dostatečný (4)',
    5: 'Nedostatečný (5)',
  };

  const percentage = Math.max(
    0,
    Math.min(100, Math.round((quizScore / 5) * 50 + Math.max(0, 50 - examErrors * 10)))
  );

  return {
    grade,
    gradeText: names[grade] || 'Hodnoceno',
    percentage,
  };
}

export function formatCZK(amount: number): string {
  return new Intl.NumberFormat('cs-CZ', {
    style: 'currency',
    currency: 'CZK',
    maximumFractionDigits: 0,
  }).format(amount || 0);
}
