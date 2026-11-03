/**
 * Unit tests for pure accounting logic
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  calculateAmounts,
  calcPayroll,
  validateOcrInvoice,
  calculateTeacherGrade,
} from './accountingLogic';

test('standard 21% VAT calculation with 100% advance', () => {
  const result = calculateAmounts({
    currency: 'CZK',
    customBaseAmount: 100000,
    vatRate: 21,
    advanceType: 'pct',
    advancePercent: 100,
  });

  assert.equal(result.base, 100000);
  assert.equal(result.vatTotal, 21000);
  assert.equal(result.grandTotal, 121000);
  assert.equal(result.advanceGross, 121000);
  assert.equal(result.advanceVat, 21000);
  assert.equal(result.advanceBase, 100000);
  assert.equal(result.remainingPayableGross, 0);
});

test('foreign currency EUR with exchange rate differences', () => {
  const result = calculateAmounts({
    currency: 'EUR',
    foreignBaseAmount: 4000,
    invoiceRate: 25.2,
    paymentRate: 25.7,
    vatRate: 0,
    advancePercent: 0,
  });

  assert.equal(result.base, 100800);
  assert.equal(result.grandTotal, 100800);
  assert.equal(result.remainingPayableGross, 100800);
  assert.equal(result.actualPaymentCzk, 102800);
  assert.equal(result.fxDifference, 2000); // 102800 - 100800 = 2000 Kč exchange rate gain
});

test('reverse charge mode (§ 92e) sets 0% invoice VAT and shifts tax to buyer', () => {
  const result = calculateAmounts({
    currency: 'CZK',
    isReverseCharge: true,
    reverseChargeCode: '4',
    reverseChargeVatRate: 21,
    customBaseAmount: 200000,
  });

  assert.equal(result.vatRate, 0);
  assert.equal(result.vatTotal, 0);
  assert.equal(result.grandTotal, 200000);
  assert.equal(result.reverseChargeVatAmount, 42000);
});

test('payroll calculation with taxpayer relief and two children generates tax bonus', () => {
  const p = calcPayroll(25000, true, 2, 0);

  assert.equal(p.gross, 25000);
  assert.equal(p.zpEmployee, 1125); // 4.5% of 25000
  assert.equal(p.spEmployee, 1775); // 7.1% of 25000
  assert.equal(p.taxBase, 25000);
  assert.equal(p.taxBefore, 3750); // 15%
  assert.equal(p.appliedTaxpayerRelief, 2570);
  // Tax after basic relief: 3750 - 2570 = 1180 Kč
  // Child relief: 1267 + 1860 = 3127 Kč
  // 3127 - 1180 = 1947 Kč tax bonus!
  assert.equal(p.taxAfterRelief, 0);
  assert.equal(p.taxBonus, 1947);
  // Net wage: 25000 - 1125 - 1775 - 0 + 1947 = 24047 Kč
  assert.equal(p.netWage, 24047);
  assert.equal(p.wageToPay, 24047);
});

test('OCR invoice validation detects invalid IČO and mismatched VAT', () => {
  const invalid = validateOcrInvoice({
    baseAmount: 10000,
    vatRate: 21,
    vatAmount: 2500, // Error: 21% of 10000 is 2100, not 2500!
    totalAmount: 12500,
    supplierIco: '12345', // Error: only 5 digits!
    variableSymbol: '',
  });

  assert.equal(invalid.isValid, false);
  assert.equal(invalid.errors.length, 2);
  assert.equal(invalid.warnings.length, 1);
});

test('teacher evaluation grading logic calculates proper mark', () => {
  const grade1 = calculateTeacherGrade(0, 5);
  assert.equal(grade1.grade, 1);

  const grade2 = calculateTeacherGrade(2, 4);
  assert.equal(grade2.grade, 2);

  const grade5 = calculateTeacherGrade(8, 1);
  assert.equal(grade5.grade, 5);
});
