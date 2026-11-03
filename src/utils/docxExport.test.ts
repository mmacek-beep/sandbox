/**
 * Unit test for docx export generation
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { generateQuizDocx } from './docxExport';

test('generates valid docx blob for quiz', async () => {
  const blob = await generateQuizDocx();
  assert.ok(blob);
  assert.ok(blob.size > 1000); // Standard docx zip container is at least a few KB
});
