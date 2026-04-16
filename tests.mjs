import assert from 'node:assert/strict';
import {
  createDefaultContext,
  evaluateDocument,
  evaluateExpression,
  formatResult,
  parseDocument,
} from './lib/engine.js';

const ctx = createDefaultContext({
  fx: {
    base: 'USD',
    fetchedAt: Date.now(),
    date: '2026-04-15',
    rates: { USD: 1, EUR: 0.9, GBP: 0.8, JPY: 150 },
  },
});

assert.equal(formatResult(evaluateExpression('20% of 50', ctx)), '10');
assert.equal(formatResult(evaluateExpression('5% on 30', ctx)), '31.5');
assert.equal(formatResult(evaluateExpression('6% off 40 EUR', ctx)), '€ 37.6');
assert.equal(formatResult(evaluateExpression('50 as a % of 100', ctx)), '50%');
assert.equal(formatResult(evaluateExpression('3 ft in cm', ctx)), '91.44 cm');
assert.equal(formatResult(evaluateExpression('1 cm in px', ctx)), '37.795 px');
assert.equal(formatResult(evaluateExpression('sqrt(16)', ctx)), '4');
assert.equal(formatResult(evaluateExpression('0xff', ctx)), '255');
assert.match(formatResult(evaluateExpression('now in Tokyo', ctx)), /AM|PM|:/);
assert.equal(
  formatResult(
    evaluateExpression(
      '$1 + 5 USD',
      createDefaultContext({
        fx: ctx.fx,
        lineValues: [{ value: evaluateExpression('10 USD', ctx), breaksRollup: false }],
      }),
    ),
  ),
  '$ 15',
);

const doc = parseDocument(['a = 10 USD', 'b = 15 USD', 'sum', '', 'avg'].join('\n'));
const evaluated = evaluateDocument(doc, ctx);
assert.equal(evaluated.lines[2].displayText, '$ 25');
assert.ok(evaluated.lines[4].error);

console.log('tests passed');
