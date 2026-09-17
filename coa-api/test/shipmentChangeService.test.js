'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { addChangeId } = require('../services/shipmentChangeService');

test('change IDs are deterministic and change when the new value changes', () => {
  const first = addChangeId({ IINDEX: 1, CHANGE_TYPE: 'MODIFIED', QTY: 2 }, '2026-08-26', '2026-08-27');
  const same = addChangeId({ QTY: 2, CHANGE_TYPE: 'MODIFIED', IINDEX: 1 }, '2026-08-26', '2026-08-27');
  const different = addChangeId({ IINDEX: 1, CHANGE_TYPE: 'MODIFIED', QTY: 3 }, '2026-08-26', '2026-08-27');
  assert.equal(first.CHANGE_ID, same.CHANGE_ID);
  assert.notEqual(first.CHANGE_ID, different.CHANGE_ID);
});
