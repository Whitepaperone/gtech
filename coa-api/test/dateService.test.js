'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { getKoreaDate } = require('../services/dateService');

test('Korea yesterday is calculated across a UTC date boundary', () => {
  const now = new Date('2026-08-30T16:00:00.000Z'); // 2026-08-31 01:00 KST
  assert.equal(getKoreaDate(0, now), '2026-08-31');
  assert.equal(getKoreaDate(-1, now), '2026-08-30');
});
