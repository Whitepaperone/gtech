'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const acknowledgementService = require('../services/acknowledgementService');

test('acknowledgements are independent per consumer and idempotent', t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'coa-ack-'));
  const logFile = path.join(directory, 'acks.ndjson');
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const changeId = 'a'.repeat(64);

  assert.equal(acknowledgementService.acknowledge('PRODUCTION_PLAN', [changeId], {}, logFile), 1);
  assert.equal(acknowledgementService.acknowledge('PRODUCTION_PLAN', [changeId], {}, logFile), 0);
  assert.equal(acknowledgementService.acknowledgedIds('PRODUCTION_PLAN', logFile).has(changeId), true);
  assert.equal(acknowledgementService.acknowledgedIds('SHIPMENT_DISPOSAL', logFile).has(changeId), false);
});
