'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const snapshotService = require('../services/snapshotService');

test('latest snapshot strictly before through-date is selected', t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'coa-snapshot-'));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  snapshotService.save('2026-08-27', [], directory);
  snapshotService.save('2026-08-30', [], directory);
  assert.equal(snapshotService.findLatestBefore('2026-08-30', directory), '2026-08-27');
});

test('historical comparison identifies past-new, modified and deleted rows', () => {
  const oldRows = [
    { IINDEX: 1, DDAY: '2026-08-20', QTY: 1 },
    { IINDEX: 2, DDAY: '2026-08-21', QTY: 2 }
  ];
  const currentRows = [
    { IINDEX: 1, DDAY: '2026-08-20', QTY: 3 },
    { IINDEX: 3, DDAY: '2026-08-22', QTY: 1 }
  ];
  const changes = snapshotService.compareHistorical(oldRows, currentRows);
  assert.deepEqual(changes.map(row => row.CHANGE_TYPE).sort(), ['DELETED', 'MODIFIED', 'PAST_NEW']);
});
