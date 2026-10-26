const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const snapshots = require('../services/snapshotService');
const service = require('../services/shipmentChangeService');
const repository = require('../repositories/shipmentRepository');
const config = require('../config');
const { getKoreaDate } = require('../services/dateService');
const ack = require('../services/acknowledgementService');
const base = { DDAY: '2026-01-01', SALESCODE: 'S1', ITEMCODE: 'P1' };
test('splitting rows, replacing indexes and changing customers are ignored', () => {
  assert.deepEqual(snapshots.compareHistorical([{ ...base, IINDEX: 1, QTY: 100 }],
    [{ ...base, IINDEX: 2, QTY: 60, CUSTOMERNAME: 'changed' }, { ...base, IINDEX: 3, QTY: 40 }]), []);
});
test('quantity totals and removed keys include old and new totals', () => {
  const changes = snapshots.compareHistorical([{ ...base, QTY: 60 }, { ...base, QTY: 40 }], [{ ...base, QTY: 120 }]);
  assert.equal(changes[0].OLD_TOTAL_QTY, 100);
  assert.equal(changes[0].NEW_TOTAL_QTY, 120);
  const moved = snapshots.compareHistorical([{ ...base, QTY: 100 }], [{ ...base, DDAY: '2026-01-02', QTY: 100 }]);
  assert.deepEqual(moved.map(x => x.CHANGE_TYPE).sort(), ['DELETED', 'PAST_NEW']);
});
test('unacknowledged deletion survives polls and days, independently per consumer', async t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'coa-events-'));
  const originalDir = config.snapshotDir, originalAck = config.ackLogFile;
  const originalQuery = repository.findThroughDate;
  config.snapshotDir = directory;
  config.ackLogFile = path.join(directory, 'acks.ndjson');
  t.after(() => { config.snapshotDir = originalDir; config.ackLogFile = originalAck; repository.findThroughDate = originalQuery; fs.rmSync(directory, { recursive: true, force: true }); });
  snapshots.save(getKoreaDate(-2), [{ ...base, QTY: 100 }]);
  repository.findThroughDate = async () => [];
  const first = await service.getChanges('PRODUCTION_PLAN');
  assert.equal(first.summary.deletedCount, 1);
  const second = await service.getChanges('PRODUCTION_PLAN');
  assert.equal(second.data[0].CHANGE_ID, first.data[0].CHANGE_ID);
  ack.acknowledge('PRODUCTION_PLAN', [first.data[0].CHANGE_ID]);
  assert.equal((await service.getChanges('PRODUCTION_PLAN')).data.length, 0);
  assert.equal((await service.getChanges('SHIPMENT_DISPOSAL')).data.length, 1);
});
