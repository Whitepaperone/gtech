'use strict';

const shipmentRepository = require('../repositories/shipmentRepository');
const snapshotService = require('./snapshotService');
const { dateOnly, getKoreaDate } = require('./dateService');

async function getChanges() {
  const currentThroughDate = getKoreaDate(-1);
  const currentRows = (await shipmentRepository.findThroughDate(currentThroughDate)).map(snapshotService.normalizeRow);
  const baseDate = snapshotService.findLatestBefore(currentThroughDate);

  // First run establishes a baseline. Historical change types cannot be inferred without an older snapshot.
  if (!baseDate) {
    if (!snapshotService.load(currentThroughDate)) snapshotService.save(currentThroughDate, currentRows);
    return buildResponse(currentThroughDate, null, [], [], true);
  }

  const snapshot = snapshotService.load(baseDate);
  const historicalRows = currentRows.filter(row => dateOnly(row.DDAY) <= baseDate);
  const newRows = currentRows
    .filter(row => {
      const date = dateOnly(row.DDAY);
      return date && date > baseDate && date <= currentThroughDate;
    })
    .map(row => ({ ...row, CHANGE_TYPE: 'NEW', CHANGED_FIELDS: [], OLD_VALUES: {} }));
  const historicalChanges = snapshotService.compareHistorical(snapshot.rows, historicalRows);

  if (!snapshotService.load(currentThroughDate)) snapshotService.save(currentThroughDate, currentRows);
  return buildResponse(currentThroughDate, baseDate, newRows, historicalChanges, false);
}

function buildResponse(currentThroughDate, baseDate, newRows, historicalChanges, baselineCreated) {
  const data = [...newRows, ...historicalChanges];
  return {
    ok: true,
    today: getKoreaDate(0),
    targetDate: currentThroughDate,
    currentThroughDate,
    snapshotBaseDate: baseDate,
    baselineCreated,
    summary: {
      newCount: newRows.length,
      modifiedCount: historicalChanges.filter(x => x.CHANGE_TYPE === 'MODIFIED').length,
      pastNewCount: historicalChanges.filter(x => x.CHANGE_TYPE === 'PAST_NEW').length,
      deletedCount: historicalChanges.filter(x => x.CHANGE_TYPE === 'DELETED').length,
      totalCount: data.length
    },
    data
  };
}

module.exports = { getChanges };
