'use strict';

const crypto = require('crypto');
const shipmentRepository = require('../repositories/shipmentRepository');
const snapshotService = require('./snapshotService');
const acknowledgementService = require('./acknowledgementService');
const { dateOnly, getKoreaDate } = require('./dateService');

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonicalize(value[key])]));
  }
  return value;
}

function addChangeId(change, baseDate, currentThroughDate) {
  const identity = canonicalize({ baseDate, currentThroughDate, change });
  return { ...change, CHANGE_ID: crypto.createHash('sha256').update(JSON.stringify(identity)).digest('hex') };
}

async function getChanges(consumer) {
  const normalizedConsumer = acknowledgementService.validateConsumer(consumer);
  const currentThroughDate = getKoreaDate(-1);
  const currentRows = (await shipmentRepository.findThroughDate(currentThroughDate))
    .map(snapshotService.normalizeRow);
  const baseDate = snapshotService.findLatestBefore(currentThroughDate);

  if (!baseDate) {
    if (!snapshotService.load(currentThroughDate)) snapshotService.save(currentThroughDate, currentRows);
    return buildResponse(currentThroughDate, null, normalizedConsumer, [], true);
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
  const allChanges = [...newRows, ...historicalChanges]
    .map(change => addChangeId(change, baseDate, currentThroughDate));
  const acknowledged = acknowledgementService.acknowledgedIds(normalizedConsumer);
  const pendingChanges = allChanges.filter(change => !acknowledged.has(change.CHANGE_ID));

  if (!snapshotService.load(currentThroughDate)) snapshotService.save(currentThroughDate, currentRows);
  return buildResponse(currentThroughDate, baseDate, normalizedConsumer, pendingChanges, false);
}

function buildResponse(currentThroughDate, baseDate, consumer, data, baselineCreated) {
  return {
    ok: true,
    today: getKoreaDate(0),
    targetDate: currentThroughDate,
    currentThroughDate,
    snapshotBaseDate: baseDate,
    consumer,
    baselineCreated,
    summary: {
      newCount: data.filter(x => x.CHANGE_TYPE === 'NEW').length,
      modifiedCount: data.filter(x => x.CHANGE_TYPE === 'MODIFIED').length,
      pastNewCount: data.filter(x => x.CHANGE_TYPE === 'PAST_NEW').length,
      deletedCount: data.filter(x => x.CHANGE_TYPE === 'DELETED').length,
      totalCount: data.length
    },
    data
  };
}

module.exports = { addChangeId, getChanges };
