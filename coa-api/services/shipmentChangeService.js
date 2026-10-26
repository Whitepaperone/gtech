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
  // Include today's stored snapshot so repeated calls detect changes since the last successful poll.
  const baseDate = snapshotService.listSnapshotDates().filter(date => date <= currentThroughDate).at(-1) || null;
  const journal = require('./changeJournalService');
  if (!baseDate) {
    snapshotService.save(currentThroughDate, currentRows);
    return buildResponse(currentThroughDate, null, normalizedConsumer,
      journal.pending(normalizedConsumer), true);
  }
  const snapshot = snapshotService.load(baseDate);
  const currentHash = crypto.createHash('sha256').update(JSON.stringify(currentRows)).digest('hex');
  const changes = snapshotService.compareHistorical(snapshot.rows, currentRows).map(change => {
    if (change.CHANGE_TYPE === 'PAST_NEW' && dateOnly(change.DDAY) > baseDate) {
      change.CHANGE_TYPE = 'NEW';
    }
    return addChangeId(change, snapshot.createdAt || baseDate,
      currentHash);
  });
  // Persist events before advancing the snapshot; a failed save can safely be retried.
  journal.append(changes);
  snapshotService.save(currentThroughDate, currentRows);
  const pendingChanges = journal.pending(normalizedConsumer);

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
