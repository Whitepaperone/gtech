'use strict';

const fs = require('fs');
const path = require('path');
const config = require('../config');

const CONSUMERS = new Set(['PRODUCTION_PLAN', 'SHIPMENT_DISPOSAL']);

function validateConsumer(value) {
  const consumer = String(value || '').trim().toUpperCase();
  if (!CONSUMERS.has(consumer)) {
    const error = new Error('consumer must be PRODUCTION_PLAN or SHIPMENT_DISPOSAL');
    error.statusCode = 400;
    throw error;
  }
  return consumer;
}

function readEntries(logFile = config.ackLogFile) {
  if (!fs.existsSync(logFile)) return [];
  return fs.readFileSync(logFile, 'utf8').split(/\r?\n/).filter(Boolean).map((line, index) => {
    try {
      return JSON.parse(line);
    } catch (error) {
      throw new Error(`Invalid acknowledgement log at line ${index + 1}`);
    }
  });
}

function acknowledgedIds(consumer, logFile = config.ackLogFile) {
  const normalizedConsumer = validateConsumer(consumer);
  return new Set(readEntries(logFile)
    .filter(entry => entry.consumer === normalizedConsumer)
    .map(entry => entry.changeId));
}

function acknowledge(consumer, changeIds, metadata = {}, logFile = config.ackLogFile) {
  const normalizedConsumer = validateConsumer(consumer);
  if (!Array.isArray(changeIds) || changeIds.length === 0 || changeIds.length > 2000) {
    const error = new Error('changeIds must be a non-empty array with at most 2000 items');
    error.statusCode = 400;
    throw error;
  }
  const uniqueIds = [...new Set(changeIds.map(value => String(value || '').trim()))];
  if (uniqueIds.some(id => !/^[a-f0-9]{64}$/.test(id))) {
    const error = new Error('changeIds contains an invalid ID');
    error.statusCode = 400;
    throw error;
  }
  const existing = acknowledgedIds(normalizedConsumer, logFile);
  const pending = uniqueIds.filter(id => !existing.has(id));
  if (!pending.length) return 0;

  fs.mkdirSync(path.dirname(logFile), { recursive: true });
  const now = new Date().toISOString();
  const safeMetadata = {
    operator: String(metadata.operator || '').slice(0, 100),
    computer: String(metadata.computer || '').slice(0, 100),
    sourceFile: path.basename(String(metadata.sourceFile || '')).slice(0, 260),
    outputFile: path.basename(String(metadata.outputFile || '')).slice(0, 260)
  };
  const lines = pending.map(changeId => JSON.stringify({
    acknowledgedAt: now, consumer: normalizedConsumer, changeId, ...safeMetadata
  })).join('\n') + '\n';
  fs.appendFileSync(logFile, lines, 'utf8');
  return pending.length;
}

module.exports = { acknowledge, acknowledgedIds, validateConsumer };
