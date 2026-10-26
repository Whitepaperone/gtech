'use strict';

const fs = require('fs');
const path = require('path');
const config = require('../config');
const { dateOnly } = require('./dateService');

const FILE_PATTERN = /^shipment_(\d{4}-\d{2}-\d{2})\.json$/;

function ensureDirectory(directory = config.snapshotDir) {
  fs.mkdirSync(directory, { recursive: true });
}

function snapshotFile(date, directory = config.snapshotDir) {
  return path.join(directory, `shipment_${date}.json`);
}

function listSnapshotDates(directory = config.snapshotDir) {
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory)
    .map(file => FILE_PATTERN.exec(file))
    .filter(Boolean)
    .map(match => match[1])
    .sort();
}

function findLatestBefore(currentThroughDate, directory = config.snapshotDir) {
  return listSnapshotDates(directory).filter(date => date < currentThroughDate).at(-1) || null;
}

function normalizeValue(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return dateOnly(value);
  if (Buffer.isBuffer(value)) return value.toString('utf8');
  return value;
}

function normalizeRow(row) {
  return Object.fromEntries(Object.entries(row).map(([key, value]) => [key, normalizeValue(value)]));
}

function save(date, rows, directory = config.snapshotDir) {
  ensureDirectory(directory);
  const normalizedRows = rows.map(normalizeRow);
  const target = snapshotFile(date, directory);
  const temporary = `${target}.${process.pid}.tmp`;
  const data = { date, throughDate: date, createdAt: new Date().toISOString(), count: normalizedRows.length, rows: normalizedRows };
  fs.writeFileSync(temporary, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(temporary, target);
  return data;
}

function load(date, directory = config.snapshotDir) {
  const file = snapshotFile(date, directory);
  if (!fs.existsSync(file)) return null;
  const snapshot = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (!snapshot || !Array.isArray(snapshot.rows)) throw new Error(`Invalid snapshot: shipment_${date}.json`);
  return snapshot;
}

function shipmentKey(row) {
  return JSON.stringify([dateOnly(row.DDAY), String(row.SALESCODE ?? '').trim(), String(row.ITEMCODE ?? '').trim()]);
}

function groupRows(rows) {
  const groups = new Map();
  for (const raw of rows) {
    const row = normalizeRow(raw);
    const key = shipmentKey(row);
    if (!groups.has(key)) groups.set(key, { ...row, QTY: 0, SOURCE_ROWS: [] });
    const group = groups.get(key);
    const quantity = Number(row.QTY ?? 0);
    if (!Number.isFinite(quantity)) throw new Error('Invalid shipment quantity');
    group.QTY += quantity;
    group.SOURCE_ROWS.push(row);
  }
  return groups;
}

function compareHistorical(snapshotRows, currentRows) {
  const oldMap = groupRows(snapshotRows);
  const currentMap = groupRows(currentRows);
  const changes = [];
  for (const [key, row] of currentMap) {
    const old = oldMap.get(key);
    if (old && Math.abs(old.QTY - row.QTY) < 1e-9) continue;
    changes.push({ ...row, CHANGE_TYPE: old ? 'MODIFIED' : 'PAST_NEW',
      CHANGED_FIELDS: old ? ['QTY'] : [], OLD_VALUES: old ? { QTY: old.QTY } : {},
      OLD_TOTAL_QTY: old ? old.QTY : 0, NEW_TOTAL_QTY: row.QTY,
      GROUP_COMPARISON: true });
  }
  for (const [key, old] of oldMap) {
    if (!currentMap.has(key)) changes.push({ ...old, CHANGE_TYPE: 'DELETED',
      CHANGED_FIELDS: [], OLD_VALUES: {}, OLD_TOTAL_QTY: old.QTY,
      NEW_TOTAL_QTY: 0, GROUP_COMPARISON: true });
  }
  return changes;
}

module.exports = {
  compareHistorical, findLatestBefore, listSnapshotDates, load, normalizeRow,
  save, snapshotFile
};
