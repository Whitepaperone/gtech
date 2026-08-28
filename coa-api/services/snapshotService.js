'use strict';

const fs = require('fs');
const path = require('path');
const config = require('../config');
const { dateOnly } = require('./dateService');

const FILE_PATTERN = /^shipment_(\d{4}-\d{2}-\d{2})\.json$/;
const COMPARE_FIELDS = [
  'DDAY', 'SALESCODE', 'ITEMCODE', 'ITEMNAME', 'DESCRIPTION',
  'QTY', 'RETURNQTY', 'RSQTY', 'UNIT', 'UNITPRICE', 'STAX',
  'CUSTOMERCODE', 'CUSTOMERNAME', 'OUTLOCATION', 'LOCATIONNAME',
  'REMARK', 'BASETYPE'
];

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
  if (row.IINDEX === null || row.IINDEX === undefined) throw new Error('Shipment row has no IINDEX');
  return String(row.IINDEX);
}

function compareHistorical(snapshotRows, currentRows) {
  const oldMap = new Map(snapshotRows.map(row => [shipmentKey(row), normalizeRow(row)]));
  const currentMap = new Map();
  const changes = [];

  for (const rawRow of currentRows) {
    const row = normalizeRow(rawRow);
    const key = shipmentKey(row);
    currentMap.set(key, row);
    const oldRow = oldMap.get(key);
    if (!oldRow) {
      changes.push({ ...row, CHANGE_TYPE: 'PAST_NEW', CHANGED_FIELDS: [], OLD_VALUES: {} });
      continue;
    }
    const changedFields = COMPARE_FIELDS.filter(field => String(oldRow[field] ?? '') !== String(row[field] ?? ''));
    if (changedFields.length) {
      const oldValues = Object.fromEntries(changedFields.map(field => [field, oldRow[field] ?? null]));
      changes.push({ ...row, CHANGE_TYPE: 'MODIFIED', CHANGED_FIELDS: changedFields, OLD_VALUES: oldValues });
    }
  }

  for (const [key, oldRow] of oldMap) {
    if (!currentMap.has(key)) changes.push({ ...oldRow, CHANGE_TYPE: 'DELETED', CHANGED_FIELDS: [], OLD_VALUES: {} });
  }
  return changes;
}

module.exports = {
  compareHistorical, findLatestBefore, listSnapshotDates, load, normalizeRow,
  save, snapshotFile
};
