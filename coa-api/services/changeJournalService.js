 'use strict';
const fs = require('fs');
const path = require('path');
const config = require('../config');
const acknowledgements = require('./acknowledgementService');
function read(file = path.join(config.snapshotDir, 'change_events.json')) {
  return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : [];
}
function append(changes, file = path.join(config.snapshotDir, 'change_events.json')) {
  if (!changes.length) return;
  const events = read(file);
  const known = new Set(events.map(event => event.CHANGE_ID));
  for (const change of changes) {
    if (!known.has(change.CHANGE_ID)) {
      events.push({ ...change, DETECTED_AT: new Date().toISOString() });
      known.add(change.CHANGE_ID);
    }
  }
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const temporary = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(temporary, JSON.stringify(events), 'utf8');
  fs.renameSync(temporary, file);
}
function pending(consumer) {
  const acknowledged = acknowledgements.acknowledgedIds(consumer);
  return read().filter(event => !acknowledged.has(event.CHANGE_ID));
}
module.exports = { read, append, pending };
