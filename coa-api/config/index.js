'use strict';

const path = require('path');

function positiveInteger(name, fallback) {
  const value = Number(process.env[name] || fallback);
  if (!Number.isInteger(value) || value <= 0) {
    throw new Error(`${name} must be a positive integer`);
  }
  return value;
}

const snapshotSetting = process.env.SNAPSHOT_DIR || 'data/snapshots';

const config = {
  port: positiveInteger('PORT', 3000),
  apiKey: process.env.API_KEY || '',
  jsonLimit: process.env.JSON_LIMIT || '10kb',
  rateLimitWindowMs: positiveInteger('RATE_LIMIT_WINDOW_MS', 60_000),
  rateLimitMax: positiveInteger('RATE_LIMIT_MAX', 100),
  snapshotDir: path.resolve(__dirname, '..', snapshotSetting),
  database: {
    host: process.env.FIREBIRD_HOST || '127.0.0.1',
    port: positiveInteger('FIREBIRD_PORT', 3050),
    database: process.env.FIREBIRD_DATABASE || '',
    user: process.env.FIREBIRD_USER || '',
    password: process.env.FIREBIRD_PASSWORD || ''
  }
};

config.validate = function validate() {
  const missing = [];
  if (!config.database.database) missing.push('FIREBIRD_DATABASE');
  if (!config.database.user) missing.push('FIREBIRD_USER');
  if (!config.database.password) missing.push('FIREBIRD_PASSWORD');
  if (!config.apiKey) missing.push('API_KEY');
  if (missing.length) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }
};

module.exports = config;
