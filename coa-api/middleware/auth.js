'use strict';

const crypto = require('crypto');
const config = require('../config');

function safeEqual(actual, expected) {
  const actualBuffer = Buffer.from(actual, 'utf8');
  const expectedBuffer = Buffer.from(expected, 'utf8');
  if (actualBuffer.length !== expectedBuffer.length) return false;
  return crypto.timingSafeEqual(actualBuffer, expectedBuffer);
}

module.exports = function auth(req, res, next) {
  const header = req.get('authorization') || '';
  const match = /^Bearer\s+(.+)$/i.exec(header);
  if (!match || !config.apiKey || !safeEqual(match[1], config.apiKey)) {
    return res.status(401).json({ ok: false, error: 'Unauthorized' });
  }
  next();
};

module.exports.safeEqual = safeEqual;
