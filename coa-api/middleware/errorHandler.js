'use strict';

module.exports = function errorHandler(err, req, res, next) {
  console.error(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`, err);
  if (res.headersSent) return next(err);
  res.status(500).json({ ok: false, error: 'Internal server error' });
};
