'use strict';

module.exports = function errorHandler(err, req, res, next) {
  console.error(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`, err);
  if (res.headersSent) return next(err);
  const statusCode = Number.isInteger(err.statusCode) ? err.statusCode : 500;
  res.status(statusCode).json({
    ok: false,
    error: statusCode < 500 ? err.message : 'Internal server error'
  });
};
