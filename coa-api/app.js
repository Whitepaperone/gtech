'use strict';

require('dotenv').config();

const express = require('express');
const helmet = require('helmet');
const { rateLimit } = require('express-rate-limit');
const config = require('./config');
const auth = require('./middleware/auth');
const errorHandler = require('./middleware/errorHandler');
const orderRoutes = require('./routes/orders');
const shipmentRoutes = require('./routes/shipments');

function createApp() {
  const app = express();

  app.disable('x-powered-by');
  app.use(helmet());
  app.use(express.json({ limit: config.jsonLimit }));
  app.use(rateLimit({
    windowMs: config.rateLimitWindowMs,
    limit: config.rateLimitMax,
    standardHeaders: 'draft-7',
    legacyHeaders: false
  }));

  app.get('/', (req, res) => {
    res.json({ status: 'ok', service: 'COA Inventory API' });
  });

  app.use(auth);
  app.use('/orders', orderRoutes);
  app.use('/shipments', shipmentRoutes);
  app.use((req, res) => res.status(404).json({ ok: false, error: 'Not found' }));
  app.use(errorHandler);

  return app;
}

if (require.main === module) {
  config.validate();
  createApp().listen(config.port, '0.0.0.0', () => {
    console.log(`COA Inventory API started on port ${config.port}`);
  });
}

module.exports = { createApp };
