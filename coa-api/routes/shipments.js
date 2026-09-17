'use strict';

const express = require('express');
const config = require('../config');
const acknowledgementService = require('../services/acknowledgementService');
const shipmentChangeService = require('../services/shipmentChangeService');
const snapshotService = require('../services/snapshotService');

const router = express.Router();

router.get('/changes', async (req, res, next) => {
  try {
    res.json(await shipmentChangeService.getChanges(req.query.consumer));
  } catch (error) {
    next(error);
  }
});

router.post('/acknowledgements', (req, res, next) => {
  try {
    const body = req.body || {};
    const acknowledgedCount = acknowledgementService.acknowledge(
      body.consumer,
      body.changeIds,
      body.metadata
    );
    res.json({ ok: true, consumer: body.consumer, acknowledgedCount });
  } catch (error) {
    next(error);
  }
});

router.get('/snapshot-status', (req, res, next) => {
  try {
    res.json({ ok: true, snapshotDirectory: config.snapshotDir, files: snapshotService.listSnapshotDates() });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
