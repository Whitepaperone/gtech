'use strict';

const express = require('express');
const config = require('../config');
const shipmentChangeService = require('../services/shipmentChangeService');
const snapshotService = require('../services/snapshotService');

const router = express.Router();

router.get('/changes', async (req, res, next) => {
  try {
    res.json(await shipmentChangeService.getChanges());
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
