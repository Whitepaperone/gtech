'use strict';

const express = require('express');
const orderRepository = require('../repositories/orderRepository');

const router = express.Router();

router.get('/', async (req, res, next) => {
  try {
    const rows = await orderRepository.findOpenOrders();
    res.json({ ok: true, count: rows.length, data: rows });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
