'use strict';

const KOREA_OFFSET_MS = 9 * 60 * 60 * 1000;

function getKoreaDate(offsetDays = 0, now = new Date()) {
  const koreaTime = new Date(now.getTime() + KOREA_OFFSET_MS);
  koreaTime.setUTCDate(koreaTime.getUTCDate() + offsetDays);
  const year = koreaTime.getUTCFullYear();
  const month = String(koreaTime.getUTCMonth() + 1).padStart(2, '0');
  const day = String(koreaTime.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function dateOnly(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return String(value).substring(0, 10);
}

module.exports = { getKoreaDate, dateOnly };
