'use strict';

const Firebird = require('node-firebird');
const config = require('./index');

function query(sql, params = []) {
  return new Promise((resolve, reject) => {
    const connectionOptions = {
      ...config.database,
      wireCrypt: Firebird.WIRE_CRYPT_ENABLE
    };

    Firebird.attach(connectionOptions, (attachError, db) => {
      if (attachError) return reject(attachError);

      db.query(sql, params, (queryError, rows) => {
        db.detach();
        if (queryError) return reject(queryError);
        resolve(rows);
      });
    });
  });
}

module.exports = { query };
