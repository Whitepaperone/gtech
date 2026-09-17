'use strict';

// Local Excel helper: reads codes from stdin and queries Firebird directly.
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { findBom } = require('./repositories/bomRepository');
const config = require('./config');
let input = '';
function emit(event) {
  process.stdout.write(JSON.stringify(event) + '\n');
}
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => { input += chunk; });
process.stdin.on('end', async () => {
  try {
    if (!config.database.database || !config.database.user || !config.database.password) {
      throw new Error('coa-api/.env DB settings are missing');
    }
    const codes = JSON.parse(input);
    if (!Array.isArray(codes) || codes.some(code => typeof code !== 'string')) {
      throw new Error('Expected an array of string codes');
    }
    const output = {};
    const uniqueCodes = [...new Set(codes)];
    const started = Date.now();
    let completed = 0;
    for (const code of uniqueCodes) {
      const itemStarted = Date.now();
      const progress = status => emit({ TYPE: 'PROGRESS', STATUS: status,
        CODE: code, COMPLETED: completed, TOTAL: uniqueCodes.length,
        ELAPSED: Math.round((Date.now() - started) / 1000),
        ITEM_SECONDS: Math.round((Date.now() - itemStarted) / 1000) });
      progress('START');
      const heartbeat = setInterval(() => progress('WAIT'), 5000);
      const timer = setTimeout(() => {
        emit({ TYPE: 'ERROR', MESSAGE: `DB timeout after 45 seconds: ${code}` });
        process.exit(1);
      }, 45000);
      try {
        const bom = await findBom({ CODE: code });
        output[code] = bom ? bom.DATA : [];
        completed += 1;
        progress('DONE');
      } finally { clearTimeout(timer); clearInterval(heartbeat); }
    }
    emit({ TYPE: 'RESULT', DATA: output });
  } catch (error) {
    emit({ TYPE: 'ERROR', MESSAGE: error.message });
    process.exitCode = 1;
  }
});
