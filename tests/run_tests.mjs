import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const codeGsPath = path.join(__dirname, '..', 'src', 'Code.gs');
const sidebarHtmlPath = path.join(__dirname, '..', 'src', 'Sidebar.html');
const manifestPath = path.join(__dirname, '..', 'src', 'appsscript.json');

function createMockAppsScriptEnv({ throwOnUserProps = false } = {}) {
  const userStore = new Map();
  const docStore = new Map();
  const scriptStore = new Map();
  const cacheStore = new Map();
  const fetchCalls = [];

  const makeProps = (store) => ({
    getProperty: (k) => (store.has(k) ? store.get(k) : null),
    setProperty: (k, v) => { store.set(k, String(v)); },
    deleteProperty: (k) => { store.delete(k); }
  });

  const makeHttpResponse = (status, bodyObj) => ({
    getResponseCode: () => status,
    getContentText: () => JSON.stringify(bodyObj)
  });

  const routeRequest = (url, opts) => {
    const payload = opts?.payload ? JSON.parse(opts.payload) : {};
    const auth = opts?.headers?.Authorization || '';
    fetchCalls.push({ url, auth, payload });

    if (auth !== 'Bearer TEST_VALID_KEY_123456') {
      return makeHttpResponse(401, {
        ok: false,
        error: { code: 'invalid_credentials', message: 'Invalid or inactive API key' }
      });
    }

    if (url.endsWith('/get_balance')) {
      return makeHttpResponse(200, {
        ok: true,
        data: {
          account_email: 'ops@proweblook.com',
          email: { amount: 5000, unlimited: false },
          phone: { amount: 1200, unlimited: false },
          whatsapp: { amount: 850, unlimited: false }
        }
      });
    }

    if (url.endsWith('/verify_email')) {
      const deliverable = payload.email === 'support@proweblook.com';
      return makeHttpResponse(200, {
        ok: true,
        data: {
          email: payload.email,
          verdict: deliverable ? 'deliverable' : 'undeliverable',
          reason: deliverable ? 'completed' : 'null_mx',
          billing_status: 'charged'
        }
      });
    }

    if (url.endsWith('/validate_phone')) {
      return makeHttpResponse(200, {
        ok: true,
        data: {
          verdict: 'valid',
          reason: 'completed',
          billing_status: 'upstream_managed',
          details: {
            status: 'VALID_CONFIRMED',
            formatted_number: '+12025550123',
            linetype: 'MOBILE',
            carrier: 'Verizon Wireless',
            countrycode: payload.country_code || 'US',
            countryname: 'United States of America'
          }
        }
      });
    }

    if (url.endsWith('/check_whatsapp')) {
      const onWa = payload.phone_number === '+14155552671';
      return makeHttpResponse(200, {
        ok: true,
        data: {
          verdict: onWa ? 'on_whatsapp' : 'not_on_whatsapp',
          reason: 'completed',
          billing_status: 'upstream_managed',
          details: {
            status: onWa ? 'VALID_CONFIRMED' : 'INVALID',
            numberstatus: onWa,
            businessnumber: true,
            formatted_number: payload.phone_number,
            countrycode: 'US'
          }
        }
      });
    }

    return makeHttpResponse(404, { ok: false, error: { code: 'unknown_operation', message: 'Not found' } });
  };

  // In-memory mock sheet for bulk verification tests
  const gridValues = Array.from({ length: 10 }, () => Array(10).fill(''));
  const gridBgs = Array.from({ length: 10 }, () => Array(10).fill(null));
  gridValues[0][1] = 'Email';
  gridValues[1][1] = 'support@proweblook.com';
  gridValues[2][1] = 'bounce@example.com';
  gridValues[3][1] = 'already@proweblook.com';
  gridValues[3][2] = 'deliverable';
  gridValues[3][3] = 'completed';

  let activeSelection = { row: 1, col: 2, numRows: 4, numCols: 1 };

  const mockSheet = {
    getName: () => 'Leads',
    getRange: (r, c, numR = 1, numC = 1) => ({
      getRow: () => r,
      getColumn: () => c,
      getNumRows: () => numR,
      getNumColumns: () => numC,
      getSheet: () => mockSheet,
      getValues: () => {
        const out = [];
        for (let i = 0; i < numR; i++) {
          out.push(gridValues[r - 1 + i].slice(c - 1, c - 1 + numC));
        }
        return out;
      },
      setValues: (vals) => {
        for (let i = 0; i < numR; i++) {
          for (let j = 0; j < numC; j++) {
            gridValues[r - 1 + i][c - 1 + j] = vals[i][j];
          }
        }
        return mockSheet.getRange(r, c, numR, numC);
      },
      setBackgrounds: (bgs) => {
        for (let i = 0; i < numR; i++) {
          for (let j = 0; j < numC; j++) {
            gridBgs[r - 1 + i][c - 1 + j] = bgs[i][j];
          }
        }
        return mockSheet.getRange(r, c, numR, numC);
      },
      setFontWeight: () => mockSheet.getRange(r, c, numR, numC),
      setBackground: () => mockSheet.getRange(r, c, numR, numC),
      setFontColor: () => mockSheet.getRange(r, c, numR, numC)
    }),
    setFrozenRows: () => {},
    autoResizeColumns: () => {}
  };

  const insertedSheets = [];
  const sandbox = {
    PropertiesService: {
      getUserProperties: () => {
        if (throwOnUserProps) {
          throw new Error('You do not have permission to call getUserProperties');
        }
        return makeProps(userStore);
      },
      getDocumentProperties: () => makeProps(docStore),
      getScriptProperties: () => makeProps(scriptStore)
    },
    CacheService: {
      getDocumentCache: () => ({
        get: (k) => (cacheStore.has(k) ? cacheStore.get(k) : null),
        put: (k, v) => { cacheStore.set(k, v); }
      }),
      getScriptCache: () => ({
        get: (k) => (cacheStore.has(k) ? cacheStore.get(k) : null),
        put: (k, v) => { cacheStore.set(k, v); }
      })
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest: (_alg, str) => Array.from(crypto.createHash('sha256').update(str, 'utf8').digest())
    },
    UrlFetchApp: {
      fetch: (url, opts) => routeRequest(url, opts),
      fetchAll: (reqs) => reqs.map((r) => routeRequest(r.url, r))
    },
    SpreadsheetApp: {
      getActiveRange: () => mockSheet.getRange(activeSelection.row, activeSelection.col, activeSelection.numRows, activeSelection.numCols),
      getActiveSpreadsheet: () => ({
        getSheetByName: (n) => (n === 'Leads' ? mockSheet : null),
        insertSheet: (n) => {
          insertedSheets.push(n);
          return mockSheet;
        },
        setActiveSheet: () => {},
        toast: () => {}
      }),
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        createMenu: () => ({
          addItem: function() { return this; },
          addSeparator: function() { return this; },
          addToUi: function() {}
        }),
        showSidebar: () => {},
        alert: () => {}
      })
    },
    ScriptApp: { AuthMode: { NONE: 'NONE' } },
    HtmlService: {
      createHtmlOutputFromFile: () => ({ setTitle: function() { return this; } })
    }
  };

  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(codeGsPath, 'utf8'), sandbox, { filename: 'Code.gs' });

  return { sandbox, fetchCalls, gridValues, gridBgs, userStore, docStore, cacheStore, insertedSheets, setSelection: (s) => { activeSelection = s; } };
}

console.log('Running ProWebLook Google Sheets Add-on test suite...\n');

// 1. Validate manifest & HTML assets
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
assert.equal(manifest.runtimeVersion, 'V8');
assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/spreadsheets.currentonly'));
assert.deepEqual(manifest.urlFetchWhitelist, ['https://proweblook.com/api/v1/zapier/']);
assert.ok(fs.readFileSync(sidebarHtmlPath, 'utf8').includes('google.script.run'));
console.log('✓ 1. Manifest least-privilege scopes, urlFetchWhitelist, and Sidebar HTML verified');

// 2. Test saveApiKey + getSidebarState (never leaks raw key)
const env = createMockAppsScriptEnv();
assert.throws(() => env.sandbox.saveApiKey('WRONG_KEY_9999999999', true), /Invalid or inactive API key/);
const savedState = env.sandbox.saveApiKey('TEST_VALID_KEY_123456', true);
assert.equal(savedState.connected, true);
assert.equal(savedState.formulasEnabled, true);
assert.equal(savedState.maskedKey, 'TEST••••••••3456');
assert.ok(!JSON.stringify(savedState).includes('TEST_VALID_KEY_123456'));
assert.equal(savedState.balances.email.amount, 5000);
console.log('✓ 2. API key validation, storage, and masked Sidebar state verified');

// 3. Test custom function execution under restricted anonymous custom-function permissions (getUserProperties throws!)
const cfEnv = createMockAppsScriptEnv({ throwOnUserProps: true });
cfEnv.docStore.set('PROWEBLOOK_API_KEY', 'TEST_VALID_KEY_123456');

const singleVerdict = cfEnv.sandbox.PROWEBLOOK_VERIFY_EMAIL('SUPPORT@proweblook.com');
assert.equal(singleVerdict, 'deliverable');
assert.equal(cfEnv.fetchCalls.length, 1);

const plain = (v) => JSON.parse(JSON.stringify(v));

// Repeat exact same call -> must hit CacheService with 0 extra network requests!
const cachedVerdict = cfEnv.sandbox.PROWEBLOOK_VERIFY_EMAIL('support@proweblook.com', true);
assert.deepEqual(plain(cachedVerdict), [['deliverable', 'completed', 'support@proweblook.com']]);
assert.equal(cfEnv.fetchCalls.length, 1, 'Expected CacheService hit with 0 extra HTTP calls');
console.log('✓ 3. PROWEBLOOK_VERIFY_EMAIL custom-function permission safety & 6-hour CacheService deduplication verified');

// 4. Test 2D range batching via UrlFetchApp.fetchAll
const rangeOut = cfEnv.sandbox.PROWEBLOOK_VERIFY_EMAIL([
  ['support@proweblook.com'], // already cached!
  ['dead@example.com'],
  [''],
  ['not-an-email']
]);
assert.deepEqual(plain(rangeOut), [
  ['deliverable'],
  ['undeliverable'],
  [''],
  ['invalid_syntax']
]);
assert.equal(cfEnv.fetchCalls.length, 2, 'Only the 1 uncached valid email should trigger an HTTP call');
console.log('✓ 4. 2D range batching, blank-cell skip, and syntax pre-check verified');

// 5. Test PROWEBLOOK_VALIDATE_PHONE & Google Sheets numeric '+' restoration + scientific notation guard
const phoneVerdict = cfEnv.sandbox.PROWEBLOOK_VALIDATE_PHONE(12025550123, '', true); // numeric cell where Sheets ate '+'
assert.deepEqual(plain(phoneVerdict), [['valid', 'MOBILE', 'Verizon Wireless', 'US', '+12025550123']]);
assert.equal(cfEnv.fetchCalls.at(-1).payload.phone_number, '+12025550123');
assert.equal('country_code' in cfEnv.fetchCalls.at(-1).payload, false);

const nationalVerdict = cfEnv.sandbox.PROWEBLOOK_VALIDATE_PHONE('(202) 555-0123', 'us', false);
assert.equal(nationalVerdict, 'valid');
assert.equal(cfEnv.fetchCalls.at(-1).payload.country_code, 'US');

const sciGuard = cfEnv.sandbox.PROWEBLOOK_VALIDATE_PHONE('1.41556E+10');
assert.equal(sciGuard, 'invalid_format');
console.log('✓ 5. PROWEBLOOK_VALIDATE_PHONE numeric + restoration, ISO normalization, and scientific-notation guard verified');

// 6. Test PROWEBLOOK_WHATSAPP & PROWEBLOOK_CHECK_WHATSAPP
const waFull = cfEnv.sandbox.PROWEBLOOK_WHATSAPP(14155552671, true);
assert.deepEqual(plain(waFull), [['on_whatsapp', 'business', '+14155552671', 'US']]);
const waAlias = cfEnv.sandbox.PROWEBLOOK_CHECK_WHATSAPP('+12025550199', false);
assert.equal(waAlias, 'not_on_whatsapp');
console.log('✓ 6. PROWEBLOOK_WHATSAPP & PROWEBLOOK_CHECK_WHATSAPP verified');

// 7. Test PROWEBLOOK_BALANCE
assert.equal(cfEnv.sandbox.PROWEBLOOK_BALANCE('email'), 5000);
assert.deepEqual(plain(cfEnv.sandbox.PROWEBLOOK_BALANCE()), [
  ['Email Credits', 5000],
  ['Phone Credits', 1200],
  ['WhatsApp Credits', 850]
]);
console.log('✓ 7. PROWEBLOOK_BALANCE verified');

// 8. Test Sidebar Bulk Enricher (header handling, skipExisting, permanent cell writes, background colors)
const bulkMeta = env.sandbox.prepareBulkSelection('verify_email', true);
assert.equal(bulkMeta.startRow, 2);
assert.equal(bulkMeta.totalRows, 3);
assert.deepEqual(env.gridValues[0].slice(2, 4), ['Email Verdict', 'Reason']);

const chunkRes = env.sandbox.processBulkChunk({
  sheetName: bulkMeta.sheetName,
  col: bulkMeta.col,
  startRow: bulkMeta.startRow,
  totalRows: bulkMeta.totalRows,
  offset: 0,
  limit: 15,
  operation: 'verify_email',
  countryCode: '',
  skipExisting: true
});
assert.equal(chunkRes.done, true);
assert.equal(chunkRes.processed, 2); // row 2 & row 3 verified; row 4 skipped because already populated!
assert.equal(chunkRes.skipped, 1);
assert.deepEqual(env.gridValues[1].slice(2, 4), ['deliverable', 'completed']);
assert.deepEqual(env.gridValues[2].slice(2, 4), ['undeliverable', 'null_mx']);
assert.equal(env.gridBgs[1][2], '#e6f4ea');
assert.equal(env.gridBgs[2][2], '#fce8e6');

const demoRes = env.sandbox.insertDemoSheet();
assert.equal(demoRes.sheetName, 'ProWebLook Demo');
console.log('✓ 8. Sidebar Bulk Enricher (headers, skipExisting, status colors, and Demo Sheet) verified');

console.log('\nALL 8 TEST SUITES PASSED!');
