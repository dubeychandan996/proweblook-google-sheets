/**
 * @OnlyCurrentDoc
 *
 * ProWebLook for Google Sheets — Official Add-on & Apps Script Starter
 * Verify emails, validate global phone numbers, and check WhatsApp registration
 * directly inside Google Sheets via custom formulas or the Bulk Enricher sidebar.
 *
 * Website: https://proweblook.com
 * API Profile: https://proweblook.com/user/my_profile
 * License: MIT
 */

var PROWEBLOOK_API_BASE = 'https://proweblook.com/api/v1/zapier';
var PROWEBLOOK_PROP_KEY = 'PROWEBLOOK_API_KEY';
var PROWEBLOOK_CACHE_TTL_SECONDS = 21600; // 6 hours (Apps Script CacheService maximum)
var PROWEBLOOK_MAX_FORMULA_BATCH = 25;
var PROWEBLOOK_PARALLEL_CHUNK = 10;

/* ==========================================================================
 * 1. ADD-ON LIFECYCLE & MENU
 * ========================================================================== */

/**
 * Runs when the add-on is installed; delegates to onOpen.
 * @param {Object} e Apps Script install event.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Adds the ProWebLook menu when the spreadsheet opens.
 * @param {Object} e Apps Script open event.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu ? ui.createAddonMenu() : ui.createMenu('ProWebLook');

  // When running as a container-bound script (Extensions -> Apps Script), also attach a top-level menu.
  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    menu.addItem('Open ProWebLook Sidebar', 'showProWebLookSidebar');
    menu.addToUi();
    return;
  }

  ui.createMenu('ProWebLook')
    .addItem('Open Sidebar (Connect & Bulk Verify)', 'showProWebLookSidebar')
    .addSeparator()
    .addItem('Bulk Verify Selected Emails', 'menuBulkVerifyEmails')
    .addItem('Bulk Validate Selected Phones', 'menuBulkValidatePhones')
    .addItem('Bulk Check Selected WhatsApp Numbers', 'menuBulkCheckWhatsApp')
    .addSeparator()
    .addItem('Check Credit Balance', 'menuShowBalanceAlert')
    .addItem('Insert Sample Lead Sheet', 'insertDemoSheet')
    .addToUi();
}

/**
 * Opens the ProWebLook HTML Sidebar inside Google Sheets.
 */
function showProWebLookSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('ProWebLook Lead Verifier');
  SpreadsheetApp.getUi().showSidebar(html);
}

function menuBulkVerifyEmails() {
  runQuickMenuBulk_('verify_email');
}

function menuBulkValidatePhones() {
  runQuickMenuBulk_('validate_phone');
}

function menuBulkCheckWhatsApp() {
  runQuickMenuBulk_('check_whatsapp');
}

function menuShowBalanceAlert() {
  var ui = SpreadsheetApp.getUi();
  var key = getApiKey_();
  if (!key) {
    ui.alert('ProWebLook Not Connected', 'Open ProWebLook -> Open Sidebar to connect your API Key from https://proweblook.com/user/my_profile.', ui.ButtonSet.OK);
    return;
  }
  var res = callProWebLookApi_('get_balance', {}, key, false);
  if (!res.ok) {
    ui.alert('ProWebLook Error', res.message, ui.ButtonSet.OK);
    return;
  }
  var d = res.data || {};
  var msg = [
    'Account: ' + (d.account_email || 'Connected'),
    '',
    'Email Verifier Credits: ' + formatBalanceAmount_(d.email),
    'Phone Validation Credits: ' + formatBalanceAmount_(d.phone),
    'WhatsApp Checker Credits: ' + formatBalanceAmount_(d.whatsapp)
  ].join('\n');
  ui.alert('ProWebLook Credit Balance', msg, ui.ButtonSet.OK);
}

/* ==========================================================================
 * 2. CUSTOM SPREADSHEET FUNCTIONS
 * ========================================================================== */

/**
 * Verifies an email address (or column range of emails) using ProWebLook.
 * Returns "deliverable", "undeliverable", or "unknown" (or 3 columns when fullDetails is TRUE).
 *
 * @param {string|Array<Array<string>>} email Email address or single-column range (e.g. A2 or A2:A20).
 * @param {boolean=} fullDetails Optional. Set TRUE to return [verdict, reason, normalized_email].
 * @return {string|Array<Array<string>>} Verification verdict or detail columns.
 * @customfunction
 */
function PROWEBLOOK_VERIFY_EMAIL(email, fullDetails) {
  var showDetails = Boolean(fullDetails);
  return executeCustomFunctionGrid_(email, function(rawVal) {
    var cleaned = sanitizeCellString_(rawVal).toLowerCase();
    if (!cleaned) {
      return { skip: true, output: showDetails ? ['', '', ''] : '' };
    }
    if (cleaned.indexOf('@') === -1 || cleaned.length > 320) {
      return { skip: true, output: showDetails ? ['invalid_syntax', 'invalid_email_format', cleaned] : 'invalid_syntax' };
    }
    return {
      skip: false,
      operation: 'verify_email',
      payload: { email: cleaned },
      format: function(res) {
        if (!res.ok) {
          var errText = 'ERROR: ' + res.message;
          return showDetails ? [errText, res.code || 'error', cleaned] : errText;
        }
        var d = res.data || {};
        var verdict = String(d.verdict || 'unknown');
        return showDetails
          ? [verdict, String(d.reason || 'completed'), String(d.email || cleaned)]
          : verdict;
      }
    };
  }, showDetails ? 3 : 1);
}

/**
 * Validates a phone number (or column range) and returns carrier & line type details.
 * Returns "valid", "invalid", or "unknown" (or 5 columns when fullDetails is TRUE).
 *
 * @param {string|number|Array<Array<string>>} phone Phone number in E.164 (+12025550123) or national format with countryCode.
 * @param {string=} countryCode Optional 2-letter ISO country code (e.g. "US", "IN", "GB") for national numbers.
 * @param {boolean=} fullDetails Optional. Set TRUE to return [verdict, linetype, carrier, country_code, e164_number].
 * @return {string|Array<Array<string>>} Phone validation verdict or detail columns.
 * @customfunction
 */
function PROWEBLOOK_VALIDATE_PHONE(phone, countryCode, fullDetails) {
  var isoCountry = normalizeCountryCode_(countryCode);
  var showDetails = Boolean(fullDetails);
  if (countryCode && String(countryCode).trim() !== '' && !isoCountry) {
    throw new Error('countryCode must be a 2-letter ISO code such as "US", "GB", or "IN".');
  }

  return executeCustomFunctionGrid_(phone, function(rawVal) {
    var rawStr = sanitizeCellString_(rawVal);
    if (!rawStr) {
      return { skip: true, output: showDetails ? ['', '', '', '', ''] : '' };
    }
    var normalized = normalizePhoneInput_(rawVal, isoCountry, !isoCountry);
    if (!normalized.ok) {
      return {
        skip: true,
        output: showDetails ? ['invalid_format', normalized.reason, '', isoCountry, rawStr] : 'invalid_format'
      };
    }
    var payload = { phone_number: normalized.phone };
    if (isoCountry) {
      payload.country_code = isoCountry;
    }
    return {
      skip: false,
      operation: 'validate_phone',
      payload: payload,
      format: function(res) {
        if (!res.ok) {
          var errText = 'ERROR: ' + res.message;
          return showDetails ? [errText, '', '', isoCountry, normalized.phone] : errText;
        }
        var d = res.data || {};
        var details = d.details || {};
        var verdict = String(d.verdict || 'unknown');
        return showDetails
          ? [
              verdict,
              String(details.linetype || ''),
              String(details.carrier || ''),
              String(details.countrycode || isoCountry || ''),
              String(details.formatted_number || normalized.phone)
            ]
          : verdict;
      }
    };
  }, showDetails ? 5 : 1);
}

/**
 * Checks whether a phone number is registered on WhatsApp (sends NO message or OTP).
 * Returns "on_whatsapp" or "not_on_whatsapp" (or 4 columns when fullDetails is TRUE).
 *
 * @param {string|number|Array<Array<string>>} phone Phone number in international E.164 format (e.g. "+14155552671").
 * @param {boolean=} fullDetails Optional. Set TRUE to return [verdict, account_type, formatted_number, country_code].
 * @return {string|Array<Array<string>>} WhatsApp registration verdict or detail columns.
 * @customfunction
 */
function PROWEBLOOK_WHATSAPP(phone, fullDetails) {
  var showDetails = Boolean(fullDetails);
  return executeCustomFunctionGrid_(phone, function(rawVal) {
    var rawStr = sanitizeCellString_(rawVal);
    if (!rawStr) {
      return { skip: true, output: showDetails ? ['', '', '', ''] : '' };
    }
    var normalized = normalizePhoneInput_(rawVal, '', true);
    if (!normalized.ok) {
      return {
        skip: true,
        output: showDetails ? ['invalid_format', normalized.reason, rawStr, ''] : 'invalid_format'
      };
    }
    return {
      skip: false,
      operation: 'check_whatsapp',
      payload: { phone_number: normalized.phone },
      format: function(res) {
        if (!res.ok) {
          var errText = 'ERROR: ' + res.message;
          return showDetails ? [errText, '', normalized.phone, ''] : errText;
        }
        var d = res.data || {};
        var details = d.details || {};
        var verdict = String(d.verdict || 'not_on_whatsapp');
        var accountType = '';
        if (verdict === 'on_whatsapp' && typeof details.businessnumber === 'boolean') {
          accountType = details.businessnumber ? 'business' : 'personal';
        }
        return showDetails
          ? [
              verdict,
              accountType,
              String(details.formatted_number || normalized.phone),
              String(details.countrycode || '')
            ]
          : verdict;
      }
    };
  }, showDetails ? 4 : 1);
}

/**
 * Alias for PROWEBLOOK_WHATSAPP — checks if a phone number has an active WhatsApp account.
 *
 * @param {string|number|Array<Array<string>>} phone Phone number in international E.164 format (e.g. "+14155552671").
 * @param {boolean=} fullDetails Optional. Set TRUE to return [verdict, account_type, formatted_number, country_code].
 * @return {string|Array<Array<string>>} WhatsApp registration verdict or detail columns.
 * @customfunction
 */
function PROWEBLOOK_CHECK_WHATSAPP(phone, fullDetails) {
  return PROWEBLOOK_WHATSAPP(phone, fullDetails);
}

/**
 * Returns remaining ProWebLook credit balances (free — consumes 0 credits).
 *
 * @param {string=} product Optional: "email", "phone", or "whatsapp". Omit to return all 3 rows.
 * @return {number|string|Array<Array<string|number>>} Credit balance(s).
 * @customfunction
 */
function PROWEBLOOK_BALANCE(product) {
  var key = getApiKey_();
  if (!key) {
    return 'Connect API key in ProWebLook -> Open Sidebar (and enable sheet formulas)';
  }
  var res = callProWebLookApi_('get_balance', {}, key, false);
  if (!res.ok) {
    return 'ERROR: ' + res.message;
  }
  var d = res.data || {};
  if (product && String(product).trim() !== '') {
    var p = String(product).trim().toLowerCase();
    if (!d[p] || typeof d[p].amount !== 'number') {
      throw new Error('product must be "email", "phone", or "whatsapp".');
    }
    return d[p].amount;
  }
  return [
    ['Email Credits', d.email ? d.email.amount : 0],
    ['Phone Credits', d.phone ? d.phone.amount : 0],
    ['WhatsApp Credits', d.whatsapp ? d.whatsapp.amount : 0]
  ];
}

/* ==========================================================================
 * 3. SIDEBAR RPC HANDLERS
 * ========================================================================== */

/**
 * Returns connection status, formula-access state, and credit balances for Sidebar initialization.
 * Never returns the raw API key to the client DOM.
 */
function getSidebarState() {
  var docKey = getDocumentApiKey_();
  var userKey = getUserApiKey_();
  var activeKey = userKey || docKey || getScriptApiKey_();
  if (!activeKey) {
    return {
      connected: false,
      formulasEnabled: false,
      maskedKey: '',
      balances: null
    };
  }
  var balanceRes = callProWebLookApi_('get_balance', {}, activeKey, false);
  return {
    connected: balanceRes.ok,
    formulasEnabled: Boolean(docKey),
    maskedKey: maskApiKey_(activeKey),
    balances: balanceRes.ok ? balanceRes.data : null,
    error: balanceRes.ok ? '' : balanceRes.message
  };
}

/**
 * Validates a ProWebLook API key via get_balance and stores it in UserProperties
 * (plus DocumentProperties when enableFormulas is true).
 *
 * @param {string} rawApiKey User's ProWebLook API key.
 * @param {boolean} enableFormulas Whether custom formulas (=PROWEBLOOK_...) should work in this sheet.
 */
function saveApiKey(rawApiKey, enableFormulas) {
  var key = sanitizeCellString_(rawApiKey);
  if (!key || !/^[A-Za-z0-9._~+\/=-]{10,255}$/.test(key)) {
    throw new Error('Please paste a valid ProWebLook API Key from https://proweblook.com/user/my_profile.');
  }
  var check = callProWebLookApi_('get_balance', {}, key, false);
  if (!check.ok) {
    throw new Error(check.message || 'Invalid ProWebLook API Key.');
  }

  PropertiesService.getUserProperties().setProperty(PROWEBLOOK_PROP_KEY, key);
  if (enableFormulas) {
    PropertiesService.getDocumentProperties().setProperty(PROWEBLOOK_PROP_KEY, key);
  } else {
    PropertiesService.getDocumentProperties().deleteProperty(PROWEBLOOK_PROP_KEY);
  }

  return {
    connected: true,
    formulasEnabled: Boolean(enableFormulas),
    maskedKey: maskApiKey_(key),
    balances: check.data
  };
}

/**
 * Toggles DocumentProperties storage so custom formulas (=PROWEBLOOK_...) can read the API key.
 * @param {boolean} enabled Enable or disable formula access in the active spreadsheet.
 */
function setFormulaAccess(enabled) {
  var key = getUserApiKey_() || getDocumentApiKey_();
  if (!key) {
    throw new Error('Connect your ProWebLook API Key first.');
  }
  if (enabled) {
    PropertiesService.getDocumentProperties().setProperty(PROWEBLOOK_PROP_KEY, key);
  } else {
    PropertiesService.getDocumentProperties().deleteProperty(PROWEBLOOK_PROP_KEY);
  }
  return { formulasEnabled: Boolean(enabled) };
}

/**
 * Removes stored ProWebLook API keys from both UserProperties and DocumentProperties.
 */
function disconnectApiKey() {
  try {
    PropertiesService.getUserProperties().deleteProperty(PROWEBLOOK_PROP_KEY);
  } catch (e) {}
  try {
    PropertiesService.getDocumentProperties().deleteProperty(PROWEBLOOK_PROP_KEY);
  } catch (e) {}
  return { connected: false, formulasEnabled: false };
}

/**
 * Inspects the user's currently highlighted range in the active sheet prior to bulk verification.
 */
function prepareBulkSelection(operation, hasHeader) {
  var key = getApiKey_();
  if (!key) {
    throw new Error('Please connect your ProWebLook API Key first.');
  }
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    throw new Error('Select a column of emails or phone numbers in the sheet first.');
  }
  if (range.getNumColumns() > 1) {
    throw new Error('Please select a single column of cells (for example B2:B100). Results will be written to the adjacent columns on the right.');
  }

  var sheet = range.getSheet();
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var col = range.getColumn();

  var headers = getBulkHeaders_(operation);
  if (hasHeader && numRows > 1) {
    sheet.getRange(startRow, col + 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f1f5f9');
    startRow += 1;
    numRows -= 1;
  }

  return {
    sheetName: sheet.getName(),
    col: col,
    startRow: startRow,
    totalRows: numRows,
    outputCols: headers.length
  };
}

/**
 * Processes a chunk of rows during Sidebar bulk verification and writes permanent values + status colors.
 */
function processBulkChunk(job) {
  var key = getApiKey_();
  if (!key) {
    throw new Error('ProWebLook API Key is disconnected.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(job.sheetName);
  if (!sheet) {
    throw new Error('Target sheet "' + job.sheetName + '" was not found.');
  }

  var rowOffset = Number(job.offset) || 0;
  var limit = Math.min(20, Math.max(1, Number(job.limit) || 15));
  var remaining = Number(job.totalRows) - rowOffset;
  if (remaining <= 0) {
    return { processed: 0, skipped: 0, done: true };
  }

  var batchCount = Math.min(limit, remaining);
  var currentStartRow = Number(job.startRow) + rowOffset;
  var col = Number(job.col);
  var outputCols = getBulkHeaders_(job.operation).length;
  var isoCountry = normalizeCountryCode_(job.countryCode || '');

  var inputValues = sheet.getRange(currentStartRow, col, batchCount, 1).getValues();
  var existingOutputs = sheet.getRange(currentStartRow, col + 1, batchCount, outputCols).getValues();

  var outValues = [];
  var outColors = [];
  var pendingCalls = [];
  var skippedCount = 0;

  for (var i = 0; i < batchCount; i++) {
    var rawCell = inputValues[i][0];
    var rawStr = sanitizeCellString_(rawCell);
    var existingVerdict = sanitizeCellString_(existingOutputs[i][0]);

    if (!rawStr) {
      outValues.push(emptyRow_(outputCols));
      outColors.push(fillRow_(outputCols, null));
      skippedCount++;
      continue;
    }

    if (job.skipExisting && existingVerdict && existingVerdict.indexOf('ERROR:') !== 0) {
      outValues.push(existingOutputs[i]);
      outColors.push(fillRow_(outputCols, verdictColor_(existingVerdict)));
      skippedCount++;
      continue;
    }

    var spec = buildBulkSpec_(job.operation, rawCell, isoCountry);
    if (spec.skip) {
      outValues.push(spec.row);
      outColors.push(fillRow_(outputCols, verdictColor_(spec.row[0])));
      skippedCount++;
    } else {
      outValues.push(null);
      outColors.push(null);
      pendingCalls.push({ index: i, operation: spec.operation, payload: spec.payload, format: spec.format });
    }
  }

  if (pendingCalls.length > 0) {
    var apiResults = executeBatchApiCalls_(pendingCalls, key, true);
    for (var j = 0; j < pendingCalls.length; j++) {
      var item = pendingCalls[j];
      var formattedRow = item.format(apiResults[j]);
      outValues[item.index] = formattedRow;
      outColors[item.index] = fillRow_(outputCols, verdictColor_(formattedRow[0]));
    }
  }

  var outRange = sheet.getRange(currentStartRow, col + 1, batchCount, outputCols);
  outRange.setValues(outValues);
  outRange.setBackgrounds(outColors);

  var nextOffset = rowOffset + batchCount;
  var done = nextOffset >= Number(job.totalRows);
  var balances = null;
  if (done) {
    var balRes = callProWebLookApi_('get_balance', {}, key, false);
    if (balRes.ok) {
      balances = balRes.data;
    }
  }

  return {
    processed: batchCount - skippedCount,
    skipped: skippedCount,
    nextOffset: nextOffset,
    done: done,
    balances: balances
  };
}

/**
 * Inserts a ready-to-test "ProWebLook Demo" worksheet with sample leads and instructions.
 */
function insertDemoSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseName = 'ProWebLook Demo';
  var name = baseName;
  var idx = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + ' ' + idx++;
  }
  var sheet = ss.insertSheet(name);
  var rows = [
    ['Lead Name', 'Email Address', 'Email Verdict', 'Email Reason', 'Phone Number (E.164)', 'Phone Verdict', 'WhatsApp Status'],
    ['Alex Rivera', 'support@proweblook.com', '', '', "'+12025550123", '', ''],
    ['Jordan Taylor', 'test@example.com', '', '', "'+14155552671", '', ''],
    ['Samir Patel', 'hello@github.com', '', '', "'+442079460958", '', '']
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight('bold')
    .setBackground('#0648B2')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
  ss.setActiveSheet(sheet);
  return { sheetName: name };
}

/* ==========================================================================
 * 4. INTERNAL HELPERS, BATCHING & CACHING
 * ========================================================================== */

function runQuickMenuBulk_(operation) {
  var ui = SpreadsheetApp.getUi();
  try {
    var job = prepareBulkSelection(operation, false);
    if (job.totalRows > 40) {
      showProWebLookSidebar();
      ui.alert('Large Selection Detected', 'Use the ProWebLook Sidebar on the right to bulk-verify ' + job.totalRows + ' rows with live progress.', ui.ButtonSet.OK);
      return;
    }
    var offset = 0;
    var totalProcessed = 0;
    while (offset < job.totalRows) {
      var res = processBulkChunk({
        sheetName: job.sheetName,
        col: job.col,
        startRow: job.startRow,
        totalRows: job.totalRows,
        offset: offset,
        limit: 15,
        operation: operation,
        countryCode: '',
        skipExisting: true
      });
      totalProcessed += res.processed;
      offset = res.nextOffset;
      if (res.done) break;
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Verified ' + totalProcessed + ' row(s).', 'ProWebLook', 5);
  } catch (err) {
    ui.alert('ProWebLook', err.message || String(err), ui.ButtonSet.OK);
  }
}

function getBulkHeaders_(operation) {
  if (operation === 'verify_email') {
    return ['Email Verdict', 'Reason'];
  }
  if (operation === 'validate_phone') {
    return ['Phone Verdict', 'Line Type', 'Carrier', 'Country', 'E.164 Number'];
  }
  return ['WhatsApp Status', 'Account Type', 'Country'];
}

function buildBulkSpec_(operation, rawCell, isoCountry) {
  if (operation === 'verify_email') {
    var email = sanitizeCellString_(rawCell).toLowerCase();
    if (email.indexOf('@') === -1 || email.length > 320) {
      return { skip: true, row: ['invalid_syntax', 'invalid_email_format'] };
    }
    return {
      skip: false,
      operation: 'verify_email',
      payload: { email: email },
      format: function(res) {
        if (!res.ok) return ['ERROR: ' + res.message, res.code || 'error'];
        var d = res.data || {};
        return [String(d.verdict || 'unknown'), String(d.reason || 'completed')];
      }
    };
  }

  var requirePlus = operation === 'check_whatsapp' || !isoCountry;
  var norm = normalizePhoneInput_(rawCell, isoCountry, requirePlus);
  if (!norm.ok) {
    return {
      skip: true,
      row: operation === 'validate_phone'
        ? ['invalid_format', norm.reason, '', isoCountry, sanitizeCellString_(rawCell)]
        : ['invalid_format', norm.reason, '']
    };
  }

  if (operation === 'validate_phone') {
    var phonePayload = { phone_number: norm.phone };
    if (isoCountry) phonePayload.country_code = isoCountry;
    return {
      skip: false,
      operation: 'validate_phone',
      payload: phonePayload,
      format: function(res) {
        if (!res.ok) return ['ERROR: ' + res.message, '', '', isoCountry, norm.phone];
        var d = res.data || {};
        var det = d.details || {};
        return [
          String(d.verdict || 'unknown'),
          String(det.linetype || ''),
          String(det.carrier || ''),
          String(det.countrycode || isoCountry || ''),
          String(det.formatted_number || norm.phone)
        ];
      }
    };
  }

  return {
    skip: false,
    operation: 'check_whatsapp',
    payload: { phone_number: norm.phone },
    format: function(res) {
      if (!res.ok) return ['ERROR: ' + res.message, '', ''];
      var d = res.data || {};
      var det = d.details || {};
      var verdict = String(d.verdict || 'not_on_whatsapp');
      var acctType = (verdict === 'on_whatsapp' && typeof det.businessnumber === 'boolean')
        ? (det.businessnumber ? 'business' : 'personal')
        : '';
      return [verdict, acctType, String(det.countrycode || '')];
    }
  };
}

function executeCustomFunctionGrid_(input, rowBuilder, width) {
  var isGrid = Array.isArray(input);
  var grid = isGrid ? input : [[input]];
  var key = getApiKey_();

  var results = [];
  var pending = [];

  for (var r = 0; r < grid.length; r++) {
    var row = Array.isArray(grid[r]) ? grid[r] : [grid[r]];
    if (row.length > 1) {
      throw new Error('Select a single-column range (for example A2:A20).');
    }
    var spec = rowBuilder(row[0]);
    if (spec.skip) {
      results.push(Array.isArray(spec.output) ? spec.output : [spec.output]);
    } else if (!key) {
      var missingMsg = 'Connect API key in ProWebLook -> Open Sidebar (enable sheet formulas)';
      var missingRow = emptyRow_(width);
      missingRow[0] = missingMsg;
      results.push(missingRow);
    } else {
      var cached = readCachedResult_(spec.operation, spec.payload);
      if (cached) {
        var formattedCached = spec.format(cached);
        results.push(Array.isArray(formattedCached) ? formattedCached : [formattedCached]);
      } else {
        results.push(null);
        pending.push({ index: r, operation: spec.operation, payload: spec.payload, format: spec.format });
      }
    }
  }

  if (pending.length > PROWEBLOOK_MAX_FORMULA_BATCH) {
    throw new Error('Range has ' + pending.length + ' uncached rows (max ' + PROWEBLOOK_MAX_FORMULA_BATCH + ' per formula). Use ProWebLook -> Open Sidebar for unlimited bulk verification.');
  }

  if (pending.length > 0) {
    var apiResponses = executeBatchApiCalls_(pending, key, true);
    for (var i = 0; i < pending.length; i++) {
      var item = pending[i];
      var out = item.format(apiResponses[i]);
      results[item.index] = Array.isArray(out) ? out : [out];
    }
  }

  if (!isGrid && width === 1) {
    return results[0][0];
  }
  return results;
}

function executeBatchApiCalls_(items, apiKey, useCache) {
  var responses = new Array(items.length);
  var toFetch = [];

  for (var i = 0; i < items.length; i++) {
    if (useCache) {
      var hit = readCachedResult_(items[i].operation, items[i].payload);
      if (hit) {
        responses[i] = hit;
        continue;
      }
    }
    toFetch.push({ pos: i, operation: items[i].operation, payload: items[i].payload });
  }

  for (var start = 0; start < toFetch.length; start += PROWEBLOOK_PARALLEL_CHUNK) {
    var slice = toFetch.slice(start, start + PROWEBLOOK_PARALLEL_CHUNK);
    var reqs = slice.map(function(entry) {
      return {
        url: PROWEBLOOK_API_BASE + '/' + entry.operation,
        method: 'post',
        contentType: 'application/json',
        headers: {
          Accept: 'application/json',
          Authorization: 'Bearer ' + apiKey
        },
        payload: JSON.stringify(entry.payload || {}),
        muteHttpExceptions: true
      };
    });

    var rawResponses = reqs.length === 1
      ? [UrlFetchApp.fetch(reqs[0].url, reqs[0])]
      : UrlFetchApp.fetchAll(reqs);

    for (var k = 0; k < slice.length; k++) {
      var parsed = parseApiResponse_(rawResponses[k]);
      if (useCache && parsed.ok) {
        writeCachedResult_(slice[k].operation, slice[k].payload, parsed);
      }
      responses[slice[k].pos] = parsed;
    }
  }

  return responses;
}

function callProWebLookApi_(operation, payload, apiKey, useCache) {
  var list = executeBatchApiCalls_([{ operation: operation, payload: payload }], apiKey, useCache);
  return list[0];
}

function parseApiResponse_(httpResponse) {
  if (!httpResponse) {
    return { ok: false, code: 'network_error', message: 'No response from ProWebLook API' };
  }
  var status = Number(httpResponse.getResponseCode ? httpResponse.getResponseCode() : 0);
  var text = String(httpResponse.getContentText ? httpResponse.getContentText() : '');
  var json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch (e) {
    json = null;
  }

  if (status >= 200 && status < 300 && json && json.ok === true && json.data) {
    return { ok: true, data: json.data };
  }

  var errObj = json && json.error ? json.error : {};
  var code = String(errObj.code || ('http_' + status));
  var message = String(errObj.message || defaultErrorMessage_(status, code));
  return { ok: false, code: code, message: message };
}

function defaultErrorMessage_(status, code) {
  if (status === 401 || code === 'invalid_credentials') return 'Invalid or inactive ProWebLook API Key';
  if (status === 402 || code === 'insufficient_credit') return 'Insufficient ProWebLook credits';
  if (status === 429 || code === 'rate_limited') return 'Rate limit reached (60/min) — retry shortly';
  if (status === 400 || code === 'invalid_input') return 'Invalid input format';
  return 'ProWebLook service temporarily unavailable';
}

function getApiKey_() {
  return getDocumentApiKey_() || getUserApiKey_() || getScriptApiKey_();
}

function getDocumentApiKey_() {
  try {
    var docProps = PropertiesService.getDocumentProperties();
    return docProps ? (docProps.getProperty(PROWEBLOOK_PROP_KEY) || '') : '';
  } catch (e) {
    return '';
  }
}

function getUserApiKey_() {
  try {
    var userProps = PropertiesService.getUserProperties();
    return userProps ? (userProps.getProperty(PROWEBLOOK_PROP_KEY) || '') : '';
  } catch (e) {
    // Custom functions execute without UserProperties permission; fail closed cleanly.
    return '';
  }
}

function getScriptApiKey_() {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    return scriptProps ? (scriptProps.getProperty(PROWEBLOOK_PROP_KEY) || '') : '';
  } catch (e) {
    return '';
  }
}

function buildCacheKey_(operation, payload) {
  var raw = 'pwl:v1:' + operation + ':' + JSON.stringify(payload || {});
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = digest.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
  return 'pwl_' + hex.slice(0, 48);
}

function getCacheStore_() {
  try {
    return CacheService.getDocumentCache() || CacheService.getScriptCache();
  } catch (e) {
    try {
      return CacheService.getScriptCache();
    } catch (e2) {
      return null;
    }
  }
}

function readCachedResult_(operation, payload) {
  var cache = getCacheStore_();
  if (!cache) return null;
  try {
    var raw = cache.get(buildCacheKey_(operation, payload));
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    return parsed && parsed.ok === true ? parsed : null;
  } catch (e) {
    return null;
  }
}

function writeCachedResult_(operation, payload, result) {
  var cache = getCacheStore_();
  if (!cache || !result || result.ok !== true) return;
  try {
    cache.put(buildCacheKey_(operation, payload), JSON.stringify(result), PROWEBLOOK_CACHE_TTL_SECONDS);
  } catch (e) {}
}

function sanitizeCellString_(val) {
  if (val === null || val === undefined) return '';
  return String(val).replace(/[\x00-\x1F\x7F]/g, '').trim();
}

function normalizeCountryCode_(countryCode) {
  var cleaned = sanitizeCellString_(countryCode).toUpperCase();
  if (!cleaned) return '';
  return /^[A-Z]{2}$/.test(cleaned) ? cleaned : '';
}

function normalizePhoneInput_(rawVal, isoCountry, requirePlus) {
  if (typeof rawVal === 'number') {
    if (!isFinite(rawVal) || Math.floor(rawVal) !== rawVal || String(rawVal).indexOf('e') !== -1 || String(rawVal).indexOf('E') !== -1) {
      return { ok: false, reason: 'scientific_notation_format_as_text' };
    }
  }
  var str = sanitizeCellString_(rawVal);
  if (/^[0-9]+(\.[0-9]+)?[eE][+-]?[0-9]+$/.test(str)) {
    return { ok: false, reason: 'scientific_notation_format_as_text' };
  }

  var compact = str.replace(/[\s().-]/g, '');
  if (compact.indexOf('00') === 0 && compact.length > 7) {
    compact = '+' + compact.slice(2);
  }

  if (compact.indexOf('+') === 0) {
    if (!/^\+[1-9][0-9]{5,14}$/.test(compact)) {
      return { ok: false, reason: 'invalid_e164_number' };
    }
    return { ok: true, phone: compact };
  }

  // Google Sheets strips leading '+' when users type +14155552671 into an unformatted cell.
  if (requirePlus) {
    if (/^[1-9][0-9]{6,14}$/.test(compact)) {
      return { ok: true, phone: '+' + compact };
    }
    return { ok: false, reason: 'missing_country_code_or_plus' };
  }

  if (!/[0-9]{5,15}/.test(compact)) {
    return { ok: false, reason: 'invalid_phone_digits' };
  }
  return { ok: true, phone: str };
}

function maskApiKey_(key) {
  if (!key || key.length < 10) return '••••••••';
  return key.slice(0, 4) + '••••••••' + key.slice(-4);
}

function formatBalanceAmount_(bucket) {
  if (!bucket || typeof bucket.amount !== 'number') return '0';
  return bucket.amount.toLocaleString();
}

function emptyRow_(len) {
  var arr = [];
  for (var i = 0; i < len; i++) arr.push('');
  return arr;
}

function fillRow_(len, val) {
  var arr = [];
  for (var i = 0; i < len; i++) arr.push(val);
  return arr;
}

function verdictColor_(verdict) {
  var v = String(verdict || '').toLowerCase();
  if (v === 'deliverable' || v === 'valid' || v === 'on_whatsapp') return '#e6f4ea';
  if (v === 'undeliverable' || v === 'invalid' || v === 'not_on_whatsapp' || v.indexOf('invalid_') === 0 || v.indexOf('error:') === 0) return '#fce8e6';
  if (v === 'unknown') return '#fef7e0';
  return null;
}
