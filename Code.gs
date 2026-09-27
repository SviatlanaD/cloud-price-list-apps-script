/**
 * Cloud price list builder.
 *
 * Pulls Azure VM prices and ECB exchange rates into the spreadsheet, builds a
 * multi-currency price list and exports it as an .xlsx file into a chosen
 * Drive folder.
 *
 * Both APIs are public and need no key.
 */

var SHEET_CONFIG = 'config';
var SHEET_RAW = 'prices_raw';
var SHEET_FX = 'fx_rates';
var SHEET_RESULT = 'price_list';

var AZURE_API = 'https://prices.azure.com/api/retail/prices';
var FX_API = 'https://api.frankfurter.dev/v1/latest';

var HOURS_PER_MONTH = 730; // Azure bills hourly; 730 is the usual month for pricing


/** Menu in the spreadsheet, so the whole thing runs without opening the editor. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Price list')
    .addItem('1. Import prices', 'importPrices')
    .addItem('2. Import FX rates', 'importRates')
    .addItem('3. Build price list', 'buildPriceList')
    .addSeparator()
    .addItem('4. Export .xlsx to Drive', 'exportToDrive')
    .addSeparator()
    .addItem('Run all steps', 'runAll')
    .addToUi();
}


function runAll() {
  importPrices();
  importRates();
  buildPriceList();
  exportToDrive();
}


/**
 * Reads the settings from the config sheet.
 * Keys are in column A, values in column B.
 */
function getConfig() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    throw new Error('Sheet "' + SHEET_CONFIG + '" not found. See README for the expected layout.');
  }

  var rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  var config = {};
  rows.forEach(function (row) {
    var key = String(row[0]).trim();
    if (key) config[key] = String(row[1]).trim();
  });

  ['service', 'region', 'currencies', 'drive_folder_id'].forEach(function (key) {
    if (!config[key]) throw new Error('config: "' + key + '" is empty');
  });

  config.currencies = config.currencies.split(',').map(function (c) {
    return c.trim().toUpperCase();
  }).filter(String);

  config.max_rows = Number(config.max_rows) || 200;
  return config;
}


/** Writes a 2D array to a sheet, replacing whatever was there. */
function writeSheet(name, rows) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);

  sheet.clear();
  if (rows.length) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}


/**
 * Azure retail prices for one service in one region.
 * The API pages 1000 items at a time and hands back the next URL.
 */
function importPrices() {
  var config = getConfig();

  var filter = "serviceName eq '" + config.service + "'" +
               " and armRegionName eq '" + config.region + "'" +
               " and priceType eq 'Consumption'";
  var url = AZURE_API + '?$filter=' + encodeURIComponent(filter);

  var items = [];
  while (url && items.length < config.max_rows) {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error('Azure API returned ' + response.getResponseCode());
    }
    var page = JSON.parse(response.getContentText());
    items = items.concat(page.Items || []);
    url = page.NextPageLink;
  }

  // Spot and low-priority SKUs are priced differently and would muddle the
  // comparison, so only pay-as-you-go meters go into the price list.
  items = items.filter(function (item) {
    return item.unitPrice > 0 &&
           item.meterName.indexOf('Spot') === -1 &&
           item.meterName.indexOf('Low Priority') === -1;
  }).slice(0, config.max_rows);

  var rows = [['sku', 'product', 'meter', 'region', 'unit', 'unit_price_usd']];
  items.forEach(function (item) {
    rows.push([
      item.armSkuName,
      item.productName,
      item.meterName,
      item.armRegionName,
      item.unitOfMeasure,
      item.unitPrice
    ]);
  });

  writeSheet(SHEET_RAW, rows);
  SpreadsheetApp.getActive().toast(items.length + ' prices imported', 'Done', 5);
}


/**
 * ECB reference rates. The API quotes against EUR, and Azure prices are in USD,
 * so every rate is divided by the USD rate to rebase the table on USD.
 */
function importRates() {
  var config = getConfig();

  var symbols = config.currencies.slice();
  if (symbols.indexOf('USD') === -1) symbols.push('USD');

  var url = FX_API + '?base=EUR&symbols=' + symbols.join(',');
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('FX API returned ' + response.getResponseCode());
  }

  var payload = JSON.parse(response.getContentText());

  // The API leaves the base currency out of the rates object, so EUR has to be
  // added back by hand before anything can be rebased on USD.
  var eurRates = payload.rates;
  eurRates.EUR = 1;

  var eurToUsd = eurRates.USD;
  if (!eurToUsd) throw new Error('FX API did not return a USD rate');

  var rows = [['currency', 'rate_per_usd', 'ecb_date']];
  config.currencies.forEach(function (currency) {
    if (!eurRates[currency]) throw new Error('No ECB rate for ' + currency);
    rows.push([currency, eurRates[currency] / eurToUsd, payload.date]);
  });

  writeSheet(SHEET_FX, rows);
  SpreadsheetApp.getActive().toast('Rates as of ' + payload.date, 'Done', 5);
}


/** Joins prices with rates and adds a monthly cost column per currency. */
function buildPriceList() {
  var config = getConfig();
  var spreadsheet = SpreadsheetApp.getActive();

  var prices = spreadsheet.getSheetByName(SHEET_RAW);
  var fx = spreadsheet.getSheetByName(SHEET_FX);
  if (!prices || !fx) throw new Error('Run the import steps first');

  var priceRows = prices.getDataRange().getValues().slice(1);
  var fxRows = fx.getDataRange().getValues().slice(1);

  var rates = {};
  fxRows.forEach(function (row) { rates[row[0]] = row[1]; });

  var header = ['sku', 'product', 'region', 'unit'];
  config.currencies.forEach(function (currency) {
    header.push('hourly_' + currency.toLowerCase());
    header.push('monthly_' + currency.toLowerCase());
  });

  var rows = [header];
  priceRows.forEach(function (row) {
    var usd = row[5];
    var line = [row[0], row[1], row[3], row[4]];
    config.currencies.forEach(function (currency) {
      var hourly = usd * rates[currency];
      line.push(round(hourly, 4));
      line.push(round(hourly * HOURS_PER_MONTH, 2));
    });
    rows.push(line);
  });

  var sheet = writeSheet(SHEET_RESULT, rows);
  sheet.autoResizeColumns(1, header.length);
  spreadsheet.toast((rows.length - 1) + ' rows in the price list', 'Done', 5);
}


function round(value, digits) {
  var factor = Math.pow(10, digits);
  return Math.round(value * factor) / factor;
}


/**
 * Exports the price list as .xlsx into the Drive folder from the config sheet.
 *
 * Apps Script cannot write .xlsx directly, and the spreadsheet export endpoint
 * ignores the gid for xlsx and returns the whole workbook. So the price list is
 * copied into a throwaway spreadsheet, that one is exported, and the temporary
 * file is deleted afterwards.
 */
function exportToDrive() {
  var config = getConfig();
  var source = SpreadsheetApp.getActive();

  var sheet = source.getSheetByName(SHEET_RESULT);
  if (!sheet) throw new Error('Build the price list first');

  var folder = DriveApp.getFolderById(config.drive_folder_id);

  var temp = SpreadsheetApp.create('tmp_price_list_export');
  var placeholder = temp.getSheets()[0];
  sheet.copyTo(temp).setName(SHEET_RESULT);
  temp.deleteSheet(placeholder);
  SpreadsheetApp.flush();

  try {
    var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
    var blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob().setName(buildFileName(config));

    var file = folder.createFile(blob);
    source.toast('Saved as ' + file.getName(), 'Exported', 6);
    return file.getUrl();
  } finally {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  }
}


function buildFileName(config) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  return 'price_list_' + config.region + '_' + stamp + '.xlsx';
}
