// ==== ГОЛОВНІ ФУНКЦІЇ ====

function updatePrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('інвестиції');
  var priceCol = getColByHeader(sheet, 'Поточна ціна');
  var lastCoinRow = findLastCoinRow(sheet);

  for (var row = 2; row <= lastCoinRow; row++) {
    var coin = sheet.getRange(row, 1).getValue().toString().trim().toUpperCase();
    if (!coin || coin.indexOf('USDT') === 0) continue;

    var price = getOkxPrice(coin + '-USDT');
    if (price === 'н/д') price = getBingxPrice(coin + '-USDT');
    if (price === 'н/д') price = getBybitPrice(coin + 'USDT');

    sheet.getRange(row, priceCol).setValue(price);
    Utilities.sleep(200);
  }
}

function updateFromExchanges() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('інвестиції');
  var monthNames = ['january','february','march','april','may','june','july','august','september','october','november','december'];
  var now = new Date();
  var currentMonth = monthNames[now.getMonth()];

  // 1. Колонка поточного місяця
  var monthCol = getColByHeader(sheet, currentMonth);
  if (!monthCol) {
    var planCol = getColByHeader(sheet, 'plan');
    sheet.insertColumnBefore(planCol);
    sheet.getRange(1, planCol).setValue(currentMonth);
    monthCol = planCol;
  }

  // 2. Зібрати угоди з обох бірж
  var allFills = getOkxFills().concat(getBingxFills());

  var monthlyTotals = {};
  var qtyTotals = {};
  var seenCoins = {};

  allFills.forEach(function(fill) {
    var coin = fill.coin;
    seenCoins[coin] = true;
    var qty = fill.qty;
    var usdValue = fill.usdValue;
    var fillTime = fill.time;

    if (fill.side === 'buy') {
      qtyTotals[coin] = (qtyTotals[coin] || 0) + qty;
      if (fillTime.getMonth() === now.getMonth() && fillTime.getFullYear() === now.getFullYear()) {
        monthlyTotals[coin] = (monthlyTotals[coin] || 0) + usdValue;
      }
    } else if (fill.side === 'sell') {
      qtyTotals[coin] = (qtyTotals[coin] || 0) - qty;
    }
  });

  // 3. Знайти рядок першого "USDT ..." (межа монет)
  var firstUsdtRow = findFirstUsdtRow(sheet);

  // 4. Додати нові токени, яких ще немає в таблиці
  var existingCoins = {};
  for (var r = 2; r < firstUsdtRow; r++) {
    var c = sheet.getRange(r, 1).getValue().toString().trim().toUpperCase();
    if (c) existingCoins[c] = true;
  }
  Object.keys(seenCoins).forEach(function(coin) {
    if (!existingCoins[coin]) {
      sheet.insertRowBefore(firstUsdtRow);
      sheet.getRange(firstUsdtRow, 1).setValue(coin);
      firstUsdtRow++; // зсув після вставки
    }
  });

  // 5. Записати суми та Qty
  var qtyCol = getColByHeader(sheet, 'Qty');
  var lastCoinRow = findLastCoinRow(sheet);
  for (var row = 2; row <= lastCoinRow; row++) {
    var coin = sheet.getRange(row, 1).getValue().toString().trim().toUpperCase();
    if (!coin) continue;
    if (monthlyTotals[coin]) {
      var existing = sheet.getRange(row, monthCol).getValue() || 0;
      sheet.getRange(row, monthCol).setValue(Math.round((existing + monthlyTotals[coin]) * 100) / 100);
    }
    if (qtyTotals[coin] !== undefined) {
      sheet.getRange(row, qtyCol).setValue(qtyTotals[coin]);
    }
  }

  updateUsdtBalances(sheet);
  Logger.log('Готово: ' + currentMonth + ', оброблено угод: ' + allFills.length);
}

// ==== БАЛАНСИ USDT ====

function updateUsdtBalances(sheet) {
  var okxRow = findRowByName(sheet, 'USDT OKX');
  var bybitRow = findRowByName(sheet, 'USDT Bybit');
  var bingxRow = findRowByName(sheet, 'USDT BingX');
  var totalCol = 2; // колонка B — Total $

  if (okxRow) sheet.getRange(okxRow, totalCol).setValue(getOkxUsdtBalance());
  if (bybitRow) sheet.getRange(bybitRow, totalCol).setValue(getBybitUsdtBalance());
  if (bingxRow) sheet.getRange(bingxRow, totalCol).setValue(getBingxUsdtBalance());
}

function getOkxUsdtBalance() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('OKX_API_KEY');
  var secretKey = props.getProperty('OKX_SECRET_KEY');
  var passphrase = props.getProperty('OKX_PASSPHRASE');
  var timestamp = new Date().toISOString();
  var requestPath = '/api/v5/account/balance?ccy=USDT';
  var prehash = timestamp + 'GET' + requestPath;
  var signature = Utilities.base64Encode(Utilities.computeHmacSha256Signature(prehash, secretKey));
  var options = {
    method: 'get',
    headers: {
      'OK-ACCESS-KEY': apiKey, 'OK-ACCESS-SIGN': signature,
      'OK-ACCESS-TIMESTAMP': timestamp, 'OK-ACCESS-PASSPHRASE': passphrase
    },
    muteHttpExceptions: true
  };
  try {
    var res = JSON.parse(UrlFetchApp.fetch('https://www.okx.com' + requestPath, options).getContentText());
    return parseFloat(res.data[0].details[0].eq) || 0;
  } catch (e) { return 0; }
}

function getBybitUsdtBalance() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BYBIT_API_KEY');
  var apiSecret = props.getProperty('BYBIT_API_SECRET');
  var timestamp = Date.now().toString();
  var recvWindow = '5000';
  var qs = 'accountType=UNIFIED';
  var prehash = timestamp + apiKey + recvWindow + qs;
  var signature = Utilities.computeHmacSha256Signature(prehash, apiSecret)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var options = {
    method: 'get',
    headers: {
      'X-BAPI-API-KEY': apiKey, 'X-BAPI-SIGN': signature,
      'X-BAPI-TIMESTAMP': timestamp, 'X-BAPI-RECV-WINDOW': recvWindow
    },
    muteHttpExceptions: true
  };
  try {
    var res = JSON.parse(UrlFetchApp.fetch('https://api.bybit.com/v5/account/wallet-balance?' + qs, options).getContentText());
    var coins = res.result.list[0].coin;
    for (var i=0;i<coins.length;i++) if (coins[i].coin==='USDT') return parseFloat(coins[i].walletBalance);
    return 0;
  } catch(e){ return 0; }
}

function getBingxUsdtBalance() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BINGX_API_KEY');
  var secretKey = props.getProperty('BINGX_SECRET_KEY');
  var timestamp = Date.now().toString();
  var params = 'timestamp=' + timestamp;
  var signature = Utilities.computeHmacSha256Signature(params, secretKey)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var url = 'https://open-api.bingx.com/openApi/spot/v1/account/balance?' + params + '&signature=' + signature;
  var options = { method: 'get', headers: { 'X-BX-APIKEY': apiKey }, muteHttpExceptions: true };
  try {
    var res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    var balances = res.data.balances;
    for (var i=0;i<balances.length;i++) if (balances[i].asset==='USDT') return parseFloat(balances[i].free);
    return 0;
  } catch(e){ return 0; }
}

// ==== ІСТОРІЯ УГОД ====

function getOkxFills() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('OKX_API_KEY');
  var secretKey = props.getProperty('OKX_SECRET_KEY');
  var passphrase = props.getProperty('OKX_PASSPHRASE');
  var timestamp = new Date().toISOString();
  var requestPath = '/api/v5/trade/fills-history?instType=SPOT';
  var prehash = timestamp + 'GET' + requestPath;
  var signature = Utilities.base64Encode(Utilities.computeHmacSha256Signature(prehash, secretKey));
  var options = {
    method: 'get',
    headers: {
      'OK-ACCESS-KEY': apiKey, 'OK-ACCESS-SIGN': signature,
      'OK-ACCESS-TIMESTAMP': timestamp, 'OK-ACCESS-PASSPHRASE': passphrase
    },
    muteHttpExceptions: true
  };
  try {
    var data = JSON.parse(UrlFetchApp.fetch('https://www.okx.com' + requestPath, options).getContentText());
    if (!data.data) return [];
    return data.data.map(function(f) {
      var coin = f.instId.split('-')[0];
      var qty = parseFloat(f.fillSz);
      var fee = Math.abs(parseFloat(f.fee));
      var netQty = (f.feeCcy === coin) ? qty - fee : qty;
      return {
        coin: coin, side: f.side, qty: netQty,
        usdValue: qty * parseFloat(f.fillPx),
        time: new Date(parseInt(f.fillTime))
      };
    });
  } catch(e) { return []; }
}

function getBingxFills() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BINGX_API_KEY');
  var secretKey = props.getProperty('BINGX_SECRET_KEY');
  var timestamp = Date.now().toString();
  var params = 'timestamp=' + timestamp;
  var signature = Utilities.computeHmacSha256Signature(params, secretKey)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var url = 'https://open-api.bingx.com/openApi/spot/v1/trade/myTrades?' + params + '&signature=' + signature;
  var options = { method: 'get', headers: { 'X-BX-APIKEY': apiKey }, muteHttpExceptions: true };
  try {
    var data = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    if (!data.data) return [];
    return data.data.map(function(f) {
      var coin = f.symbol.split('-')[0];
      var qty = parseFloat(f.qty);
      var fee = Math.abs(parseFloat(f.commission || 0));
      var netQty = (f.commissionAsset === coin) ? qty - fee : qty;
      return {
        coin: coin, side: f.isBuyer ? 'buy' : 'sell', qty: netQty,
        usdValue: qty * parseFloat(f.price),
        time: new Date(parseInt(f.time))
      };
    });
  } catch(e) { return []; }
}

// ==== ЦІНИ (публічні, без ключа) ====

function getOkxPrice(instId) {
  try {
    var d = JSON.parse(UrlFetchApp.fetch('https://www.okx.com/api/v5/market/ticker?instId=' + instId, {muteHttpExceptions:true}).getContentText());
    return (d.data && d.data.length) ? parseFloat(d.data[0].last) : 'н/д';
  } catch(e){ return 'н/д'; }
}
function getBingxPrice(instId) {
  try {
    var d = JSON.parse(UrlFetchApp.fetch('https://open-api.bingx.com/openApi/spot/v1/ticker/24hr?symbol=' + instId, {muteHttpExceptions:true}).getContentText());
    return (d.data && d.data.length) ? parseFloat(d.data[0].lastPrice) : 'н/д';
  } catch(e){ return 'н/д'; }
}
function getBybitPrice(symbol) {
  try {
    var d = JSON.parse(UrlFetchApp.fetch('https://api.bybit.com/v5/market/tickers?category=spot&symbol=' + symbol, {muteHttpExceptions:true}).getContentText());
    return (d.result && d.result.list && d.result.list.length) ? parseFloat(d.result.list[0].lastPrice) : 'н/д';
  } catch(e){ return 'н/д'; }
}

// ==== ДОПОМІЖНІ ====

function getColByHeader(sheet, name) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var idx = headers.indexOf(name);
  return idx === -1 ? null : idx + 1;
}
function findFirstUsdtRow(sheet) {
  var col = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  for (var i = 1; i < col.length; i++) {
    if (col[i][0].toString().indexOf('USDT') === 0) return i + 1;
  }
  return sheet.getLastRow();
}
function findLastCoinRow(sheet) {
  return findFirstUsdtRow(sheet) - 1;
}
function findRowByName(sheet, name) {
  var col = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  for (var i = 0; i < col.length; i++) {
    if (col[i][0] === name) return i + 1;
  }
  return null;
}

// ==== ТРИГЕР ====

function createDailyTrigger() {
  ScriptApp.newTrigger('updatePrices').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('updateFromExchanges').timeBased().everyDays(1).atHour(9).create();
}

function createDailyTrigger() {
  ScriptApp.newTrigger('updatePrices').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('updateFromExchanges').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('updateQtyFromBalances').timeBased().everyDays(1).atHour(9).create();
}

function debugBalances() {
  Logger.log('--- OKX ---');
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('OKX_API_KEY');
  var secretKey = props.getProperty('OKX_SECRET_KEY');
  var passphrase = props.getProperty('OKX_PASSPHRASE');
  var timestamp = new Date().toISOString();
  var requestPath = '/api/v5/account/balance?ccy=USDT';
  var prehash = timestamp + 'GET' + requestPath;
  var signature = Utilities.base64Encode(Utilities.computeHmacSha256Signature(prehash, secretKey));
  var options = {
    method: 'get',
    headers: { 'OK-ACCESS-KEY': apiKey, 'OK-ACCESS-SIGN': signature, 'OK-ACCESS-TIMESTAMP': timestamp, 'OK-ACCESS-PASSPHRASE': passphrase },
    muteHttpExceptions: true
  };
  Logger.log(UrlFetchApp.fetch('https://www.okx.com' + requestPath, options).getContentText());

  Logger.log('--- BINGX FILLS ---');
  var bxApiKey = props.getProperty('BINGX_API_KEY');
  var bxSecret = props.getProperty('BINGX_SECRET_KEY');
  var bxTimestamp = Date.now().toString();
  var params = 'timestamp=' + bxTimestamp;
  var bxSignature = Utilities.computeHmacSha256Signature(params, bxSecret)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var url = 'https://open-api.bingx.com/openApi/spot/v1/trade/myTrades?' + params + '&signature=' + bxSignature;
  var bxOptions = { method: 'get', headers: { 'X-BX-APIKEY': bxApiKey }, muteHttpExceptions: true };
  Logger.log(UrlFetchApp.fetch(url, bxOptions).getContentText());
}

function getBingxUsdtBalance() {
  var balances = getBingxBalances();
  return balances['USDT'] || 0;
}

function getBingxBalances() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BINGX_API_KEY');
  var secretKey = props.getProperty('BINGX_SECRET_KEY');
  var timestamp = Date.now().toString();
  var params = 'timestamp=' + timestamp;
  var signature = Utilities.computeHmacSha256Signature(params, secretKey)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var url = 'https://open-api.bingx.com/openApi/spot/v1/account/balance?' + params + '&signature=' + signature;
  var options = { method: 'get', headers: { 'X-BX-APIKEY': apiKey }, muteHttpExceptions: true };
  var result = {};
  try {
    var res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    Logger.log('BingX balance raw: ' + JSON.stringify(res));
    var list = (res.data && res.data.balances) ? res.data.balances : [];
    list.forEach(function(b) {
      var total = parseFloat(b.free) + parseFloat(b.locked);
      if (total > 0) result[b.asset] = total;
    });
  } catch(e) { Logger.log('BingX balance error: ' + e.message); }
  return result;
}

function getOkxBalances() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('OKX_API_KEY');
  var secretKey = props.getProperty('OKX_SECRET_KEY');
  var passphrase = props.getProperty('OKX_PASSPHRASE');
  var timestamp = new Date().toISOString();
  var requestPath = '/api/v5/account/balance';
  var prehash = timestamp + 'GET' + requestPath;
  var signature = Utilities.base64Encode(Utilities.computeHmacSha256Signature(prehash, secretKey));
  var options = {
    method: 'get',
    headers: { 'OK-ACCESS-KEY': apiKey, 'OK-ACCESS-SIGN': signature, 'OK-ACCESS-TIMESTAMP': timestamp, 'OK-ACCESS-PASSPHRASE': passphrase },
    muteHttpExceptions: true
  };
  var result = {};
  try {
    var res = JSON.parse(UrlFetchApp.fetch('https://www.okx.com' + requestPath, options).getContentText());
    var details = (res.data && res.data[0] && res.data[0].details) ? res.data[0].details : [];
    details.forEach(function(d) {
      var eq = parseFloat(d.eq);
      if (eq > 0) result[d.ccy] = eq;
    });
  } catch(e) { Logger.log('OKX balance error: ' + e.message); }
  return result;
}

function getBybitBalances() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BYBIT_API_KEY');
  var apiSecret = props.getProperty('BYBIT_API_SECRET');
  var timestamp = Date.now().toString();
  var recvWindow = '5000';
  var qs = 'accountType=UNIFIED';
  var prehash = timestamp + apiKey + recvWindow + qs;
  var signature = Utilities.computeHmacSha256Signature(prehash, apiSecret)
    .map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?'0'+v:v;}).join('');
  var options = {
    method: 'get',
    headers: { 'X-BAPI-API-KEY': apiKey, 'X-BAPI-SIGN': signature, 'X-BAPI-TIMESTAMP': timestamp, 'X-BAPI-RECV-WINDOW': recvWindow },
    muteHttpExceptions: true
  };
  var result = {};
  try {
    var rawResponse = UrlFetchApp.fetch('https://api.bybit.com/v5/account/wallet-balance?' + qs, options).getContentText();
Logger.log('Bybit raw: ' + rawResponse);
var res = JSON.parse(rawResponse);
    var coins = (res.result && res.result.list && res.result.list[0]) ? res.result.list[0].coin : [];
    coins.forEach(function(c) {
      var bal = parseFloat(c.walletBalance);
      if (bal > 0) result[c.coin] = bal;
    });
  } catch(e) { Logger.log('Bybit balance error: ' + e.message); }
  return result;
}

function updateQtyFromBalances() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('інвестиції');
  var qtyCol = getColByHeader(sheet, 'Qty');
  var lastCoinRow = findLastCoinRow(sheet);

  var okx = getOkxBalances();
  var bybit = {};
  var bingx = getBingxBalances();

  for (var row = 2; row <= lastCoinRow; row++) {
    var coin = sheet.getRange(row, 1).getValue().toString().trim().toUpperCase();
    if (!coin) continue;
    var total = (okx[coin] || 0) + (bybit[coin] || 0) + (bingx[coin] || 0);
    sheet.getRange(row, qtyCol).setValue(total);
  }

  updateUsdtBalances(sheet);
  Logger.log('Qty оновлено з реальних балансів бірж');
}
