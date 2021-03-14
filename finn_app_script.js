function fetchFinnData(finn_url) {
  /** fetchFinnData('https://www.finn.no/realestate/homes/ad.html?finnkode=116229744'); */
  if (!finn_url || finn_url.indexOf('?finnkode=') === -1) {
    return;
  }
  var API_URL = 'https://<YOUR_HEROKU_APP>.herokuapp.com/'
  var param = finn_url.match(/\?finnkode=.+/)[0]
  
  var url = API_URL + param;
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  var data = JSON.parse(response);
  
  return data;
}

function getColumnMapping(sheet) {
  var columnHeaders = sheet.getRange(1, 2, 1, 30);
  var values = columnHeaders.getValues()[0];
  
  var mapping = {};
  for (var col in values) {
    if( !values[col] ) {
      continue;
    }
    var key = values[col];
    mapping[key] = {key: key, index: parseInt(col, 10) + 2};
  }

  return mapping;
}

function onFinnCellEdit(e) {
  /* If you paste an interesting finn.no URL in the finn.no column, then update matching values and autofill some cells */  
  if (!e.value || e.value.indexOf('https://www.finn.no/realestate/homes/ad.html?finnkode=') === -1) {
    Logger.log('skipping not a URL: ' + e.value);
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var headerValue = sheet.getRange(1, e.range.getColumn()).getValue();
  if (headerValue !== 'finn.no') {
    Logger.log('skipping wrong header: ' + headerValue);
    return;
  }
  Logger.log(["found URL @", e.range.getRow(), e.range.getColumn()]);
  
  /* Fetch and prepare data */
  var colMap = getColumnMapping(sheet);
  var finnData = fetchFinnData(e.value);
  
  /* Update cells with value from finn API */
  for (var key in finnData.ad) {
    var val = finnData.ad[key];
    
    if (key in colMap) {
      var matchingRange = sheet.getRange(e.range.getRow(), colMap[key].index);
      matchingRange.setValue(val);
    }
  }
  
  /* Autofill certain columns */
  const autoFillRanges = ['autofill_1', 'autofill_2', 'autofill_3', 'autofill_4'];
  const rowAbove = e.range.getRow() - 1;
  for (var i in autoFillRanges) {
    Logger.log('Autofilling named range ' + autoFillRanges[i]); 
    var autoFillCols = sheet.getRange(autoFillRanges[i]);
    var srcRange = sheet.getRange(rowAbove, autoFillCols.getColumn(), 1, autoFillCols.getNumColumns())
    var dstRange = sheet.getRange(rowAbove, autoFillCols.getColumn(), 2, autoFillCols.getNumColumns())
    srcRange.autoFill(dstRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  e.range.setNote('Sist oppdatert: ' + new Date());
}