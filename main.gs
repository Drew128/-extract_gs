function formula() {
  var formula_list = sheetToJson("formula", emptyKey = false, emptyValue = false)
  for (var i = 0; i < formula_list.length; i++) {
    if (formula_list[i]['active'] != true) {continue;}
    var pageName = formula_list[i]['page']
    var columnName = formula_list[i]['column']
    var sheet = findPage(pageName)
    var sheetLastRow = sheet.getLastRow()
    var range = sheet.getRange(columnName+"1:"+columnName).getValues()
    for (var r = range.length-1; r >= 0 ; r--) {if (range[r] != '') {break}}
    var rangeLastRow = r+1
    if (sheetLastRow != rangeLastRow) {
    sheet.getRange(columnName+rangeLastRow).autoFill(sheet.getRange(columnName+rangeLastRow+':'+columnName+sheetLastRow), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
    log(['autofill on page', pageName, 'range', columnName+rangeLastRow+':'+columnName+sheetLastRow])
    }
  }
}


function processor() {
  
  var request_list = sheetToJson("requests", emptyKey = false, emptyValue = false)
  for (var i = 0; i < request_list.length; i++) {
    var request = {};
    if (request_list[i]['active'] != true) {continue;}
    request['url'] = request_list[i]['url'];
    request['method'] = request_list[i]['method'];
    if (typeof request_list[i]['headers'] != 'undefined') {request['headers'] = JSON.parse(request_list[i]['headers'].split("'").join('"'))};   
    log(["request", request])
    var response = UrlFetchApp.fetchAll([request]);  
    var jsonResponse = JSON.parse(response[0].getContentText())                          //     var jsonResponse = JSON.parse(response);
    
  if (request_list[i]['response_format'] == 'json_of_json') {
      jsonResponse_new = []
      for (j in jsonResponse) {
        row = jsonResponse[j]
        row["_key_"] = j
        jsonResponse_new.push(row)
      }
    var jsonResponse = jsonResponse_new;
    }
    var page_output = request_list[i]['page_output'];
    log(['response legth', jsonResponse.length])
    safeInsert(page_output, jsonResponse, insertTime=true)  
  }
};


function findPage(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {log("(sheet.getLastRow() == 0 | sheet == null)"); return {}}
  return sheet
};


function headerAndlastRow(sheet) {
  var lastRow = sheet.getLastRow();  
  var lastColumn = sheet.getLastColumn();
  if (lastRow == 0) {return {header:{}, newRow:lastRow+2}}
  var header = sheet.getRange(1, 1, 1, lastColumn);    
  var valueKeyColumns = rowToKV(header);
  header = addField(valueKeyColumns)
  return {header:valueKeyColumns, newRow:lastRow+1}
};


function addField(header, field=null) {
  if (field != null) {header[field] = 0}
  for (column in header) {
    header[column] = parseInt(header[column])+1
  }
  return header
};


function safeInsert(page, arrayOfObj, insertTime=true) {
  var sheet = findPage(page);
  var headerAndRow = headerAndlastRow(sheet);
  var header = headerAndRow['header'];
  var newRow = headerAndRow['newRow'];
  if(insertTime == true && typeof header['insertTime'] == 'undefined') {
  var header = addField(header, 'insertTime')              // move right to compansate indexing from zero
  sheet.getRange(1, 1).setValue('insertTime')
  };
  
  maxColumn = function (h=header) {return Math.max.apply(null, Object.values(header))}
  
  arrayOfObj.reverse()
  for (var i = 0; i < arrayOfObj.length; i++) {
    obj = arrayOfObj[i]
    if (insertTime == true) {obj['insertTime'] = dateFormat()}
    var columnNames = Object.keys(obj)
    for (n in columnNames) {
      name = columnNames[n]
      var columnN = header[name];
      if (typeof header[name] == 'undefined') {   // if there is no colmn with KEY name
        columnN = maxColumn()+1
        sheet.getRange(1, columnN).setValue(name)
        header[name] = columnN        
      }
      sheet.getRange(newRow, columnN).setValue(obj[name])  
    }
    newRow += 1;
  }
};



function rowToKV(row) {
  var rangeArray = row.getValues();
  var keyValueColumns = Object.assign({}, rangeArray[0]);
  var valueKeyColumns = swap(keyValueColumns);
return valueKeyColumns
};



function log(string) {
  var sheet = findPage('log');
  var lastRow = sheet.getLastRow();
  var currentDate = new Date();
  var currentDate_format = dateFormat(currentDate);
  sheet.getRange(lastRow+1, 1).setValue(currentDate_format);
  sheet.getRange(lastRow+1, 2).setValue(JSON.stringify(string));
};


function swap(json){
  var ret = {};
  for(var key in json){
    ret[json[key]] = key;
  }
  return ret;
};

function dateFormat(d=new Date()) {
  var format = (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear() +' '+d.getHours()+':'+d.getMinutes()+':'+d.getSeconds()
  return format
};


function valuesToJson(values, emptyKey = false, emptyValue = true, columnNamesRowN = 1){
  var valuesArray = [];
  var keys = values[columnNamesRowN-1];
  for (var row = columnNamesRowN; row < values.length; row++) {
    rowObj = {}
    for (var elemPos = 0; elemPos < values[row].length; elemPos++) {
      if ((keys[elemPos] != '' || emptyKey) && (values[row][elemPos] != '' || emptyValue))
      {rowObj[keys[elemPos]] = values[row][elemPos]}
    };
    valuesArray.push(rowObj)
  }
  return valuesArray
};


function sheetToJson(sheetName, emptyKey = false, emptyValue = true){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null || sheet.getLastRow() == 0) {log("(sheet.getLastRow() == 0 | sheet == null)"); return {}}
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var json = valuesToJson(values, emptyKey, emptyValue);
  return json
};


function date_miliseconds() {return Date.now()}


function from_miliseconds(ms) {
  d = new Date(ms)
  var format = dateFormat(d)
  return format
};

function parse_date(d) {
  var arr = d.split(".")
  var date = arr[1]+"/"+arr[0]+"/"+arr[2]
  return date};
