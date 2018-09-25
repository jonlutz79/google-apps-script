/* Geocoder */

// REFERENCE: http://willgeary.github.io/data/2016/11/04/Geocoding-with-Google-Sheets.html

// PARAMS

var ADDRESS_COL = 1;
var LAT_COL = 2;
var LNG_COL = 3;

var MAX_ROWS_TO_FILL = 500;

// TODO: Make this a property
var region = PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'us';
var geocoder = Maps.newGeocoder().setRegion(region);
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();      

// EVENT HANDLERS

function onOpen() {
  // Create menu items
  var menuItems = [{
    name: 'Fill coordinates',
    functionName: 'fillCoordinates'
  }];
  
  // Add menu
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Geocode', menuItems);
};

function onEdit(e) {
  fillCoordinates(e.range);
}

function onFillCoordinatesMenuItemClicked() {
  fillCoordinates(SpreadsheetApp.getActiveSheet().getActiveRange());
}

// FUNCTIONS

function clearCoordinates(row) {
  sheet.getRange(row, LAT_COL).setValue('');
  sheet.getRange(row, LNG_COL).setValue('');
}

function fillCoordinates(range) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
    
  var firstRow = range.getRow();
  var lastRow = range.getLastRow();
  var rowCount = lastRow - firstRow;
    
  if (sheet.getName() === 'Projects' && firstRow > 1) {
    // Only update coordinates if execution time tolerable
    if (rowCount > MAX_ROWS_TO_FILL) {
      showAlert('A maximum of ' + MAX_ROWS_TO_FILL + ' rows can be filled at a time.');
      return;
    }
    
    // Loop through each row in range
    for (var row = firstRow; row <= lastRow; row++) {      
      // Get address, remove quotes to prevent error
      var address = sheet.getRange(row, ADDRESS_COL).getValue().replace(/\'+/g,"");
      
      if (address == '') {
        clearCoordinates(row);
        continue;
      }
            
      // Geocode address
      Logger.log("Geocoding address '" + address + "'...");
      var location = geocoder.geocode(address);
    
      // Check result validity
      if (location.status != 'OK') {
        // Invalid response, clear coordinates
        clearCoordinates(row);
        return;
      } 
      
      // Valid response, update coordinates ...
      
      //showAlert(JSON.stringify(location));
      
      lat = location.results[0].geometry.location.lat;
      lng = location.results[0].geometry.location.lng;
      
      sheet.getRange(row, LAT_COL).setValue(lat);
      sheet.getRange(row, LNG_COL).setValue(lng);
    }
  }
}
