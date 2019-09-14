/*
 * script to export data of the named sheet as an individual csv file
 * sheet downloaded to Google Drive and then downloaded as a CSV file
 * file named according to the name of the sheet
 * added | as delimiter
 * usage of the same folder
 * original author: Michael Derazon (https://gist.github.com/mderazon/9655893)
*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Download", functionName: "saveAsCSV"}];
  ss.addMenu("Download as CSV", csvMenuEntries);
};

function saveAsCSV() {
  // gets Active Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // gets Active Sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  folderName = ss.getName().toLowerCase().replace(/ /g,'_') + '_csv'

  // create a folder from the name of the spreadsheet
  var folder, folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    folder = folders.next();
  }
  else {
    folder =  DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv');
  }


  // append timestamp and csv" extension to the sheet name
  fileName = sheet.getName() + "_" + new Date().getTime() + ".csv";
  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(fileName, sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  //File downlaod
  var downloadURL = file.getDownloadUrl().slice(0, -8);

}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = data[row][col];
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join("|") + "\n";
        }
        else {
          csv += data[row].join("|");
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
