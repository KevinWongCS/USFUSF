/*  
name: Kevin Wong
date: 8/7/2022
file: getITSnum
desc: Uses Optical Character Recognition(OCR) to get the ITS-[number] from an Asset Tag and creates a link to ServiceNow to asset page.
  Requires image link in a column from a image hosting site like imgur.
*/


function readTextFromImage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"); // get sheet by name
  var lastRow = sheet.getLastRow(); //get last row

  for(var i = 2; i <= lastRow; i++){ 
    var url = sheet.getRange(i, 1).getValue();  //gets Values("URLS") from i-th Row, column 1
    var imageBlob = UrlFetchApp.fetch(url).getBlob();
    var resource = {
      title : imageBlob.getName(),
      mimeType : imageBlob.getContentType()
    };
    var options = {
      ocr : true
    };
    var docFile = Drive.Files.insert(resource, imageBlob, options);
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();
    var textITS = "";

      for(var j = 65; j < text.length; j++){    //Gets String with just "ITS-[number]"
        textITS += text[j];
      }
  
    sheet.getRange(i, 3).setValue(textITS); //Inserts ITS-[number] into i-th Row, column 3
    Drive.Files.remove(docFile.id);

  }
}

