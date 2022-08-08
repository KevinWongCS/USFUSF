/*  
name: Kevin Wong
date: 8/8/2022
file: getITSnum
desc: Uses Optical Character Recognition(OCR) to get the ITS-[number] from an Asset Tag and creates a link to ServiceNow to asset page.
  credit: https://gist.github.com/kltng/c25422538e15e155bccef0e289ea3faa
  original fork: https://gist.github.com/rob0tca/b7fd4488d84a49e5ca87536048629406 
*/

function listFilesInFolder() {

  //note: Change the folder ID below to reflect your folder you are working in.
  var folder = DriveApp.getFolderById("1YIC3wTmyhtaE0Acx0xWOwVauMrslh2xB");
  var images = folder.getFiles();
  var counter = 2; //counter for while loop

  //Google Sheet setup
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  sheet.appendRow(["File Name", "Date", "Size(bytes)", "URL", "ITS-[number]", "ServiceNow Link", "Asset Status"]);
  
  //main loop
  while (images.hasNext()) {
    var image = images.next();
    var data = [
        image.getName(),
        image.getDateCreated(),
        image.getSize(),
        image.getUrl()       
    ]
    sheet.appendRow(data);

    //Creates a Google doc with image file and pulls the OCR text into Google Sheets
    var imageName = image.getName();
    var docName = imageName.split("\.")[0];
    var file = {
        title: docName,
        mimeType: "image/jpeg"
    }
    

    //Store new Google Doc in project folder
    Drive.Files.insert(file, image, { ocr: true }); //have to do it this way, can't create a file directly into the folder...
    var newFile = DriveApp.getFilesByName(docName).next();
    folder.addFile(newFile);
    var doc = DocumentApp.openById(newFile.getId());
    var body = doc.getBody().getText().slice(doc.getBody().getText().lastIndexOf("\n"));

    //counter appends the new data into the next row
    sheet.getRange(counter, 5).setValue(body)
    counter++;  //increment the counter

    //delete doc from folder and then drive
    folder.removeFile(newFile);
    Drive.Files.remove(newFile.getId());
  }
}
