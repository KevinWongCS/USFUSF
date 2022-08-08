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
  var folder = DriveApp.getFolderById("1M62Bblkpu79JAsjf-J_1Im9AxkBIr0Az");
  var PDFs = folder.getFiles();
  var counter = 2; //counter for while loop

  //Google Sheet setup
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  sheet.appendRow(["File Name", "Date", "Size(bytes)", "URL", "ITS-[number]", "ServiceNow Link", "Asset Status"]);
  
  //main loop
  while (PDFs.hasNext()) {
    var pdf = PDFs.next();
    var data = [
        pdf.getName(),
        pdf.getDateCreated(),
        pdf.getSize(),
        pdf.getUrl()       
    ]
    sheet.appendRow(data);

  // Use OCR to convert PDF to a temporary Google Document
  var docName = pdf.getName().replace(/\.pdf$/, '');
  var file = {
    title: docName,
    mimeType: pdf.getMimeType() || 'application/pdf'
  }
  var image = pdf.getBlob()

  Drive.Files.insert(file, image, { ocr: true }); //have to do it this way, can't create a file directly into the folder...
    var newFile = DriveApp.getFilesByName(docName).next();
    var doc = DocumentApp.openById(newFile.getId());
    var body = doc.getBody().getText().slice(doc.getBody().getText()); //.lastIndexOf("\n"));

    //counter appends the new data into the next row
    sheet.getRange(counter, 5).setValue(body)
    counter++;  //increment the counter

    //delete doc from drive
    Drive.Files.remove(newFile.getId());
  }
}
