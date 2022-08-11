/*  
name: Kevin Wong
date: 8/10/2022
file: DateStandardizationFormatTest.gs
desc: Standardizing the format of dates in a ticket summary Google sheet.
*/

function DateStandardizationFormatTest() {
  var sheet = SpreadsheetApp.getActiveSheet();


  // var dataRange = sheet.getDataRange().getValues();
  // console.log("something:", dataRange);

  //all the data in column A that we want to modify
  var dataRange2 = sheet.getRange("A1:A49").getValues();  //edit the variable in "getRange" to loop through entire column
  //console.log(dataRange2.length)

  //mainloop
  for (let i = 1; i <= dataRange2.length; i++){ //edit and change where the data first begins
    //get the data
    var cellData = sheet.getRange(i, 1).getValue().toString();
    //console.log("cellData: " + cellData);

    //get date 
    var dateDelimiter = cellData.indexOf("-");
    var date = cellData.slice(0, dateDelimiter).trim();
    var dateNoYear = date.slice(0, date.length - 2);
    var dateUpdated = dateNoYear.replaceAll(".", "/");
    // console.log("Date: " + "#" + i + " " + dateUpdated);

    //get message
    var message = cellData.slice(dateDelimiter, cellData.length).trim();
    
    //reappend the updated date to the message
    var newMessage = dateUpdated + "2022" + " " + message;
    console.log("New Message" + "#" +  i + " " + newMessage);

    //Input corrected dates
    sheet.getRange(i, 1).setValue(newMessage);
    
  } //end of mainloop
  
  //////////////////////////////////////////////////
  //Test Bench
  // console.log("something:");
  // console.log(dataRange2);
  // console.log(dataRange2.length)
  // console.log(dataRange2[dataRange2.length - 1]);

}
