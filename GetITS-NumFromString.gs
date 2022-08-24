/*  
name: Kevin Wong
date: 8/19/2022
file: DateStandardizationFormatTest.gs
desc: Gettings ITS-[number]'s from Strings.
*/

function DateStandardizationFormatTest() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var newSheet = SpreadsheetApp.create(sheet);
  const TicketITSArray = [];

  //all the data in column A that we want to modify
  var dataRange = sheet.getRange("A2:A1343").getValues(); //2:1343
  //all the data in column B that we want to modify
  var dataRange = sheet.getRange("B2:B1343").getValues();
  //console.log(dataRange2.length)
  //all the dates in column D that we want
  var dataRange = sheet.getRange("C2:C1343").getValues();

  //mainloop
  for (let i = 2; i <= dataRange.length; i++){ //edit and change where the data first begins
    
    //counter
    // console.log("Counter: " + i);

    //get Ticket Number
    var ticketNum = sheet.getRange(i, 1).getValue().toString();
    // console.log("cellData: " + cellData);

    //get Date
    var ticketDate = sheet.getRange(i, 4).getValue().toString().substring(4, 16);
    // console.log("cellData: " + ticketDate);

    //get Short Description
    var cellData = sheet.getRange(i, 2).getValue().toString();
    // console.log("cellData: " + cellData);

    //check if string contains "ITS-"
    if(cellData.includes("ITS-")) {
      ITSArray = [];

      //turn cellData into an array of words via "split()"
      var cellDataArray = cellData.split(" ");

      //create ITS-[number] array
      for(var j = 0; j < cellDataArray.length; j++){
        if(cellDataArray[j].includes("ITS-") == true){
          ITSArray.push(cellDataArray[j]);
        }
      }

      //create dictionary;
      const item = {ticketNum: ticketNum, ticketDate: ticketDate, ITSnums: ITSArray};  //"java objects for documentation: https://www.w3schools.com/js/js_objects.aspf"
      console.log("Counter: " + i + " : " + item.ticketNum + " - " + item.ticketDate + " : " + item.ITSnums);

      //Pipe ITS-[number]'s into new sheet
      newSheetTab = newSheet.getSheetByName("Sheet1");
      // newSheetTab.append(["Ticket Number", "Date", "ITS-[number]", "RITM link", "ITS link" ]);
      

      for(var k = 0; k < item.ITSnums.length; k++){
        /////////// TURN THESE INTO LINKS AND THEN RUN A WEB SCRAPPER/CRAWLER ///////////// 8/19: failed webscrapper doesn't work because of a redirect
        var lastRow = newSheetTab.getLastRow();
        newSheetTab.getRange(lastRow + 1, 1).setValue(item.ticketNum);
        newSheetTab.getRange(lastRow + 1, 2).setValue(item.ticketDate);
        newSheetTab.getRange(lastRow + 1, 3).setValue(item.ITSnums[k].replace(",", ""));
        newSheetTab.getRange(lastRow + 1, 4).setValue("= hyperlink( \"https://usf.service-now.com/nav_to.do?uri=%2F$sn_global_search_results.do%3Fsysparm_search%3D\" & ".concat("A", lastRow + 1, ",", "A", lastRow + 1, " )"));
        newSheetTab.getRange(lastRow + 1, 5).setValue("= hyperlink( \"https://usf.service-now.com/nav_to.do?uri=%2F$sn_global_search_results.do%3Fsysparm_search%3D\" & ".concat("C", lastRow + 1, ",", "C", lastRow + 1, " )"));
    
      }
    
    } else {
      console.log("Counter: " + i + " : " + ticketNum + " - " + ticketDate + " : " + "NULL");
    }//end of if-else statement
  } //end of mainloop
  
  //////////////////////////////////////////////////////////////////////////////////////////////
  //Test Bench
}
