// Copyright (c) 2024 Soren Vinson (sorenv654@gmail.com)
// This code is available under the MIT License: https://opensource.org/license/mit

  function ClientBooking() {
  
    //1  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1371:H1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('AW1371:AY1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
    
    
  
  
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
  
  
  
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
  
  
              var AllNotes = "In Notes: " + checkInNotes;
  
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
  
          }
      }
  
  //1   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
  
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
  
  
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
  
              var AllNotes = " Out Notes: " + checkOutNotes;
  
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
  
              var lastDateString = lastDate1.toString();
  
  
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
          
          }
      }
  
  //1  //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
  
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
  
  
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
  
  
              var AllNotes = "Turn Notes: " + turnNotes;
  
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
  
          }
      }
  
  
  
  //1  //--START-- Mark houses whose entries have already been created  --START--//   
  
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('B1371:B1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
  
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "AW" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('D1371:D1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "AX" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('F1371:F1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "AY" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      //2  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('I1371:P1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('BB1371:BD1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
    
    
  
  
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
  
  
  
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
  
  
              var AllNotes = "In Notes: " + checkInNotes;
  
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
  
  
          }
      }
  
  //2   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
  
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
  
  
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
  
              var AllNotes = " Out Notes: " + checkOutNotes;
  
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
  
              var lastDateString = lastDate1.toString();
  
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
          }
      }
  
  //2  //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
  
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
  
  
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
  
  
              var AllNotes = "Turn Notes: " + turnNotes;
  
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
  
          }
      }
  
  
  
  //2  //--START-- Mark houses whose entries have already been created  --START--//   
  
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('J1371:J1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
  
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BB" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('L1371:L1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BC" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('N1371:N1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BD" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      //3  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Q1371:X1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('BG1371:BI1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
    
    
     
  
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
  
  
  
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
  
  
              var AllNotes = "In Notes: " + checkInNotes;
  
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
  
          }
      }
  
  //3   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
  
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
  
  
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
  
              var AllNotes = " Out Notes: " + checkOutNotes;
  
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
  
              var lastDateString = lastDate1.toString();
  
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
          }
      }
  
  //3  //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
  
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
  
  
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
  
  
              var AllNotes = "Turn Notes: " + turnNotes;
  
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
  
          }
      }
  
  
  
  //3  //--START-- Mark houses whose entries have already been created  --START--//   
  
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('R1371:R1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
  
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BG" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('T1371:T1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BH" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('V1371:V1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BI" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  //4  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--// 
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Y1371:AF1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('BL1371:BN1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
    
    
    
  
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
  
  
  
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
  
  
              var AllNotes = "In Notes: " + checkInNotes;
  
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
  
          }
      }
  
  //4   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
  
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
  
  
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
  
              var AllNotes = " Out Notes: " + checkOutNotes;
  
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
  
              var lastDateString = lastDate1.toString();
  
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
  
              }
          }
      
  
  //4  //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
  
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
  
  
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
  
  
              var AllNotes = "Turn Notes: " + turnNotes;
  
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
          }
      }
  
  
  
  //4  //--START-- Mark houses whose entries have already been created  --START--//   
  
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Z1371:Z1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
  
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BL" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('AB1371:AB1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BM" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('AD1371:AD1828').getValues();
  
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "BN" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
  

  var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BM1371:BT1828').getValues();
  var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('CD1371:CF1828').getValues();
  var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
  
  
  //6  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  
  
  for (i = 0; i < values.length; i++) {
      if (values[i][1] != "" && matrix[i][0] == "") {
  
  
  
          var checkInDate = (values[i][0]);
          var checkInNotes = (values[i][7]);
  
  
          var AllNotes = "In Notes: " + checkInNotes;
  
          if (values[i][7] != "") {
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN - Notes", checkInDate, { description: AllNotes });
          } else { 
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN ", checkInDate );
          }
  
      }
  }
  
  //6   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
  
  for (i = 0; i < values.length; i++) {
      if (values[i][3] != "" && matrix[i][1] == "") {
          for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
  
  
          var checkOutDate = (values[i][0]);
          var nextCheckInDate = (values[i + j][0]);
          var checkOutNotes = (values[i][7]);
  
          var AllNotes = " Out Notes: " + checkOutNotes;
  
          var string = nextCheckInDate.toString();
          var split1 = string.split(' ', 3);
          var conc = (split1[1]) + " " + (split1[2]);
  
          var lastDateString = lastDate1.toString();
  
          if (string == lastDateString && values[i][7] != "") {
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT - Notes ", checkOutDate, { description: AllNotes });
          } 
          if (string == lastDateString && values[i][7] == "") {
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT 2927 E Plaimor", checkOutDate);
          }
          if (string != lastDateString && values[i][7] != "") { 
             CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
          }
          if (string != lastDateString && values[i][7] == "") { 
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT " + " - " + conc, checkOutDate);
           }
      }
  }
  
  //6  //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
  
  for (i = 0; i < values.length; i++) {
      if (values[i][5] != "" && matrix[i][2] == "") {
  
  
          var turnDate = (values[i][0]);
          var turnNotes = (values[i][7]);
  
  
          var AllNotes = "Turn Notes: " + turnNotes;
  
          if (values[i][7] != ""){
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN - Notes ", turnDate, { description: AllNotes });
          } else { 
              CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN ", turnDate );
          }
  
      }
    }
  
  
  
  //5 //--START-- Mark houses whose entries have already been created  --START--//   
  
  var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
  var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
  var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BN1371:BN1828').getValues();
  
  for (i = 0; i < markedValues.length; i++) {
  
      if (markedValues[i] != "") {
          var cell = i + 1371;
          var cell2 = "CD" + cell;
          ssM.getRange(cell2).setValue("x");
      }
  }
  
  var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BP1371:BP1828').getValues();
  
  for (i = 0; i < markedValues.length; i++) {
      if (markedValues[i] != "") {
          var cell = i + 1371;
          var cell2 = "CE" + cell;
          ssM.getRange(cell2).setValue("x");
      }
  }
  
  var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BR1371:BR1828').getValues();
  
  for (i = 0; i < markedValues.length; i++) {
      if (markedValues[i] != "") {
          var cell = i + 1371;
          var cell2 = "CF" + cell;
          ssM.getRange(cell2).setValue("x");
      }


    }


    //6  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  

    
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BU1371:CB1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('CG1371:CI1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
            
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
      
      
      
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
      
      
              var AllNotes = "In Notes: " + checkInNotes;
      
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House NAme - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House NAme ", checkInDate );
              }
      
          }
      }
      
      //6   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
      
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
      
      
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
      
              var AllNotes = " Out Notes: " + checkOutNotes;
      
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
      
              var lastDateString = lastDate1.toString();
      
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House NAme - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House NAme", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House NAme - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House NAme " + " - " + conc, checkOutDate);
               }
          }
      }
      
      //6 //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
      
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
      
      
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
      
      
              var AllNotes = "Turn Notes: " + turnNotes;
      
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House NAme - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House NAme ", turnDate );
              }
      
          }
      }
      
      
      
      //6 //--START-- Mark houses whose entries have already been created  --START--//   
      
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BV1371:BV1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
      
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CG" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BX1371:BX1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CH" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('BZ1371:BZ1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CI" + cell;
              ssM.getRange(cell2).setValue("x");
          }
  }
  
      //7  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  

    
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('CC1371:CJ1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('CJ1371:CL1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
            
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
      
      
      
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
      
      
              var AllNotes = "In Notes: " + checkInNotes;
      
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
      
          }
      }
      
      //7   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
      
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
      
      
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
      
              var AllNotes = " Out Notes: " + checkOutNotes;
      
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
      
              var lastDateString = lastDate1.toString();
      
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
          }
      }
      
      //7 //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
      
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
      
      
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
      
      
              var AllNotes = "Turn Notes: " + turnNotes;
      
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
      
          }
      }
      
      
      
      //7 //--START-- Mark houses whose entries have already been created  --START--//   
      
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('CD1371:CD1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
      
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CJ" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('CF1371:CF1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CK" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('CH1371:CH1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "CL" + cell;
              ssM.getRange(cell2).setValue("x");
          }
  }
 
      //8  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  

    
      var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('ck1371:CP1828').getValues();
      var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('CM1371:CO1828').getValues();
      var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
            
      for (i = 0; i < values.length; i++) {
          if (values[i][1] != "" && matrix[i][0] == "") {
      
      
      
              var checkInDate = (values[i][0]);
              var checkInNotes = (values[i][7]);
      
      
              var AllNotes = "In Notes: " + checkInNotes;
      
              if (values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name - Notes", checkInDate, { description: AllNotes });
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN House Name ", checkInDate );
              }
      
          }
      }
      
      //8   //--START-- Search for CHECK OUT DATES - NEXT CHECK IN DATE - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//   
      
      for (i = 0; i < values.length; i++) {
          if (values[i][3] != "" && matrix[i][1] == "") {
              for (j = 0; j < (values.length - (i + 1)) && values[i + j][1] == ""; j++) { }
      
      
              var checkOutDate = (values[i][0]);
              var nextCheckInDate = (values[i + j][0]);
              var checkOutNotes = (values[i][7]);
      
              var AllNotes = " Out Notes: " + checkOutNotes;
      
              var string = nextCheckInDate.toString();
              var split1 = string.split(' ', 3);
              var conc = (split1[1]) + " " + (split1[2]);
      
              var lastDateString = lastDate1.toString();
      
              if (string == lastDateString && values[i][7] != "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes ", checkOutDate, { description: AllNotes });
              } 
              if (string == lastDateString && values[i][7] == "") {
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name", checkOutDate);
              }
              if (string != lastDateString && values[i][7] != "") { 
                 CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name - Notes " + " - " + conc, checkOutDate, { description: AllNotes });
              }
              if (string != lastDateString && values[i][7] == "") { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("OUT House Name " + " - " + conc, checkOutDate);
               }
          }
      }
      
      //8 //--START-- Search for TURN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES dates  --START--//
      
      for (i = 0; i < values.length; i++) {
          if (values[i][5] != "" && matrix[i][2] == "") {
      
      
              var turnDate = (values[i][0]);
              var turnNotes = (values[i][7]);
      
      
              var AllNotes = "Turn Notes: " + turnNotes;
      
              if (values[i][7] != ""){
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name - Notes ", turnDate, { description: AllNotes });
                  
              } else { 
                  CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("TURN House Name ", turnDate );
              }
      
          }
      }
      
      
      
      //8 //--START-- Mark houses whose entries have already been created  --START--//   
      
      var ss = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1')
      var ssM = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1')
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Cl1371:Cl1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
      
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "Cm" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Cn1371:Cn1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "Cn" + cell;
              ssM.getRange(cell2).setValue("x");
          }
      }
      
      var markedValues = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('Cp1371:Cp1828').getValues();
      
      for (i = 0; i < markedValues.length; i++) {
          if (markedValues[i] != "") {
              var cell = i + 1371;
              var cell2 = "Co" + cell;
              ssM.getRange(cell2).setValue("x");
          }
  }
  
        //9  // --START-- Search for CHECK IN DATES - CHECK IN NOTES - CHECK OUT NOTES - TURN NOTES  --START--//  

    
        var values = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('cs1371:Cz1828').getValues();
        var matrix = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('matrixSheet1').getRange('Cp1371:Cr1828').getValues();
        var lastDate1 = SpreadsheetApp.openById('//GoogleSheetName').getSheetByName('Sheet1').getRange('A1828:A1828').getValue();
              
        for (i = 0; i < values.length; i++) {
            if (values[i][1] != "" && matrix[i][0] == "") {
        
        
        
                var checkInDate = (values[i][0]);
                var checkInNotes = (values[i][7]);
        
        
                var AllNotes = "In Notes: " + checkInNotes;
        
                if (values[i][7] != "") {
                    CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN 45745 Camino Del Rey - Notes", checkInDate, { description: AllNotes });
                } else { 
                    CalendarApp.getCalendarById("usern@group.calendar.google.com").createAllDayEvent("IN 45745 Camino Del Rey ", checkInDate );
                }
        
            }
        }
        
        
