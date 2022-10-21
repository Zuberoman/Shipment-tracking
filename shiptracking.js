function EditCell(e) {
    var sheet = SpreadsheetApp.openById('*******').getSheetByName("2021/2022: Reception"); 
    var AdressSheet = SpreadsheetApp.openById('*******').getSheetByName("Address Book");
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var AdressColumn = sheet.getRange("C2:C30").getValues(); 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var column = ss.getActiveCell().getColumn(); //kolumna w której wystąpił event
    var row = ss.getActiveCell().getRow(); //rząd w którym wystąpił event
    var cell = ss.value; //Wartość komórki w której nastąpił event
    var zaznaczenie = ss.getSheetByName("2021/2022: Reception").getActiveCell().getRowIndex();
    var kolumna = ss.getSheetByName("2021/2022: Reception").getActiveCell().getColumn();
    var komorka = ss.getSheetByName("2021/2022: Reception").getRange(zaznaczenie,kolumna).getValue();
    
    
    var subject = "CZE Protoshop shipment tracking:" + sheet.getRange(row, column-12).getValue();
    
    var body = '\n\n' + "Zarejestrowano przesyłkę od: " + sheet.getRange(row, column-1).getValue() + '<br />' + "Reference: " + sheet.getRange(row, column-10).getValue() + '<br />' +  'Rev.: ' + sheet.getRange(row, column-9).getValue() + '<br />' + "Zawierającą: " + sheet.getRange(row, column-11).getValue() + '<br />' + 'Ilość: ' + sheet.getRange(row, column-8).getValue() ;  //2018-11-16
    var options = {}
    options.htmlBody = body + '<br />' + '<br />' + '<a href=\"' + 'https://docs.google.com/spreadsheets/d/1dBq-CWvz61K9TIbZCJLPOIpysxmSiuDEwTqJZBpRGL0/edit#gid=2055216124' + '">Link do pliku</a>'  + '<br />' + '<br />' + '<i> Wiadomość wygenerowana automatycznie, proszę na nią nie odpowiadać. </i>' + '<br />' + '<br />';
    
   var body2 = '\n\n' + "Zarejestrowano przesyłkę od: " + sheet.getRange(row, column-1).getValue() + '<br />' + "Reference: " + sheet.getRange(row, column-10).getValue() + '<br />' +  'Rev.: ' + sheet.getRange(row, column-9).getValue() + '<br />' + "Zawierającą: " + sheet.getRange(row, column-11).getValue() + '<br />' + 'Ilość: ' + sheet.getRange(row, column-8).getValue() + '<br><br>' + '<a href=\"' + 'https://docs.google.com/spreadsheets/d/1dBq-CWvz61K9TIbZCJLPOIpysxmSiuDEwTqJZBpRGL0/edit#gid=2055216124' + '">Link do pliku</a>'  + '<br />' + '<br />' + '<i> Wiadomość wygenerowana automatycznie, proszę na nią nie odpowiadać. </i>' + '<br />' + '<br />';


    var namiary = AdressSheet.getRange("A2:A30").getDisplayValues();
    

    if(activeSheet.getName() == "2021/2022: Reception" && cell !== "" && column==15){
    
// SZUKANIE ADRESÓW W KSIĄŻCE MAILOWEJ
         
         var adresat = [];

         for(var i =0; i<namiary.length; i++){
           if(namiary[i] == komorka){
             adresat.push(AdressSheet.getRange(i+2, 3).getValue());
           }
         }

          }
          Logger.log(adresat[0])
//KONIEC SZUKANIA ADRESÓW W KSIĄŻCE MAILOWEJ
          MailApp.sendEmail({to:adresat[0], subject:subject, htmlBody:body2});
      var msg = "Wysłano powidomienie o przesyłce do ";
      var contmsg = msg.concat(adresat);
  ss.toast(contmsg,"Shipment Tracking");             
}

