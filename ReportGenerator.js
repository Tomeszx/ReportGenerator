function invoicing_reports() {
  //Acces to the sheets
  const start_time_script = new Date().getTime();

  const invoicing_data_ID = '1GWEZlPTAcURQIgjSCWuvRkFMWWtbhjp1qTPAzgiyNt4'
  const invoicing_data = SpreadsheetApp.openById(invoicing_data_ID).getSheetByName('Data for Invoicing').getRange("3:3").getValues();
  const template_id = DriveApp.getFileById(invoicing_data[0][3]);
  const folder_reports = DriveApp.getFolderById(invoicing_data[0][2]);
  const folder_with_data = DriveApp.getFolderById(invoicing_data[0][1])

  const reports_queue_ID = '1GWEZlPTAcURQIgjSCWuvRkFMWWtbhjp1qTPAzgiyNt4'
  const reports_queue = SpreadsheetApp.openById(reports_queue_ID).getSheetByName(invoicing_data[0][4]);
  const queue_data = reports_queue.getRange("A2:Q").getValues();

 //find file
  var detalied_spreadsheets = folder_with_data.getFiles();
  while (detalied_spreadsheets.hasNext()) {
    // Checking the duration of the script
    var end_time_script = ((new Date().getTime() - start_time_script)/60000).toFixed(2);
    if (end_time_script > 15){
      console.log("End script")
      
      var partners_folders2 = folder_reports.getFolders();
      var l = 1;
      while (partners_folders2.hasNext()) {
        l += 1;
        var current_folder = partners_folders2.next();
        var folder_name = current_folder.getName();
        var folder_link = current_folder.getUrl();
        reports_queue.getRange(l, 7,1,2).setValues([[folder_name, folder_link]]);
      }
      return;
    }
  
    var start_time_file = new Date().getTime();
    var sheet = detalied_spreadsheets.next(); //find spreadsheet
    var sheet_id = sheet.getId();
    var spreadsheet_all = SpreadsheetApp.openById(sheet_id)
    var detalied_name_spread = sheet.getName();
    
    var spreads_done = reports_queue.getRange("I:I").getValues().filter(String);
    if(spreads_done.filter(function (e) { return e == detalied_name_spread}).length > 0){continue};
    reports_queue.getRange(spreads_done.length + 1, 9).setValue(detalied_name_spread);SpreadsheetApp.flush();
    console.log(detalied_name_spread);

    //Get data
    var all_data = pull_data_request([[sheet_id, spreadsheet_all]]);
    var title_width = new Array;
    for(var a = 0;  all_data[0].length > a; a++){title_width.push(all_data[0][a] + "AAA")}; //add row to later do good resize of the columns

    // Find Partner to generate report
    for (var queue_c = 2;  queue_data.length + 1 >= queue_c; queue_c++){
      if(queue_data[queue_c - 2][0] == ""){continue}; 

      //Reports data shop name, etc....
      var queue_legal_name = queue_data[queue_c - 2][1];
      var queue_shop_name = queue_data[queue_c - 2][0];
      var all_shop_ID = queue_data[queue_c - 2][2].split(",");
      if(queue_legal_name != ""){var queue_partner_name = queue_legal_name;}else{var queue_partner_name = queue_shop_name};
      var report_title = queue_partner_name + " " + detalied_name_spread;

      // Check if file already exist in folder
      var is_folder_exist = DriveApp.getFoldersByName(queue_partner_name);
      if(is_folder_exist.hasNext()){
        var is_file_exist = is_folder_exist.next().getFilesByName(report_title)
        if(is_file_exist.hasNext()){continue}
      }

      // filter data from reports
      var first_id = []
      for (shop_id in all_shop_ID){
        if(first_id.length <1){var first_id = all_data.filter(function (e) { return e[7] == all_shop_ID[shop_id]})}
        else{var next_id = all_data.filter(function (e) { return e[7] == all_shop_ID[shop_id]}); if(next_id.length >0 ){first_id = first_id.concat(next_id);} };}
      var filter_data = first_id;
      if (filter_data.length < 1 || queue_partner_name == ""){continue};

      //Find Partners folder
      var create_new_folder = true;
      var partners_folders = folder_reports.getFolders();
      while (partners_folders.hasNext()) {
        var folder = partners_folders.next();
        if(folder.getName() == queue_partner_name){
          create_new_folder = false;
          var new_folder =  DriveApp.getFolderById(folder.getId());
          break;
        };
      };

      //If can`t find, create a new one
      if(create_new_folder == true){var new_folder = DriveApp.getFolderById(folder_reports.createFolder(queue_partner_name).getId())};

      // Create copy of template in Partners folder
      try{var new_template = template_id.makeCopy(report_title, new_folder);}
      catch{SpreadsheetApp.flush(); var new_template = template_id.makeCopy(report_title, new_folder);}
      var new_template_id = new_template.getId();
      var partners_report = SpreadsheetApp.openById(new_template_id).getSheets()[0]; 

      partners_report.getRange(1,1,1, all_data[0].length).setValues([all_data[0]]);
      partners_report.getRange(2,1,filter_data.length, filter_data[0].length).setValues(filter_data);
      prepare_inv_report(partners_report, filter_data.length + 2, [all_data[0]], title_width);
    }
    console.log("The script finished the file in" , ((new Date().getTime() - start_time_file)/60000).toFixed(2), "min.");
  }
}


function prepare_inv_report(sheet, sheetLastRow, titles, title_width){
  sheet.getRange(sheetLastRow,1,1,titles[0].length).setValues(new Array(title_width)); // add one row for resize columns

  sheet.getRange(1,1,sheetLastRow + 2,titles[0].length).createFilter();
  sheet.autoResizeColumns(1,titles[0].length);
  sheet.deleteRow(sheetLastRow);
}


function pull_data_request(sheets_ranges){
  var token = ScriptApp.getOAuthToken();

  // get all ranges for every 20K rows
  
  var ranges = []
  for(x in sheets_ranges){
    var spreadsheetId = sheets_ranges[x][0]
    var report = sheets_ranges[x][1]

    var data_last_row = report.getLastRow()
    var data_last_col = columnToLetter(report.getLastColumn())
    for(var s = 1; s < data_last_row; s+=20000){
      if (s + 19999 > data_last_row){
        e = report.getName() + "!A" + String(s) + ":" + data_last_col + String(data_last_row)}
      else{
        e = report.getName() + "!A" + String(s) + ":" + data_last_col + String(s + 19999)}
      ranges.push(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${e}?majorDimension=ROWS`)
    }
  }

  //Make a request
  var requests = ranges.map(function(url) {
    return {
      method: "get",
      url: url,
      headers: {Authorization: "Bearer " + token},
      muteHttpExceptions: true,
    }
  });

  var res = UrlFetchApp.fetchAll(requests);
  var values = res.reduce(function(ar, e) {
    Array.prototype.push.apply(ar, JSON.parse(e.getContentText()).values);
    return ar;
  }, []);

  // fill blank spaces in data
  values = values.map(r => r.length == values[0].length ? r : r
    .concat(Array(values[0].length - r.length).fill("")));

  return values
}


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
