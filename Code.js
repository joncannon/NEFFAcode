//Test!

function onOpen() {
  // Add a menu with some items, some separators, and a sub-menu.
  SpreadsheetApp.getUi().createMenu('NEFFA')
      .addItem('Print active column', 'print_column')
      .addItem('Import new applications', 'import_applications')
      .addItem('Import performer data', 'get_data')
      .addToUi();
  console.log("added menu item")
}

// Every time a data cell is edited, 
function onEdit(e) {
  console.log("onEdit");
  var range = e.range;
  var cellToAlter;
  
  
  // if you're adding or changing an application number in a room (not including extra conflict checks)
  var ThisSpreadsheet=SpreadsheetApp.getActiveSpreadsheet();
  var ThisSheet=ThisSpreadsheet.getActiveSheet();
  var day=ThisSheet.getName();
  
  
  var no_change = false;
  
  if(day=="Friday"){
    day_index=0;
  } else if(day=="Saturday"){
    day_index=1;
  }else if(day=="Sunday"){
    day_index=2;
  }else{
    no_change=true;
  }
  console.log(no_change);
  if(no_change==false){
    console.log("here");
    var data_column_constants = ["FridayRooms", "SaturdayRooms", "SundayRooms"];
    var data_column_list=ThisSpreadsheet.getRangeByName(data_column_constants[day_index]).getValues();
    var isdanceroom_column_constants = ["FridayIsDance", "SaturdayIsDance", "SundayIsDance"];
    var is_dance_room=ThisSpreadsheet.getRangeByName(isdanceroom_column_constants[day_index]).getValues();
    var cellToAlter;
    console.log("column: "+range.getColumn());
    
    
    var foundit=false;
    var column_index;
    var add_line=false;
    var i;
    for(i=0;i<data_column_list.length;i++){
      if(data_column_list[i][0]==range.getColumn() & foundit==false){
        foundit=true;
        column_index=i;
        add_line=true;
      }else if(data_column_list[i][0]+1==range.getColumn() & is_dance_room[i][0] & foundit==false){
        foundit=true;
        column_index=i;
      }
    }
    
    if(foundit){
      var rangetocolor;
    
      if(is_dance_room[column_index][0]){
        cellToAlter = range.offset(0, 2);
        rangetocolor=ThisSheet.getRange(1, data_column_list[column_index][0], 2, 3);
      }else{
        cellToAlter = range.offset(0, 1);
        rangetocolor=ThisSheet.getRange(1, data_column_list[column_index][0], 2, 2)
      }
    
    
      console.log("column to alter: "+ cellToAlter.getColumn());
    
    
      console.log("found: "+foundit);
//  var first_extra_column = ThisSpreadsheet.getRangeByName("ExtraGridCol").getValue()
//  console.error(first_extra_column)

      
      if(add_line){
        if (range.isBlank()){
          // if you're deleting an application number, clear the top border
          cellToAlter.setBorder(false, null, null, null, null, null);
          console.log("cleared");
        }else {
          //if you're adding one, add a top border
          cellToAlter.setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
          console.log("added");
        }
      }
      
      rangetocolor.setBackground("blue");
      console.log("setbg" + rangetocolor);
    }
  }
}

function print_column() {

  var ThisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var ThisSheet=ThisSpreadsheet.getActiveSheet();
  var day=ThisSheet.getName();
  var day_index;
  console.log("day: "+day)  
  var Grid=ThisSpreadsheet.getSheetByName(day + " AUX4");
  var Grid2=ThisSpreadsheet.getSheetByName(day + " AUX3");
    console.log("Grid2: "+Grid2)  
  var AppSheet=ThisSpreadsheet.getSheetByName("Full applications");
  var WebGrid=ThisSpreadsheet.getSheetByName("Webgrid");
  var Constants=ThisSpreadsheet.getSheetByName("Constants");
  var Useful=ThisSpreadsheet.getSheetByName("Useful application data");
  var PDFGrid=ThisSpreadsheet.getSheetByName("PDF display draft");
  
  var active_column = ThisSpreadsheet.getActiveCell().getColumn();
  console.log("active column: "+active_column)
  var room = Grid2.getRange(1, active_column).getValue()
  console.log("room: "+room)
  if(day=="Friday"){
    day_index=0;
  } else if(day=="Saturday"){
    day_index=1;
  }else if(day=="Sunday"){
    day_index=2;
  }else{
    throw new Error("Bad sheet");
  }
  

  var data_column_constants = ["FridayRooms", "SaturdayRooms", "SundayRooms"];
  var day_codes = ["F", "S", "U"];
  var data_column_list=ThisSpreadsheet.getRangeByName(data_column_constants[day_index]).getValues();

  var foundit=false;
  var column_index;
  var i;
  for(i=0;i<data_column_list.length;i++){
    if(data_column_list[i][0]==active_column){
      foundit=true;
      column_index=i;
    }
  }
  
  if(foundit){
  
    var last_row = WebGrid.getLastRow();
    var check_row=1;
    var i
    Logger.log("last row: "+last_row)
    for (i=1; i<=last_row; i++){
      Logger.log(WebGrid.getRange(check_row, 6).getValue().toString());
      if(WebGrid.getRange(check_row, 6).getValue().toString()== room & WebGrid.getRange(check_row, 5).getValue().toString()== day_codes[day_index]){
        WebGrid.deleteRow(check_row)
      }else{
        check_row++;
      }
    }
    var start_data_row_constants = ["FridayStartRow", "SaturdayStartRow", "SundayStartRow"];
    var end_data_row_constants = ["FridayEndRow", "SaturdayEndRow", "SundayEndRow"];
  
    var start_data_row=ThisSpreadsheet.getRangeByName(start_data_row_constants[day_index]).getValue();
    var end_data_row=ThisSpreadsheet.getRangeByName(end_data_row_constants[day_index]).getValue();
  
    var num_rows = end_data_row - start_data_row;
    
    var start_print_row_constants = ["FridayStartPrint","SaturdayStartPrint","SundayStartPrint"];
    var start_print_row = ThisSpreadsheet.getRangeByName(start_print_row_constants[day_index]).getValue();
    var start_print_column = ThisSpreadsheet.getRangeByName("StartPrintColumn").getValue();;
    var print_column=start_print_column+2*column_index;
    
    var print_range = PDFGrid.getRange(start_print_row, print_column, num_rows, 2);
    print_range.breakApart();
    print_range.clear();
    
    var title_column=ThisSpreadsheet.getRangeByName("EventCol").getValue();
    var description_column = ThisSpreadsheet.getRangeByName("DescriptionCol").getValue();
    var type_column = ThisSpreadsheet.getRangeByName("TypeCol").getValue();
    var level_column = ThisSpreadsheet.getRangeByName("LevelCol").getValue();
    var duo_performer_column = ThisSpreadsheet.getRangeByName("DuoIDCol").getValue()
    var busy_performer_column = ThisSpreadsheet.getRangeByName("BusyIDCol").getValue();
    var main_performer_column = ThisSpreadsheet.getRangeByName("MainIDCol").getValue();
    var display_name_column = ThisSpreadsheet.getRangeByName("DisplayNameCol").getValue();
    
    var data_column_constants = ["FridayRooms", "SaturdayRooms", "SundayRooms"];
    var isdanceroom_column_constants = ["FridayIsDance", "SaturdayIsDance", "SundayIsDance"];
    
    var is_dance_room=ThisSpreadsheet.getRangeByName(isdanceroom_column_constants[day_index]).getValues();
    console.log(is_dance_room)
    
    var current_app= Grid.getRange(start_data_row, active_column).getValue();
    var current_app_2;
    var current_app_offset=0;
    var current_app_row=start_data_row;
    var row_count;

    
    for(row_count = 0; row_count < num_rows+1; row_count++){
      var data_row = start_data_row+row_count;
      var next_app=Grid.getRange(data_row, active_column).getValue();
      console.log("row: "+row_count);
      
      if(next_app!=current_app || row_count==num_rows){
        var num_total_cells = row_count-current_app_offset;
  //      console.log(Grid2.getRange(current_app_row, 2).getValue());
//        console.log(Utilities.formatDate(Grid2.getRange(current_app_row, 2).getValue(), "GMT-05:00", "HHmm"));
 //       var time = Utilities.formatDate(Grid2.getRange(current_app_row, 2).getValue(), "GMT-05:00", "HHmm");
        var time = Grid2.getRange(current_app_row, 2).getValue();
        console.log("current app offset: "+current_app_offset)
        console.log("first print row: "+start_print_row+current_app_offset)
        var all_cells = PDFGrid.getRange(start_print_row+current_app_offset, print_column, num_total_cells, 2);

        all_cells.getCell(1,1).setFontSize(10);
        
        if(current_app=="---"){
          WebGrid.appendRow(["","","","",day_codes[day_index],room,time,""]);
          
          
          all_cells.merge();
          all_cells.setBorder(true, true,true,true,null,null);
          
          all_cells.getCell(1,1).setValue("");
          console.log("3")
        }else{
          var num_title_cells = Math.floor(num_total_cells/2);
          var num_performer_cells = num_total_cells-num_title_cells;
          
          var room = Grid2.getRange(1, active_column).getValue();
          
          var app_sheet_row = Grid2.getRange(current_app_row, active_column).getValue();
          
          var title = AppSheet.getRange(app_sheet_row, title_column).getValue();
          var description = AppSheet.getRange(app_sheet_row, description_column).getValue();
          var type = AppSheet.getRange(app_sheet_row, type_column).getValue();
          var level = AppSheet.getRange(app_sheet_row, level_column).getValue();
          var mainID = Useful.getRange(app_sheet_row, main_performer_column).getValue();
          var display_name = Useful.getRange(app_sheet_row, display_name_column).getValue();
          
          var title_cells;
          
          if(num_title_cells>0){
            title_cells = PDFGrid.getRange(start_print_row+current_app_offset, print_column, num_title_cells, 2);
            title_cells.merge();
            
            title_cells.getCell(1,1).setValue(title);
            title_cells.getCell(1,1).setWrap(true);
            title_cells.getCell(1,1).setFontWeight("bold");
            title_cells.getCell(1,1).setVerticalAlignment("top");
            title_cells.getCell(1,1).setHorizontalAlignment("left");
          }
          
          var performer_cells = PDFGrid.getRange(start_print_row+current_app_offset+num_title_cells, print_column, num_performer_cells, 1);
          
          var code_cells = PDFGrid.getRange(start_print_row+current_app_offset+num_total_cells-1, print_column+1);
          
          performer_cells.merge();
//          code_cells.merge();
          
          var output_row = [current_app, title, description, type+" "+level, day_codes[day_index], room, time, mainID];
          
          var duo_performer = AppSheet.getRange(app_sheet_row, duo_performer_column).getValue();
          if (duo_performer!==""){
            output_row.push(duo_performer);
          }
          var busy_performer = AppSheet.getRange(app_sheet_row, busy_performer_column).getValue();
          if (busy_performer!==""){
            output_row.push(busy_performer);
          }
          
          
          if (is_dance_room[column_index][0]){
            
            var active_column_2 = active_column + 1;
            var app_sheet_row_2 = Grid2.getRange(current_app_row, active_column_2).getValue();
            
            current_app_2=Grid.getRange(current_app_row, active_column_2).getValue();

            if(current_app_2!="---"){
              var main_performer_2 = Useful.getRange(app_sheet_row_2, main_performer_column).getValue();
              output_row.push(main_performer_2);
              console.log("main 2: "+main_performer_2);
              
              var duo_performer_2 = AppSheet.getRange(app_sheet_row_2, duo_performer_column).getValue();
              if (duo_performer_2!==""){
                output_row.push(duo_performer_2);
                console.log("duo 2: "+duo_performer_2);
              }
              var busy_performer_2= AppSheet.getRange(app_sheet_row_2, busy_performer_column).getValue();
              if (busy_performer_2!==""){
                output_row.push(busy_performer_2);
              }
              
              if (display_name.trim() !== ""){
                display_name = display_name + "\n";
              }
              display_name = display_name + Useful.getRange(app_sheet_row_2, display_name_column).getValue();
            }
          }
          
          var code;
          if (level=="NA"){
            code=type;
          }else{
            code=type+level;
          }
          
          all_cells.setBackground("white");
          all_cells.setBorder(true, true,true,true, false,false);
          
          performer_cells.getCell(1,1).setValue(display_name);
          performer_cells.getCell(1,1).setWrap(true);
          performer_cells.getCell(1,1).setFontWeight("normal");
          performer_cells.getCell(1,1).setVerticalAlignment("bottom");  
          performer_cells.getCell(1,1).setHorizontalAlignment("left");
          
          code_cells.getCell(1,1).setValue(code);
          code_cells.getCell(1,1).setFontWeight("bold");
          code_cells.getCell(1,1).setHorizontalAlignment("right");
          code_cells.getCell(1,1).setVerticalAlignment("bottom");
          WebGrid.appendRow(output_row);
        }
        current_app = next_app;
        current_app_offset = row_count;
        current_app_row = data_row;
      }
    }
    
    var cells_to_color;
    if (is_dance_room[column_index][0]){
      cells_to_color=3;
    }else{
      cells_to_color=2;
    }
    
    ThisSheet.getRange(1, active_column, 2, cells_to_color).setBackground("#666666");
    console.log("coloring "+active_column+", "+cells_to_color+" cells");
//    for(i=0; i<num_rows; i++){
  //    if(i%8 < 4){
        
    //  }
    //}
  }else{
    throw new Error("Bad column");
  }
}


function get_data() {
  console.log('getting data')
  
  var options = {};
  options.headers = {"Authorization": "Basic " + Utilities.base64Encode("TheBoom" + ":" + "TheHook")};
  var response = UrlFetchApp.fetch('https://www.neffa.org/cgi-bin/pcheck/dumppn.pl', options);
  var filetext = response.getContentText().replace("\"", "\'");
  var csvData = Utilities.parseCsv(filetext, "\t")
  var PerformerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("performers")
  PerformerSheet.clearContents().clearFormats();
  PerformerSheet.getRange(1,1).setValue("Perf Number");
  PerformerSheet.getRange(1,2).setValue("Perf Name");
  PerformerSheet.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData);

  var response = UrlFetchApp.fetch('https://cgi.neffa.org//pcheck/dumpmem.pl', options);
  var filetext = response.getContentText();
  var csvData = Utilities.parseCsv(filetext, "\t")
  var MemberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("members")
  MemberSheet.clearContents().clearFormats();
  MemberSheet.getRange(1,1).setValue("Group Number");
  MemberSheet.getRange(1,2).setValue("Member Number");
  MemberSheet.getRange(1,3).setValue("type");
  MemberSheet.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData);
}



function import_applications() {
  console.log('getting applications')
  
  var ThisSpreadsheet=SpreadsheetApp.getActiveSpreadsheet()
  var AppIDcol = ThisSpreadsheet.getRangeByName('AppIDcol').getValue();
  
  var options = {"muteHttpExceptions": true,
                 "method"  : "GET"} ;
  var response = UrlFetchApp.fetch('https://apply.pacew.org/download.php?direct_download=1YXfx4Qmyx', options);

  var filetext = response.getContentText();
  
//  filetext=filetext.replace(/\n,/g, 'NPLACEHOLDER');
//  filetext=filetext.replace(/\r,/g, 'RPLACEHOLDER');
//  filetext=filetext.replace(/\n/g, ' ');
//  filetext=filetext.replace(/\r/g, ' ');
//  filetext=filetext.replace(/NPLACEHOLDER/g, '\n,');
//  filetext=filetext.replace(/RPLACEHOLDER/g, '\r,');
//  filetext=filetext+'\n"\n';
  var filetext = filetext.replace(/(?=["'])(?:"[^"\\]*(?:\\[\s\S][^"\\]*)*"|'[^'\\]\r\n(?:\\[\s\S][^'\\]\r\n)*')/g, function(match) { return match.replace(/\r\n/g,"\r\n")}); 
  //var filetext = filetext.replace(/(?=["'])(?:"[^"\\]*(?:\[\s\S][^"\])"|'[^'\]\r\n(?:\[\s\S][^'\]\r\n)')/g, function(match) { return match.replace(/\r\n/g,"\r\n")} );
  
  console.log(filetext)
  var csvData = Utilities.parseCsv(filetext, ",")

  var AppSheet = ThisSpreadsheet.getSheetByName("Full applications")
  var allAppData = AppSheet.getRange(2,1,AppSheet.getLastRow(), AppSheet.getLastColumn());
  
  console.log(AppSheet.getLastRow());
  
  allAppData.sort(AppIDcol);
  
  var currentRow=2;
  var currentID=AppSheet.getRange(currentRow, AppIDcol).getValue();
  console.log(csvData.length);
  console.log(currentID);
  for(i=0; i<csvData.length; i++){
    var line = csvData[i];
    console.log(currentID+" , "+line[AppIDcol-1]);
    while((currentID<line[AppIDcol-1] | currentID == "")&& currentRow<AppSheet.getLastRow()){
      currentRow++;
      currentID = AppSheet.getRange(currentRow, AppIDcol).getValue();
      console.log(currentID+" , "+line[AppIDcol-1]);
    }
    if(currentID>line[AppIDcol-1] | currentRow==AppSheet.getLastRow()){
      AppSheet.appendRow(line);
    }
    
  }
}