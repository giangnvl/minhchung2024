//###########################################################
var tenCTDT = ""
var inputFileName = "BM01_" + tenCTDT + ".xlsx";  // Tên file Excel biểu mẫu 01
var outputFileName = "BM03_" + tenCTDT; // Tên file Google Sheet biểu mẫu 03 được xuất ra
var bm1Folder = ""; // ID của folder chứa các file Excel biểu mẫu 01
var bm3Folder = ""; // ID của folder chứa các file Google Sheet biểu mẫu 03 được xuất ra.
//###########################################################

var data = []; // Mảng chứa dữ liệu đọc từ file Excel biểu mẫu 01
var rows = []; // Mảng chứa dữ liệu đã được xử lý để ghi vào biểu mẫu 03. VD [["Tiêu chuẩn", "Tiêu chí", "Mã hóa minh chứng", "ID1(nếu có)"]]

function main() {
  try {
    convertExcelToArray(); // Excel -> Google Sheets tạm -> data[]
    // Gọi function chuyển data[] --> rows[]
    for (var i = 0; i < data.length; i++) {
      processData(data[i][0], data[i][1]);
    }
    // Chuyển row[] --> Google Sheets biểu mẫu 03
    exportData();
    Logger.log("Chương trình đã hoàn thành")
  } catch (e) {
    Logger.log("Xảy ra lỗi: " + e);
    return;
  }
}

// Xuất dữ liệu từ rows -> Google Sheets biểu mẫu 03
function exportData() {
  // Tạo một Google Sheet mới trong thư mục có ID là folderId
  var config = {
    title: outputFileName,
    parents: [{ id: bm3Folder }],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var spreadsheet = Drive.Files.insert(config);
  var sheet = SpreadsheetApp.openById(spreadsheet.id).getSheets()[0];
  // Ghi dòng header vào file
  sheet.appendRow(["Tiêu chuẩn", "Tiêu chí", "Mã hóa minh chứng", "ID1 (nếu có)", "ID2 (nếu có)", "ID3 (nếu có)", "ID4 (nếu có)", "ID5 (nếu có)", "ID6 (nếu có)", "ID7 (nếu có)", "ID8 (nếu có)", "ID9 (nếu có)", "ID10 (nếu có)"]);
  // Ghi nội dung chính vào file
  for (var i = 0; i < rows.length; i++)
    sheet.appendRow(rows[i])
  // Chỉnh độ rộng các cột
  sheet.setColumnWidths(1, 2, 90);
  sheet.setColumnWidths(3, 3, 140);
  sheet.setColumnWidths(4, 13, 90);
  // Format Google Sheet vừa được tạo
  sheet.setRowHeights(1, 800, 27);
  sheet.getRange('A1:P1').setFontWeight('bold')
  sheet.getDataRange().setBorder(true, true, true, true, true, true, "#4285f4", SpreadsheetApp.BorderStyle.SOLID)
  sheet.getDataRange().setHorizontalAlignment("center")
  sheet.getDataRange().setVerticalAlignment("middle")
  sheet.setFrozenRows(1);
}

// Xử lý dữ liệu từ mảng data để xuất vào file Google Sheets
function processData(criteria, num) { // processData("3.1", 5)
  var tmp = criteria.split(".");
  Logger.log("Đang xử lý: " + tmp[0] + "." + tmp[1]);
  var preindex = tmp[0] < 10 ? "0" + tmp[0] : tmp[0];
  for (var i = 1; i <= num; i++) { // Duyệt qua số lượng mã hóa minh chứng của từng tiêu chí
    var index = i < 10 ? "0" + i : i
    var envString = "H" + tmp[0] + "." + preindex + ".0" + tmp[1] + "." + index;// 2.1 -> H2.02.01.index
    rows.push(["Tiêu chuẩn " + tmp[0], "Tiêu chí " + criteria, envString, "", "", "", "", "", "", ""])
  }
}

// Function chuyển đổi file excel tại bm1Folder sang mảng data
function convertExcelToArray() {
  var inputFile = DriveApp.getFolderById(bm1Folder).getFilesByName(inputFileName);
  if (!inputFile.hasNext()) {
    Logger.log("Không thể tìm thấy file " + inputFileName);
    return;
  }
  var excelFile = inputFile.next();
  var blob = excelFile.getBlob();
  var config = {
    title: "[TMP] " + excelFile.getName(), //sets the title of the converted file
    parents: [{ id: excelFile.getParents().next().getId() }],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var spreadsheet = Drive.Files.insert(config, blob);
  // Đọc dữ liệu và ghi vào data[]
  var sheet = SpreadsheetApp.openById(spreadsheet.id).getSheets()[0];
  data = sheet.getRange("A2:B51").getValues()
  // Xóa file Google Sheet tạm sau khi xử lý xong
  Drive.Files.trash(spreadsheet.id);
}
