//###########################################################\
var rootFolder = ""; // ID cây thư mục của CTĐT
var tenCTDT = ""
var beginRow =   // 
var proofFolder = ""; // ID của folder chứa kho minh chứng
var dataFolderID = "" // ID folder Data, chứa danh sách các file CSV được tạo ra bởi script 03.LayDanhSachFolderID
var bm1Folder = ""; // ID của folder chứa các file Excel biểu mẫu 01
var bm3Folder = ""; // ID của folder chứa các file Google Sheet biểu mẫu 03 được xuất ra.
var csvListFolderID = tenCTDT + "_ListFolderID.csv"; // File ListFolderID của từng CTĐT
var listFolderID = [] // Chứa danh sách FolderID (nội dung file CSV)
//###########################################################

function main() {
  if (!checkData()) {
    Logger.log("KHÔNG PASS ĐƯỢC HÀM KIỂM TRA DỮ LIỆU ĐẦU VÀO")
    return
  }
  processData()
  Logger.log("HOÀN TẤT")
  //readData(1);
}

// Function đọc dữ liệu biểu mẫu 03 (sau khi đã test lỗi) và tạo shortcut
function processData() {
  var bm03 = DriveApp.getFolderById(bm3Folder).getFilesByName("BM03_" + tenCTDT).next(); // Biểu mẫu 03 dạng Google Sheet
  var tmp = DriveApp.getFolderById(dataFolderID).getFilesByName(csvListFolderID);
  listFolderID = Utilities.parseCsv(tmp.next().getBlob().getDataAsString(), ";"); // Mảng chứa danh sách folder id trong file CSV

  // Đọc dữ liệu trong file BM03
  var sheet = SpreadsheetApp.openById(bm03.getId()).getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();
  // Lặp qua từng dòng trong biểu mẫu 3, bỏ qua dòng header
  for (var row = beginRow; row < data.length; row++) { // Row bắt đầu từ 1, bỏ dòng header
    // Tìm ID của folder chứa mã hóa minh chứng trong file CSV
    var tmp = listFolderID.filter(function (tmprow) {
      return tmprow[0] == data[row][4]; // return tmprow[0] == data[row][2];
    })
    if (tmp.length == 0) {
      // NOTE: sửa riêng cho trường hợp khoa CNTT từ data[row][2] -> data[row][4]
      Logger.log("Không tìm thấy ID của folder " + data[row][2] + " trên cây thư mục minh chứng CTĐT") // giá trị gốc data[row][2]
      continue
    }
    var maHoaMCID = tmp[0][1];
    for (var col = 5; col < data[row].length; col++) { // Col bắt đầu từ 3, bỏ 3 cột Tiêu chuẩn và Tiêu chí, Mã hóa MC  
      proof = data[row][col];
      // Bỏ qua các ô rỗng trong BM03
      if (proof.toString().trim().length == 0)
        continue
      // Bỏ qua các ô không đúng format xxxxx trong BM03
      if (!isValid(proof)) {
        Logger.log("Lỗi minh chứng tại dòng " + (row + 1) + ", cột " + (col + 2) + ". Giá trị đọc được: " + proof)
        continue
      }
      // Lần lượt đọc từng id trở đi và tạo shortcut tại destFolder
      var srcFile = searchFiles(proof);
      if (srcFile == null) {
        Logger.log("### Không tìm thấy minh chứng " + proof + " trong ngân hàng minh chứng TQA.");
        sheet.getRange(row+1, col+1).setBackground("blue") // gán màu xanh cho các ô bị lỗi
        continue;
      }
      createShortcutv2(srcFile, proof, maHoaMCID);
      Logger.log("--> Đã tạo shortcut MC " + proof + " trong folder " + data[row][4]) //data[row][2]
    }
  }
}

// Function kiểm tra format mã minh chứng có đúng 5 số?
function isValid(value) {
  var regex = /^\d{5}$/;
  return regex.test(value);
}

// Hàm kiểm tra tồn tại của BM03, file CSV và format BM03 đúng cú pháp
function checkData() {
  var result = true
  var tmp = DriveApp.getFolderById(bm3Folder).getFilesByName("BM03_" + tenCTDT);
  // Kiểm tra file biểu mẫu 03 có tồn tại?
  if (!tmp.hasNext()) {
    Logger.log("Không tìm thấy biểu mẫu 03: BM03_" + tenCTDT + " tại folder Output BM3");
    return false
  }
  var bm03 = tmp.next();
  // Kiểm tra xem còn file biểu mẫu 03 bị duplicate không?
  if (tmp.hasNext()) {
    Logger.log("Có nhiều hơn 1 file biểu mẫu BM03_" + tenCTDT);
    return false
  }
  // Kiểm tra tenCTDT_ListFolderID.csv có tồn tại?
  tmp = DriveApp.getFolderById(dataFolderID).getFilesByName(csvListFolderID);
  if (!tmp.hasNext()) {
    Logger.log("Không tìm thấy file ListFolderID, vui lòng chạy script 03.LayDanhSachFolderID trước");
    return false
  }

  // Đọc dữ liệu trong file BM03
  var sheet = SpreadsheetApp.openById(bm03.getId()).getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();
  // Lặp qua từng dòng trong biểu mẫu 3, bỏ qua dòng header
  for (var row = 1; row < data.length; row++) { // Row bắt đầu từ 1, bỏ dòng header
    for (var col = 5; col < data[row].length; col++) { // Col bắt đầu từ 3, bỏ 3 cột Tiêu chuẩn và Tiêu chí, Mã hóa MC
      proof = data[row][col];
      if (proof.toString().trim().length == 0) // Nếu dữ liệu trong ô rỗng thì bỏ qua
        continue
      if (!isValid(proof)) {
        Logger.log("Lỗi minh chứng tại dòng " + (row + 1) + ", cột " + (col + 2) + ". Giá trị đọc được: " + proof)
        sheet.getRange(row+1, col+1).setBackground("red")
        result = false
        continue
      }
    }
  }
  return result
}

// Tìm ID của file MC trong Ngân hàng minh chứng dựa vào filename
function searchFiles(fileName) {
  var folder = DriveApp.getFolderById(proofFolder); // folder Id
  var result = folder.searchFiles('title contains "' + fileName + '"');
  while (result.hasNext()) {
    return result.next().getId();
  }
  return null;
}

// Tạo shortcut phiên bản 2
function createShortcutv2(srcFileID, shorcutName, destFolderID) {
  var folder = DriveApp.getFolderById(destFolderID);
  // Kiểm tra trong destFolderID đã có shortcut chưa, nếu có thì không tạo nữa để tránh bị duplicate
  var result = folder.searchFiles('title contains "' + shorcutName + '"');
  if (result.hasNext()) {
    Logger.log("Đã tồn tại shortcut " + shorcutName + " trong folder " + folder.getName());
    return;
  }
  var shortcut = DriveApp.createShortcut(srcFileID);
  folder.addFile(shortcut);
}
