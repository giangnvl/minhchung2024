//#################################################################################
var rootFolder = ""; // ID của folder chứa minh chứng từng CTĐT
var tenCTDT = ""
var bm1Folder = ""; // ID của folder chứa các file Excel biểu mẫu 01
var inputFileName = "BM01_" + tenCTDT + ".xlsx";  // Tên file Excel biểu mẫu 01
var outputFileName = "BM01_" + tenCTDT; // Tên file CSV được xuất ra, có format BM01_TenCTDT.csv
//#################################################################################

let listRootCriteria = []; // Danh sách ID của từng folder theo từng tiêu chuẩn: Tieu chuan 1; Tieu chuan 2

function main(){
  convertExcelToCSV()
  createStructure(); // Tao cau truc 11 thu muc tuong ung 11 tieu chuan
  readData(1); // Đọc file CSV tạm, bỏ dòng header
  Drive.Files.trash(DriveApp.getFolderById(bm1Folder).getFilesByName(outputFileName+".csv").next().getId()); // Xóa file CSV tạm
}

// Fucntion đọc file Excel biểu mẫu 01 -> file CSV
function convertExcelToCSV(){
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
    //parents: [{id: DriveApp.getFolderById(bm3Folder).getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var spreadsheet = Drive.Files.insert(config, blob);
  // Đọc dữ liệu và ghi vào data[]
  var sheet = SpreadsheetApp.openById(spreadsheet.id).getSheets()[0];
  data = sheet.getRange("A1:B51").getValues()

  // Mở Google Sheets và chuyển đổi thành CSV
  var csvString = '';
  for(var row = 0; row<data.length; row++)
    csvString+=data[row][0] + ";" + data[row][1] + '\n'
  config = {
    title: outputFileName+".csv",
    parents: [{ id: bm1Folder }],
    mimeType: MimeType.CSV
  };
  Drive.Files.insert(config, Utilities.newBlob(csvString)); // Tạo file CSV
  Drive.Files.trash(spreadsheet.id); // Xóa file Google Sheet tạm
}

// Tạo cấu trúc thư mục chứa minh chứng trong rootFolder, gồm 11 folders.
function createStructure(){
  var root = DriveApp.getFolderById(rootFolder);
  listRootCriteria.push("SampleDataForPosition0");
  for(var i=1; i<=11; i++){
    var childFolder = root.createFolder("Tieu chuan " + i);
    listRootCriteria.push(childFolder.getId());
  }
}

// Tạo các thư mục mã hoá minh chứng
function createFolders(criteria, numofFolders){ // createFolders("3.1", 5)
  var data = criteria.split(".");
  Logger.log("Đang xử lý: " + data[0] + "." + data[1]);
  var root = DriveApp.getFolderById(listRootCriteria[data[0]]); // Chọn thư mục Tiêu chuẩn x làm root
  Logger.log("Chọn thư mục root: " + root);
  var criteriaFolder = root.createFolder("Tieu chi " + criteria);
  var preindex = data[0] < 10 ? "0" + data[0] : data[0]; 
  for(var i=1; i<=numofFolders; i++){
    var index = i<10 ? "0"+i : i;
    var folderName =  "H" + data[0] + "." + preindex + ".0" + data[1] + "." + index;// 2.1 -> H2.02.01.index
    //var folderName =  "H1" + "." + preindex + ".0" + data[1] + "." + index;// 2.1 -> H2.02.01.index
    Logger.log("Đang tạo thư mục: " + folderName);
    var tmp = criteriaFolder.createFolder(folderName);
  }
}

// Đọc file dữ liệu CSV
function readData(beginRow) {
  var file = DriveApp.getFolderById(bm1Folder).getFilesByName(outputFileName+".csv")
  if (!file.hasNext()){
    Logger.log("Không thể mở được file CSV")
    return false;
  }
  var data = Utilities.parseCsv(file.next().getBlob().getDataAsString(), ";");
  for (var i = beginRow; i < data.length; i++) {
    createFolders(data[i][0], data[i][1]);
  }
}
