// ==========================================
// 1. Database Setup (Run once)
// ==========================================
const SPREADSHEET_ID = "1UIFTJZ5vxw6MXh1mGMS4yQzdher9JpFyFF6q0KM9cew";

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function setupDatabase() {
  var ss = getSS();
  
  var sheetsConfig = {
    "Users": ["User ID", "Role", "Username", "Password", "Name"],
    "Courses": ["Course ID", "Name", "Description", "Number of Sessions", "Course Folder URL", "Course Folder ID"],
    "Groups": ["Group ID", "Course", "Group Name", "Group Code", "Instructor", "Start Date", "End Date", "Students"],
    "Students": ["Student ID", "Name", "Group Code", "Phone"],
    "Sessions": ["Session ID", "Course", "Session Number", "Topic"],
    "Assignments": ["Submission ID", "Student Name", "Course", "Group", "Session", "Upload Date", "Drive Folder", "Files", "Status"],
    "Materials": ["Material ID", "Course", "Week", "Title", "Type", "Size", "Description", "URL", "File ID", "Upload Date"],
    "Permissions": ["Material ID", "Group Code"],
    "Announcements": ["Announcement ID", "Title", "Content", "Date", "Group Code"],
    "Tracking": ["Tracking ID", "Course", "Group", "Session Number", "Session Topic", "What was explained", "Homework", "Required Files", "Notes", "Attendance Notes", "Date"],
    "ActivityLog": ["Log ID", "Date", "User", "Action", "Details"],
    "Settings": ["Key", "Value"]
  };
  
  for (var sheetName in sheetsConfig) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    var headers = sheetsConfig[sheetName];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  
  // Add default admin user if not exists
  var usersSheet = ss.getSheetByName("Users");
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow(["U-1", "Admin", "admin", "admin", "System Admin"]);
  }

  Logger.log("Database successfully set up!");
  return "Database Setup Complete";
}

function getSheetDataAsObjects(sheetName) {
  try {
    var sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return [];
    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    
    var headers = rows[0];
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = rows[i][j];
      }
      data.push(obj);
    }
    return data;
  } catch (e) {
    return [];
  }
}

// ==========================================
// 2. GET Requests (Fetch Data)
// ==========================================
function doGet(e) {
  var action = e.parameter.action;
  var responseData = {};

  try {
    if (action === 'getInitialData') {
      responseData = { 
        status: "success", 
        courses: getSheetDataAsObjects("Courses"),
        groups: getSheetDataAsObjects("Groups"),
        materials: getSheetDataAsObjects("Materials"),
        permissions: getSheetDataAsObjects("Permissions"),
        tracking: getSheetDataAsObjects("Tracking")
      };
    } else {
      responseData = { status: "success", message: "API is working. Please specify a valid action." };
    }
  } catch (error) {
    responseData = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 3. POST Requests (Send Data & Upload)
// ==========================================
function doPost(e) {
  var responseData = {};
  
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;

    if (action === 'submitTask') {
      responseData = processTaskSubmission(postData);
    } else if (action === 'uploadMaterial') {
      responseData = processMaterialUpload(postData);
    } else if (action === 'addCourse') {
      responseData = addRecord("Courses", postData.data);
    } else if (action === 'addGroup') {
      responseData = addRecord("Groups", postData.data);
    } else if (action === 'addTracking') {
      responseData = addRecord("Tracking", postData.data);
    } else {
      responseData = { status: "error", message: "Action not supported in POST" };
    }
  } catch (error) {
    responseData = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 4. Drive Logic
// ==========================================
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function processTaskSubmission(data) {
  var studentName = data.studentName || "Unknown Student";
  var groupCode = data.groupName || "Unknown Group";
  var sessionName = data.sessionName || "Unknown Session";
  
  // Find course from group
  var groups = getSheetDataAsObjects("Groups");
  var groupObj = groups.find(function(g) { return g["Group Code"] === groupCode; });
  var courseName = groupObj ? groupObj["Course"] : "General Course";
  
  // Assignments > Course > Group > Student Name > Session
  var rootFolder = getOrCreateFolder(DriveApp, "Assignments");
  var courseFolder = getOrCreateFolder(rootFolder, courseName);
  var groupFolder = getOrCreateFolder(courseFolder, groupCode);
  var studentFolder = getOrCreateFolder(groupFolder, studentName);
  var sessionFolder = getOrCreateFolder(studentFolder, sessionName);
  
  var uploadedFilesLog = [];
  var files = data.files || [];
  
  for (var i = 0; i < files.length; i++) {
    var fileData = files[i];
    var contentType = fileData.mimeType;
    var decodedData = Utilities.base64Decode(fileData.base64);
    var blob = Utilities.newBlob(decodedData, contentType);
    
    var fileExtension = fileData.filename.substring(fileData.filename.lastIndexOf("."));
    var customSuffix = files.length > 1 ? "_" + fileData.filename.split(".")[0] : "_" + sessionName;
    var finalFileName = studentName + "_" + groupCode + customSuffix + fileExtension;
    
    blob.setName(finalFileName);
    var file = sessionFolder.createFile(blob);
    uploadedFilesLog.push(finalFileName);
  }
  
  var sheet = getSS().getSheetByName("Assignments");
  var submissionId = "SUB-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  sheet.appendRow([
    submissionId,
    studentName,
    courseName,
    groupCode,
    sessionName,
    new Date(),
    sessionFolder.getUrl(),
    uploadedFilesLog.join(" , "),
    "Submitted"
  ]);
  
  return { status: "success", submissionId: submissionId };
}

function processMaterialUpload(data) {
  // Course > Materials > Week
  var rootFolder = getOrCreateFolder(DriveApp, "Courses");
  var courseFolder = getOrCreateFolder(rootFolder, data.courseName);
  var materialsFolder = getOrCreateFolder(courseFolder, "Materials");
  var weekFolder = getOrCreateFolder(materialsFolder, data.week);
  
  var fileData = data.file;
  var contentType = fileData.mimeType;
  var decodedData = Utilities.base64Decode(fileData.base64);
  var blob = Utilities.newBlob(decodedData, contentType);
  blob.setName(fileData.filename);
  var file = weekFolder.createFile(blob);
  
  var materialId = "MAT-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  var sheet = getSS().getSheetByName("Materials");
  sheet.appendRow([
    materialId,
    data.courseName,
    data.week,
    data.title,
    data.type,
    file.getSize(),
    data.description,
    file.getUrl(),
    file.getId(),
    new Date()
  ]);
  
  // Save Permissions
  var permSheet = getSS().getSheetByName("Permissions");
  var groups = data.groups || [];
  for(var i = 0; i < groups.length; i++) {
    permSheet.appendRow([materialId, groups[i]]);
  }
  
  return { status: "success", materialId: materialId, url: file.getUrl() };
}

// Helper to add a generic record
function addRecord(sheetName, recordData) {
  var sheet = getSS().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " not found");
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for(var i = 0; i < headers.length; i++) {
    // If column name is ID, auto generate if missing
    if (headers[i].indexOf("ID") !== -1 && !recordData[headers[i]]) {
      row.push("REC-" + Utilities.getUuid().substring(0, 8).toUpperCase());
    } else {
      row.push(recordData[headers[i]] || "");
    }
  }
  
  sheet.appendRow(row);
  return { status: "success" };
}