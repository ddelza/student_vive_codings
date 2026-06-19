var GAS_URL = "https://script.google.com/macros/s/AKfycbzyeRDa2Swq71UlKMysPkDRxviqf6qTHNAJvwVRZxbT0d5KKn9_H2ehU8WU8xjjUzNeoQ/exec";
var SHEET_ID = "19MC5UCVgJWCNTDk83g0Af30iUkDX7s6pXwZGWIHu6-w";

function doGet(e) {
  var page   = (e && e.parameter) ? e.parameter.page   : '';
  var action = (e && e.parameter) ? e.parameter.action : '';

  // GitHub Pages → GAS JSON API
  if (action === 'getProjects') {
    return ContentService
      .createTextOutput(JSON.stringify(getStudentProjects()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'gallery') {
    return HtmlService.createHtmlOutputFromFile('Gallery')
        .setTitle('학생 작품 갤러리')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('학생 작품 제출 시스템')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// GitHub Pages fetch() 요청 처리 (Content-Type: text/plain → preflight 없음)
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;

    if (action === 'submitProject') {
      result = submitProjectData(data);
    } else if (action === 'addLike') {
      var updatedLikes = addLike(data.rowIndex, data.likerName);
      result = { success: true, likes: updatedLikes };
    } else {
      result = { success: false, message: '알 수 없는 action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 1. 작품 제출 (GAS 내부: google.script.run 경유)
function processForm(formObject) {
  try {
    var sheet = getSheet();
    ensureHeader(sheet);
    var fileUrl = saveFile(formObject.txtFile);
    sheet.appendRow([
      new Date(), formObject.grade, formObject.classNo, formObject.number,
      formObject.studentName, formObject.projectName,
      formObject.codeText || "파일 제출로 대체됨", fileUrl, ""
    ]);
    return { success: true, message: "성공적으로 제출되었습니다! 수고하셨습니다." };
  } catch (error) {
    return { success: false, message: "오류가 발생했습니다: " + error.toString() };
  }
}

// 1-b. 작품 제출 (GitHub Pages fetch() 경유 — 파일은 base64로 전달)
function submitProjectData(data) {
  try {
    var sheet = getSheet();
    ensureHeader(sheet);

    var fileUrl = "없음";
    if (data.fileBase64 && data.fileName) {
      var decoded  = Utilities.base64Decode(data.fileBase64);
      var blob     = Utilities.newBlob(decoded, 'text/plain', data.fileName);
      var folders  = DriveApp.getFoldersByName("학생작품제출_TXT파일");
      var folder   = folders.hasNext() ? folders.next() : DriveApp.createFolder("학생작품제출_TXT파일");
      var file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(), data.grade, data.classNo, data.number,
      data.studentName, data.projectName,
      data.codeText || "파일 제출로 대체됨", fileUrl, ""
    ]);
    return { success: true, message: "성공적으로 제출되었습니다! 수고하셨습니다." };
  } catch (error) {
    return { success: false, message: "오류가 발생했습니다: " + error.toString() };
  }
}

// 2. 갤러리 데이터 반환
function getStudentProjects() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var projects = [];

  for (var i = 1; i < data.length; i++) {
    var codeData  = data[i][6];
    var isHtmlCode = codeData && codeData !== "파일 제출로 대체됨" && codeData !== "";
    projects.push({
      rowIndex:    i + 1,
      studentInfo: data[i][1] + "학년 " + data[i][2] + "반 " + data[i][3] + "번 " + data[i][4],
      projectName: data[i][5],
      codeText:    codeData,
      fileUrl:     data[i][7],
      isHtmlCode:  isHtmlCode,
      likes:       data[i][8] ? data[i][8].toString() : ""
    });
  }
  return projects.reverse();
}

// 3. 좋아요 추가
function addLike(rowIndex, likerName) {
  var sheet = getSheet();
  if (sheet.getRange(1, 9).getValue() === "") {
    sheet.getRange(1, 9).setValue("좋아요 명단");
  }
  var currentLikes = sheet.getRange(rowIndex, 9).getValue();
  var likesArray   = currentLikes ? currentLikes.toString().split(',') : [];
  var isAlreadyLiked = likesArray.some(function(name) {
    return name.replace(/\s/g, '') === likerName.replace(/\s/g, '');
  });
  if (!isAlreadyLiked) {
    likesArray.push(likerName);
    var updatedLikes = likesArray.join(',');
    sheet.getRange(rowIndex, 9).setValue(updatedLikes);
    return updatedLikes;
  } else {
    throw new Error("이미 좋아요를 누르셨습니다.");
  }
}

// ── 내부 유틸 ──
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["타임스탬프","학년","반","번호","이름","작품명","제출된 코드","첨부파일 링크","좋아요 명단"]);
  }
}

function saveFile(fileBlob) {
  if (!fileBlob || fileBlob.size === 0) return "없음";
  var folders = DriveApp.getFoldersByName("학생작품제출_TXT파일");
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder("학생작품제출_TXT파일");
  var file    = folder.createFile(fileBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
