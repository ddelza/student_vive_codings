var GAS_URL = "https://script.google.com/macros/s/AKfycbzyeRDa2Swq71UlKMysPkDRxviqf6qTHNAJvwVRZxbT0d5KKn9_H2ehU8WU8xjjUzNeoQ/exec";
var SHEET_ID = "19MC5UCVgJWCNTDk83g0Af30iUkDX7s6pXwZGWIHu6-w";
var PADLET_ROSTER_SHEET_ID = "1yvm2kKox1CenK2VLJ43HOF3Ik4spWxd9D87A6upHClU"; // 전교생 학번/성명 명단 (padlet 로그인 검증용)
var PADLET_ROSTER_GID = 62759410; // 명단이 들어있는 탭의 gid

var CONTEST_ROSTER_SHEET_ID = "1FZxiI3H_Yw1MD0CKfWKsenRP4D0PWA2pdVReTYNvGQE"; // 학번/성명/구글계정 명단 (대회 제출페이지 구글 로그인 검증용)
var CONTEST_GOOGLE_CLIENT_ID = "YOUR_GOOGLE_CLIENT_ID"; // Google Cloud Console에서 발급받은 OAuth 클라이언트 ID로 교체할 것

function doGet(e) {
  var page   = (e && e.parameter) ? e.parameter.page   : '';
  var action = (e && e.parameter) ? e.parameter.action : '';

  // GitHub Pages → GAS JSON API
  if (action === 'getProjects') {
    return ContentService
      .createTextOutput(JSON.stringify(getStudentProjects()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getPadletPosts') {
    return ContentService
      .createTextOutput(JSON.stringify(getPadletPosts(e.parameter.studentId, e.parameter.studentName)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'verifyPadletLogin') {
    return ContentService
      .createTextOutput(JSON.stringify({ valid: verifyPadletStudent_(e.parameter.studentId, e.parameter.studentName) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'verifyPadletAdmin') {
    return ContentService
      .createTextOutput(JSON.stringify({ valid: verifyPadletAdmin_(e.parameter.password) }))
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
    } else if (action === 'submitPadletPost') {
      result = submitPadletPost(data);
    } else if (action === 'editPadletPost') {
      result = editPadletPost(data);
    } else if (action === 'togglePadletReaction') {
      result = togglePadletReaction(data);
    } else if (action === 'addPadletComment') {
      result = addPadletComment(data);
    } else if (action === 'deletePadletComment') {
      result = deletePadletComment(data);
    } else if (action === 'togglePadletPin') {
      result = togglePadletPin(data);
    } else if (action === 'togglePadletDefaultCollapse') {
      result = togglePadletDefaultCollapse(data);
    } else if (action === 'deletePadletPost') {
      result = deletePadletPost(data);
    } else if (action === 'verifyContestLogin') {
      result = verifyContestLogin(data);
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

// ===================== padlet.html (디지털 시민성 바이브코딩 아이디어 패들렛) =====================
// 학번/이름은 별도 명단으로 검증하지 않고 그대로 받아 기록한다(전교 대상, 교사는 학번란에 0000 입력).
var PADLET_SHEET_NAME = '디지털시민성패들렛';
var PADLET_BOARDS = ['idea', 'complaint', 'ask', 'disclosure', 'tips', 'link'];
var PADLET_TEACHER_ID = '0000';
var PADLET_ADMIN_PASSWORD = '7683101'; // 관리자 모드 비밀번호 (게시물 고정/전체 수정·삭제 권한)

// 1회성 수복 함수 — Apps Script 편집기에서 딱 한 번 직접 실행하면 됨.
// 교사 게시물의 학번('0000')이 구글시트에 저장될 때 숫자로 자동 변환되면서
// 앞의 0이 사라져 0으로 저장된 기존 행들을 다시 '0000' 문자로 되돌린다.
// (이 문제 때문에 교사 게시물 상단고정, 수정, 삭제가 "본인 글이 아닙니다"로 실패할 수 있음)
function fixPadletTeacherIds() {
  var sheet = getPadletSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '수정할 행이 없습니다.';

  // 앞으로도 다시 숫자로 변환되지 않도록 학번 열 전체를 텍스트 서식으로 고정
  sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('@');

  var ids = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var fixed = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === '0') {
      sheet.getRange(i + 2, 3).setValue('0000');
      fixed++;
    }
  }
  Logger.log('학번 0 -> 0000 으로 복구된 행: ' + fixed + '개');
  return fixed;
}

function getPadletSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PADLET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PADLET_SHEET_NAME);
    sheet.appendRow(['게시ID', '제출시각', '작성자ID', '작성자이름', '파트', '작성자유형', '제목', '내용', '좋아요수', '반응기록JSON', '댓글JSON', '고정여부', '기본접기여부']);
  }
  return sheet;
}

function padletSafeJson_(str, fallback) {
  try {
    if (!str) return fallback;
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

// 학번+이름 조합을 식별키로 사용 — 교사는 전부 학번이 '0000'으로 같으므로,
// 이름까지 합쳐야 교사 여러 명이 서로의 좋아요/댓글 삭제 권한을 침범하지 않는다.
function padletAuthorKey_(id, name) {
  return String(id || '').trim() + '||' + String(name || '').trim();
}

function padletIsTeacher_(id) {
  return String(id || '').trim() === PADLET_TEACHER_ID;
}

function padletNormalize_(s) {
  return String(s || '').replace(/\s+/g, '');
}

// 관리자 모드 비밀번호 확인 (게시물 고정, 전체 수정/삭제 권한에 사용)
function verifyPadletAdmin_(password) {
  return String(password || '') === PADLET_ADMIN_PASSWORD;
}

// 명단 탭을 이름이 아니라 gid(탭 고유ID)로 찾음 — 탭 이름이 바뀌어도 안전하게 동작
function getPadletRosterSheet_() {
  var ss = SpreadsheetApp.openById(PADLET_ROSTER_SHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === PADLET_ROSTER_GID) return sheets[i];
  }
  return sheets[0];
}

// 학번+이름이 전교생 명단과 일치하는지 확인 (교사 계정 '0000'은 명단 검증 없이 통과)
function verifyPadletStudent_(studentId, studentName) {
  if (padletIsTeacher_(studentId)) return true;

  var id = String(studentId || '').trim();
  var name = padletNormalize_(studentName);
  if (!id || !name) return false;

  var sheet = getPadletRosterSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A열: 학번, B열: 성명
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      return padletNormalize_(data[i][1]) === name;
    }
  }
  return false;
}

// 게시ID로 행 번호를 찾을 때 ID열(A열)만 읽음 — 댓글이 쌓여 행이 커져도 조회 속도가 느려지지 않게 함
function findPadletRow_(sheet, postId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(postId)) return i + 2; // 1-based, 헤더행(1) 보정
  }
  return -1;
}

// 클라이언트(padlet.html)에서 호출: 전체 게시물 + 내 반응/소유 상태 포함해 반환
function getPadletPosts(myId, myName) {
  var sheet = getPadletSheet_();
  var data = sheet.getDataRange().getValues();
  var myKey = padletAuthorKey_(myId, myName);
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var reactions = padletSafeJson_(row[9], {});
    // 반응기록 JSON의 키는 "학번||이름" 형태(padletAuthorKey_) — 이름만 뽑아서
    // "누가 좋아요를 눌렀는지" 마우스오버로 보여줄 때 씀
    var likerNames = Object.keys(reactions).map(function (k) {
      var parts = k.split('||');
      return parts[1] || parts[0];
    });
    var comments = padletSafeJson_(row[10], []).map(function (c) {
      c.isTeacher = padletIsTeacher_(c.studentId);
      c.isMine = padletAuthorKey_(c.studentId, c.studentName) === myKey;
      return c;
    });
    posts.push({
      id: row[0],
      timestamp: Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      authorId: row[2],
      authorName: row[3],
      board: row[4],
      authorType: row[5],
      title: row[6],
      content: row[7],
      likeCount: Number(row[8]) || 0,
      likerNames: likerNames,
      myLiked: !!reactions[myKey],
      comments: comments,
      isTeacher: padletIsTeacher_(row[2]),
      isMine: padletAuthorKey_(row[2], row[3]) === myKey,
      pinned: row[11] === true || String(row[11]).toUpperCase() === 'TRUE',
      defaultCollapsed: row[12] === true || String(row[12]).toUpperCase() === 'TRUE'
    });
  }
  posts.sort(function (a, b) {
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1; // 관리자가 고정한 게시물이 항상 최상단
    if (a.isTeacher !== b.isTeacher) return a.isTeacher ? -1 : 1; // 그 다음 교사 게시물
    return a.id < b.id ? 1 : -1; // 그 외엔 최신 게시물이 위로
  });
  return { valid: true, posts: posts };
}

// 게시물 등록
function submitPadletPost(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();
  if (!studentId || !studentName) {
    return { success: false, message: '학번(또는 0000)과 이름을 입력해주세요.' };
  }
  if (!verifyPadletStudent_(studentId, studentName)) {
    return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
  }
  if (PADLET_BOARDS.indexOf(data.board) === -1) {
    return { success: false, message: '파트를 올바르게 선택해주세요.' };
  }
  if (!data.content || data.content.trim().length < 3) {
    return { success: false, message: '내용을 3자 이상 작성해주세요.' };
  }

  var isTeacher = padletIsTeacher_(studentId);
  var authorType = isTeacher ? '교사' : '학생';
  if (data.board === 'complaint') {
    if (['학생', '학부모', '교사'].indexOf(data.authorType) === -1) {
      return { success: false, message: '작성자 유형(학생/학부모/교사)을 선택해주세요.' };
    }
    authorType = data.authorType;
  }

  var sheet = getPadletSheet_();
  var id = 'P' + new Date().getTime() + Math.floor(Math.random() * 1000);
  // 학번(특히 교사용 '0000')이 숫자로 자동 변환되어 앞의 0이 사라지는 걸 막기 위해,
  // appendRow로 값이 들어갈 다음 행의 C열(학번)을 미리 텍스트 서식으로 고정해둔다.
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 3).setNumberFormat('@');
  sheet.appendRow([
    id, new Date(), studentId, studentName, data.board, authorType,
    (data.title || '').trim(), data.content.trim(), 0, '{}', '[]', false, false
  ]);
  return { success: true, id: id };
}

// 게시물 수정 — 본인(학번+이름 조합)이 작성한 게시물만 가능 (단, 관리자 비밀번호가 맞으면 누구 글이든 수정 가능)
function editPadletPost(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();
  var isAdmin = verifyPadletAdmin_(data.adminPassword);
  if (!isAdmin && !verifyPadletStudent_(studentId, studentName)) {
    return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
  }
  if (!data.content || data.content.trim().length < 3) {
    return { success: false, message: '내용을 3자 이상 작성해주세요.' };
  }

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  if (!isAdmin) {
    var ownerRow = sheet.getRange(rowNum, 3, 1, 2).getValues()[0];
    if (padletAuthorKey_(ownerRow[0], ownerRow[1]) !== padletAuthorKey_(studentId, studentName)) {
      return { success: false, message: '본인이 작성한 게시물만 수정할 수 있습니다.' };
    }
  }

  var board = sheet.getRange(rowNum, 5).getValue();
  if (board === 'complaint' && ['학생', '학부모', '교사'].indexOf(data.authorType) !== -1) {
    sheet.getRange(rowNum, 6).setValue(data.authorType);
  }
  sheet.getRange(rowNum, 7, 1, 2).setValues([[(data.title || '').trim(), data.content.trim()]]);
  return { success: true };
}

// 좋아요 토글 — 학번+이름 조합 1개당 게시물 1개에 좋아요 1개만 가능, 다시 누르면 취소
function togglePadletReaction(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();
  if (!verifyPadletStudent_(studentId, studentName)) {
    return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
  }

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  var row = sheet.getRange(rowNum, 9, 1, 2).getValues()[0];
  var likeCount = Number(row[0]) || 0;
  var reactions = padletSafeJson_(row[1], {});
  var key = padletAuthorKey_(studentId, studentName);

  var liked = !!reactions[key];
  if (liked) {
    delete reactions[key];
    likeCount--;
  } else {
    reactions[key] = true;
    likeCount++;
  }
  likeCount = Math.max(0, likeCount);

  sheet.getRange(rowNum, 9, 1, 2).setValues([[likeCount, JSON.stringify(reactions)]]);
  return { success: true, likeCount: likeCount, myLiked: !liked };
}

// 댓글/대댓글 추가
function addPadletComment(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();
  if (!verifyPadletStudent_(studentId, studentName)) {
    return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
  }
  if (!data.text || data.text.trim().length < 1) {
    return { success: false, message: '댓글 내용을 입력해주세요.' };
  }

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  var comments = padletSafeJson_(sheet.getRange(rowNum, 11).getValue(), []);
  var comment = {
    id: 'C' + new Date().getTime() + Math.floor(Math.random() * 1000),
    parentId: data.parentId || null,
    studentId: studentId,
    studentName: studentName,
    text: data.text.trim(),
    time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
  comments.push(comment);
  sheet.getRange(rowNum, 11).setValue(JSON.stringify(comments));

  var result = comments.map(function (c) {
    return {
      id: c.id, parentId: c.parentId, studentId: c.studentId, studentName: c.studentName,
      text: c.text, time: c.time,
      isTeacher: padletIsTeacher_(c.studentId),
      isMine: padletAuthorKey_(c.studentId, c.studentName) === padletAuthorKey_(studentId, studentName)
    };
  });
  return { success: true, comments: result };
}

// 본인(학번+이름 조합) 댓글/대댓글 삭제 — 대댓글이 달린 댓글을 지우면 대댓글도 함께 삭제
// (관리자 비밀번호가 맞으면 누구 댓글이든 삭제 가능)
function deletePadletComment(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();
  var isAdmin = verifyPadletAdmin_(data.adminPassword);
  if (!isAdmin && !verifyPadletStudent_(studentId, studentName)) {
    return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
  }
  var myKey = padletAuthorKey_(studentId, studentName);

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  var comments = padletSafeJson_(sheet.getRange(rowNum, 11).getValue(), []);
  var target = null;
  for (var i = 0; i < comments.length; i++) {
    if (comments[i].id === data.commentId) { target = comments[i]; break; }
  }
  if (!target) return { success: false, message: '댓글을 찾을 수 없습니다.' };
  if (!isAdmin && padletAuthorKey_(target.studentId, target.studentName) !== myKey) {
    return { success: false, message: '본인이 작성한 댓글만 삭제할 수 있습니다.' };
  }

  comments = comments.filter(function (c) {
    return c.id !== data.commentId && c.parentId !== data.commentId;
  });

  sheet.getRange(rowNum, 11).setValue(JSON.stringify(comments));

  var result = comments.map(function (c) {
    return {
      id: c.id, parentId: c.parentId, studentId: c.studentId, studentName: c.studentName,
      text: c.text, time: c.time,
      isTeacher: padletIsTeacher_(c.studentId),
      isMine: padletAuthorKey_(c.studentId, c.studentName) === myKey
    };
  });
  return { success: true, comments: result };
}

// 관리자 전용: 게시물 상단 고정/해제 토글
function togglePadletPin(data) {
  if (!verifyPadletAdmin_(data.adminPassword)) {
    return { success: false, message: '관리자 비밀번호가 일치하지 않습니다.' };
  }

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  var current = sheet.getRange(rowNum, 12).getValue();
  var pinned = !(current === true || String(current).toUpperCase() === 'TRUE');
  sheet.getRange(rowNum, 12).setValue(pinned);
  return { success: true, pinned: pinned };
}

// 관리자 전용: 특정 게시물을 기본적으로 접어서(제목/본문 숨김) 보여줄지 토글
// (게시판 종류와 상관없이 이 글 하나만 접어두고 싶을 때 사용)
function togglePadletDefaultCollapse(data) {
  if (!verifyPadletAdmin_(data.adminPassword)) {
    return { success: false, message: '관리자 비밀번호가 일치하지 않습니다.' };
  }

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  var current = sheet.getRange(rowNum, 13).getValue();
  var collapsed = !(current === true || String(current).toUpperCase() === 'TRUE');
  sheet.getRange(rowNum, 13).setValue(collapsed);
  return { success: true, defaultCollapsed: collapsed };
}

// 관리자 전용: 게시물 전체 삭제 (댓글 포함, 행 자체를 지움)
// 본인(학번+이름 조합)이 작성한 게시물은 스스로 삭제 가능. 관리자 비밀번호가 맞으면 누구 글이든 삭제 가능.
function deletePadletPost(data) {
  var isAdmin = verifyPadletAdmin_(data.adminPassword);
  var studentId = String(data.studentId || '').trim();
  var studentName = String(data.studentName || '').trim();

  var sheet = getPadletSheet_();
  var rowNum = findPadletRow_(sheet, data.postId);
  if (rowNum === -1) return { success: false, message: '게시물을 찾을 수 없습니다.' };

  if (!isAdmin) {
    if (!verifyPadletStudent_(studentId, studentName)) {
      return { success: false, message: '학번 또는 이름이 명단과 일치하지 않습니다.' };
    }
    var ownerRow = sheet.getRange(rowNum, 3, 1, 2).getValues()[0];
    if (padletAuthorKey_(ownerRow[0], ownerRow[1]) !== padletAuthorKey_(studentId, studentName)) {
      return { success: false, message: '본인이 작성한 게시물만 삭제할 수 있습니다.' };
    }
  }

  sheet.deleteRow(rowNum);
  return { success: true };
}

// ===================== contest.html (대회 제출페이지 구글 로그인) =====================

function contestNormalize_(s) {
  return String(s || '').replace(/\s+/g, '');
}

// 명단 시트: A열=학번, B열=성명, C열=구글계정 (전교생, 첫 번째 탭)
function getContestRosterSheet_() {
  return SpreadsheetApp.openById(CONTEST_ROSTER_SHEET_ID).getSheets()[0];
}

// 구글 로그인으로 받은 credential(JWT)이 진짜 구글이 발급한 것인지 서버에서 재확인하고 이메일을 꺼냄
function verifyGoogleCredential_(credential) {
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return null;

  var payload = JSON.parse(res.getContentText());
  if (payload.aud !== CONTEST_GOOGLE_CLIENT_ID) return null; // 우리 앱용으로 발급된 토큰이 아니면 거부

  return payload.email || null;
}

// 학번+이름+구글이메일이 명단의 같은 행에서 전부 일치해야 통과
function verifyContestLogin(data) {
  var studentId = String(data.studentId || '').trim();
  var studentName = contestNormalize_(data.studentName);
  if (!studentId || !studentName) {
    return { success: false, message: '학번과 이름을 입력해주세요.' };
  }
  if (!data.credential) {
    return { success: false, message: '구글 로그인 정보가 없습니다.' };
  }

  var email = verifyGoogleCredential_(data.credential);
  if (!email) {
    return { success: false, message: '구글 로그인 확인에 실패했습니다. 다시 시도해주세요.' };
  }

  var sheet = getContestRosterSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: '명단을 불러올 수 없습니다.' };

  var data2d = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A~C열
  for (var i = 0; i < data2d.length; i++) {
    var rowId = String(data2d[i][0] || '').trim();
    var rowName = contestNormalize_(data2d[i][1]);
    var rowEmail = String(data2d[i][2] || '').trim().toLowerCase();

    if (rowEmail && rowEmail === email.toLowerCase()) {
      if (rowId === studentId && rowName === studentName) {
        return { success: true, studentId: rowId, studentName: data2d[i][1] };
      }
      return { success: false, message: '입력한 학번 또는 이름이 구글 계정과 일치하지 않습니다.' };
    }
  }

  return { success: false, message: '명단에서 이 구글 계정을 찾을 수 없습니다.' };
}
