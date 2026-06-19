function doGet(e) {
  // 파라미터가 정상적으로 전달되었는지 안전하게 확인합니다.
  var page = (e && e.parameter) ? e.parameter.page : '';

  // 주소 끝에 ?page=gallery 가 붙어있으면 갤러리 화면(Gallery.html)을 렌더링
  if (page === 'gallery') {
    return HtmlService.createHtmlOutputFromFile('Gallery')
        .setTitle('학생 작품 갤러리')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // 그 외의 경우 기본 제출 폼 화면(Index.html)을 렌더링
  else {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('학생 작품 제출 시스템')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// 1. 학생들의 작품을 제출받아 스프레드시트에 저장하는 함수
function processForm(formObject) {
  try {
    var sheetId = "19MC5UCVgJWCNTDk83g0Af30iUkDX7s6pXwZGWIHu6-w";
    var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];

    // 헤더(제목 행)가 없는 경우 최초 생성 (I열에 '좋아요 명단' 추가)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["타임스탬프", "학년", "반", "번호", "이름", "작품명", "제출된 코드", "첨부파일 링크", "좋아요 명단"]);
    }

    var fileUrl = "없음";

    // 파일이 첨부된 경우 드라이브에 저장 처리
    if (formObject.txtFile && formObject.txtFile.size > 0) {
      var fileBlob = formObject.txtFile;
      var folderName = "학생작품제출_TXT파일";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var file = folder.createFile(fileBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    // 시트에 새 데이터 행 추가 (좋아요 컬럼은 처음엔 비워둠)
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

// 2. 갤러리에 띄울 학생 작품 데이터 목록을 가져오는 함수
function getStudentProjects() {
  var sheetId = "19MC5UCVgJWCNTDk83g0Af30iUkDX7s6pXwZGWIHu6-w";
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var projects = [];

  // 2번째 줄(인덱스 1)부터 반복하여 데이터 추출
  for (var i = 1; i < data.length; i++) {
    var codeData = data[i][6];
    var isHtmlCode = codeData && codeData !== "파일 제출로 대체됨" && codeData !== "";

    projects.push({
      rowIndex: i + 1, // 시트의 실제 행 번호 (데이터 업데이트 시 필요)
      studentInfo: data[i][1] + "학년 " + data[i][2] + "반 " + data[i][3] + "번 " + data[i][4],
      projectName: data[i][5],
      codeText: codeData,
      fileUrl: data[i][7],
      isHtmlCode: isHtmlCode,
      likes: data[i][8] ? data[i][8].toString() : "" // I열(9번째 열)의 좋아요 내역
    });
  }

  // 최신 제출물이 맨 위로 오게 배열을 뒤집어서 반환
  return projects.reverse();
}

// 3. 좋아요 버튼 클릭 시 시트에 이름을 추가하는 함수
function addLike(rowIndex, likerName) {
  var sheetId = "19MC5UCVgJWCNTDk83g0Af30iUkDX7s6pXwZGWIHu6-w";
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];

  // I열(9번째 열)의 헤더가 비어있으면 이름 지정
  if (sheet.getRange(1, 9).getValue() === "") {
    sheet.getRange(1, 9).setValue("좋아요 명단");
  }

  // 해당 학생의 현재 좋아요 목록 가져오기
  var currentLikes = sheet.getRange(rowIndex, 9).getValue();
  var likesArray = currentLikes ? currentLikes.toString().split(',') : [];

  // 중복 좋아요 방지 (공백 제거 후 이름만 비교)
  var isAlreadyLiked = likesArray.some(function(name) {
    return name.replace(/\s/g, '') === likerName.replace(/\s/g, '');
  });

  if (!isAlreadyLiked) {
    likesArray.push(likerName);
    var updatedLikes = likesArray.join(',');

    // 시트에 업데이트된 이름 목록 저장
    sheet.getRange(rowIndex, 9).setValue(updatedLikes);

    return updatedLikes; // 성공 시 프론트엔드로 업데이트된 문자열 반환
  } else {
    // 이미 있는 이름이면 에러 발생시켜 프론트엔드에 알림
    throw new Error("이미 좋아요를 누르셨습니다.");
  }
}
