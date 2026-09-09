/**
 * 다트비 스터디 인증 — Google Apps Script 백엔드
 *
 * 두 곳에 씁니다:
 *  1) "log" 시트 — 제출 1건당 1행. 응답자/진행일자/사진 등 모든 상세
 *     정보를 그대로 남깁니다. HTML의 "인증 내역 확인" 화면(doGet)은
 *     이 시트만 읽습니다.
 *  2) "출석현황" 시트 — 학회에서 이미 쓰던 그 시트(스터디 x 주차
 *     매트릭스)입니다. 제출이 들어오면 해당 스터디 행 / 해당 주차
 *     블록의 대면·참·결·사진 인증·지연 인증 칸을 자동으로 채웁니다.
 *     운영진은 기존에 보던 이 화면을 그대로 보면 됩니다.
 *
 * ── 설정 방법 ──
 * 1. script.google.com → 새 프로젝트 → 이 코드 전체를 붙여넣기
 * 2. 아래 SHEET_ID, PHOTO_FOLDER_ID 값을 본인 것으로 교체
 *    - SHEET_ID: 스프레드시트 URL 중 .../d/여기부분/edit 의 "여기부분"
 *      ("2026-2 정규스터디 시트"를 업로드/붙여넣은 그 스프레드시트,
 *      탭 이름은 반드시 "출석현황" 그대로 유지)
 *    - PHOTO_FOLDER_ID: 인증 사진을 모아둘 구글드라이브 폴더 URL 중
 *      .../folders/여기부분 의 "여기부분"
 * 3. "log" 탭은 없으면 스크립트가 알아서 만듭니다 (직접 안 만들어도 됨).
 * 4. 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 계정: 나(Me)
 *    - 액세스 권한이 있는 사용자: 전체
 * 5. 배포 후 나오는 웹 앱 URL을 HTML 파일의 APPS_SCRIPT_URL에 붙여넣기
 * 6. 코드를 수정했다면 "새 배포"가 아니라 기존 배포를
 *    "배포 관리 → 수정 → 새 버전"으로 갱신해야 같은 URL이 유지됩니다.
 */

const SHEET_ID          = "1BryEAzUtDGQpi8eMNRaUYigrJE82xAZ7DbyBz0n5Nyo";
const PHOTO_FOLDER_ID   = "1cVmDfFoiAk7J2-Mt4qd1gpISPIS2nAPc";
const STATUS_SHEET_NAME = "출석현황";
const LOG_SHEET_NAME    = "log";

// "출석현황" 시트의 고정 레이아웃 (2026-2 정규스터디 시트 기준)
const STATUS_HEADER_ROW  = 6;  // No./유형/스터디명/대면/총/참/결/사진 인증/지연 인증 행
const STATUS_FIRST_ROW   = 7;  // 첫 번째 스터디 데이터 행
const STATUS_NAME_COL    = 3;  // "스터디명" 열 (C)
const STATUS_WEEK1_COL   = 4;  // 1주차 블록 시작 열 (D)
const STATUS_BLOCK_WIDTH = 6;  // 블록 하나당 열 개수: 대면,총,참,결,사진 인증,지연 인증
// 블록 내 오프셋 (0-based, STATUS_WEEK1_COL 기준)
const OFFSET = { 대면: 0, 총: 1, 참: 2, 결: 3, 사진인증: 4, 지연인증: 5 };

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    let photoUrl = "";
    if (data.photoBase64) {
      photoUrl = _savePhoto(data.photoBase64, data.photoName);
    }

    _appendLog(data, photoUrl);
    _updateStatusMatrix(data, photoUrl);

    return _json({ status: "success" });
  } catch (err) {
    return _json({ status: "error", message: String(err) });
  }
}

function doGet(e) {
  try {
    const sheet = _getLogSheet();
    const values = sheet.getDataRange().getValues();
    values.shift(); // 헤더 행 제거

    const records = values
      .filter(row => row.some(cell => cell !== "" && cell !== null))
      .map(row => ({
        submitTimestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
        generation:      row[1],
        respondent:      row[2],
        studyName:       row[3],
        studyDate:       _fmtDate(row[4]),
        round:           row[5],
        mode:            row[6],
        participants:    row[7],
        photoUrl:        row[8],
        isLate:          String(row[9]).toUpperCase() === "TRUE",
        submitDate:      _fmtDate(row[10]),
      }));

    return _json({ status: "success", records: records });
  } catch (err) {
    return _json({ status: "error", message: String(err) });
  }
}

// ── "log" 시트: 제출 1건당 1행, 상세 기록 전체 보관 ──
function _getLogSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow([
      "제출시각", "기수", "응답자", "스터디명", "진행일자", "주차",
      "진행방식", "참여인원", "사진URL", "지연여부", "제출일",
    ]);
  }
  return sheet;
}

function _appendLog(data, photoUrl) {
  _getLogSheet().appendRow([
    new Date(),
    data.generation || "",
    data.respondent || "",
    data.studyName || "",
    data.studyDate || "",
    data.round || "",
    data.mode || "",
    data.participants || "",
    photoUrl,
    data.isLate ? "TRUE" : "FALSE",
    data.submitDate || "",
  ]);
}

// ── "출석현황" 매트릭스 시트: 해당 스터디 행 × 해당 주차 블록만 갱신 ──
function _updateStatusMatrix(data, photoUrl) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(STATUS_SHEET_NAME);
  if (!sheet) throw new Error(`"${STATUS_SHEET_NAME}" 시트를 찾을 수 없습니다.`);

  const row = _findStudyRow(sheet, data.studyName);
  if (row === -1) throw new Error(`"${data.studyName}" 스터디를 출석현황 시트에서 찾을 수 없습니다.`);

  const weekNum = _parseWeekNum(data.round);
  if (!weekNum) throw new Error(`주차 정보를 인식할 수 없습니다: ${data.round}`);

  const blockStart = STATUS_WEEK1_COL + (weekNum - 1) * STATUS_BLOCK_WIDTH;

  // 대면/비대면
  sheet.getRange(row, blockStart + OFFSET.대면)
    .setValue(data.mode === "offline" ? "대면" : "비대면");

  // 참석 인원 + 결석 인원(총 - 참, "총" 값이 이미 채워져 있다는 전제)
  sheet.getRange(row, blockStart + OFFSET.참).setValue(Number(data.participants) || 0);
  const totalCell = sheet.getRange(row, blockStart + OFFSET.총);
  const total = Number(totalCell.getValue());
  if (total) {
    sheet.getRange(row, blockStart + OFFSET.결).setValue(Math.max(total - Number(data.participants || 0), 0));
  }

  // 사진 인증 — 응답자 이름이 보이는 링크로 (클릭하면 사진으로 이동)
  const label = (data.respondent || "인증사진").replace(/"/g, '""');
  sheet.getRange(row, blockStart + OFFSET.사진인증)
    .setFormula(`=HYPERLINK("${photoUrl}","${label}")`);

  // 지연 인증
  sheet.getRange(row, blockStart + OFFSET.지연인증)
    .setValue(data.isLate ? "지연" : "");
}

function _findStudyRow(sheet, studyName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < STATUS_FIRST_ROW) return -1;
  const names = sheet.getRange(STATUS_FIRST_ROW, STATUS_NAME_COL, lastRow - STATUS_FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === String(studyName).trim()) {
      return STATUS_FIRST_ROW + i;
    }
  }
  return -1;
}

function _parseWeekNum(roundLabel) {
  const m = String(roundLabel || "").match(/(\d+)\s*주차/);
  return m ? Number(m[1]) : null;
}

function _savePhoto(base64, fileName) {
  const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, "image/jpeg", fileName || ("photo_" + Date.now() + ".jpg"));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // 브라우저에서 <img>로 바로 렌더링되는 썸네일 URL 형태
  return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
}

function _fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
