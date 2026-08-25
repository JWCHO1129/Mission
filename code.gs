// =============================================
// Google Apps Script - 단교선 신청 웹앱
// 시트명: 단교선신청
// =============================================

const SHEET_NAME = '단교선신청';

/**
 * GET 요청 처리 (웹앱 배포 시 기본 진입점)
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('단교선 신청')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST 요청 처리 - 폼 데이터를 시트에 저장
 * @param {Object} e - 요청 이벤트 객체
 */
function doPost(e) {
  try {
    const params = e.parameter;

    const name          = params['성명']               || '';
    const schoolGrade   = params['학교/학년(또는 전공)'] || '';
    const phone         = params['휴대폰']              || '';
    const guardianPhone = params['보호자 폰번호']       || '';
    const program       = params['희망교육프로그램']    || '';

    appendToSheet(name, schoolGrade, phone, guardianPhone, program);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', message: '신청이 완료되었습니다.' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * google.script.run 방식으로 호출되는 함수
 * @param {Object} formData - 폼 데이터 객체
 */
function submitForm(formData) {
  try {
    appendToSheet(
      formData['성명'],
      formData['학교/학년(또는 전공)'],
      formData['휴대폰'],
      formData['보호자 폰번호'],
      formData['희망교육프로그램']
    );
    return { result: 'success', message: '신청이 완료되었습니다.' };
  } catch (err) {
    return { result: 'error', message: err.message };
  }
}

/**
 * 스프레드시트에 데이터 행 추가
 */
function appendToSheet(name, schoolGrade, phone, guardianPhone, program) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // 시트가 없으면 생성하고 헤더 추가
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['성명', '학교/학년(또는 전공)', '휴대폰', '보호자 폰번호', '희망교육프로그램']);
  }

  // 헤더가 없는 경우 첫 행 확인 후 추가
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['성명', '학교/학년(또는 전공)', '휴대폰', '보호자 폰번호', '희망교육프로그램']);
  }

  sheet.appendRow([name, schoolGrade, phone, guardianPhone, program]);
}
