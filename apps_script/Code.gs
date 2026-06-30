/**
 * 일매출 관리 시스템 - Apps Script 모음
 * 스프레드시트 확장 메뉴(Apps Script 편집기)에 이 파일 내용을 붙여넣고 사용한다.
 * Extensions > Apps Script 에서 새 스크립트 파일로 추가.
 */

const SHEET_STORES = 'stores';
const SHEET_SALES_LOGS = 'sales_logs';
const FORM_ENTRY_ID = 'YOUR_ENTRY_ID_HERE'; // Form 미리보기 URL에서 store_id 질문의 entry.XXXXXXXXX 값으로 교체
const FORM_BASE_URL = 'https://docs.google.com/forms/d/e/YOUR_FORM_ID/viewform'; // 실제 배포된 Form URL로 교체

/**
 * stores 탭 H열(입력링크)에 매장별 pre-filled Google Form 링크를 자동 생성한다.
 * 메뉴: 일매출관리 > 1. 매장별 입력링크 생성
 */
function generateStoreLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A열 store_id
  const links = ids.map(row => {
    const storeId = row[0];
    if (!storeId) return [''];
    const url = `${FORM_BASE_URL}?usp=pp_url&${FORM_ENTRY_ID}=${encodeURIComponent(storeId)}`;
    return [url];
  });
  sheet.getRange(2, 8, links.length, 1).setValues(links); // H열
  SpreadsheetApp.getUi().alert(`${links.length}개 매장 링크 생성 완료`);
}

/**
 * Google Form 응답이 sales_logs A~F열에 들어올 때마다 J열(확인여부)을
 * 기본값 FALSE로 채운다. 폼 제출 트리거(onFormSubmit)에 연결해서 사용.
 * 트리거 설정: 편집 > 현재 프로젝트의 트리거 > 추가 > onFormSubmit / 양식 제출 시
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SALES_LOGS);
  const range = e.range; // 새로 추가된 응답 행
  const row = range.getRow();

  // J열(확인여부)이 비어 있으면 기본값 FALSE
  const jCell = sheet.getRange(row, 10);
  if (jCell.getValue() === '') {
    jCell.setValue(false);
  }

  checkDuplicateAndNotify_(sheet, row);
}

/**
 * 새로 제출된 행이 같은 store_id + 영업일 기준으로 중복인지 확인하고,
 * 중복이면 관리자에게 알림(Spreadsheet 알림 + 선택적으로 이메일)을 보낸다.
 * I열(중복여부) 수식이 이미 TRUE/FALSE를 계산하므로 그 결과를 그대로 읽는다.
 */
function checkDuplicateAndNotify_(sheet, row) {
  const storeId = sheet.getRange(row, 2).getValue();
  const saleDate = sheet.getRange(row, 3).getValue();
  const isDup = sheet.getRange(row, 9).getValue(); // I열, 수식이 채워질 시간이 필요할 수 있음

  if (isDup === true) {
    const companyName = sheet.getRange(row, 8).getValue();
    const msg = `[중복 입력 감지] ${companyName} (${storeId}) - ${saleDate}`;
    Logger.log(msg);
    // 필요 시 이메일 알림 활성화:
    // MailApp.sendEmail('admin@example.com', '중복 입력 감지', msg);
  }
}

/**
 * missing_check 탭 A5의 HSTACK/LET 수식이 일부 Sheets 환경에서
 * 배열 조건(COUNTIFS에 범위 인자로 배열을 넣는 방식)을 지원하지 않아
 * 오류가 날 경우 사용하는 Apps Script 기반 대체 구현.
 * 메뉴: 일매출관리 > 2. 미제출 매장 새로고침(스크립트 방식)
 */
function refreshMissingCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storesSheet = ss.getSheetByName(SHEET_STORES);
  const logsSheet = ss.getSheetByName(SHEET_SALES_LOGS);
  const checkSheet = ss.getSheetByName('missing_check');

  const checkDate = checkSheet.getRange('B1').getValue();
  const checkDateStr = Utilities.formatDate(new Date(checkDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const storesLastRow = storesSheet.getLastRow();
  const storesData = storesSheet.getRange(2, 1, storesLastRow - 1, 8).getValues();
  // A store_id, B 담당컨설턴트, C 업체명, D 점주명, E 연락처, F 지역, G 상태, H 입력링크

  const logsLastRow = logsSheet.getLastRow();
  const submittedIds = new Set();
  if (logsLastRow >= 2) {
    const logsData = logsSheet.getRange(2, 2, logsLastRow - 1, 2).getValues(); // B store_id, C 영업일
    logsData.forEach(([storeId, saleDate]) => {
      if (!storeId || !saleDate) return;
      const d = Utilities.formatDate(new Date(saleDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (d === checkDateStr) submittedIds.add(storeId);
    });
  }

  const missingRows = storesData
    .filter(row => row[6] === 'active' && !submittedIds.has(row[0]))
    .map(row => [row[1], row[2], row[4], checkDateStr, row[7], row[6]]);

  // 5행부터 기존 결과 영역 초기화 후 다시 작성
  const clearRows = Math.max(checkSheet.getLastRow() - 4, 0);
  if (clearRows > 0) {
    checkSheet.getRange(5, 1, clearRows, 6).clearContent();
  }
  if (missingRows.length === 0) {
    checkSheet.getRange('A5').setValue('오늘 미제출 매장 없음');
  } else {
    checkSheet.getRange(5, 1, missingRows.length, 6).setValues(missingRows);
  }

  checkSheet.getRange('B2').setValue(submittedIds.size);
  checkSheet.getRange('D2').setValue(
    storesData.filter(r => r[6] === 'active').length - submittedIds.size
  );
}

/**
 * 스프레드시트를 열 때 커스텀 메뉴를 추가한다.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('일매출관리')
    .addItem('1. 매장별 입력링크 생성', 'generateStoreLinks')
    .addItem('2. 미제출 매장 새로고침(스크립트 방식)', 'refreshMissingCheck')
    .addToUi();
}
