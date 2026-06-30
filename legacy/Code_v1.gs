/**
 * 일매출 관리 시스템 - Apps Script 모음
 * 스프레드시트 확장 메뉴(Apps Script 편집기)에 이 파일 내용을 붙여넣고 사용한다.
 * Extensions > Apps Script 에서 새 스크립트 파일로 추가.
 */

const SHEET_STORES = 'stores';
const SHEET_SALES_LOGS = 'sales_logs';
const PROP_STORE_ID_ITEM_ID = 'STORE_ID_ITEM_ID'; // createSalesForm() 실행 후 자동 저장됨

/**
 * 실제 Google Form을 코드로 생성하고, 응답 대상을 이 스프레드시트로 연결한 뒤,
 * 응답 탭 이름을 sales_logs로 맞추고 G/H/I 수식까지 다시 채워준다.
 * AI는 구글 로그인을 할 수 없으므로, 이 스크립트를 동진님 계정에서 한 번 실행하면
 * 동진님 계정 권한으로 실제 Form이 만들어진다 (로그인 위임이 필요 없는 방식).
 *
 * 실행 전 주의: sales_logs 탭에 테스트 데이터가 있으면 백업 후 비우고 실행할 것.
 * (응답 탭이 새로 생성되면서 기존 sales_logs를 대체하기 때문)
 *
 * 메뉴: 일매출관리 > 0. Google Form 실제로 생성하기
 */
function createSalesForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldLogs = ss.getSheetByName(SHEET_SALES_LOGS);
  if (oldLogs && oldLogs.getLastRow() > 1) {
    throw new Error('sales_logs에 데이터가 있습니다. 백업 후 헤더만 남기고 다시 실행하세요.');
  }

  const form = FormApp.create('일매출 입력');
  form.setDescription('매장별 일매출을 입력해주세요. 제출 후 수정은 불가합니다.');
  form.setAllowResponseEdits(false);
  form.setConfirmationMessage('오늘 매출이 정상적으로 등록됐습니다.');

  const storeIdItem = form.addTextItem().setTitle('매장코드').setRequired(true);
  form.addDateItem().setTitle('영업일').setRequired(true);

  const salesPage = form.addPageBreakItem().setTitle('매출 입력');
  form.addTextItem().setTitle('일 총매출 (원)').setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumber().build());

  const memoSharedPage = form.addPageBreakItem().setTitle('특이사항');
  form.addParagraphTextItem().setTitle('특이사항').setRequired(false);

  // "영업 여부" 항목은 분기 대상 페이지보다 먼저 추가된 페이지들을 참조해야 하므로
  // 마지막에 추가하고, 두 선택지 모두 동일한 "특이사항" 페이지로 합류시킨다.
  const statusItem = form.addMultipleChoiceItem();
  statusItem.setTitle('영업 여부').setRequired(true);
  statusItem.setChoices([
    statusItem.createChoice('영업', salesPage),
    statusItem.createChoice('휴무', memoSharedPage),
  ]);
  // 항목 순서를 매장코드 → 영업일 → 영업여부 → (분기) 순으로 맞추기 위해 재배치
  form.moveItem(statusItem.getIndex(), 2);

  if (oldLogs) ss.deleteSheet(oldLogs);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const newLogs = ss.getSheets().find(s => s.getName() !== SHEET_SALES_LOGS
    && /^(Form Responses|설문지 응답)/.test(s.getName()));
  if (!newLogs) throw new Error('응답 탭을 찾지 못했습니다. 수동으로 sales_logs로 이름을 바꿔주세요.');
  newLogs.setName(SHEET_SALES_LOGS);

  newLogs.getRange('G1:K1').setValues([['log_id', '업체명', '중복여부', '확인여부', '수정메모']]);
  newLogs.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","","L"&TEXT(ROW(A2:A)-1,"0000")))');
  newLogs.getRange('H2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,stores!$A:$C,3,0),"store_id 오류")))');
  newLogs.getRange('I2').setFormula(
    '=MAP(B2:B1000,C2:C1000,ROW(B2:B1000),LAMBDA(store,date,rownum,' +
    'IF(OR(store="",date=""),FALSE,COUNTIFS($B$2:INDEX($B:$B,rownum),store,' +
    '$C$2:INDEX($C:$C,rownum),date)>1)))'
  );
  newLogs.getRange('J2').setValue(false);

  PropertiesService.getScriptProperties().setProperty(PROP_STORE_ID_ITEM_ID, String(storeIdItem.getId()));
  PropertiesService.getScriptProperties().setProperty('FORM_PUBLISHED_URL', form.getPublishedUrl());
  PropertiesService.getScriptProperties().setProperty('FORM_EDIT_URL', form.getEditUrl());

  generateStoreLinks();

  SpreadsheetApp.getUi().alert(
    'Form 생성 완료\n응답 URL: ' + form.getPublishedUrl() +
    '\n편집 URL: ' + form.getEditUrl() +
    '\nstores!H열에 매장별 pre-filled 링크가 자동으로 채워졌습니다.'
  );
}

/**
 * createSalesForm()이 만든 실제 Form을 기준으로 매장별 pre-filled 링크를
 * stores 탭 H열에 채운다. Form.createResponse()로 실제 prefilled URL을 만들기 때문에
 * 가짜 URL이 아니라 클릭하면 바로 해당 매장코드가 채워진 채 입력 화면이 열린다.
 * 메뉴: 일매출관리 > 1. 매장별 입력링크 생성
 */
function generateStoreLinks() {
  const itemId = PropertiesService.getScriptProperties().getProperty(PROP_STORE_ID_ITEM_ID);
  if (!itemId) {
    throw new Error('먼저 "0. Google Form 실제로 생성하기"를 실행하세요.');
  }
  const editUrl = PropertiesService.getScriptProperties().getProperty('FORM_EDIT_URL');
  const form = FormApp.openByUrl(editUrl);
  const storeIdItem = form.getItemById(Number(itemId)).asTextItem();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const links = ids.map(([storeId]) => {
    if (!storeId) return [''];
    const response = form.createResponse();
    response.withItemResponse(storeIdItem.createResponse(storeId));
    return [response.toPrefilledUrl()];
  });
  sheet.getRange(2, 8, links.length, 1).setValues(links);
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
    .addItem('0. Google Form 실제로 생성하기', 'createSalesForm')
    .addItem('1. 매장별 입력링크 다시 생성', 'generateStoreLinks')
    .addItem('2. 미제출 매장 새로고침(스크립트 방식)', 'refreshMissingCheck')
    .addToUi();
}
