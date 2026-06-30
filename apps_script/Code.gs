/**
 * 일매출 관리 시스템 v2 - Apps Script
 *
 * v1과 다른 점: monthly_summary/missing_check의 집계와 sales_logs의 중복체크를
 * 더 이상 LET/HSTACK/MAP 같은 고급 시트 수식에 맡기지 않고, 전부 이 스크립트가
 * 계산해서 "값"으로 써넣는다. Form 제출 → 트리거 → 전체 재계산까지 자동으로 이어진다.
 *
 * 설치: 확장 프로그램 > Apps Script 편집기 → 이 파일 내용을 통째로 붙여넣고 저장.
 * 시작: 시트 새로고침 → 메뉴 [일매출관리 > 0. 전체 자동 설정] 실행.
 */

const SHEET_STORES = 'stores';
const SHEET_SALES_LOGS = 'sales_logs';
const SHEET_MONTHLY = 'monthly_summary';
const SHEET_MISSING = 'missing_check';
const SHEET_DASHBOARD = 'dashboard';
const PROP_STORE_ID_ITEM_ID = 'STORE_ID_ITEM_ID';
const PROP_FORM_EDIT_URL = 'FORM_EDIT_URL';
const ADMIN_EMAIL = ''; // 비워두면 이메일 발송 안 함. 채우면 매일 미제출 매장 요약 발송.

// ===========================================================================
// 0. 전체 자동 설정 (Form 생성 + 트리거 설치) — 처음 한 번만 실행
// ===========================================================================
function setupAll() {
  if (!PropertiesService.getScriptProperties().getProperty(PROP_FORM_EDIT_URL)) {
    createSalesForm();
  }
  installTriggers();
  refreshAll();
  SpreadsheetApp.getUi().alert('전체 자동 설정 완료. Form/트리거/초기 집계가 모두 준비되었습니다.');
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onFormSubmitHandler', 'dailyAutoRefresh'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('dailyAutoRefresh').timeBased().atHour(23).nearMinute(50).everyDays(1).create();
}

// ===========================================================================
// 1. Google Form 실제 생성
// ===========================================================================
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

  const statusItem = form.addMultipleChoiceItem();
  statusItem.setTitle('영업 여부').setRequired(true);
  statusItem.setChoices([
    statusItem.createChoice('영업', salesPage),
    statusItem.createChoice('휴무', memoSharedPage),
  ]);
  form.moveItem(statusItem.getIndex(), 2);

  if (oldLogs) ss.deleteSheet(oldLogs);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const newLogs = ss.getSheets().find(s => s.getName() !== SHEET_SALES_LOGS
    && /^(Form Responses|설문지 응답)/.test(s.getName()));
  if (!newLogs) throw new Error('응답 탭을 찾지 못했습니다. 수동으로 sales_logs로 이름을 바꿔주세요.');
  newLogs.setName(SHEET_SALES_LOGS);

  newLogs.getRange('G1:K1').setValues([['log_id', '업체명', '중복여부', '확인여부', '수정메모']]);

  PropertiesService.getScriptProperties().setProperty(PROP_STORE_ID_ITEM_ID, String(storeIdItem.getId()));
  PropertiesService.getScriptProperties().setProperty(PROP_FORM_EDIT_URL, form.getEditUrl());

  generateStoreLinks();
  recomputeSalesLogs_();

  SpreadsheetApp.getUi().alert(
    'Form 생성 완료\n응답 URL: ' + form.getPublishedUrl() +
    '\nstores!H열에 매장별 pre-filled 링크가 자동으로 채워졌습니다.'
  );
}

function generateStoreLinks() {
  const itemId = PropertiesService.getScriptProperties().getProperty(PROP_STORE_ID_ITEM_ID);
  const editUrl = PropertiesService.getScriptProperties().getProperty(PROP_FORM_EDIT_URL);
  if (!itemId || !editUrl) throw new Error('먼저 "0. 전체 자동 설정"을 실행하세요.');

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

// ===========================================================================
// 2. sales_logs 파생 컬럼(log_id/업체명/중복여부/확인여부) 전체 재계산
// ===========================================================================
function recomputeSalesLogs_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SALES_LOGS);
  const storesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const data = sheet.getRange(2, 1, n, 6).getValues(); // A~F
  const storeMap = {};
  storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 3).getValues()
    .forEach(([id, , name]) => { if (id) storeMap[id] = name; });

  const seen = {}; // key: storeId|dateStr -> count
  const logIds = [], names = [], dups = [], checks = [];

  data.forEach((row, i) => {
    const [, storeId, saleDate] = row;
    if (!storeId) {
      logIds.push(['']); names.push(['']); dups.push(['']);
      checks.push([sheet.getRange(i + 2, 10).getValue() || false]);
      return;
    }
    logIds.push(['L' + String(i + 1).padStart(4, '0')]);
    names.push([storeMap[storeId] || 'store_id 오류']);

    const dateStr = saleDate ? Utilities.formatDate(new Date(saleDate), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    const key = storeId + '|' + dateStr;
    seen[key] = (seen[key] || 0) + 1;
    dups.push([seen[key] > 1]);

    const existing = sheet.getRange(i + 2, 10).getValue();
    checks.push([existing === '' || existing === null ? false : existing]);
  });

  sheet.getRange(2, 7, n, 1).setValues(logIds);
  sheet.getRange(2, 8, n, 1).setValues(names);
  sheet.getRange(2, 9, n, 1).setValues(dups);
  sheet.getRange(2, 10, n, 1).setValues(checks);
}

// ===========================================================================
// 3. monthly_summary 전체 재계산 (C1=기준연도, E1=기준월 기준)
// ===========================================================================
function rebuildMonthlySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MONTHLY);
  const storesSheet = ss.getSheetByName(SHEET_STORES);
  const logsSheet = ss.getSheetByName(SHEET_SALES_LOGS);

  const year = sheet.getRange('C1').getValue();
  const month = sheet.getRange('E1').getValue();
  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);
  const today = new Date();
  const cmpEnd = today < monthEnd ? today : monthEnd;
  const cmpDays = cmpEnd < monthStart ? 0 : cmpEnd.getDate();
  const tz = Session.getScriptTimeZone();
  const todayKey = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

  const stores = storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 8).getValues()
    .filter(r => r[0] && r[6] === 'active');

  const logsLastRow = logsSheet.getLastRow();
  const logs = logsLastRow >= 2
    ? logsSheet.getRange(2, 1, logsLastRow - 1, 9).getValues() // A~I
    : [];

  // store_id -> { dateKey -> {status, sales, dup} } (dup인 행은 status/sales 집계에서 제외하되 표시는 남김)
  const byStore = {};
  logs.forEach(row => {
    const [, storeId, saleDate, status, sales, , , , dup] = row;
    if (!storeId || !saleDate) return;
    const dateKey = Utilities.formatDate(new Date(saleDate), tz, 'yyyy-MM-dd');
    byStore[storeId] = byStore[storeId] || {};
    const cell = byStore[storeId][dateKey] || { hasNonDup: false, status: null, sales: 0, hasDup: false };
    if (dup === true) {
      cell.hasDup = true;
    } else {
      cell.hasNonDup = true;
      cell.status = status;
      cell.sales = (cell.sales || 0) + (Number(sales) || 0);
    }
    byStore[storeId][dateKey] = cell;
  });

  function monthSalesAndCounts(storeId, fromDate, toDate) {
    const days = byStore[storeId] || {};
    let sales = 0, workDays = 0, offDays = 0, uniqueInputDays = 0;
    Object.keys(days).forEach(dateKey => {
      const d = new Date(dateKey);
      if (d < fromDate || d > toDate) return;
      const cell = days[dateKey];
      if (!cell.hasNonDup) return;
      uniqueInputDays++;
      if (cell.status === '영업') { workDays++; sales += cell.sales; }
      else if (cell.status === '휴무') { offDays++; }
    });
    return { sales, workDays, offDays, uniqueInputDays };
  }

  const out = [];
  stores.forEach(s => {
    const storeId = s[0], consultant = s[1], name = s[2];
    if (monthStart > today) {
      out.push({ storeId, consultant, name, inputRate: 0, missingDays: 0, offDays: 0, workDays: 0,
        monthSales: 0, dailyAvg: '', mom: '', days: Array(31).fill('') });
      return;
    }
    const cur = monthSalesAndCounts(storeId, monthStart, cmpEnd);
    const inputRate = cmpDays > 0 ? cur.uniqueInputDays / cmpDays : 0;
    const missingDays = cmpDays - cur.uniqueInputDays;
    const dailyAvg = cur.workDays > 0 ? cur.sales / cur.workDays : '';

    const prevMonthStart = new Date(year, month - 2, 1);
    const prevMonthEnd0 = new Date(year, month - 1, 0);
    const prevEnd = new Date(year, month - 2, Math.min(cmpDays, prevMonthEnd0.getDate()));
    const prev = monthSalesAndCounts(storeId, prevMonthStart, prevEnd);
    const mom = prev.sales === 0 ? '데이터부족' : (cur.sales - prev.sales) / prev.sales;

    const days = [];
    for (let day = 1; day <= 31; day++) {
      const d = new Date(year, month - 1, day);
      if (d.getMonth() !== month - 1) { days.push('---'); continue; }
      if (d > today) { days.push(''); continue; }
      const dateKey = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      const cell = (byStore[storeId] || {})[dateKey];
      if (!cell) { days.push('미입력'); continue; }
      if (cell.hasDup) { days.push('중복확인'); continue; }
      if (!cell.hasNonDup) { days.push('미입력'); continue; }
      if (cell.status === '휴무') { days.push('휴무'); continue; }
      days.push(cell.sales);
    }

    out.push({ storeId, consultant, name, inputRate, missingDays, offDays: cur.offDays,
      workDays: cur.workDays, monthSales: cur.sales, dailyAvg, mom, days });
  });

  const startRow = 4;
  const rows = out.map(r => [
    r.storeId, r.consultant, r.name, r.inputRate, r.missingDays, r.offDays,
    r.workDays, r.monthSales, r.dailyAvg, r.mom, ...r.days,
  ]);
  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
  // 매장 수가 줄어든 경우를 대비해 남는 행 정리
  const lastRow = sheet.getLastRow();
  if (lastRow > startRow + rows.length - 1) {
    sheet.getRange(startRow + rows.length, 1, lastRow - (startRow + rows.length) + 1, 41).clearContent();
  }
}

// ===========================================================================
// 4. missing_check 재계산
// ===========================================================================
function refreshMissingCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storesSheet = ss.getSheetByName(SHEET_STORES);
  const logsSheet = ss.getSheetByName(SHEET_SALES_LOGS);
  const checkSheet = ss.getSheetByName(SHEET_MISSING);
  const tz = Session.getScriptTimeZone();

  const checkDate = checkSheet.getRange('B1').getValue();
  const checkDateStr = Utilities.formatDate(new Date(checkDate), tz, 'yyyy-MM-dd');

  const storesData = storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 8).getValues();

  const logsLastRow = logsSheet.getLastRow();
  const submittedIds = new Set();
  let todaySales = 0;
  if (logsLastRow >= 2) {
    const logsData = logsSheet.getRange(2, 2, logsLastRow - 1, 8).getValues(); // B~I
    logsData.forEach(([storeId, saleDate, status, sales, , , , dup]) => {
      if (!storeId || !saleDate) return;
      const d = Utilities.formatDate(new Date(saleDate), tz, 'yyyy-MM-dd');
      if (d !== checkDateStr) return;
      submittedIds.add(storeId);
      if (dup !== true && status === '영업') todaySales += Number(sales) || 0;
    });
  }

  const missingRows = storesData
    .filter(row => row[6] === 'active' && !submittedIds.has(row[0]))
    .map(row => [row[1], row[2], row[4], checkDateStr, row[7], row[6]]);

  const lastRow = checkSheet.getLastRow();
  if (lastRow >= 5) checkSheet.getRange(5, 1, lastRow - 4, 6).clearContent();
  if (missingRows.length === 0) {
    checkSheet.getRange('A5').setValue('오늘 미제출 매장 없음');
  } else {
    checkSheet.getRange(5, 1, missingRows.length, 6).setValues(missingRows);
  }

  const activeCount = storesData.filter(r => r[6] === 'active').length;
  checkSheet.getRange('B2').setValue(submittedIds.size);
  checkSheet.getRange('D2').setValue(activeCount - submittedIds.size);
  checkSheet.getRange('F2').setValue(todaySales);

  return { activeCount, submitted: submittedIds.size, missing: missingRows, todaySales, checkDateStr };
}

// ===========================================================================
// 5. dashboard 재계산 (missing_check 결과 + monthly_summary 결과를 재사용)
// ===========================================================================
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEET_DASHBOARD);
  const missing = refreshMissingCheck();

  dash.getRange('B3').setValue(new Date(missing.checkDateStr));
  dash.getRange('A6').setValue(missing.activeCount);
  dash.getRange('B6').setValue(missing.submitted);
  dash.getRange('C6').setValue(missing.missing.length);
  dash.getRange('D6').setValue(missing.todaySales);

  const monthlySheet = ss.getSheetByName(SHEET_MONTHLY);
  const monthlyLastRow = monthlySheet.getLastRow();
  let monthTotal = 0;
  const byConsultant = {};
  if (monthlyLastRow >= 4) {
    const rows = monthlySheet.getRange(4, 1, monthlyLastRow - 3, 8).getValues(); // A~H
    rows.forEach(([storeId, consultant, , , , , , monthSales]) => {
      if (!storeId) return;
      monthTotal += Number(monthSales) || 0;
      byConsultant[consultant] = byConsultant[consultant] || { stores: 0, sales: 0 };
      byConsultant[consultant].stores++;
      byConsultant[consultant].sales += Number(monthSales) || 0;
    });
  }
  dash.getRange('E6').setValue(monthTotal);
  dash.getRange('F6').setValue(missing.activeCount > 0 ? missing.submitted / missing.activeCount : 0);
  dash.getRange('F6').setNumberFormat('0%');

  const storesSheet = ss.getSheetByName(SHEET_STORES);
  const storesData = storesSheet.getRange(2, 1, Math.max(storesSheet.getLastRow() - 1, 0), 7).getValues();
  const missingByConsultant = {};
  missing.missing.forEach(([consultant]) => {
    missingByConsultant[consultant] = (missingByConsultant[consultant] || 0) + 1;
  });
  const consultants = [...new Set(storesData.filter(r => r[6] === 'active').map(r => r[1]))];

  const rows = consultants.map(c => {
    const storeCount = storesData.filter(r => r[6] === 'active' && r[1] === c).length;
    const missingCount = missingByConsultant[c] || 0;
    return [c, storeCount, storeCount - missingCount, missingCount, (byConsultant[c] || {}).sales || 0];
  });

  const lastRow = dash.getLastRow();
  if (lastRow >= 10) dash.getRange(10, 1, lastRow - 9, 5).clearContent();
  if (rows.length > 0) dash.getRange(10, 1, rows.length, 5).setValues(rows);
}

// ===========================================================================
// 6. 통합 재계산 (메뉴/트리거 공용)
// ===========================================================================
function refreshAll() {
  recomputeSalesLogs_();
  rebuildMonthlySummary();
  refreshMissingCheck();
  refreshDashboard();
}

// ===========================================================================
// 7. 트리거 핸들러
// ===========================================================================
function onFormSubmitHandler(e) {
  refreshAll();
  notifyIfDuplicate_();
}

function dailyAutoRefresh() {
  const missingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MISSING);
  missingSheet.getRange('B1').setValue(new Date());
  refreshAll();
  sendMissingSummaryEmail_();
}

function notifyIfDuplicate_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SALES_LOGS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const isDup = sheet.getRange(lastRow, 9).getValue();
  if (isDup === true) {
    const storeId = sheet.getRange(lastRow, 2).getValue();
    const companyName = sheet.getRange(lastRow, 8).getValue();
    Logger.log(`[중복 입력 감지] ${companyName} (${storeId})`);
  }
}

function sendMissingSummaryEmail_() {
  if (!ADMIN_EMAIL) return;
  const checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MISSING);
  const lastRow = checkSheet.getLastRow();
  if (lastRow < 5) return;
  const rows = checkSheet.getRange(5, 1, lastRow - 4, 6).getValues();
  if (rows.length === 1 && rows[0][0] === '오늘 미제출 매장 없음') return;
  const body = rows.map(r => `${r[0]} | ${r[1]} | ${r[2]}`).join('\n');
  MailApp.sendEmail(ADMIN_EMAIL, '[일매출] 오늘 미제출 매장 안내', body);
}

// ===========================================================================
// 8. 메뉴
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('일매출관리')
    .addItem('0. 전체 자동 설정 (Form 생성 + 트리거 설치)', 'setupAll')
    .addSeparator()
    .addItem('1. monthly_summary 새로고침', 'rebuildMonthlySummary')
    .addItem('2. missing_check 새로고침', 'refreshMissingCheck')
    .addItem('3. dashboard 새로고침', 'refreshDashboard')
    .addItem('4. 전체 새로고침', 'refreshAll')
    .addSeparator()
    .addItem('5. 매장별 입력링크 다시 생성', 'generateStoreLinks')
    .addToUi();
}
