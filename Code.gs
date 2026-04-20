/**
 * 텀블벅 후원 트래커 - Google Apps Script 백엔드
 *
 * 동작 원리:
 *  - 이 스크립트가 붙어있는 스프레드시트의 "projects" 시트를 DB처럼 사용
 *  - 웹앱으로 배포되면 /exec URL이 생성되고, Index.html이 UI로 서빙됨
 *  - UI는 google.script.run.* 호출로 이 파일의 함수를 부름
 */

const SHEET_NAME = 'projects';

// 시트 컬럼 순서 - 이 순서대로 저장됩니다.
const COLUMNS = [
  'id', 'title', 'creator', 'amount', 'category',
  'status', 'paymentDate', 'deliveryDate', 'downloaded',
  'url', 'notes', 'createdAt', 'updatedAt'
];

// =========================================================
// 웹앱 진입점
// =========================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('텀블벅 후원 트래커')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 메뉴 추가 (시트에서 바로 실행할 수 있게)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('텀블벅 트래커')
    .addItem('시트 초기화', 'setupSheet')
    .addItem('웹앱 URL 보기', 'showWebAppUrl')
    .addToUi();
}

// =========================================================
// 시트 초기화 - 최초 1회 실행
// =========================================================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // 헤더 행이 없거나 비었을 경우에만 추가
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.getRange(1, 1, 1, COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#1f1a14')
      .setFontColor('#fbf7ee');
    sheet.setFrozenRows(1);

    // 보기 좋게 컬럼 너비 조정
    sheet.setColumnWidth(1, 260);  // id
    sheet.setColumnWidth(2, 260);  // title
    sheet.setColumnWidth(3, 140);  // creator
    sheet.setColumnWidth(4, 100);  // amount
    sheet.setColumnWidth(5, 100);  // category
    sheet.setColumnWidth(6, 100);  // status
    sheet.setColumnWidth(7, 110);  // paymentDate
    sheet.setColumnWidth(8, 110);  // deliveryDate
    sheet.setColumnWidth(9, 90);   // downloaded
    sheet.setColumnWidth(10, 220); // url
    sheet.setColumnWidth(11, 240); // notes
  }

  try {
    SpreadsheetApp.getUi().alert('시트가 준비되었어요. 이제 웹앱으로 사용하실 수 있습니다.');
  } catch (e) {
    // UI 컨텍스트가 없을 수 있음 (직접 실행 시)
  }
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const msg = url
    ? '웹앱 URL:\n\n' + url + '\n\n모바일/PC 어디서든 이 주소로 접속하세요.'
    : '아직 웹앱으로 배포되지 않았어요.\n[배포] → [새 배포] → 유형 "웹 앱" 으로 배포해주세요.';
  SpreadsheetApp.getUi().alert(msg);
}

// =========================================================
// 공통 헬퍼
// =========================================================
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}

function rowToObject_(row) {
  const obj = {};
  COLUMNS.forEach((col, i) => {
    let v = row[i];
    // 날짜 객체면 YYYY-MM-DD 문자열로
    if (v instanceof Date && !isNaN(v)) {
      v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (col === 'amount') v = Number(v) || 0;
    if (col === 'downloaded') v = v === true || v === 'TRUE' || v === 'true' || v === 1;
    obj[col] = v == null ? '' : v;
  });
  return obj;
}

function objectToRow_(obj) {
  return COLUMNS.map(col => {
    const v = obj[col];
    if (v === undefined || v === null) return '';
    if (col === 'downloaded') return !!v;
    return v;
  });
}

function findRowIndexById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based
  }
  return -1;
}

function newId_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

// =========================================================
// 클라이언트에서 호출하는 API
// =========================================================

/**
 * 모든 프로젝트 조회
 */
function listProjects() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  return values
    .filter(row => row[0]) // id 있는 행만
    .map(rowToObject_);
}

/**
 * 새 프로젝트 추가 (단일)
 */
function addProject(data) {
  const sheet = getSheet_();
  const project = {
    id: data.id || newId_(),
    title: String(data.title || '').trim(),
    creator: String(data.creator || '').trim(),
    amount: Number(data.amount) || 0,
    category: String(data.category || ''),
    status: data.status || 'pending',
    paymentDate: data.paymentDate || '',
    deliveryDate: data.deliveryDate || '',
    downloaded: !!data.downloaded,
    url: String(data.url || '').trim(),
    notes: String(data.notes || '').trim(),
    createdAt: data.createdAt || nowIso_(),
    updatedAt: nowIso_(),
  };
  sheet.appendRow(objectToRow_(project));
  return project;
}

/**
 * 여러 개 일괄 추가 (붙여넣기 파싱 결과용)
 */
function addProjects(list) {
  if (!Array.isArray(list) || list.length === 0) return { added: 0, items: [] };
  const sheet = getSheet_();
  const existing = listProjects();
  const existingKeys = new Set(existing.map(p => p.title.trim() + '::' + p.amount));
  const toInsert = [];
  const items = [];
  list.forEach(data => {
    const key = String(data.title || '').trim() + '::' + (Number(data.amount) || 0);
    if (existingKeys.has(key)) return;
    existingKeys.add(key);
    const project = {
      id: newId_(),
      title: String(data.title || '').trim(),
      creator: String(data.creator || '').trim(),
      amount: Number(data.amount) || 0,
      category: String(data.category || ''),
      status: data.status || 'pending',
      paymentDate: data.paymentDate || '',
      deliveryDate: data.deliveryDate || '',
      downloaded: !!data.downloaded,
      url: String(data.url || '').trim(),
      notes: String(data.notes || '').trim(),
      createdAt: nowIso_(),
      updatedAt: nowIso_(),
    };
    toInsert.push(objectToRow_(project));
    items.push(project);
  });
  if (toInsert.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toInsert.length, COLUMNS.length).setValues(toInsert);
  }
  return { added: items.length, items };
}

/**
 * 프로젝트 수정
 */
function updateProject(data) {
  if (!data || !data.id) throw new Error('id가 필요합니다.');
  const sheet = getSheet_();
  const rowIdx = findRowIndexById_(sheet, data.id);
  if (rowIdx < 0) throw new Error('해당 ID의 프로젝트를 찾을 수 없어요: ' + data.id);

  const existingValues = sheet.getRange(rowIdx, 1, 1, COLUMNS.length).getValues()[0];
  const existing = rowToObject_(existingValues);
  const merged = Object.assign({}, existing, data, { updatedAt: nowIso_() });
  sheet.getRange(rowIdx, 1, 1, COLUMNS.length).setValues([objectToRow_(merged)]);
  return merged;
}

/**
 * 다운로드 체크박스 토글
 */
function toggleDownloaded(id) {
  const sheet = getSheet_();
  const rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx < 0) throw new Error('프로젝트 없음');
  const col = COLUMNS.indexOf('downloaded') + 1;
  const current = sheet.getRange(rowIdx, col).getValue();
  const next = !(current === true || current === 'TRUE' || current === 'true' || current === 1);
  sheet.getRange(rowIdx, col).setValue(next);
  const updatedCol = COLUMNS.indexOf('updatedAt') + 1;
  sheet.getRange(rowIdx, updatedCol).setValue(nowIso_());
  return { id, downloaded: next };
}

/**
 * 프로젝트 삭제
 */
function deleteProject(id) {
  const sheet = getSheet_();
  const rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx < 0) throw new Error('프로젝트 없음');
  sheet.deleteRow(rowIdx);
  return { id, deleted: true };
}
