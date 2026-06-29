// ====================================================
// 소재 인덱싱 시스템 - Code.gs
// ====================================================

const SHEET_ID           = '1S74kKyOO3ATqk12nmQR860VT5ms8s5FS33T3gF5kd7I';
const MASTER_SHEET_NAME  = '소재_마스터';
const SETTINGS_SHEET_NAME = '설정';
const HIERARCHY_SHEET_NAME = '매체_계층';
const RAW_SHEET_NAME     = '매체_RAW';
const DETECT_SHEET_NAME  = '신규소재감지';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소재 인덱싱 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// --------------------------------------------------
// 시트 초기화 (최초 1회 실행)
//
// [설정 시트] 세로형 — 1행: 카테고리명, 2행~: 값
//   광고유형 | 소재유형 | 소구포인트 | 후킹방식 | 이미지유형 | 모델유형 | 보종
//
// [매체_계층 시트] — 매체/캠페인/그룹/소재이름 계층 정의
//   매체 | 캠페인 | 그룹 | 소재이름  (행마다 1개 경로)
// --------------------------------------------------
function initializeSheets() {
  const ss = getSpreadsheet();

  // 설정 시트 (기타 드롭다운 — 계층 항목 제외)
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    const categories = ['광고유형', '소재유형', '소구포인트', '후킹방식', '이미지유형', '모델유형', '보종'];
    settingsSheet.getRange(1, 1, 1, categories.length).setValues([categories]);
    settingsSheet.getRange(1, 1, 1, categories.length).setFontWeight('bold');
    settingsSheet.setFrozenRows(1);
  }

  // 매체_계층 시트
  let hierarchySheet = ss.getSheetByName(HIERARCHY_SHEET_NAME);
  if (!hierarchySheet) {
    hierarchySheet = ss.insertSheet(HIERARCHY_SHEET_NAME);
    hierarchySheet.getRange(1, 1, 1, 4).setValues([['매체', '캠페인', '그룹', '소재이름']]);
    hierarchySheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    hierarchySheet.setFrozenRows(1);
  }

  // 소재_마스터 시트
  let masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
    const headers = [['이미지코드', '등록일자', '매체', '캠페인', '그룹', '소재이름', '보종',
      '광고유형', '소재유형', '소구포인트', '후킹방식', '소구상세', '이미지유형', '모델유형', '이미지URL']];
    masterSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    masterSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold');
    masterSheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(RAW_SHEET_NAME))    ss.insertSheet(RAW_SHEET_NAME);
  if (!ss.getSheetByName(DETECT_SHEET_NAME)) ss.insertSheet(DETECT_SHEET_NAME);

  return { success: true, message: '시트 초기화 완료' };
}

// --------------------------------------------------
// 기타 드롭다운 설정 읽기 (세로형)
// 1행 = 카테고리명, 2행~ = 값
// --------------------------------------------------
function getSettings() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return {};

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const settings = {};

  headers.forEach((header, colIdx) => {
    if (!header) return;
    settings[header] = [];
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
      const val = data[rowIdx][colIdx];
      if (val !== '' && val !== null && val !== undefined) {
        settings[header].push(String(val));
      }
    }
  });
  return settings;
}

// --------------------------------------------------
// 기타 드롭다운 설정 저장 (세로형으로 덮어씀)
// --------------------------------------------------
function saveSettings(settingsData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) return { error: true, message: '설정 시트가 없습니다.' };

    sheet.clearContents();
    const valid = settingsData.filter(item => item.name);
    if (!valid.length) return { success: true };

    sheet.getRange(1, 1, 1, valid.length).setValues([valid.map(i => i.name)]);
    sheet.getRange(1, 1, 1, valid.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    valid.forEach((item, colIdx) => {
      const vals = item.values.filter(v => v !== '');
      if (vals.length > 0) {
        sheet.getRange(2, colIdx + 1, vals.length, 1).setValues(vals.map(v => [v]));
      }
    });
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// 매체 계층 읽기
// 반환: [{ 매체, 캠페인, 그룹, 소재이름 }, ...]
// --------------------------------------------------
function getHierarchy() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(HIERARCHY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data
    .filter(r => r[0] || r[1] || r[2] || r[3])
    .map(r => ({
      매체:     String(r[0] || ''),
      캠페인:   String(r[1] || ''),
      그룹:     String(r[2] || ''),
      소재이름: String(r[3] || '')
    }));
}

// --------------------------------------------------
// 매체 계층 저장
// rows: [{ 매체, 캠페인, 그룹, 소재이름 }, ...]
// --------------------------------------------------
function saveHierarchy(rows) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(HIERARCHY_SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(HIERARCHY_SHEET_NAME);

    sheet.clearContents();
    sheet.getRange(1, 1, 1, 4).setValues([['매체', '캠페인', '그룹', '소재이름']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);

    const valid = rows.filter(r => r.매체 || r.캠페인 || r.그룹 || r.소재이름);
    if (valid.length > 0) {
      sheet.getRange(2, 1, valid.length, 4)
        .setValues(valid.map(r => [r.매체, r.캠페인, r.그룹, r.소재이름]));
    }
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// 중복 체크 (매체+캠페인+그룹+소재이름 조합)
// --------------------------------------------------
function checkDuplicate(매체, 캠페인, 그룹, 소재이름) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 4).getValues();
  return data.some(r => r[0] === 매체 && r[1] === 캠페인 && r[2] === 그룹 && r[3] === 소재이름);
}

// --------------------------------------------------
// 이미지코드 생성: IMG + YYMMDD + 3자리 순번
// --------------------------------------------------
function generateImageCode() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const today = new Date();
  const dateStr = String(today.getFullYear()).slice(2)
    + String(today.getMonth() + 1).padStart(2, '0')
    + String(today.getDate()).padStart(2, '0');
  const prefix = 'IMG' + dateStr;
  let seq = 1;
  if (sheet && sheet.getLastRow() >= 2) {
    const codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    seq = codes.filter(c => c && String(c).startsWith(prefix)).length + 1;
  }
  return prefix + String(seq).padStart(3, '0');
}

// --------------------------------------------------
// Drive 폴더 생성/조회
// --------------------------------------------------
function createDriveFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('DRIVE_FOLDER_ID');
  if (folderId) {
    try { DriveApp.getFolderById(folderId); return folderId; } catch (e) {}
  }
  const existing = DriveApp.getFoldersByName('소재_이미지');
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder('소재_이미지');
  folderId = folder.getId();
  props.setProperty('DRIVE_FOLDER_ID', folderId);
  return folderId;
}

// --------------------------------------------------
// Drive 업로드 → 공개 URL 반환
// --------------------------------------------------
function uploadImageToDrive(base64Data, fileName, mimeType) {
  const folder = DriveApp.getFolderById(createDriveFolder());
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// --------------------------------------------------
// 소재 저장
// --------------------------------------------------
function saveCreative(data) {
  try {
    if (!data.forceSave && checkDuplicate(data.매체, data.캠페인, data.그룹, data.소재이름))
      return { duplicate: true };

    let imageUrl = '';
    if (data.fileData) imageUrl = uploadImageToDrive(data.fileData, data.fileName, data.mimeType);

    const imageCode = generateImageCode();
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sheet.appendRow([
      imageCode, dateStr, data.매체, data.캠페인, data.그룹, data.소재이름,
      data.보종, data.광고유형, data.소재유형, data.소구포인트,
      data.후킹방식, data.소구상세, data.이미지유형, data.모델유형, imageUrl
    ]);

    return { success: true, imageCode, imageUrl };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// 신규 소재 감지
// --------------------------------------------------
function detectNewCreatives() {
  try {
    const ss = getSpreadsheet();
    const rawSheet    = ss.getSheetByName(RAW_SHEET_NAME);
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const detectSheet = ss.getSheetByName(DETECT_SHEET_NAME);

    if (!rawSheet || rawSheet.getLastRow() < 2)
      return { count: 0, items: [], message: '매체_RAW 시트에 데이터가 없습니다.' };

    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 4).getValues();

    const masterSet = new Set();
    if (masterSheet && masterSheet.getLastRow() >= 2) {
      masterSheet.getRange(2, 3, masterSheet.getLastRow() - 1, 4).getValues()
        .forEach(r => masterSet.add(r.join('\x00')));
    }

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const newItems = rawData.filter(r => r[0] && !masterSet.has(r.join('\x00')));

    detectSheet.clearContents();
    if (newItems.length > 0) {
      detectSheet.getRange(1, 1, 1, 5).setValues([['매체', '캠페인', '그룹', '소재이름', '감지일시']]);
      detectSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      detectSheet.getRange(2, 1, newItems.length, 5)
        .setValues(newItems.map(r => [r[0], r[1], r[2], r[3], now]));
    }

    return {
      count: newItems.length,
      items: newItems.map(r => ({ 매체: r[0], 캠페인: r[1], 그룹: r[2], 소재이름: r[3] })),
      message: `신규 소재 ${newItems.length}건 발견`
    };
  } catch (e) {
    return { error: true, message: e.message };
  }
}
