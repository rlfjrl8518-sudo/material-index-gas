// ====================================================
// 소재 인덱싱 시스템 - Code.gs
// ====================================================

const SHEET_ID           = '1S74kKyOO3ATqk12nmQR860VT5ms8s5FS33T3gF5kd7I';
const MASTER_SHEET_NAME  = '소재_마스터';
const SETTINGS_SHEET_NAME = '설정';
const HIERARCHY_SHEET_NAME = '매체_계층';
const RAW_SHEET_NAME     = '매체_RAW';
const DETECT_SHEET_NAME  = '신규소재감지';
const DGPM_SHEET_NAME    = 'DG_PM_광고단위';

// 광고 단위에 여러 이미지가 포함되는 매체 (1:N 구조)
const DGPM_MEDIA = ['디멘드젠', '피맥스'];

// DG_PM_광고단위 시트 헤더 (소재_마스터 컬럼 순서 기준)
const DGPM_HEADERS = [
  '광고단위코드',
  '등록일시', '최근수정일시',
  '매체', '캠페인', '그룹', '소재이름',
  '보종', '광고유형',
  '소재유형', '소구포인트', '후킹방식', '소구상세',
  '이미지유형목록', '모델유형목록',
  '이미지수', '이미지코드목록',
  '번들URL'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소재 인덱싱 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
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
      '광고유형', '소재유형', '소구포인트', '후킹방식', '소구상세', '이미지유형', '모델유형', '이미지URL', '파일해시']];
    masterSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    masterSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold');
    masterSheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(RAW_SHEET_NAME))    ss.insertSheet(RAW_SHEET_NAME);
  if (!ss.getSheetByName(DETECT_SHEET_NAME)) ss.insertSheet(DETECT_SHEET_NAME);

  // DG_PM_광고단위 시트
  if (!ss.getSheetByName(DGPM_SHEET_NAME)) {
    const dgpmSheet = ss.insertSheet(DGPM_SHEET_NAME);
    dgpmSheet.getRange(1, 1, 1, DGPM_HEADERS.length).setValues([DGPM_HEADERS]).setFontWeight('bold');
    dgpmSheet.setFrozenRows(1);
  }

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
// 소재_마스터의 이미지코드 목록 반환 (선택 드롭다운용)
// 중복 제거 후 최신순 정렬, 대표 소재이름 포함
// --------------------------------------------------
function getImageCodes() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
  // col: 0=이미지코드, 1=등록일자, 2=매체, 5=소재이름, 6=보종, 8=소재유형, 14=이미지URL
  const seen = new Map();
  data.forEach(row => {
    const code = String(row[0] || '').trim();
    if (!code || seen.has(code)) return;
    seen.set(code, {
      code,
      등록일자: row[1] ? String(row[1]).slice(0, 10) : '',
      매체:     String(row[2] || ''),
      소재이름: String(row[5] || ''),
      보종:     String(row[6] || ''),
      소재유형: String(row[8] || ''),
      imageUrl: String(row[14] || '')
    });
  });
  return [...seen.values()].reverse();
}

// --------------------------------------------------
// 디버그: 이미지 코드 원시값 확인
// --------------------------------------------------
function debugImageCodes() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const result = data.map((row, i) => ({
    행: i + 2,
    코드: JSON.stringify(String(row[0] || '')), // 공백·특수문자 포함해서 보여줌
    길이: String(row[0] || '').length
  }));
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// --------------------------------------------------
// 이미지코드로 소재 정보 조회 (기존 코드 선택 시 폼 자동 채우기)
// --------------------------------------------------
function getCreativeByImageCode(imageCode) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  const rows = data.filter(r => String(r[0]).trim() === imageCode);
  if (!rows.length) return null;
  const row = rows[rows.length - 1]; // 가장 최근 행

  return {
    매체:       String(row[2]  || ''),
    캠페인:     String(row[3]  || ''),
    그룹:       String(row[4]  || ''),
    소재이름:   String(row[5]  || ''),
    보종:       String(row[6]  || ''),
    광고유형:   String(row[7]  || ''),
    소재유형:   String(row[8]  || ''),
    소구포인트: String(row[9]  || ''),
    후킹방식:   String(row[10] || ''),
    소구상세:   String(row[11] || ''),
    이미지유형: String(row[12] || ''),
    모델유형:   String(row[13] || ''),
    이미지URL:  String(row[14] || '')
  };
}

// --------------------------------------------------
// 동일 이미지 파일 조회 (파일해시 기준)
// 이미 등록된 이미지면 imageCode + imageUrl 반환, 없으면 null
// --------------------------------------------------
function checkExistingImage(fileHash) {
  if (!fileHash) return null;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const lastCol = sheet.getLastColumn();
  if (lastCol < 16) return null; // 파일해시 열 없음

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  const found = data.find(row => row[15] && row[15] === fileHash);
  if (!found) return null;

  return { imageCode: String(found[0]), imageUrl: String(found[14]) };
}

// --------------------------------------------------
// 중복 체크 (매체+캠페인+그룹+소재이름 조합)
// DG/PM은 동일 소재이름에 여러 이미지가 허용되므로 체크 제외
// --------------------------------------------------
function checkDuplicate(매체, 캠페인, 그룹, 소재이름) {
  if (DGPM_MEDIA.includes(매체)) return false;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 4).getValues();
  return data.some(r => r[0] === 매체 && r[1] === 캠페인 && r[2] === 그룹 && r[3] === 소재이름);
}

// --------------------------------------------------
// 소재코드 생성
//   이미지: IMG + YYMMDD + 3자리 순번 (예: IMG260630001)
//   동영상: VID + YYMMDD + 3자리 순번 (예: VID260630001)
// --------------------------------------------------
function generateMediaCode(mimeType) {
  const isVideo = mimeType && mimeType.startsWith('video/');
  const prefix = (isVideo ? 'VID' : 'IMG');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const today = new Date();
  const dateStr = String(today.getFullYear()).slice(2)
    + String(today.getMonth() + 1).padStart(2, '0')
    + String(today.getDate()).padStart(2, '0');
  const fullPrefix = prefix + dateStr;
  let seq = 1;
  if (sheet && sheet.getLastRow() >= 2) {
    const codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    seq = codes.filter(c => c && String(c).startsWith(fullPrefix)).length + 1;
  }
  return fullPrefix + String(seq).padStart(3, '0');
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
// Drive 직접 업로드용 OAuth 토큰 + 폴더 ID 반환
// 브라우저가 Drive API로 직접 업로드할 때 사용
// --------------------------------------------------
function getUploadToken() {
  return {
    token:    ScriptApp.getOAuthToken(),
    folderId: createDriveFolder()
  };
}

// --------------------------------------------------
// Drive 파일 공개 설정 → 공개 URL 반환
// 브라우저 직접 업로드 후 fileId만 받아 공개 처리
// --------------------------------------------------
function setFilePublic(fileId) {
  const file = DriveApp.getFileById(fileId);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // 동영상: Drive 내장 플레이어 URL (uc?export=view는 동영상 스트리밍 불가)
  if (file.getMimeType().startsWith('video/')) {
    return 'https://drive.google.com/file/d/' + fileId + '/preview';
  }
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

// --------------------------------------------------
// Drive 업로드 (base64) → 공개 URL 반환
// 직접 업로드가 불가한 환경의 폴백용
// --------------------------------------------------
function uploadImageToDrive(base64Data, fileName, mimeType) {
  const folder = DriveApp.getFolderById(createDriveFolder());
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// --------------------------------------------------
// 소재_마스터에서 같은 (이미지코드 + 매체/캠페인/그룹/소재이름) 행 찾기
// 반환: 시트 행번호(1-based), 없으면 -1
// --------------------------------------------------
function _findMasterRowToUpdate(imageCode, 매체, 캠페인, 그룹, 소재이름) {
  if (!imageCode) return -1;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return -1;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    const [code, , m, c, g, n] = data[i];
    if (String(code) === imageCode && m === 매체 && c === 캠페인 && g === 그룹 && n === 소재이름)
      return i + 2; // 헤더(1행) + 데이터 오프셋
  }
  return -1;
}

// --------------------------------------------------
// 소재_마스터 기존 행 덮어쓰기 (이미지코드·등록일자 유지)
// --------------------------------------------------
function _updateMasterRow(rowIndex, data, imageCode, imageUrl) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  // 등록일자는 최초 등록일 유지, 수정일자 별도 컬럼 없으므로 그대로 둠
  const origDate = sheet.getRange(rowIndex, 2).getValue();
  const dateStr  = origDate
    ? Utilities.formatDate(new Date(origDate), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  sheet.getRange(rowIndex, 1, 1, 16).setValues([[
    imageCode, dateStr, data.매체, data.캠페인, data.그룹, data.소재이름,
    data.보종, data.광고유형, data.소재유형, data.소구포인트,
    data.후킹방식, data.소구상세, data.이미지유형, data.모델유형,
    imageUrl, data.fileHash || ''
  ]]);
}

// --------------------------------------------------
// 소재 저장 (수정 판단 포함)
//
// UPDATE 조건: 같은 이미지코드 + 같은 (매체, 캠페인, 그룹, 소재이름)
//   → 속성(유형·소구 등)만 바꾼 수정으로 판단, 기존 행 덮어씀
//
// INSERT 조건: 위 조건 불일치 (새 이미지 or 같은 이미지를 다른 지면에 등록)
// --------------------------------------------------
function saveCreative(data) {
  try {
    const isDGPM = DGPM_MEDIA.includes(data.매체);

    // 이미지코드 결정 우선순위:
    // 1) 사용자가 직접 선택한 코드
    // 2) 동일 파일 해시로 기존 코드 자동 매칭
    // 3) 신규 코드 생성
    let imageCode, imageUrl = '', reused = false;

    if (data.selectedImageCode) {
      imageCode = data.selectedImageCode;
      reused    = true;
    } else {
      const existing = checkExistingImage(data.fileHash);
      if (existing) {
        imageCode = existing.imageCode;
        imageUrl  = existing.imageUrl;
        reused    = true;
      } else {
        imageCode = generateMediaCode(data.mimeType);
      }
    }

    // URL 결정: 브라우저 직접 업로드 → base64 폴백 → 기존 URL 유지
    if (data.directUploadUrl) {
      imageUrl = data.directUploadUrl;              // 브라우저가 Drive에 직접 업로드한 경우
    } else if (data.fileData) {
      imageUrl = uploadImageToDrive(data.fileData, data.fileName, data.mimeType); // 폴백
    } else if (data.existingImageUrl) {
      imageUrl = data.existingImageUrl;
    }

    // ── UPDATE 판단 ──
    // 같은 이미지코드 + 같은 (매체, 캠페인, 그룹, 소재이름) 행이 있으면 수정
    const updateRow = _findMasterRowToUpdate(imageCode, data.매체, data.캠페인, data.그룹, data.소재이름);
    if (updateRow !== -1) {
      _updateMasterRow(updateRow, data, imageCode, imageUrl);
      let 광고단위코드 = null;
      if (isDGPM) 광고단위코드 = _updateDGPMUnit(data, imageCode);
      return { success: true, imageCode, imageUrl, reused, updated: true, 광고단위코드 };
    }

    // ── INSERT 판단 ──
    // 비DG/PM: 같은 (매체, 캠페인, 그룹, 소재이름)에 다른 이미지 존재 → 중복 경고
    if (!isDGPM && !data.forceSave && checkDuplicate(data.매체, data.캠페인, data.그룹, data.소재이름))
      return { duplicate: true };

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(MASTER_SHEET_NAME);
      const headers = ['이미지코드','등록일자','매체','캠페인','그룹','소재이름','보종',
        '광고유형','소재유형','소구포인트','후킹방식','소구상세','이미지유형','모델유형','이미지URL','파일해시'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sheet.appendRow([
      imageCode, dateStr, data.매체, data.캠페인, data.그룹, data.소재이름,
      data.보종, data.광고유형, data.소재유형, data.소구포인트,
      data.후킹방식, data.소구상세, data.이미지유형, data.모델유형,
      imageUrl, data.fileHash || ''
    ]);

    let 광고단위코드 = null;
    if (isDGPM) 광고단위코드 = _updateDGPMUnit(data, imageCode);

    return { success: true, imageCode, imageUrl, reused, updated: false, 광고단위코드 };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// DG/PM 광고단위 시트 조회/업데이트
// 동일 (매체, 캠페인, 그룹, 소재이름)이면 이미지코드목록에 추가,
// 없으면 새 광고단위 행 생성. 광고단위코드 반환.
// --------------------------------------------------
// C = 열 인덱스 (0-based), DGPM_HEADERS 순서와 동기화
const DGPM_COL = {};
DGPM_HEADERS.forEach((h, i) => { DGPM_COL[h] = i; });

function _updateDGPMUnit(data, imageCode) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(DGPM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DGPM_SHEET_NAME);
    sheet.getRange(1, 1, 1, DGPM_HEADERS.length).setValues([DGPM_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const now = new Date();
  const totalCols = DGPM_HEADERS.length;

  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalCols).getValues();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row[DGPM_COL['매체']]    === data.매체    &&
          row[DGPM_COL['캠페인']]  === data.캠페인  &&
          row[DGPM_COL['그룹']]    === data.그룹    &&
          row[DGPM_COL['소재이름']] === data.소재이름) {

        // 이미지코드 추가
        const codes = _splitList(row[DGPM_COL['이미지코드목록']]);
        if (!codes.includes(imageCode)) codes.push(imageCode);

        // 이미지유형, 모델유형 목록 갱신 (새 이미지 값 추가)
        const imgTypes  = _splitList(row[DGPM_COL['이미지유형목록']]);
        const modelTypes = _splitList(row[DGPM_COL['모델유형목록']]);
        if (data.이미지유형 && !imgTypes.includes(data.이미지유형))   imgTypes.push(data.이미지유형);
        if (data.모델유형   && !modelTypes.includes(data.모델유형))   modelTypes.push(data.모델유형);

        const rowNum = i + 2;
        sheet.getRange(rowNum, DGPM_COL['이미지유형목록'] + 1).setValue(imgTypes.join(','));
        sheet.getRange(rowNum, DGPM_COL['모델유형목록']   + 1).setValue(modelTypes.join(','));
        sheet.getRange(rowNum, DGPM_COL['이미지수']       + 1).setValue(codes.length);
        sheet.getRange(rowNum, DGPM_COL['이미지코드목록'] + 1).setValue(codes.join(','));
        sheet.getRange(rowNum, DGPM_COL['최근수정일시']   + 1).setValue(now);
        const unitCode = String(row[DGPM_COL['광고단위코드']]);
        if (!row[DGPM_COL['번들URL']]) {
          sheet.getRange(rowNum, DGPM_COL['번들URL'] + 1).setValue(_getBundleUrl(unitCode));
        }
        return unitCode;
      }
    }
  }

  // 신규 광고단위 생성 — 공통 속성은 첫 이미지 기준
  const unitCode = _generateDGPMCode(data.매체, sheet);
  const newRow = new Array(totalCols).fill('');
  newRow[DGPM_COL['광고단위코드']]   = unitCode;
  newRow[DGPM_COL['등록일시']]       = now;
  newRow[DGPM_COL['최근수정일시']]   = now;
  newRow[DGPM_COL['매체']]           = data.매체;
  newRow[DGPM_COL['캠페인']]         = data.캠페인;
  newRow[DGPM_COL['그룹']]           = data.그룹;
  newRow[DGPM_COL['소재이름']]       = data.소재이름;
  newRow[DGPM_COL['보종']]           = data.보종        || '';
  newRow[DGPM_COL['광고유형']]       = data.광고유형    || '';
  newRow[DGPM_COL['소재유형']]       = data.소재유형    || '';
  newRow[DGPM_COL['소구포인트']]     = data.소구포인트  || '';
  newRow[DGPM_COL['후킹방식']]       = data.후킹방식    || '';
  newRow[DGPM_COL['소구상세']]       = data.소구상세    || '';
  newRow[DGPM_COL['이미지유형목록']] = data.이미지유형  || '';
  newRow[DGPM_COL['모델유형목록']]   = data.모델유형    || '';
  newRow[DGPM_COL['이미지수']]       = 1;
  newRow[DGPM_COL['이미지코드목록']] = imageCode;
  newRow[DGPM_COL['번들URL']]        = _getBundleUrl(unitCode);
  sheet.appendRow(newRow);
  return unitCode;
}

function _splitList(val) {
  return val ? String(val).split(',').map(s => s.trim()).filter(Boolean) : [];
}

function _getBundleUrl(unitCode) {
  try {
    const url = ScriptApp.getService().getUrl();
    return url ? url + '?unit=' + encodeURIComponent(unitCode) : '';
  } catch(e) {
    return '';
  }
}

function _generateDGPMCode(매체, sheet) {
  const prefix = 매체 === '피맥스' ? 'PM' : 'DG';
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd');
  const fullPrefix = prefix + dateStr;
  let max = 0;
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().forEach(v => {
      const s = String(v);
      if (s.startsWith(fullPrefix)) {
        const seq = parseInt(s.slice(-3), 10);
        if (!isNaN(seq) && seq > max) max = seq;
      }
    });
  }
  return fullPrefix + String(max + 1).padStart(3, '0');
}

// --------------------------------------------------
// 소재_마스터 기존 DG/PM 소재 → DG_PM_광고단위 재구성
// 소재_마스터의 DG/PM 행을 (매체,캠페인,그룹,소재이름) 기준으로 그룹핑하여
// DG_PM_광고단위 시트를 완전히 재작성
// --------------------------------------------------
function rebuildDGPMUnits() {
  try {
    const ss = getSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet || masterSheet.getLastRow() < 2)
      return { success: true, count: 0, message: '소재_마스터에 데이터가 없습니다.' };

    // 소재_마스터 전체 읽기 (모든 열)
    // 열: 0=이미지코드, 1=등록일자, 2=매체, 3=캠페인, 4=그룹, 5=소재이름,
    //     6=보종, 7=광고유형, 8=소재유형, 9=소구포인트, 10=후킹방식,
    //     11=소구상세, 12=이미지유형, 13=모델유형, 14=이미지URL, 15=파일해시
    const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 16).getValues();

    // DG/PM 행만 필터
    const dgpmRows = data.filter(r => DGPM_MEDIA.includes(String(r[2])));

    // (매체, 캠페인, 그룹, 소재이름) 기준 그룹핑
    const orderMap = [];
    const groupMap = {};

    dgpmRows.forEach(r => {
      const [imgCode, regDate, 매체, 캠페인, 그룹, 소재이름,
             보종, 광고유형, 소재유형, 소구포인트, 후킹방식, 소구상세, 이미지유형, 모델유형] = r;
      const key = [매체, 캠페인, 그룹, 소재이름].join('\x00');

      if (!groupMap[key]) {
        groupMap[key] = {
          매체, 캠페인, 그룹, 소재이름,
          // 공통 속성: 첫 번째 이미지 기준
          보종: String(보종 || ''),
          광고유형: String(광고유형 || ''),
          소재유형: String(소재유형 || ''),
          소구포인트: String(소구포인트 || ''),
          후킹방식: String(후킹방식 || ''),
          소구상세: String(소구상세 || ''),
          // 이미지별 속성: 고유값 수집
          imgTypes: [],
          modelTypes: [],
          codes: [],
          firstDate: regDate
        };
        orderMap.push(key);
      }
      const g = groupMap[key];
      if (imgCode && !g.codes.includes(String(imgCode))) g.codes.push(String(imgCode));
      const it = String(이미지유형 || '');
      const mt = String(모델유형   || '');
      if (it && !g.imgTypes.includes(it))   g.imgTypes.push(it);
      if (mt && !g.modelTypes.includes(mt)) g.modelTypes.push(mt);
    });

    // DG_PM_광고단위 시트 초기화 후 재작성
    let dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
    if (!dgpmSheet) {
      dgpmSheet = ss.insertSheet(DGPM_SHEET_NAME);
    } else {
      dgpmSheet.clearContents();
    }
    dgpmSheet.getRange(1, 1, 1, DGPM_HEADERS.length).setValues([DGPM_HEADERS]).setFontWeight('bold');
    dgpmSheet.setFrozenRows(1);

    if (!orderMap.length) return { success: true, count: 0, message: 'DG/PM 소재가 없습니다.' };

    const now = new Date();
    orderMap.forEach(key => {
      const g = groupMap[key];
      const unitCode = _generateDGPMCode(g.매체, dgpmSheet);
      const newRow = new Array(DGPM_HEADERS.length).fill('');
      newRow[DGPM_COL['광고단위코드']]   = unitCode;
      newRow[DGPM_COL['등록일시']]       = g.firstDate || now;
      newRow[DGPM_COL['최근수정일시']]   = now;
      newRow[DGPM_COL['매체']]           = g.매체;
      newRow[DGPM_COL['캠페인']]         = g.캠페인;
      newRow[DGPM_COL['그룹']]           = g.그룹;
      newRow[DGPM_COL['소재이름']]       = g.소재이름;
      newRow[DGPM_COL['보종']]           = g.보종;
      newRow[DGPM_COL['광고유형']]       = g.광고유형;
      newRow[DGPM_COL['소재유형']]       = g.소재유형;
      newRow[DGPM_COL['소구포인트']]     = g.소구포인트;
      newRow[DGPM_COL['후킹방식']]       = g.후킹방식;
      newRow[DGPM_COL['소구상세']]       = g.소구상세;
      newRow[DGPM_COL['이미지유형목록']] = g.imgTypes.join(',');
      newRow[DGPM_COL['모델유형목록']]   = g.modelTypes.join(',');
      newRow[DGPM_COL['이미지수']]       = g.codes.length;
      newRow[DGPM_COL['이미지코드목록']] = g.codes.join(',');
      newRow[DGPM_COL['번들URL']]        = _getBundleUrl(unitCode);
      dgpmSheet.appendRow(newRow);
    });

    return { success: true, count: orderMap.length, message: `광고단위 ${orderMap.length}개 재구성 완료 (소재 속성 포함)` };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// 기존 영상 URL 일괄 변환
// uc?export=view → /file/d/ID/preview (VID 코드 행만)
// --------------------------------------------------
function migrateVideoUrls() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, message: '변환할 데이터 없음' };

    const lastRow = sheet.getLastRow();
    const codes = sheet.getRange(2, 1,  lastRow - 1, 1).getValues();
    const urls  = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
    let updated = 0;

    for (let i = 0; i < codes.length; i++) {
      const code = String(codes[i][0] || '');
      const url  = String(urls[i][0]  || '');
      if (!code.startsWith('VID')) continue;          // VID 코드만
      if (url.includes('/file/d/'))  continue;         // 이미 /preview 형식
      const m = url.match(/[?&]id=([^&]+)/);
      if (!m) continue;
      sheet.getRange(i + 2, 15).setValue(
        'https://drive.google.com/file/d/' + m[1] + '/preview'
      );
      updated++;
    }
    return { success: true, message: updated + '개 영상 URL 변환 완료' };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// DG/PM 광고단위 목록 반환 (분석/조회용)
// --------------------------------------------------
function getDGPMList(매체필터) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(DGPM_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const totalCols = DGPM_HEADERS.length;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalCols).getValues();
  return data
    .filter(r => r[DGPM_COL['광고단위코드']] && (!매체필터 || r[DGPM_COL['매체']] === 매체필터))
    .map(r => ({
      광고단위코드:   String(r[DGPM_COL['광고단위코드']]),
      매체:           String(r[DGPM_COL['매체']]),
      캠페인:         String(r[DGPM_COL['캠페인']]),
      그룹:           String(r[DGPM_COL['그룹']]),
      소재이름:       String(r[DGPM_COL['소재이름']]),
      보종:           String(r[DGPM_COL['보종']]),
      광고유형:       String(r[DGPM_COL['광고유형']]),
      소재유형:       String(r[DGPM_COL['소재유형']]),
      이미지코드목록: String(r[DGPM_COL['이미지코드목록']]),
      이미지수:       r[DGPM_COL['이미지수']] || 0,
      번들URL:        String(r[DGPM_COL['번들URL']] || ''),
      등록일시:       r[DGPM_COL['등록일시']] ? Utilities.formatDate(new Date(r[DGPM_COL['등록일시']]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      최근수정일시:   r[DGPM_COL['최근수정일시']] ? Utilities.formatDate(new Date(r[DGPM_COL['최근수정일시']]), 'Asia/Seoul', 'yyyy-MM-dd') : ''
    }));
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

