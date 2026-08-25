// ====================================================
// 소재 인덱싱 시스템 - Code.gs
// ====================================================

const SHEET_ID                    = '1S74kKyOO3ATqk12nmQR860VT5ms8s5FS33T3gF5kd7I';
const MASTER_SHEET_NAME           = '전매체 인덱스';
const SETTINGS_SHEET_NAME         = '설정';
const HIERARCHY_SHEET_NAME        = '매체 구조';
const RAW_SHEET_NAME              = '매체_RAW';
const DETECT_SHEET_NAME           = '신규소재감지';
const DGPM_SHEET_NAME             = '구글DA 인덱스';
const CONSOLIDATED_RAW_SHEET_NAME = '소재_통합RAW';
const AB_TEST_SHEET_NAME          = 'AB테스트';
const TARGETING_AB_SHEET_NAME     = '타겟팅AB테스트';
const SAVED_REPORT_SHEET_NAME     = '저장된보고서';

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

// 소재_통합RAW 시트 헤더 (A~U, 21열)
const CONSOLIDATED_RAW_HEADERS = [
  '매체', '일', '캠페인', '광고그룹', '소재이름',
  '보종', '광고유형', '소재유형', '소구포인트', '후킹방식', '소구상세',
  '이미지유형', '모델유형', '이미지URL',
  '노출수', '클릭수', 'CTR', '비용', '전환', 'CVR', 'CPA'
];

// AB테스트 시트 헤더
// 매체/캠페인/그룹은 뒤에 이어붙인 컬럼 — 기존에 저장된 시트(이 3개 컬럼이
// 없던 버전)와의 호환을 위해 중간에 끼워넣지 않고 항상 맨 끝에 추가한다.
// (_getOrCreateABTestSheet가 기존 시트를 열 때 헤더 행을 이 길이에 맞춰 늘림)
const AB_TEST_HEADERS = [
  '테스트ID', '테스트명', '가설', '변경요소', '타겟',
  '시작일', '종료일', '소재코드목록', '결론메모',
  '등록일시', '최근수정일시',
  '매체', '캠페인', '그룹'
];

// 저장된보고서 시트 헤더 — 보고서 설정(행/열 기준, 지표, 필터, 기간 등)은
// 통째로 JSON 문자열 하나에 담아 저장한다 (구조가 자주 확장될 수 있어 컬럼을
// 미리 다 나누기보다 유연하게 가져가는 편이 유지보수하기 쉽다)
const SAVED_REPORT_HEADERS = ['보고서ID', '보고서명', '설정JSON', '등록일시', '최근수정일시'];

// 타겟팅AB테스트 시트 헤더 — 소재 AB테스트와 달리 비교 단위가 소재가 아니라
// (매체,캠페인,그룹) 조합이라 슬롯 2~4개를 통째로 JSON 배열([{매체,캠페인,그룹,메모}])
// 하나에 담아 저장한다(저장된보고서와 동일한 패턴).
const TARGETING_AB_HEADERS = [
  '테스트ID', '테스트명', '가설', '시작일', '종료일', '슬롯JSON', '결론메모',
  '등록일시', '최근수정일시'
];

// --------------------------------------------------
// 웹앱 진입점
// --------------------------------------------------
function doGet(e) {
  const unitCode = e && e.parameter && e.parameter.unit;
  // 대시보드에서 번들 썸네일/아이콘을 클릭할 때, 소재_통합RAW에 박혀있는 unitCode가
  // 이미 재구성으로 stale해졌을 수 있어 _media/_campaign/_group/_name도 같이 실어보낸다
  // (getBundleData 주석 참고). 이게 없으면 bundleData가 null이 되어 번들 뷰 대신
  // 기본 화면(소재등록)으로 떨어지는 버그가 있었다(2026-07-30).
  const ctx = (e && e.parameter) ? {
    매체:     e.parameter._media    || '',
    캠페인:   e.parameter._campaign || '',
    그룹:     e.parameter._group    || '',
    소재이름: e.parameter._name     || ''
  } : null;
  const template = HtmlService.createTemplateFromFile('Index');
  template.bundleData = unitCode ? getBundleData(unitCode, ctx) : null;
  template.unitCode   = unitCode || '';
  return template.evaluate()
    .setTitle(unitCode ? '광고단위 소재' : '운영 소재 분석')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// doGet에서 Blob을 직접 반환하는 방식은 이 프로젝트에서 "지원되는 반환 형식이 아닙니다" 오류로
// 동작하지 않아(2026-08-06 실측), 검수 결과 썸네일은 base64 data URI를 문자열로 반환해 클라이언트에서
// <img src="data:..."> 로 바로 그리는 방식을 쓴다.
function getImageDataUri(fileId) {
  if (!fileId) return null;
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return null;
  }
}

// 보고서 탭 엑셀 다운로드에서 "이미지 URL" 열을 실제 그림으로 박아넣을 때 쓴다.
// 브라우저가 Drive 썸네일 URL을 직접 fetch하면 CORS에 막혀 바이트를 못 읽는 경우가 있어,
// 서버가 대신 UrlFetchApp으로 받아 base64로 돌려준다(getImageDataUri와 같은 패턴, 다만
// 이쪽은 fileId가 아니라 URL 자체를 받고, 여러 개를 한 번에 처리한다).
// 반환값: { [url]: { base64, contentType } } — 실패한 URL은 결과에서 그냥 빠진다.
function fetchReportImages(urls) {
  const out = {};
  (urls || []).forEach(url => {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return;
      const blob = res.getBlob();
      out[url] = {
        base64: Utilities.base64Encode(blob.getBytes()),
        contentType: blob.getContentType() || 'image/jpeg'
      };
    } catch (e) {
      // 개별 이미지 실패는 건너뛰고 나머지는 계속 처리
    }
  });
  return out;
}

// --------------------------------------------------
// 커스텀 메뉴 (스프레드시트 열릴 때 자동 등록)
// [통합 적재] 버튼 대체 수단으로도 활용
// --------------------------------------------------
function openDashboard() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('웹앱이 배포되지 않았습니다.\n배포 후 다시 시도하세요.');
    return;
  }
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`
  ).setHeight(1).setWidth(1);
  SpreadsheetApp.getUi().showModalDialog(html, '대시보드 열기...');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('소재 인덱싱')
    .addItem('📊 대시보드 열기', 'openDashboard')
    .addSeparator()
    .addItem('통합 적재', 'consolidateRawData')
    .addSeparator()
    .addItem('시트 초기화', 'initializeSheets')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('이미지 검수')
    .addItem('🔍 검수 대시보드 열기 (소재 검수 탭)', 'openDashboard')
    .addSeparator()
    .addItem('검수 기준 시트 열기', 'openInspectionCriteriaSheet')
    .addItem('API 키 설정', 'promptSetOpenAIKey')
    .addSeparator()
    .addItem('검수 시트 초기화', 'initInspectionSheets')
    .addToUi();
}

// --------------------------------------------------
// 외부 요청 권한 확인용 (편집기에서 한 번 실행 → UrlFetchApp 권한 동의)
// --------------------------------------------------
function testExternalFetch() {
  const res = UrlFetchApp.fetch('https://httpbin.org/get', { muteHttpExceptions: true });
  Logger.log('status: ' + res.getResponseCode());
  return 'ok: ' + res.getResponseCode();
}

// --------------------------------------------------
// 번들 뷰용 광고단위 데이터 조회
//
// ctx(선택): {매체, 캠페인, 그룹, 소재이름} — unitCode로 못 찾을 때의 폴백 매칭용.
// "광고단위 재구성"을 실행하면 예전엔 코드가 전부 새로 발급됐는데, 소재_통합RAW의
// 이미지URL(?unit=코드)은 append-only라 예전 코드가 그대로 남아있어 더 이상 코드만으로는
// 못 찾는 경우가 있었다(2026-07-30, "구글DA 소재 호버 시 데이터 없음" 버그). 이때 같은
// (매체,캠페인,그룹,소재이름) 조합을 대신 찾는다 — azCreativeContextKey와 동일한 식별 기준.
// --------------------------------------------------
function getBundleData(unitCode, ctx, opts) {
  const ss = getSpreadsheet();
  const dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
  if (!dgpmSheet || dgpmSheet.getLastRow() < 2) return null;

  const rows = dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues();
  let unitRow = rows.find(r => String(r[DGPM_COL['광고단위코드']]) === unitCode);
  if (!unitRow && ctx && ctx.매체 && ctx.캠페인 && ctx.그룹 && ctx.소재이름) {
    unitRow = rows.find(r =>
      String(r[DGPM_COL['매체']])     === ctx.매체     &&
      String(r[DGPM_COL['캠페인']])   === ctx.캠페인   &&
      String(r[DGPM_COL['그룹']])     === ctx.그룹     &&
      String(r[DGPM_COL['소재이름']]) === ctx.소재이름
    );
  }
  if (!unitRow) return null;

  const imageCodes = _splitList(unitRow[DGPM_COL['이미지코드목록']]);

  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const images = [];
  if (masterSheet && masterSheet.getLastRow() >= 2 && imageCodes.length) {
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 15).getValues();
    const codeMap = {};
    masterData.forEach(r => {
      const code = String(r[0] || '').trim();
      if (code && !codeMap[code]) {
        codeMap[code] = { code, url: String(r[14] || ''), 소재이름: String(r[5] || '') };
      }
    });
    imageCodes.forEach(code => { if (codeMap[code]) images.push(codeMap[code]); });
  }

  // Drive 공개 썸네일(drive.google.com/thumbnail?id=...) 핫링크는 간헐적으로 막혀서
  // "구글DA 번들 이미지가 안 보인다"는 문제가 있었다(2026-07-28) — 브라우저가 매번
  // drive.google.com에 별도로 요청해야 하는데, 파일 소유자/뷰어 인증 상태나 Google의
  // 핫링크 방지 정책에 따라 종종 실패한다. 번들 뷰는 이미지 개수가 적으니(광고단위 1개
  // 분량) 여기서 직접 blob을 읽어 base64 data URI로 내려준다 — Drive에 별도 요청이
  // 전혀 필요 없어져 훨씬 안정적이다. 동영상(/preview URL)은 원래도 iframe 임베드라 대상
  // 아니고, blob 조회가 실패하면(권한 등) 기존 URL로 조용히 폴백한다.
  //
  // opts.skipDataUri: 소재명 호버 미리보기(azPreviewShow)는 이 dataUri를 전혀 쓰지 않고
  // driveThumb() 핫링크로만 그리는데도, 예전엔 여기서 매번 이미지 전체를 Drive에서 읽어
  // base64로 인코딩해놓고 그 결과를 그냥 버리고 있었다 — 호버가 느렸던 주 원인이었다
  // (2026-07-30). 번들 전체보기(doGet)만 dataUri를 실제로 쓰므로 그때는 계속 계산한다.
  if (!(opts && opts.skipDataUri)) images.forEach(img => {
    if (!img.url || img.url.indexOf('/preview') !== -1) return;
    try {
      const m = img.url.match(/[?&]id=([^&]+)/) || img.url.match(/\/d\/([^/?]+)/);
      if (!m) return;
      const blob = DriveApp.getFileById(m[1]).getBlob();
      img.dataUri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      // 폴백: dataUri 없이 반환 — 클라이언트가 기존 driveThumb(img.url) 핫링크 방식 사용
    }
  });

  return {
    광고단위코드: String(unitRow[DGPM_COL['광고단위코드']]),
    매체:         String(unitRow[DGPM_COL['매체']]),
    캠페인:       String(unitRow[DGPM_COL['캠페인']]),
    그룹:         String(unitRow[DGPM_COL['그룹']]),
    소재이름:     String(unitRow[DGPM_COL['소재이름']]),
    보종:         String(unitRow[DGPM_COL['보종']]),
    광고유형:     String(unitRow[DGPM_COL['광고유형']]),
    소재유형:     String(unitRow[DGPM_COL['소재유형']]),
    이미지수:     Number(unitRow[DGPM_COL['이미지수']] || 0),
    images
  };
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

  // 소재_통합RAW 시트
  if (!ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME)) {
    const cSheet = ss.insertSheet(CONSOLIDATED_RAW_SHEET_NAME);
    cSheet.getRange(1, 1, 1, CONSOLIDATED_RAW_HEADERS.length)
      .setValues([CONSOLIDATED_RAW_HEADERS]).setFontWeight('bold');
    cSheet.setFrozenRows(1);
  }

  // AB테스트 시트
  if (!ss.getSheetByName(AB_TEST_SHEET_NAME)) {
    const abSheet = ss.insertSheet(AB_TEST_SHEET_NAME);
    abSheet.getRange(1, 1, 1, AB_TEST_HEADERS.length)
      .setValues([AB_TEST_HEADERS]).setFontWeight('bold');
    abSheet.setFrozenRows(1);
  }

  // 저장된보고서 시트
  if (!ss.getSheetByName(SAVED_REPORT_SHEET_NAME)) {
    const srSheet = ss.insertSheet(SAVED_REPORT_SHEET_NAME);
    srSheet.getRange(1, 1, 1, SAVED_REPORT_HEADERS.length)
      .setValues([SAVED_REPORT_HEADERS]).setFontWeight('bold');
    srSheet.setFrozenRows(1);
  }

  // 타겟팅AB테스트 시트
  if (!ss.getSheetByName(TARGETING_AB_SHEET_NAME)) {
    const tabSheet = ss.insertSheet(TARGETING_AB_SHEET_NAME);
    tabSheet.getRange(1, 1, 1, TARGETING_AB_HEADERS.length)
      .setValues([TARGETING_AB_HEADERS]).setFontWeight('bold');
    tabSheet.setFrozenRows(1);
  }

  initInspectionSheets();

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
  const seen = new Map();

  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
    // col: 0=이미지코드, 1=등록일자, 2=매체, 3=캠페인, 4=그룹, 5=소재이름, 6=보종, 8=소재유형, 14=이미지URL
    data.forEach(row => {
      const code = String(row[0] || '').trim();
      if (!code || seen.has(code)) return;
      seen.set(code, {
        code,
        등록일자: row[1] ? String(row[1]).slice(0, 10) : '',
        매체:     String(row[2] || ''),
        캠페인:   String(row[3] || ''),
        그룹:     String(row[4] || ''),
        소재이름: String(row[5] || ''),
        보종:     String(row[6] || ''),
        소재유형: String(row[8] || ''),
        imageUrl: String(row[14] || '')
      });
    });
  }

  // 디멘드젠/피맥스는 광고단위(번들) 단위로 구글DA 인덱스에 등록되는데, 소재 등록
  // 탭의 이미지 업로드를 거치지 않고 등록된 경우 전매체 인덱스에 개별 이미지코드가
  // 없다 — 이런 소재는 AB 테스트 등 이미지코드 기반 기능에서 아예 선택할 수 없었다
  // (2026-08-20 확인). 광고단위코드를 코드처럼 취급해 목록에 함께 포함시킨다.
  const dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
  if (dgpmSheet && dgpmSheet.getLastRow() >= 2) {
    const dgpmRows = dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues();
    dgpmRows.forEach(row => {
      const code = String(row[DGPM_COL['광고단위코드']] || '').trim();
      if (!code || seen.has(code)) return;
      const firstImageCode = _splitList(row[DGPM_COL['이미지코드목록']])[0];
      const bundleImageUrl = (firstImageCode && seen.has(firstImageCode))
        ? seen.get(firstImageCode).imageUrl
        : String(row[DGPM_COL['번들URL']] || '');
      seen.set(code, {
        code,
        등록일자: row[DGPM_COL['등록일시']] ? String(row[DGPM_COL['등록일시']]).slice(0, 10) : '',
        매체:     String(row[DGPM_COL['매체']]     || ''),
        캠페인:   String(row[DGPM_COL['캠페인']]   || ''),
        그룹:     String(row[DGPM_COL['그룹']]     || ''),
        소재이름: String(row[DGPM_COL['소재이름']] || ''),
        보종:     String(row[DGPM_COL['보종']]     || ''),
        소재유형: String(row[DGPM_COL['소재유형']] || ''),
        imageUrl: bundleImageUrl
      });
    });
  }

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
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
    const rows = data.filter(r => String(r[0]).trim() === imageCode);
    if (rows.length) {
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
  }

  // 전매체 인덱스에 없으면 디멘드젠/피맥스 광고단위코드로 폴백한다 — 소재 등록 탭의
  // 이미지 업로드를 거치지 않고 등록된 DG/PM 소재는 개별 이미지코드가 전매체 인덱스에
  // 없고 구글DA 인덱스 시트에 광고단위(번들) 단위로만 있어(getImageCodes()가 이런
  // 소재를 픽커 목록에 광고단위코드로 포함시키는 것과 같은 이유), 여기서도 못 찾으면
  // "정보 불러오기"가 아무 반응 없이 조용히 실패했다(2026-08-25 확인).
  const dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
  if (dgpmSheet && dgpmSheet.getLastRow() >= 2) {
    const dgpmRows = dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues();
    const dgpmRow = dgpmRows.find(r => String(r[DGPM_COL['광고단위코드']] || '').trim() === imageCode);
    if (dgpmRow) {
      const firstImageCode = _splitList(dgpmRow[DGPM_COL['이미지코드목록']])[0];
      let resolvedImageUrl = '';
      if (firstImageCode && sheet && sheet.getLastRow() >= 2) {
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
        const match = data.find(r => String(r[0]).trim() === firstImageCode);
        if (match) resolvedImageUrl = String(match[14] || '');
      }
      return {
        매체:       String(dgpmRow[DGPM_COL['매체']]     || ''),
        캠페인:     String(dgpmRow[DGPM_COL['캠페인']]   || ''),
        그룹:       String(dgpmRow[DGPM_COL['그룹']]     || ''),
        소재이름:   String(dgpmRow[DGPM_COL['소재이름']] || ''),
        보종:       String(dgpmRow[DGPM_COL['보종']]     || ''),
        광고유형:   String(dgpmRow[DGPM_COL['광고유형']] || ''),
        소재유형:   String(dgpmRow[DGPM_COL['소재유형']] || ''),
        소구포인트: String(dgpmRow[DGPM_COL['소구포인트']] || ''),
        후킹방식:   String(dgpmRow[DGPM_COL['후킹방식']]   || ''),
        소구상세:   String(dgpmRow[DGPM_COL['소구상세']]   || ''),
        이미지유형: _splitList(dgpmRow[DGPM_COL['이미지유형목록']])[0] || '',
        모델유형:   _splitList(dgpmRow[DGPM_COL['모델유형목록']])[0] || '',
        이미지URL:  resolvedImageUrl || String(dgpmRow[DGPM_COL['번들URL']] || '')
      };
    }
  }

  return null;
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
        _splitList(data.모델유형).forEach(mt => { if (!modelTypes.includes(mt)) modelTypes.push(mt); });

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
      if (it && !g.imgTypes.includes(it)) g.imgTypes.push(it);
      _splitList(모델유형).forEach(mt => { if (!g.modelTypes.includes(mt)) g.modelTypes.push(mt); });
    });

    // DG_PM_광고단위 시트 초기화 후 재작성
    let dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);

    // 지우기 전에 기존 (매체,캠페인,그룹,소재이름) → 광고단위코드 매핑을 저장해둔다.
    // 소재_통합RAW의 이미지URL(?unit=코드)은 append-only라 예전 코드가 그대로 남아있는데,
    // 재구성 때마다 코드를 새로 발급하면 그 링크가 전부 끊긴다(2026-07-30 실제 발생 —
    // "구글DA 소재 호버 시 데이터 없음"). 이미 있던 그룹은 코드를 그대로 유지한다.
    const existingCodeMap = {};
    if (dgpmSheet && dgpmSheet.getLastRow() >= 2) {
      dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues().forEach(r => {
        const key  = [r[DGPM_COL['매체']], r[DGPM_COL['캠페인']], r[DGPM_COL['그룹']], r[DGPM_COL['소재이름']]].join('\x00');
        const code = String(r[DGPM_COL['광고단위코드']] || '');
        if (code && !existingCodeMap[key]) existingCodeMap[key] = code;
      });
    }

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
      const unitCode = existingCodeMap[key] || _generateDGPMCode(g.매체, dgpmSheet);
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
// DG/PM 광고단위코드 → 대표 이미지URL 맵
// 소재_통합RAW의 이미지URL은 개별 이미지가 아니라 번들(?unit=코드) 링크라서,
// 그 unitCode로 DG_PM_광고단위의 이미지코드목록 중 첫 번째 코드를 찾고
// 소재_마스터에서 그 코드의 실제 이미지URL을 붙여 대표 이미지로 쓴다.
// (소재이름으로 매칭하면 RAW의 소재이름과 소재_마스터 소재이름이 다를 수 있어 불안정하므로
//  같은 코드베이스 안에서 생성된 광고단위코드로 정확히 매칭한다)
// --------------------------------------------------
function getDGPMThumbMap() {
  const ss = getSpreadsheet();
  const dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!dgpmSheet || dgpmSheet.getLastRow() < 2) return {};

  const codeToUrl = {};
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 15).getValues();
    data.forEach(row => {
      const code = String(row[0]  || '').trim();
      const url  = String(row[14] || '').trim();
      if (code && url && !codeToUrl[code]) codeToUrl[code] = url;
    });
  }

  const rows = dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues();
  const result = {};
  rows.forEach(row => {
    const unitCode = String(row[DGPM_COL['광고단위코드']] || '').trim();
    if (!unitCode) return;
    const codes = _splitList(row[DGPM_COL['이미지코드목록']]);
    for (const code of codes) {
      if (codeToUrl[code]) { result[unitCode] = codeToUrl[code]; break; }
    }
  });
  return result;
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

// --------------------------------------------------
// 소재 분석용: 소재_통합RAW 전체 데이터 반환 (1행=헤더 포함)
// --------------------------------------------------
// 연결 확인 및 시트 상태 진단 (Index.html에서 testPing() 호출)
function testPing() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  return {
    status: 'ok',
    sheetFound: sheet !== null,
    lastRow: sheet ? sheet.getLastRow() : -1,
    lookingFor: CONSOLIDATED_RAW_SHEET_NAME,
    allSheets: ss.getSheets().map(s => s.getName())
  };
}

// 소재_통합RAW의 '일' 컬럼 인덱스 (CONSOLIDATED_RAW_HEADERS 기준) — Date 객체 변환 시
// 이 컬럼만 검사하면 되므로, 21개 컬럼 전부를 매번 instanceof 체크하던 것보다 훨씬 가볍다.
const CONSOLIDATED_RAW_DATE_COL = 1;

function getAllData() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const raw = sheet.getRange(1, 1, lastRow, CONSOLIDATED_RAW_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  // google.script.run은 Date 객체 직렬화 실패 시 null 반환 → 문자열로 변환 ('일' 컬럼만 검사)
  for (let i = 1; i < raw.length; i++) { // 0번째(헤더 행)는 건너뜀
    const cell = raw[i][CONSOLIDATED_RAW_DATE_COL];
    if (cell instanceof Date) raw[i][CONSOLIDATED_RAW_DATE_COL] = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
  }
  return raw;
}

// getAllDataSince의 경계 행(캐시가 마지막으로 갖고 있는 행)이 지금도 그때와 같은
// 내용인지 대조하기 위한 경량 지문. Date는 두 호출(getAllData/getAllDataSince) 모두
// 이미 'yyyy-MM-dd' 문자열로 바꿔서 클라이언트에 내려주므로, 여기서도 같은 방식으로
// 맞춰서 비교해야 한다.
function _rowFingerprint_(rowValues) {
  const tz = Session.getScriptTimeZone();
  return rowValues
    .map(v => (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v))
    .join('');
}

// 이전에 불러온 시트 행 번호(sinceLastRow, 헤더 포함) 이후로 새로 추가된 행만 반환.
// 소재_통합RAW은 기존 행을 수정하지 않고 뒤에만 추가(append-only)하는 구조라서
// 클라이언트가 이미 가진 데이터에 이어붙이기만 하면 되고, 그러면 데이터가 아무리
// 쌓여도 로드 시간은 "지난번 이후 늘어난 양"에만 비례하게 된다.
//
// sinceFingerprint: 클라이언트가 캐시에 저장해둔 sinceLastRow번째 행의 지문
// (azFingerprint, Index.html). append-only 전제가 실제로 지켜졌다면 지금 시트의
// sinceLastRow번째 행도 그 지문과 같아야 한다.
function getAllDataSince(sinceLastRow, sinceFingerprint) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  if (!sheet) return { rows: [], lastRow: 0 };
  const lastRow = sheet.getLastRow();
  // 시트 행 수가 클라이언트가 마지막으로 알고 있던 값보다 줄었다는 건 중간에 행이
  // 삭제됐다는 뜻이다 — "이후 늘어난 만큼만" 델타 로직은 append-only를 전제하므로
  // 이 경우 새로 쌓인 데이터가 있어도 조용히 빈 델타를 반환해 대시보드가 오래된
  // 캐시에 영구히 멈춰버린다(2026-07-31, 디멘드젠 데이터 삭제 후 새로 적재한 피맥스
  // 데이터가 대시보드에 하나도 안 보이는 문제로 확인). 이 경우를 명시적으로 알려서
  // 클라이언트가 캐시를 버리고 전체를 다시 받도록 한다.
  if (lastRow < sinceLastRow) return { rows: [], lastRow, invalidated: true };

  // 위 행 수 비교만으로는 "중간에서 몇 행이 지워지고 그 이상(혹은 같은 수)이 다시
  // 쌓여 lastRow가 우연히 줄지 않은" 경우를 놓친다 — 이때도 append-only 전제가
  // 깨졌으므로 캐시의 나머지 부분(2행~sinceLastRow행)이 실제 시트와 더 이상 대응하지
  // 않는데, 행 수만 보면 정상적인 증가로 보여 조용히 잘못된 데이터가 이어붙여진다.
  // sinceLastRow 위치의 실제 내용을 캐시가 기억하는 지문과 대조해 이 케이스를 잡는다.
  if (sinceLastRow >= 2 && sinceFingerprint) {
    const boundaryRow = sheet.getRange(sinceLastRow, 1, 1, CONSOLIDATED_RAW_HEADERS.length).getValues()[0];
    if (_rowFingerprint_(boundaryRow) !== sinceFingerprint) {
      return { rows: [], lastRow, invalidated: true };
    }
  }

  if (lastRow < 2 || lastRow <= sinceLastRow) return { rows: [], lastRow: Math.max(lastRow, 0) };

  const startRow = Math.max(sinceLastRow + 1, 2); // 1행은 헤더라 항상 건너뜀
  const numRows = lastRow - startRow + 1;
  const raw = sheet.getRange(startRow, 1, numRows, CONSOLIDATED_RAW_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 0; i < raw.length; i++) {
    const cell = raw[i][CONSOLIDATED_RAW_DATE_COL];
    if (cell instanceof Date) raw[i][CONSOLIDATED_RAW_DATE_COL] = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
  }
  return { rows: raw, lastRow };
}

// --------------------------------------------------
// 소재 분석용: 전매체 인덱스 전체 데이터 반환 (1행=헤더 포함, 생애주기 분석용)
// --------------------------------------------------
function getSpreadsheetUrl() {
  return getSpreadsheet().getUrl();
}

// --------------------------------------------------
function getMasterData() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, sheet.getLastRow(), 16).getValues();
}

// --------------------------------------------------
// API 키 관리 (스크립트 속성)
// --------------------------------------------------
const MANAGED_API_KEYS = [
  { id: 'OPENAI_API_KEY', label: 'OpenAI API Key', description: 'AI 인사이트 분석 (소재 분석 탭) · 소재 검수 이미지 분석 (소재 검수 탭)' }
];

function getApiKeys() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return MANAGED_API_KEYS.map(k => {
    const val = props[k.id] || '';
    return {
      id:          k.id,
      label:       k.label,
      description: k.description,
      isSet:       !!val,
      masked:      val ? val.slice(0, 4) + '••••••••' + val.slice(-4) : ''
    };
  });
}

function saveApiKey(keyId, value) {
  const allowed = MANAGED_API_KEYS.map(k => k.id);
  if (!allowed.includes(keyId)) return { error: true, message: '허용되지 않은 키입니다.' };
  try {
    const props = PropertiesService.getScriptProperties();
    if (value && value.trim()) {
      props.setProperty(keyId, value.trim());
    } else {
      props.deleteProperty(keyId);
    }
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// OpenAI 인사이트 생성
// PropertiesService 'OPENAI_API_KEY' → gpt-5 호출 → 텍스트 반환
// --------------------------------------------------
function getOpenAIInsight(aggregatedData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) return {
    error: true,
    message: 'OPENAI_API_KEY가 설정되지 않았습니다.\nApps Script 편집기 > 프로젝트 설정 > 스크립트 속성에서 OPENAI_API_KEY를 추가하세요.'
  };

  // 매체/캠페인/광고그룹/소재이름/보종 필터가 걸려 있으면 그 맥락에 맞는 운영
  // 인사이트를 별도 섹션으로 요청한다. 프론트에서 넘기는 필터컨텍스트.유형은
  // "현재단계별_다음단계"(예: "캠페인별_광고그룹") 형식이라, 어떤 계층 조합이
  // 오든 그 두 라벨만으로 지시문을 만들 수 있다. 보종 케이스만 계층이 달라 예외 처리.
  const fc = aggregatedData.필터컨텍스트;
  let filterInstruction = '';
  if (fc && fc.유형 === '보종별_매체전략') {
    filterInstruction = `현재 "${fc.보종}" 보종으로 필터가 걸려 있습니다. 데이터의 매체별 성과를 바탕으로, ` +
      `이 보종을 운영할 때 매체별로 예산 배분이나 소재 전략을 어떻게 다르게 가져가야 하는지 2~3개 작성하세요.`;
  } else if (fc && fc.유형 && fc.유형.indexOf('별_') !== -1) {
    const [curLabel, nextLabel] = fc.유형.split('별_');
    const curVal = fc[curLabel];
    filterInstruction = `현재 "${curVal}" ${curLabel}(으)로 필터가 걸려 있습니다. ` +
      `데이터의 ${nextLabel}별 성과를 바탕으로, 이 ${curLabel} 안에서 ${nextLabel}별로 어떤 운영·예산·소재 전략이 필요한지 2~3개 작성하세요.`;
  }

  const sectionInstruction = fc
    ? `응답은 반드시 아래 형식을 그대로 지켜서 두 섹션으로 나누어 작성하세요 (마크다운, 설명 문구 없이):
[[필터인사이트]]
${filterInstruction}
번호를 붙여 줄바꿈으로 구분하세요.
[[/필터인사이트]]
[[소재인사이트]]
아래 소재유형·소구포인트·후킹방식·이미지유형·모델유형 데이터를 바탕으로, 소재 자체의 운영 인사이트를 3~5개 작성하세요.
번호를 붙여 줄바꿈으로 구분하세요.
[[/소재인사이트]]`
    : `응답은 반드시 아래 형식을 그대로 지켜서 작성하세요 (마크다운, 설명 문구 없이):
[[소재인사이트]]
아래 소재유형·소구포인트·후킹방식·이미지유형·모델유형 데이터를 바탕으로, 소재 자체의 운영 인사이트를 3~5개 작성하세요.
번호를 붙여 줄바꿈으로 구분하세요.
[[/소재인사이트]]`;

  // 2026-07-28: 데이터에 일자별추이(또는 기간이 길면 주간 버킷)가 포함되어 있으니, 선택
  // 기간의 총합/평균만 보고 판단하지 말고 추세(상승/하락, 특정 구간의 급변, 변동성)를
  // 반영하라고 명시적으로 지시한다 — 전엔 총합만 넘겨서 "최근 며칠새 CTR이 떨어지고
  // 있다" 같은 판단을 AI가 아예 할 수 없었다.
  const trendInstruction = (aggregatedData.일자별추이 && aggregatedData.일자별추이.length > 1)
    ? `\n\n일자별추이 데이터가 포함되어 있습니다 — 기간 전체의 총합/평균만 보지 말고, 추세(상승/하락 흐름, 특정 구간의 급격한 변화, 변동성)를 반드시 반영해서 분석하세요. "최근 들어 ~가 개선/악화되고 있어 ~를 검토하세요" 같은 시계열 근거를 최소 1개 이상 포함하세요.`
    : '';

  const prompt = `당신은 한화손해보험 DA 광고 소재 전략 전문가입니다.
아래는 현재 필터 조건에서 집계된 성과 데이터입니다.
광고 운영자가 즉시 활용할 수 있도록, 단순 수치 나열이 아닌 "~하기 때문에 ~를 검토하세요" 형식으로 한국어로 작성하세요.${trendInstruction}

${sectionInstruction}

${JSON.stringify(aggregatedData, null, 2)}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: 'gpt-5-mini',
      messages: [{ role: 'user', content: prompt }],
      // gpt-5 계열은 max_tokens가 아니라 max_completion_tokens를 쓰고, temperature도 기본값(1)
      // 외엔 거부한다(2026-07-28 실제 호출로 확인: 둘 다 넣으면 400 에러). reasoning_effort를
      // 안 주면 기본 reasoning에 토큰을 다 쓰고 정작 답변 내용은 비어서 나오는 것도 실측
      // 확인했다 — 데이터 분석·전략 제안이라 분류 작업보다는 약간 여유를 둬서 'low'로 설정.
      max_completion_tokens: 1600,
      reasoning_effort: 'low'
    }),
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    const json = JSON.parse(resp.getContentText());
    if (json.error) return { error: true, message: json.error.message };
    const text = json.choices[0].message.content;
    // reasoning 모델이 토큰 예산을 전부 reasoning에 써버리면 content가 빈 문자열인 채로
    // finish_reason='length'가 온다 — 이걸 그냥 success:true로 넘기면 화면엔 빈 결과만
    // 남아 "답이 안 나온다"로 보인다(2026-07-28 채팅에서 실제로 겪은 문제). 명확한 에러로 변환.
    if (!text && json.choices[0].finish_reason === 'length') {
      return { error: true, message: '응답이 비어 있습니다(토큰 한도 초과) — 다시 시도해보세요.' };
    }
    return { success: true, text };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// --------------------------------------------------
// AI 인사이트 챗봇 — 생성된 인사이트에 대한 후속 질문
// messages: [{role:'system'|'user'|'assistant', content:'...'}] 형태의 전체 대화 기록을
// 그대로 받아 OpenAI에 전달한다 (컨텍스트는 클라이언트가 system 메시지로 구성해서 보냄)
// --------------------------------------------------
function getOpenAIChatReply(messages) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) return {
    error: true,
    message: 'OPENAI_API_KEY가 설정되지 않았습니다.\nApps Script 편집기 > 프로젝트 설정 > 스크립트 속성에서 OPENAI_API_KEY를 추가하세요.'
  };
  if (!Array.isArray(messages) || !messages.length) return { error: true, message: '대화 내용이 없습니다.' };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: 'gpt-5-mini',
      messages: messages,
      // "질문하면 답이 안 나온다"는 신고를 실제 API 호출로 재현했다(2026-07-28): reasoning_effort
      // 'low' + max_completion_tokens 900이면 짧은 후속 질문에도 모델이 900 토큰을 전부
      // reasoning에 써버려서 finish_reason=length, 실제 답변 내용은 빈 문자열로 나왔다.
      // 'minimal'로 낮추고 여유분도 늘려 같은 조건에서 정상 답변 나오는 것 확인.
      max_completion_tokens: 1200,
      reasoning_effort: 'minimal'
    }),
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    const json = JSON.parse(resp.getContentText());
    if (json.error) return { error: true, message: json.error.message };
    const text = json.choices[0].message.content;
    // reasoning 모델이 토큰 예산을 전부 reasoning에 써버리면 content가 빈 문자열인 채로
    // finish_reason='length'가 온다 — 이걸 그냥 success:true로 넘기면 화면엔 빈 결과만
    // 남아 "답이 안 나온다"로 보인다(2026-07-28 채팅에서 실제로 겪은 문제). 명확한 에러로 변환.
    if (!text && json.choices[0].finish_reason === 'length') {
      return { error: true, message: '응답이 비어 있습니다(토큰 한도 초과) — 다시 시도해보세요.' };
    }
    return { success: true, text };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// RAW_ 소스 시트마다 내보내는 방식이 달라(구글 광고 UI 직접 붙여넣기 vs 다른 도구 export
// 등) 같은 날짜·숫자도 표현이 다를 수 있다 — 일(B)이 Date 객체/"2026-07-30"/"2026.7.30"/
// 시리얼 넘버로 제각각 들어오거나, 비용 등 숫자 컬럼이 "1,234" 텍스트로 들어오는 식.
// 이 차이를 그대로 두면 중복 체크 키(_rowKey)가 문자열 비교라 실제로는 같은 날짜인데
// 다른 값으로 찍혀 중복 행이 쌓이고, 그 행의 비용·전환이 겹쳐 합산돼 CPA가 실제보다
// 크게(또는 작게) 집계되는 원인이 된다. 적재 시점에 정규화해서 통일된 형식으로 저장한다
// (2026-07-31, "구글DA 랭킹 CPA가 비정상적으로 높게 나온다" 문의 확인 중 발견).
function _normDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400000));
  const s = String(val || '').trim();
  const m = s.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const m2 = s.match(/^(\d{4})(\d{2})(\d{2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
}

function _normNum(val) {
  if (typeof val === 'number') return val;
  const n = Number(String(val || '').replace(/[,\s]/g, ''));
  return isNaN(n) ? val : n;
}

// 일(B)·숫자 컬럼(노출수/클릭수/비용/전환, O/P/R/S)·주요 텍스트 필드의 형식을 통일한다.
// row는 RAW_ 시트에서 읽은 21열(A~U) 원본 행.
function _normalizeRawRow(row) {
  const out = row.slice();
  [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].forEach(i => {
    if (typeof out[i] === 'string') out[i] = out[i].trim();
  });
  out[1] = _normDate(out[1]);
  [14, 15, 17, 18].forEach(i => { out[i] = _normNum(out[i]); });
  return out;
}

// 중복 체크 키: 일(정규화된 날짜 문자열) + 매체 + 캠페인 + 광고그룹 + 소재이름
// → 같은 소재가 여러 캠페인/그룹에 운영되는 경우에도 행별로 구분
function _rowKey(row) {
  const d = _normDate(row[1]);
  const dateStr = (d instanceof Date && !isNaN(d.getTime()))
    ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(row[1] || '').trim();
  return [dateStr, String(row[0] || '').trim(), String(row[2] || '').trim(),
          String(row[3] || '').trim(), String(row[4] || '').trim()].join('\x00');
}

// --------------------------------------------------
// 통합 적재
// RAW_ 로 시작하는 모든 시트의 A~U열 데이터를 소재_통합RAW에 병합
// 중복 키: 일(B) + 매체(A) + 캠페인(C) + 광고그룹(D) + 소재이름(E)
// → 같은 소재가 여러 캠페인/그룹에 운영되는 경우에도 행별로 구분
// --------------------------------------------------
function consolidateRawData() {
  try {
    const ss = getSpreadsheet();

    // 소재_통합RAW 시트 확보
    let targetSheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(CONSOLIDATED_RAW_SHEET_NAME);
      targetSheet.getRange(1, 1, 1, CONSOLIDATED_RAW_HEADERS.length)
        .setValues([CONSOLIDATED_RAW_HEADERS]).setFontWeight('bold');
      targetSheet.setFrozenRows(1);
    }

    // 기존 데이터로 중복 체크 Set 구성 (날짜 등 표현 차이를 흡수하도록 _rowKey로 정규화)
    const existingSet = new Set();
    if (targetSheet.getLastRow() >= 2) {
      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 5).getValues()
        .forEach(r => existingSet.add(_rowKey(r)));
    }

    // RAW_ 로 시작하는 모든 소스 시트 수집
    const rawSheets = ss.getSheets().filter(s => s.getName().startsWith('RAW_'));

    const rowsToAppend = [];

    rawSheets.forEach(srcSheet => {
      if (srcSheet.getLastRow() < 2) return;
      const data = srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 21).getValues();
      data.forEach(row => {
        // 매체(A) 또는 소재이름(E) 중 하나라도 없으면 불완전한 행으로 스킵
        if (!row[0] || !row[4]) return;
        const normalized = _normalizeRawRow(row);
        const key = _rowKey(normalized);
        if (existingSet.has(key)) return;
        existingSet.add(key);
        rowsToAppend.push(normalized);
      });
    });

    if (rowsToAppend.length > 0) {
      targetSheet
        .getRange(targetSheet.getLastRow() + 1, 1, rowsToAppend.length, 21)
        .setValues(rowsToAppend);
    }

    SpreadsheetApp.getUi().alert(`${rowsToAppend.length}건 적재 완료`);
    return { success: true, count: rowsToAppend.length };
  } catch (e) {
    SpreadsheetApp.getUi().alert('오류: ' + e.message);
    return { error: true, message: e.message };
  }
}

// ====================================================
// AB 테스트
// 소재 2~4개를 골라 조건(가설/변경요소/타겟)을 기록하고,
// 지정한 기간(시작일~종료일) 동안의 실적만 집계해 비교한다.
// ====================================================
function _getOrCreateABTestSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(AB_TEST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AB_TEST_SHEET_NAME);
    sheet.getRange(1, 1, 1, AB_TEST_HEADERS.length)
      .setValues([AB_TEST_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < AB_TEST_HEADERS.length) {
    // 매체/캠페인/그룹 컬럼이 없던 이전 버전 시트 — 뒤에 이어붙인다
    const from = sheet.getLastColumn();
    sheet.getRange(1, from + 1, 1, AB_TEST_HEADERS.length - from)
      .setValues([AB_TEST_HEADERS.slice(from)]).setFontWeight('bold');
  }
  return sheet;
}

// 소재_통합RAW(21열)에서 AB테스트 실적 계산에 실제로 쓰는 컬럼만 읽는다 —
// 컨텍스트(A~E: 매체,일,캠페인,광고그룹,소재이름)와 실적(O~S: 노출수,클릭수,CTR,비용,전환) 뿐이고
// 그 사이 F~N(보종~모델유형) 9개 열은 안 쓰는데도 통째로 읽어오고 있었다. 수천 행까지 쌓인
// 시트에서 이 9개 열을 매번 같이 읽어오던 게 AB테스트 실적 조회를 느리게 만드는 원인이었다
// (2026-08-20) — 두 좁은 범위로 나눠 읽고, 기존 코드가 쓰던 인덱스(0~4, 14/15/17/18)에 맞춰
// 다시 조립해서 반환한다(호출부 로직은 그대로 두기 위함).
function _readAbRawRows_() {
  const ss = getSpreadsheet();
  const rawSheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  if (!rawSheet || rawSheet.getLastRow() < 2) return [];
  const n = rawSheet.getLastRow() - 1;
  const ctx  = rawSheet.getRange(2, 1, n, 5).getValues();   // A~E
  const perf = rawSheet.getRange(2, 15, n, 5).getValues();  // O~S: 노출수,클릭수,CTR,비용,전환
  return ctx.map((row, i) => {
    const r = row.slice();
    r[14] = perf[i][0];
    r[15] = perf[i][1];
    r[17] = perf[i][3];
    r[18] = perf[i][4];
    return r;
  });
}

function _abDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(val).trim();
  const m = s.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return s.slice(0, 10);
}

// 저장된 AB테스트 목록 (최신 등록순)
function getABTests() {
  const sheet = _getOrCreateABTestSheet();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, AB_TEST_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.map(r => ({
    테스트ID:     String(r[0] || ''),
    테스트명:     String(r[1] || ''),
    가설:         String(r[2] || ''),
    변경요소:     String(r[3] || ''),
    타겟:         String(r[4] || ''),
    시작일:       _abDateStr(r[5]),
    종료일:       _abDateStr(r[6]),
    소재코드목록: _splitList(r[7]),
    결론메모:     String(r[8] || ''),
    등록일시:     r[9]  ? Utilities.formatDate(new Date(r[9]),  tz, 'yyyy-MM-dd HH:mm') : '',
    최근수정일시: r[10] ? Utilities.formatDate(new Date(r[10]), tz, 'yyyy-MM-dd HH:mm') : '',
    // 이 테스트가 비교하는 매체/캠페인/그룹 범위. 예전에 저장된 테스트는 이 3개가
    // 비어 있는데, getABTestPerformance가 그 경우 범위 제한 없이(과거 방식대로)
    // 소재코드 하나에 얽힌 모든 조합의 실적을 합산하는 쪽으로 자동 대체된다.
    매체:         String(r[11] || ''),
    캠페인:       String(r[12] || ''),
    그룹:         String(r[13] || '')
  })).reverse();
}

// AB테스트 저장 (신규 생성 또는 기존 수정 — data.테스트ID 유무로 판단)
function saveABTest(data) {
  try {
    const codes = (data.소재코드목록 || []).map(c => String(c || '').trim()).filter(Boolean);
    if (codes.length < 2)  return { error: true, message: '소재는 최소 2개 이상 선택해야 합니다.' };
    if (codes.length > 4)  return { error: true, message: '소재는 최대 4개까지 비교할 수 있습니다.' };
    if (!data.테스트명)     return { error: true, message: '테스트명을 입력하세요.' };
    // 매체/캠페인/그룹이 비어 있으면 같은 소재코드가 다른 캠페인·그룹에서도
    // 집행된 실적까지 전부 합쳐져서 비교 자체가 무의미해지므로, 신규 테스트는
    // 이 셋을 필수로 받는다. 기존 테스트 수정(결론메모만 다시 저장하는 경우
    // 포함)은 여기서 막지 않는다 — 예전에 저장된, 아직 범위가 없는 테스트도
    // 계속 열람·수정할 수 있어야 한다. 편집 폼 자체는 클라이언트에서 별도로
    // 범위 선택을 요구한다.
    if (!data.테스트ID && (!data.매체 || !data.캠페인 || !data.그룹)) {
      return { error: true, message: '비교 범위(매체·캠페인·그룹)를 모두 선택하세요.' };
    }

    const sheet = _getOrCreateABTestSheet();
    const now = new Date();
    const rowValues = [
      data.테스트명, data.가설 || '', data.변경요소 || '', data.타겟 || '',
      data.시작일 || '', data.종료일 || '', codes.join(','), data.결론메모 || ''
    ];
    const scopeValues = [data.매체, data.캠페인, data.그룹];

    if (data.테스트ID) {
      const ids = sheet.getLastRow() >= 2
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String)
        : [];
      const idx = ids.indexOf(String(data.테스트ID));
      if (idx === -1) return { error: true, message: '테스트를 찾을 수 없습니다.' };
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 2, 1, rowValues.length).setValues([rowValues]);
      sheet.getRange(rowNum, 11).setValue(now);
      sheet.getRange(rowNum, 12, 1, 3).setValues([scopeValues]);
      return { success: true, 테스트ID: data.테스트ID };
    }

    const testId = 'AB' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmmss');
    sheet.appendRow([testId, ...rowValues, now, now, ...scopeValues]);
    return { success: true, 테스트ID: testId };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

function deleteABTest(testId) {
  try {
    const sheet = _getOrCreateABTestSheet();
    if (sheet.getLastRow() < 2) return { success: true };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(testId));
    if (idx === -1) return { error: true, message: '테스트를 찾을 수 없습니다.' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// ====================================================
// 타겟팅 AB 테스트 — 소재 단위가 아니라 (매체,캠페인,그룹) 단위로 비교한다.
// 저장/조회/삭제는 소재 AB테스트와 같은 시트 CRUD 패턴, 슬롯(2~4개)만
// JSON 배열 문자열 하나에 담아 저장한다.
// ====================================================
function _getOrCreateTargetingABSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(TARGETING_AB_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TARGETING_AB_SHEET_NAME);
    sheet.getRange(1, 1, 1, TARGETING_AB_HEADERS.length)
      .setValues([TARGETING_AB_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 저장된 타겟팅 AB테스트 목록 (최신 등록순)
function getTargetingABTests() {
  const sheet = _getOrCreateTargetingABSheet();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TARGETING_AB_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.map(r => {
    let slots = [];
    try { slots = JSON.parse(r[5] || '[]'); } catch (e) { slots = []; }
    return {
      테스트ID:     String(r[0] || ''),
      테스트명:     String(r[1] || ''),
      가설:         String(r[2] || ''),
      시작일:       _abDateStr(r[3]),
      종료일:       _abDateStr(r[4]),
      슬롯:         slots,
      결론메모:     String(r[6] || ''),
      등록일시:     r[7] ? Utilities.formatDate(new Date(r[7]), tz, 'yyyy-MM-dd HH:mm') : '',
      최근수정일시: r[8] ? Utilities.formatDate(new Date(r[8]), tz, 'yyyy-MM-dd HH:mm') : ''
    };
  }).reverse();
}

// 타겟팅 AB테스트 저장 (신규 생성 또는 기존 수정 — data.테스트ID 유무로 판단)
function saveTargetingABTest(data) {
  try {
    const slots = (data.슬롯 || []).map(s => ({
      매체:   String(s.매체   || '').trim(),
      캠페인: String(s.캠페인 || '').trim(),
      그룹:   String(s.그룹   || '').trim(),
      메모:   String(s.메모   || '').trim()
    })).filter(s => s.매체 && s.캠페인 && s.그룹);

    if (slots.length < 2) return { error: true, message: '타겟팅은 최소 2개 이상 선택해야 합니다.' };
    if (slots.length > 4) return { error: true, message: '타겟팅은 최대 4개까지 비교할 수 있습니다.' };
    if (!data.테스트명)    return { error: true, message: '테스트명을 입력하세요.' };

    const sheet = _getOrCreateTargetingABSheet();
    const now = new Date();
    const rowValues = [
      data.테스트명, data.가설 || '', data.시작일 || '', data.종료일 || '',
      JSON.stringify(slots), data.결론메모 || ''
    ];

    if (data.테스트ID) {
      const ids = sheet.getLastRow() >= 2
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String)
        : [];
      const idx = ids.indexOf(String(data.테스트ID));
      if (idx === -1) return { error: true, message: '테스트를 찾을 수 없습니다.' };
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 2, 1, rowValues.length).setValues([rowValues]);
      sheet.getRange(rowNum, 9).setValue(now);
      return { success: true, 테스트ID: data.테스트ID };
    }

    const testId = 'TAB' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmmss');
    sheet.appendRow([testId, ...rowValues, now, now]);
    return { success: true, 테스트ID: testId };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

function deleteTargetingABTest(testId) {
  try {
    const sheet = _getOrCreateTargetingABSheet();
    if (sheet.getLastRow() < 2) return { success: true };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(testId));
    if (idx === -1) return { error: true, message: '테스트를 찾을 수 없습니다.' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// 슬롯([{매체,캠페인,그룹,메모}, ...])별로 테스트 기간 내 실적을 집계해 비교 반환.
// 소재 AB테스트(getABTestPerformance)와 달리 소재이름으로 좁히지 않고 그 매체·
// 캠페인·그룹 전체(그 안의 모든 소재 합산) 실적을 비교한다 — "타겟팅" 자체의
// 효율을 보는 것이라 소재 단위로 볼 필요가 없다.
function getTargetingABTestPerformance(slots, startDate, endDate) {
  const rawRows = _readAbRawRows_();

  return slots.map(slot => {
    const matched = rawRows.filter(r => {
      if (String(r[0] || '') !== slot.매체)   return false;
      if (String(r[2] || '') !== slot.캠페인) return false;
      if (String(r[3] || '') !== slot.그룹)   return false;
      const d = _abDateStr(r[1]);
      if (startDate && d < startDate) return false;
      if (endDate   && d > endDate)   return false;
      return true;
    });

    const sum = matched.reduce((a, r) => ({
      imp:  a.imp  + (Number(r[14]) || 0),
      clk:  a.clk  + (Number(r[15]) || 0),
      cost: a.cost + (Number(r[17]) || 0),
      conv: a.conv + (Number(r[18]) || 0)
    }), { imp: 0, clk: 0, cost: 0, conv: 0 });

    const creativeCount = new Set(matched.map(r => String(r[4] || '')).filter(Boolean)).size;

    return {
      매체: slot.매체, 캠페인: slot.캠페인, 그룹: slot.그룹, 메모: slot.메모 || '',
      소재수: creativeCount,
      imp: sum.imp, clk: sum.clk, cost: sum.cost, conv: sum.conv,
      ctr: sum.imp  > 0 ? sum.clk  / sum.imp  * 100 : 0,
      cvr: sum.clk  > 0 ? sum.conv / sum.clk  * 100 : 0,
      cpa: sum.conv > 0 ? sum.cost / sum.conv        : 0,
      notFound: matched.length === 0
    };
  });
}

// --------------------------------------------------
// 저장된 보고서 — 보고서 탭의 행/열/지표/필터/기간 설정을 이름 붙여 저장하고
// 나중에 목록에서 다시 불러올 수 있게 한다. AB테스트와 동일한 시트 CRUD 패턴.
// --------------------------------------------------
function _getOrCreateSavedReportSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SAVED_REPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SAVED_REPORT_SHEET_NAME);
    sheet.getRange(1, 1, 1, SAVED_REPORT_HEADERS.length)
      .setValues([SAVED_REPORT_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 저장된 보고서 목록 (최신 등록순), 설정JSON은 파싱해서 반환
function getSavedReports() {
  const sheet = _getOrCreateSavedReportSheet();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, SAVED_REPORT_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.map(r => {
    let config = null;
    try { config = JSON.parse(r[2] || '{}'); } catch (e) { config = {}; }
    return {
      보고서ID:     String(r[0] || ''),
      보고서명:     String(r[1] || ''),
      설정:         config,
      등록일시:     r[3] ? Utilities.formatDate(new Date(r[3]), tz, 'yyyy-MM-dd HH:mm') : '',
      최근수정일시: r[4] ? Utilities.formatDate(new Date(r[4]), tz, 'yyyy-MM-dd HH:mm') : ''
    };
  }).reverse();
}

// 보고서 저장 (신규 생성 또는 기존 수정 — reportId 유무로 판단)
function saveReportConfig(name, config, reportId) {
  try {
    if (!name) return { error: true, message: '보고서명을 입력하세요.' };
    const configJson = JSON.stringify(config || {});
    const sheet = _getOrCreateSavedReportSheet();
    const now = new Date();

    if (reportId) {
      const ids = sheet.getLastRow() >= 2
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String)
        : [];
      const idx = ids.indexOf(String(reportId));
      if (idx === -1) return { error: true, message: '보고서를 찾을 수 없습니다.' };
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 2, 1, 2).setValues([[name, configJson]]);
      sheet.getRange(rowNum, 5).setValue(now);
      return { success: true, 보고서ID: reportId };
    }

    const newId = 'RPT' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmmss');
    sheet.appendRow([newId, name, configJson, now, now]);
    return { success: true, 보고서ID: newId };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

function deleteSavedReport(reportId) {
  try {
    const sheet = _getOrCreateSavedReportSheet();
    if (sheet.getLastRow() < 2) return { success: true };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(reportId));
    if (idx === -1) return { error: true, message: '보고서를 찾을 수 없습니다.' };
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// AB테스트 "비교 범위" 선택용 — 소재_통합RAW에 실제로 실적이 있는 매체·캠페인·
// 그룹 조합만 distinct하게 뽑아 돌려준다. 조합 개수만 반환하므로 RAW가 아무리
// 커도(수만 행) 응답 크기는 항상 작다.
function getRawMediaCampaignGroups() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); // 매체,일,캠페인,광고그룹
  const seen = new Set();
  const out = [];
  rows.forEach(r => {
    const 매체   = String(r[0] || '').trim();
    const 캠페인 = String(r[2] || '').trim();
    const 그룹   = String(r[3] || '').trim();
    if (!매체 || !캠페인 || !그룹) return;
    const key = [매체, 캠페인, 그룹].join('\x00');
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ 매체, 캠페인, 그룹 });
  });
  return out;
}

// 진단용: getRawMediaCampaignGroups()가 왜 비어 보이는지 확인.
// Apps Script 편집기 상단 함수 선택 드롭다운에서 이 함수를 고른 뒤 ▶ 실행하고,
// "실행 로그"(왼쪽 시계 아이콘 또는 보기 > 실행 기록)에서 결과를 확인한다.
function debugABScopeData() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  if (!sheet) {
    const result = { error: '시트를 찾을 수 없음: ' + CONSOLIDATED_RAW_SHEET_NAME, allSheets: ss.getSheets().map(s => s.getName()) };
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    const result = { error: '데이터 없음', lastRow };
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // 매체,일,캠페인,광고그룹
  let hasMedia = 0, hasMediaCampaign = 0, hasAll3 = 0;
  const sampleIncomplete = [];
  rows.forEach((r, i) => {
    const 매체   = String(r[0] || '').trim();
    const 캠페인 = String(r[2] || '').trim();
    const 그룹   = String(r[3] || '').trim();
    if (매체) hasMedia++;
    if (매체 && 캠페인) hasMediaCampaign++;
    if (매체 && 캠페인 && 그룹) hasAll3++;
    if (매체 && (!캠페인 || !그룹) && sampleIncomplete.length < 8) {
      sampleIncomplete.push({ row: i + 2, 매체, 캠페인: 캠페인 || '(빈칸)', 그룹: 그룹 || '(빈칸)' });
    }
  });
  const result = {
    총행수: rows.length,
    매체만있음: hasMedia,
    매체_캠페인까지있음: hasMediaCampaign,
    매체_캠페인_그룹_모두있음: hasAll3,
    캠페인또는그룹빠진샘플: sampleIncomplete
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// 선택한 소재코드들의 "테스트 기간(startDate~endDate)" 실적만 집계해 비교 반환.
// scope({매체,캠페인,그룹})가 주어지면 그 조합에서 집행된 실적만 집계한다 —
// 이게 없으면 같은 이미지코드가 여러 캠페인/그룹에서 재사용된 경우 전부 합쳐져서
// "이 캠페인·그룹 안에서 소재 A/B 성과 비교"라는 AB테스트 본래 목적이 깨진다.
// scope가 없는(과거에 저장된) 테스트는 예전 방식대로 이 코드에 연결된 모든
// 조합의 실적을 합산한다(하위호환).
// 소재_통합RAW은 이미지코드가 아니라 (매체,캠페인,그룹,소재이름) 기준이라
// 소재_마스터에서 각 코드의 배치 정보를 먼저 찾은 뒤 그 키로 매칭한다.
function getABTestPerformance(codes, startDate, endDate, scope) {
  const ss = getSpreadsheet();
  const rawRows = _readAbRawRows_();

  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const masterRows = (masterSheet && masterSheet.getLastRow() >= 2)
    ? masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 16).getValues()
    : [];

  // 디멘드젠/피맥스 번들 소재(전매체 인덱스에 개별 이미지코드가 없는 경우)를 위한
  // 폴백 — getImageCodes()가 광고단위코드를 코드처럼 내려주므로, 여기서도 그
  // 광고단위코드로 구글DA 인덱스에서 직접 컨텍스트를 찾는다.
  const dgpmSheet = ss.getSheetByName(DGPM_SHEET_NAME);
  const dgpmRows = (dgpmSheet && dgpmSheet.getLastRow() >= 2)
    ? dgpmSheet.getRange(2, 1, dgpmSheet.getLastRow() - 1, DGPM_HEADERS.length).getValues()
    : [];

  const hasScope = scope && (scope.매체 || scope.캠페인 || scope.그룹);

  return codes.map(code => {
    // 같은 이미지코드가 여러 매체/캠페인/그룹에 재사용될 수 있다. scope가 있으면
    // 그 조합과 정확히 일치하는 등록 건만 남기고, scope가 없으면(과거 테스트)
    // 이 이미지코드에 연결된 모든 조합을 모아 그 중 하나라도 일치하면 실적에
    // 포함시킨다(하위호환 — 예전엔 이 필터가 아예 없어서 매번 이렇게 동작했다).
    let rowsForCode = masterRows.filter(r => String(r[0] || '').trim() === code);

    if (!rowsForCode.length) {
      const dgpmRow = dgpmRows.find(r => String(r[DGPM_COL['광고단위코드']] || '').trim() === code);
      if (dgpmRow) {
        // masterRows와 같은 컬럼 배치(0=코드,2=매체,3=캠페인,4=그룹,5=소재이름,
        // 6=보종,8=소재유형,14=이미지URL)로 맞춘 가짜 행 하나로 아래 로직을 그대로 재사용한다.
        // 이미지URL은 번들URL(대시보드 갤러리 링크 — 이미지 파일이 아니라 썸네일로 못 씀)을
        // 무조건 쓰지 않고, getImageCodes()와 똑같이 번들 안의 이미지코드목록 중 전매체
        // 인덱스에 실제로 있는 첫 번째 것의 이미지URL을 먼저 찾는다 — 이걸 빠뜨려서
        // AB테스트 실적 비교표에서 피맥스/디멘드젠 번들 소재 썸네일이 안 뜨고 있었다
        // (2026-08-21, 사용자가 캡처로 확인해줌).
        const bundleImageCodes = _splitList(dgpmRow[DGPM_COL['이미지코드목록']]);
        let resolvedImageUrl = '';
        for (let i = 0; i < bundleImageCodes.length; i++) {
          const found = masterRows.find(r => String(r[0] || '').trim() === bundleImageCodes[i]);
          if (found && found[14]) { resolvedImageUrl = String(found[14]); break; }
        }
        const pseudo = [];
        pseudo[0]  = code;
        pseudo[2]  = dgpmRow[DGPM_COL['매체']];
        pseudo[3]  = dgpmRow[DGPM_COL['캠페인']];
        pseudo[4]  = dgpmRow[DGPM_COL['그룹']];
        pseudo[5]  = dgpmRow[DGPM_COL['소재이름']];
        pseudo[6]  = dgpmRow[DGPM_COL['보종']];
        pseudo[8]  = dgpmRow[DGPM_COL['소재유형']];
        pseudo[14] = resolvedImageUrl || String(dgpmRow[DGPM_COL['번들URL']] || '');
        rowsForCode = [pseudo];
      }
    }

    if (hasScope) {
      rowsForCode = rowsForCode.filter(r =>
        (!scope.매체   || String(r[2] || '') === scope.매체)   &&
        (!scope.캠페인 || String(r[3] || '') === scope.캠페인) &&
        (!scope.그룹   || String(r[4] || '') === scope.그룹)
      );
    }
    if (!rowsForCode.length) return { code, notFound: true, scopeMismatch: !!hasScope };

    const contexts = rowsForCode.map(r => ({
      매체: String(r[2] || ''), 캠페인: String(r[3] || ''), 그룹: String(r[4] || ''), 소재이름: String(r[5] || '')
    }));
    const latest = rowsForCode[rowsForCode.length - 1]; // 표시용 메타(보종/소재유형 등)는 최신 값 사용

    const matched = rawRows.filter(r => {
      const media = String(r[0] || ''), campaign = String(r[2] || ''), group = String(r[3] || ''), name = String(r[4] || '');
      if (!contexts.some(c => c.매체 === media && c.캠페인 === campaign && c.그룹 === group && c.소재이름 === name)) return false;
      const d = _abDateStr(r[1]);
      if (startDate && d < startDate) return false;
      if (endDate   && d > endDate)   return false;
      return true;
    });

    const sum = matched.reduce((a, r) => ({
      imp:  a.imp  + (Number(r[14]) || 0),
      clk:  a.clk  + (Number(r[15]) || 0),
      cost: a.cost + (Number(r[17]) || 0),
      conv: a.conv + (Number(r[18]) || 0)
    }), { imp: 0, clk: 0, cost: 0, conv: 0 });

    // 매체/소재이름은 "가장 최근 등록된 조합"이 아니라, 실제로 이 테스트 기간에
    // 실적이 잡힌 조합을 기준으로 표시한다. 이미지코드 하나에 여러 조합이 걸려
    // 있을 때 최근 등록 조합과 실제 집행 조합이 달라, 화면에 나오는 이름이
    // 테스트 기간 필터 기준 실적과 안 맞아 보이는 문제가 있었다.
    const displayRow  = matched.length ? matched[matched.length - 1] : null;
    const displayMedia = displayRow ? String(displayRow[0] || '') : String(latest[2] || '');
    const displayName  = displayRow ? String(displayRow[4] || '') : String(latest[5] || '');

    return {
      code,
      매체:     displayMedia,
      소재이름: displayName,
      보종:     String(latest[6]  || ''),
      소재유형: String(latest[8]  || ''),
      모델유형: String(latest[13] || ''),
      imageUrl: String(latest[14] || ''),
      imp: sum.imp, clk: sum.clk, cost: sum.cost, conv: sum.conv,
      ctr: sum.imp  > 0 ? sum.clk  / sum.imp  * 100 : 0,
      cvr: sum.clk  > 0 ? sum.conv / sum.clk  * 100 : 0,
      cpa: sum.conv > 0 ? sum.cost / sum.conv        : 0
    };
  });
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }
