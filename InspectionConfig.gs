// ====================================================
// 소재 검수 - InspectionConfig.gs
// 검수기준/검수결과/검수이력 시트 정의, 기준값 읽기, API 키 관리
// ====================================================

const INSPECTION_CRITERIA_SHEET = '검수기준';
const INSPECTION_RESULT_SHEET   = '검수결과';
const INSPECTION_HISTORY_SHEET  = '검수이력';

const INSPECTION_RESULT_HEADERS = [
  '검수번호', '검수일시', '파일명', '이미지 링크',
  '상품명 결과', '상품명 인식값', '상품명 차이',
  '심의필 결과', '심의필 인식값', '심의필 차이',
  '기타 문구 결과', '기타 문구 상세',
  '로고 결과', '로고 상세',
  '폰트 결과', '폰트 상세',
  '최종 결과', '수정 요청 내용'
];

const INSPECTION_HISTORY_HEADERS = [
  '검수번호', '실행일시', '업로드 파일 수', '정상 수', '불일치 수', '확인 필요 수', '실행 사용자', '상태'
];

// 검수기준 시트 셀 좌표 (A=항목, B=사용여부, C=입력값)
const CRITERIA_ROW = {
  PRODUCT_NAME: 2,
  REVIEW_NUMBER: 3,
  EXTRA_1: 4,
  EXTRA_2: 5,
  LOGO: 6,
  FONT: 7,
  SPACING_STRICT: 10,
  IGNORE_LINEBREAK: 11,
  CASE_SENSITIVE: 12,
  SPECIAL_CHAR_STRICT: 13,
  OCR_CONFIDENCE: 14
};

// --------------------------------------------------
// 시트 초기화 (없을 때만 생성 — 기존 데이터 보존)
// --------------------------------------------------
function initInspectionSheets() {
  const ss = getSpreadsheet();

  if (!ss.getSheetByName(INSPECTION_CRITERIA_SHEET)) {
    const sheet = ss.insertSheet(INSPECTION_CRITERIA_SHEET);

    sheet.getRange(1, 1, 1, 3).setValues([['항목', '사용 여부', '입력값']]).setFontWeight('bold');
    const rows = [
      ['상품명·보종명',        false, ''],
      ['심의필',               false, ''],
      ['기타 필수 문구 1',      false, ''],
      ['기타 필수 문구 2',      false, ''],
      ['로고 (기준 이미지 Drive 링크 또는 파일ID)', false, ''],
      ['폰트 (기준 폰트명)',    false, '']
    ];
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 2, rows.length, 1).insertCheckboxes();

    sheet.getRange(9, 1, 1, 3).setValues([['추가 설정', '값', '설명']]).setFontWeight('bold');
    const opts = [
      ['띄어쓰기 정확히 비교',   false, '켜면 공백 차이도 불일치로 판정 (기본: 공백 무시)'],
      ['줄바꿈 무시',            true,  '켜면 줄바꿈을 공백으로 바꿔서 비교'],
      ['영문 대소문자 구분',      false, '켜면 대소문자가 다르면 불일치 (기본: 무시)'],
      ['특수문자 정확히 비교',    false, '심의필의 숫자·하이픈은 이 설정과 무관하게 항상 정확히 비교됩니다'],
      ['OCR 신뢰도 기준 (0~1)',  0.7,   '인식 신뢰도가 이 값보다 낮으면 확인 필요로 처리']
    ];
    sheet.getRange(10, 1, opts.length, 3).setValues(opts);
    sheet.getRange(10, 2, 4, 1).insertCheckboxes();

    sheet.getRange(15, 1, 1, 3)
      .setValues([['OpenAI API 키 입력 위치', '', '대시보드(소재 검수 탭 상단 링크) > 설정 관리 > API 키에서 OpenAI API Key를 등록하세요. 시트에 직접 입력하지 마세요. (같은 키를 소재 분석 탭의 AI 인사이트 기능과 함께 사용합니다)']]);

    sheet.setColumnWidths(1, 1, 260);
    sheet.setColumnWidths(2, 1, 90);
    sheet.setColumnWidths(3, 1, 420);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(INSPECTION_RESULT_SHEET)) {
    const sheet = ss.insertSheet(INSPECTION_RESULT_SHEET);
    sheet.getRange(1, 1, 1, INSPECTION_RESULT_HEADERS.length)
      .setValues([INSPECTION_RESULT_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(INSPECTION_HISTORY_SHEET)) {
    const sheet = ss.insertSheet(INSPECTION_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, INSPECTION_HISTORY_HEADERS.length)
      .setValues([INSPECTION_HISTORY_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return { success: true, message: '검수 시트 초기화 완료' };
}

// --------------------------------------------------
// 검수 기준 읽기 — 클라이언트 표시 + 검수 로직 양쪽에서 사용
// --------------------------------------------------
function getInspectionCriteria() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(INSPECTION_CRITERIA_SHEET);
  if (!sheet) return { error: true, message: '검수기준 시트가 없습니다. 메뉴에서 [검수 시트 초기화]를 먼저 실행하세요.' };

  const get = (row) => sheet.getRange(row, 1, 1, 3).getValues()[0];
  const field = (row) => { const r = get(row); return { use: r[1] === true, value: String(r[2] || '').trim() }; };

  const ocrRow = get(CRITERIA_ROW.OCR_CONFIDENCE);
  const ocrConfidence = Number(ocrRow[1]);

  return {
    productName:  field(CRITERIA_ROW.PRODUCT_NAME),
    reviewNumber: field(CRITERIA_ROW.REVIEW_NUMBER),
    extra1:       field(CRITERIA_ROW.EXTRA_1),
    extra2:       field(CRITERIA_ROW.EXTRA_2),
    logo:         field(CRITERIA_ROW.LOGO),
    font:         field(CRITERIA_ROW.FONT),
    options: {
      spacingStrict:     get(CRITERIA_ROW.SPACING_STRICT)[1] === true,
      ignoreLineBreak:   get(CRITERIA_ROW.IGNORE_LINEBREAK)[1] === true,
      caseSensitive:     get(CRITERIA_ROW.CASE_SENSITIVE)[1] === true,
      specialCharStrict: get(CRITERIA_ROW.SPECIAL_CHAR_STRICT)[1] === true,
      ocrConfidence: isNaN(ocrConfidence) ? 0.7 : ocrConfidence
    },
    apiKeyConfigured: !!getOpenAIApiKey()
  };
}

// --------------------------------------------------
// OpenAI API 키 — Script Properties에만 저장 (시트에 저장 금지)
// --------------------------------------------------
function getOpenAIApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
}

// 대시보드(설정 관리 > API 키)에서도 같은 OPENAI_API_KEY를 등록/변경할 수 있다.
// 이 함수는 대시보드를 열지 않고 스프레드시트에서 바로 등록하고 싶을 때 쓰는 대체 경로다.
function promptSetOpenAIKey() {
  const ui = SpreadsheetApp.getUi();
  const existing = getOpenAIApiKey();
  const resp = ui.prompt(
    'OpenAI API 키 설정',
    (existing ? '현재 키가 등록되어 있습니다. 새 값을 입력하면 교체됩니다.\n' : '') +
    '이 키는 소재 검수와 소재 분석 탭의 AI 인사이트 기능이 함께 사용합니다.\nAPI 키를 입력하세요 (sk-...):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const key = resp.getResponseText().trim();
  if (!key) { ui.alert('빈 값은 저장하지 않았습니다.'); return; }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', key);
  ui.alert('API 키가 저장되었습니다.');
}

function openInspectionCriteriaSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(INSPECTION_CRITERIA_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('먼저 [검수 시트 초기화]를 실행하세요.'); return; }
  ss.setActiveSheet(sheet);
}
