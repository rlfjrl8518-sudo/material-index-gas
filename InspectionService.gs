// ====================================================
// 소재 검수 - InspectionService.gs
// 검수 오케스트레이션 + 대량 처리(배치/트리거) + 결과/이력 조회
//
// 대량 처리 설계 메모:
// 진행 상태(state)는 { folderId, total, ok, mismatch, needCheck }만 Properties에 저장한다.
// 처리 대상 큐 자체는 저장하지 않고, 매 배치마다 원본 Drive 폴더를 스캔해서
// "검수결과 시트에 아직 없는 이미지"를 다음 배치 대상으로 삼는다.
// (PropertiesService 값 하나당 저장 용량 제한이 있어 파일 목록 전체를 담지 않기 위함)
// ====================================================

const INSPECTION_BATCH_SIZE = 20;
const ALLOWED_IMAGE_MIME = ['image/jpeg', 'image/png', 'image/webp'];

// --------------------------------------------------
// 업로드 준비 — 새 검수번호 발급 + Drive 폴더 생성 + 브라우저 직접 업로드용 토큰
// (기존 getUploadToken()과 동일한 "브라우저 → Drive 직접 업로드" 패턴 재사용)
// --------------------------------------------------
function getInspectionUploadContext() {
  const id = _generateInspectionId();
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const root = _getOrCreateFolder(DriveApp.getRootFolder(), '이미지소재검수');
  const runFolder = _getOrCreateFolder(root, dateStr + '_' + id);
  const originalFolder = _getOrCreateFolder(runFolder, '원본');
  _getOrCreateFolder(runFolder, '검수결과'); // 예약 폴더(추후 주석/결과이미지용) — 자동 삭제하지 않음

  PropertiesService.getScriptProperties().setProperty('INSPECTION_FOLDER_' + id, originalFolder.getId());

  return { inspectionId: id, token: ScriptApp.getOAuthToken(), folderId: originalFolder.getId() };
}

function _getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function _generateInspectionId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty('INSPECTION_SEQ') || '0') + 1;
    props.setProperty('INSPECTION_SEQ', String(n));
    return 'INSP-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd') + '-' + String(n).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

// --------------------------------------------------
// 검수 시작 — 업로드가 끝난 파일들(fileRefs)을 받아 처리를 개시한다
// fileRefs: [{ fileId, fileName, isZip }]
// --------------------------------------------------
function startInspection(inspectionId, fileRefs) {
  const criteria = getInspectionCriteria();
  if (criteria.error) return criteria;
  if (!criteria.apiKeyConfigured) {
    return { error: true, message: 'OpenAI API 키가 설정되지 않았습니다. 스프레드시트 메뉴 [이미지 검수 > API 키 설정]에서 등록하세요.' };
  }

  const folderId = PropertiesService.getScriptProperties().getProperty('INSPECTION_FOLDER_' + inspectionId);
  if (!folderId) return { error: true, message: '검수 폴더를 찾을 수 없습니다. 새 검수를 다시 시작하세요.' };
  const originalFolder = DriveApp.getFolderById(folderId);

  let total = 0;
  let skipped = [];
  (fileRefs || []).forEach(function (ref) {
    if (ref.isZip) {
      const res = extractImagesFromZip(ref.fileId, originalFolder);
      if (res.error) { skipped.push({ name: ref.fileName, reason: res.message }); return; }
      total += res.images.length;
      skipped = skipped.concat(res.skipped || []);
    } else {
      total += 1;
    }
  });

  if (total === 0) {
    return { error: true, message: '처리할 이미지가 없습니다.' + (skipped.length ? ' (' + skipped.length + '개 파일 제외됨)' : '') };
  }

  const state = { folderId: folderId, total: total, ok: 0, mismatch: 0, needCheck: 0 };
  PropertiesService.getScriptProperties().setProperty('INSPECTION_STATE_' + inspectionId, JSON.stringify(state));

  _appendHistoryRow(inspectionId, total, 0, 0, 0, '진행중');

  _processInspectionBatch(inspectionId);

  const progress = getInspectionProgress(inspectionId);
  progress.skipped = skipped;
  return progress;
}

// --------------------------------------------------
// 배치 1회 처리 (최대 20장) — 동시 실행 방지를 위해 LockService 사용
// --------------------------------------------------
function _processInspectionBatch(inspectionId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { busy: true };

  try {
    const props = PropertiesService.getScriptProperties();
    const stateKey = 'INSPECTION_STATE_' + inspectionId;
    const raw = props.getProperty(stateKey);
    if (!raw) return { done: true, missing: true };
    const state = JSON.parse(raw);

    const folder = DriveApp.getFolderById(state.folderId);
    const existingRows = getInspectionResults(inspectionId);
    const processedNames = {};
    existingRows.forEach(function (r) { processedNames[r['파일명']] = true; });

    const pending = _listPendingImages(folder, processedNames, INSPECTION_BATCH_SIZE);
    const resultSheet = getSpreadsheet().getSheetByName(INSPECTION_RESULT_SHEET);

    pending.forEach(function (item) {
      let outcome;
      try {
        const publicUrl = setFilePublic(item.fileId);
        const processed = _processOneImage(inspectionId, item.fileId, item.fileName, state.criteria || getInspectionCriteria(), publicUrl);
        resultSheet.appendRow(processed.row);
        outcome = processed.finalResult;
      } catch (e) {
        resultSheet.appendRow([
          inspectionId, new Date(), item.fileName, '',
          '-', '', '', '-', '', '', '-', '', '-', '', '-', '',
          '확인 필요', '처리 중 오류로 확인 필요 처리됨: ' + e.message
        ]);
        outcome = '확인 필요';
      }
      if (outcome === '정상') state.ok++;
      else if (outcome === '불일치') state.mismatch++;
      else state.needCheck++;
      Utilities.sleep(300); // API 호출 간격
    });

    const processedCount = existingRows.length + pending.length;
    const done = pending.length === 0 || processedCount >= state.total;

    _updateHistoryRow(inspectionId, state.ok, state.mismatch, state.needCheck, done ? '완료' : '진행중');

    if (done) {
      props.deleteProperty(stateKey);
      props.deleteProperty('INSPECTION_ACTIVE_ID');
      _deleteTriggersFor('continueInspectionBatch');
    } else {
      props.setProperty(stateKey, JSON.stringify(state));
      props.setProperty('INSPECTION_ACTIVE_ID', inspectionId);
      _deleteTriggersFor('continueInspectionBatch');
      _createContinueTrigger();
    }

    return { total: state.total, processed: processedCount, ok: state.ok, mismatch: state.mismatch, needCheck: state.needCheck, done: done };
  } finally {
    lock.releaseLock();
  }
}

// 시간 기반 트리거 핸들러 — 다음 배치를 이어서 처리한다
function continueInspectionBatch() {
  const id = PropertiesService.getScriptProperties().getProperty('INSPECTION_ACTIVE_ID');
  if (!id) { _deleteTriggersFor('continueInspectionBatch'); return; }
  _processInspectionBatch(id);
}

// ScriptApp.getProjectTriggers/newTrigger는 script.scriptapp 권한이 필요한데,
// appsscript.json에 스코프를 추가해도 배포 계정이 다시 승인(재인증)하기 전까지는
// "호출할 수 있는 권한이 없습니다" 예외가 난다. 이 권한 하나 때문에 이미 끝난
// 이미지 분석 결과(검수결과 시트에 이미 기록됨)까지 통째로 "검수 시작 실패"로
// 덮이는 게 문제라, 트리거 정리/등록 실패는 삼키고 로그만 남긴다 — 재인증 전까지는
// 20장 단위 배치 자동 이어처리만 안 되고(수동으로 다시 시작하면 됨), 검수 자체는 진행된다.
function _deleteTriggersFor(handlerName) {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
    });
  } catch (e) {
    Logger.log('트리거 정리 실패(권한 재인증 필요 가능성): ' + e.message);
  }
}

function _createContinueTrigger() {
  try {
    ScriptApp.newTrigger('continueInspectionBatch').timeBased().after(5000).create();
  } catch (e) {
    Logger.log('다음 배치 트리거 등록 실패(권한 재인증 필요 가능성): ' + e.message);
  }
}

function _listPendingImages(folder, processedNamesMap, limit) {
  const it = folder.getFiles();
  const pending = [];
  while (it.hasNext()) {
    const f = it.next();
    if (ALLOWED_IMAGE_MIME.indexOf(f.getMimeType()) === -1) continue;
    const name = f.getName();
    if (processedNamesMap[name]) continue;
    pending.push({ fileId: f.getId(), fileName: name });
    if (pending.length >= limit) break;
  }
  return pending;
}

// --------------------------------------------------
// 이미지 1장 분석+비교 → 검수결과 시트 1행 데이터
// --------------------------------------------------
function _processOneImage(inspectionId, fileId, fileName, criteria, imagePublicUrl) {
  const analysis = analyzeImageWithOpenAI(fileId, criteria);
  const opts = criteria.options;
  const failed = !!analysis.error;

  const productResult = judgeField(criteria.productName.value, criteria.productName.use, failed ? null : analysis.productText, opts, false);
  const reviewResult = judgeField(criteria.reviewNumber.value, criteria.reviewNumber.use, failed ? null : analysis.reviewText, opts, true);
  const extra1Result = _judgeExtraPhrase(criteria.extra1, failed ? null : analysis, opts);
  const extra2Result = _judgeExtraPhrase(criteria.extra2, failed ? null : analysis, opts);
  const extraCombined = _combineExtraResults(extra1Result, extra2Result);
  const logoResult = judgeLogo(criteria.logo, failed ? null : analysis.logo, opts.ocrConfidence);
  const fontResult = judgeFont(criteria.font, failed ? null : analysis.font);

  const finalResult = computeFinalResult([productResult, reviewResult, extraCombined.overall, logoResult]);

  const mismatchItems = [];
  if (productResult && productResult.result === '불일치') {
    mismatchItems.push({ label: '상품명', base: criteria.productName.value, recognizedText: productResult.recognizedText });
  }
  if (reviewResult && reviewResult.result === '불일치') {
    mismatchItems.push({ label: '심의필', base: criteria.reviewNumber.value, recognizedText: reviewResult.recognizedText });
  }
  if (extra1Result && extra1Result.result === '불일치') {
    mismatchItems.push({ label: '기타 필수 문구 1', base: criteria.extra1.value, recognizedText: extra1Result.recognizedText });
  }
  if (extra2Result && extra2Result.result === '불일치') {
    mismatchItems.push({ label: '기타 필수 문구 2', base: criteria.extra2.value, recognizedText: extra2Result.recognizedText });
  }
  if (logoResult && logoResult.result === '불일치') {
    mismatchItems.push({ label: '로고', base: '(기준 로고)', recognizedText: logoResult.detail || '로고 미확인' });
  }

  const revisionMessage = buildRevisionMessage(fileName, mismatchItems);

  const row = [
    inspectionId, new Date(), fileName, imagePublicUrl,
    productResult ? productResult.result : '-', productResult ? productResult.recognizedText : '', productResult ? productResult.diff : '',
    reviewResult ? reviewResult.result : '-', reviewResult ? reviewResult.recognizedText : '', reviewResult ? reviewResult.diff : '',
    extraCombined.resultLabel, extraCombined.detail,
    logoResult ? logoResult.result : '-', logoResult ? logoResult.detail : '',
    fontResult ? fontResult.result : '-', fontResult ? fontResult.detail : '',
    finalResult, revisionMessage
  ];

  return { row: row, finalResult: finalResult };
}

// 기타 필수 문구는 개별 confidence가 없어 allTexts 전체 목록에 포함되는지로 판정한다
function _judgeExtraPhrase(extraCriteria, analysis, opts) {
  if (!extraCriteria || !extraCriteria.use) return null;
  if (!analysis || analysis.error) return { result: '확인 필요', recognizedText: '', reason: 'API 응답 실패' };

  const allTexts = Array.isArray(analysis.allTexts) ? analysis.allTexts : [];
  if (allTexts.length === 0) return { result: '확인 필요', recognizedText: '', reason: '문구 인식 실패' };

  const normBase = normalizeText(extraCriteria.value, opts, false);
  if (!normBase) return { result: '확인 필요', recognizedText: '', reason: '기준값 없음' };

  const found = allTexts.some(function (t) { return normalizeText(t, opts, false).indexOf(normBase) !== -1; });
  return found
    ? { result: '일치', recognizedText: extraCriteria.value }
    : { result: '불일치', recognizedText: allTexts.join(' / ').slice(0, 200) };
}

function _combineExtraResults(r1, r2) {
  const list = [r1, r2].filter(Boolean);
  if (list.length === 0) return { overall: null, resultLabel: '-', detail: '' };
  let resultLabel;
  if (list.some(function (r) { return r.result === '불일치'; })) resultLabel = '불일치';
  else if (list.some(function (r) { return r.result === '확인 필요'; })) resultLabel = '확인 필요';
  else resultLabel = '일치';
  const detail = list.map(function (r, i) {
    return '문구' + (i + 1) + ': ' + r.result + (r.recognizedText ? ' (' + r.recognizedText + ')' : '');
  }).join(' / ');
  return { overall: { result: resultLabel }, resultLabel: resultLabel, detail: detail };
}

// --------------------------------------------------
// 진행률 / 결과 / 이력 조회
// --------------------------------------------------
function getInspectionProgress(inspectionId) {
  const raw = PropertiesService.getScriptProperties().getProperty('INSPECTION_STATE_' + inspectionId);
  if (raw) {
    const state = JSON.parse(raw);
    const processed = getInspectionResults(inspectionId).length;
    return { total: state.total, processed: processed, ok: state.ok, mismatch: state.mismatch, needCheck: state.needCheck, done: false };
  }
  const hist = _findHistoryRow(inspectionId);
  if (!hist) return { error: true, message: '검수 이력을 찾을 수 없습니다.' };
  return { total: hist.uploaded, processed: hist.uploaded, ok: hist.ok, mismatch: hist.mismatch, needCheck: hist.needCheck, done: true };
}

function getInspectionResults(inspectionId) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_RESULT_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, INSPECTION_RESULT_HEADERS.length).getValues();
  const rows = [];
  values.forEach(function (r, idx) {
    if (String(r[0]) === inspectionId) {
      const obj = {};
      INSPECTION_RESULT_HEADERS.forEach(function (h, i) { let v = r[i]; if (v instanceof Date) v = v.toISOString(); obj[h] = v; });
      obj._row = idx + 2;
      rows.push(obj);
    }
  });
  return rows;
}

function _appendHistoryRow(inspectionId, uploaded, ok, mismatch, needCheck, status) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_HISTORY_SHEET);
  let user = '(알 수 없음)';
  try { user = Session.getActiveUser().getEmail() || user; } catch (e) {}
  sheet.appendRow([inspectionId, new Date(), uploaded, ok, mismatch, needCheck, user, status]);
}

function _findHistoryRowIndex(inspectionId) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_HISTORY_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === inspectionId) return i + 2;
  }
  return -1;
}

function _updateHistoryRow(inspectionId, ok, mismatch, needCheck, status) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_HISTORY_SHEET);
  const rowIdx = _findHistoryRowIndex(inspectionId);
  if (rowIdx === -1) return;
  sheet.getRange(rowIdx, 4, 1, 3).setValues([[ok, mismatch, needCheck]]);
  sheet.getRange(rowIdx, 8).setValue(status);
}

function _findHistoryRow(inspectionId) {
  const rowIdx = _findHistoryRowIndex(inspectionId);
  if (rowIdx === -1) return null;
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_HISTORY_SHEET);
  const vals = sheet.getRange(rowIdx, 1, 1, INSPECTION_HISTORY_HEADERS.length).getValues()[0];
  return { id: vals[0], uploaded: vals[2], ok: vals[3], mismatch: vals[4], needCheck: vals[5], status: vals[7] };
}

function getRecentInspections(limit) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_HISTORY_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const n = Math.min(limit || 20, lastRow - 1);
  const values = sheet.getRange(lastRow - n + 1, 1, n, INSPECTION_HISTORY_HEADERS.length).getValues();
  return values.reverse().map(function (r) {
    return { id: r[0], date: (r[1] instanceof Date) ? r[1].toISOString() : r[1], uploaded: r[2], ok: r[3], mismatch: r[4], needCheck: r[5], user: r[6], status: r[7] };
  });
}

// --------------------------------------------------
// 실패(확인 필요)한 이미지만 재처리 — 기존 행을 그대로 갱신한다
// --------------------------------------------------
function retryFailedImages(inspectionId) {
  const criteria = getInspectionCriteria();
  if (criteria.error) return criteria;
  if (!criteria.apiKeyConfigured) return { error: true, message: 'API 키가 설정되지 않았습니다.' };

  const sheet = getSpreadsheet().getSheetByName(INSPECTION_RESULT_SHEET);
  const results = getInspectionResults(inspectionId);
  const targets = results.filter(function (r) { return r['최종 결과'] === '확인 필요'; });
  if (targets.length === 0) return { retried: 0, message: '재처리할 항목이 없습니다.' };

  let okDelta = 0, mismatchDelta = 0, needCheckDelta = 0;

  targets.forEach(function (r) {
    const fileId = _extractDriveFileId(r['이미지 링크']);
    if (!fileId) { return; }
    try {
      const publicUrl = setFilePublic(fileId);
      const processed = _processOneImage(inspectionId, fileId, r['파일명'], criteria, publicUrl);
      sheet.getRange(r._row, 1, 1, INSPECTION_RESULT_HEADERS.length).setValues([processed.row]);
      needCheckDelta--;
      if (processed.finalResult === '정상') okDelta++;
      else if (processed.finalResult === '불일치') mismatchDelta++;
      else needCheckDelta++;
    } catch (e) {
      // 실패 시 '확인 필요' 상태 그대로 유지
    }
    Utilities.sleep(300);
  });

  const hist = _findHistoryRow(inspectionId);
  if (hist) {
    _updateHistoryRow(inspectionId, Number(hist.ok) + okDelta, Number(hist.mismatch) + mismatchDelta, Number(hist.needCheck) + needCheckDelta, '완료');
  }

  return { retried: targets.length };
}

// --------------------------------------------------
// 선택 결과 삭제 (메뉴 [선택 결과 삭제] / 결과 화면에서 사용)
// --------------------------------------------------
function deleteInspectionResultRows(rowNumbers) {
  const sheet = getSpreadsheet().getSheetByName(INSPECTION_RESULT_SHEET);
  const sorted = (rowNumbers || []).slice().sort(function (a, b) { return b - a; });
  sorted.forEach(function (rowNum) {
    if (rowNum > 1) sheet.deleteRow(rowNum);
  });
  return { deleted: sorted.length };
}
