// ====================================================
// 소재 등록 자동화 - Registration.gs
// 등록 대기열(RAW에 있는데 인덱스에 없는 소재) · AI 태그 추천 · 일괄 저장 · 심의필 만료 관리
// ====================================================

const MASTER_REVIEW_COL = 17;      // 전매체 인덱스 Q열: 심의필 문구(AI 판독)
const MASTER_REVIEW_END_COL = 18;  // 전매체 인덱스 R열: 심의필 만료일(yyyy-MM-dd)
const TAG_FIELDS = ['보종', '소재유형', '소구포인트', '후킹방식', '이미지유형', '모델유형'];
const QUEUE_LOOKBACK_DAYS = 60;    // 대기열은 최근 이 기간에 집행된 소재만 본다(예전에 끝난 미등록 소재까지 쌓이지 않게)

// --------------------------------------------------
// 등록 대기열 — RAW_* 시트(통합 적재 전에도 최신)에서 전매체 인덱스·구글DA 인덱스에 없는
// (매체, 캠페인, 그룹, 소재이름)을 찾아 광고비 큰 순으로 돌려준다.
// --------------------------------------------------
function getRegistrationQueue() {
  const ss = getSpreadsheet();
  const registered = _registeredCreativeKeys(ss);
  const cutoff = new Date(Date.now() - QUEUE_LOOKBACK_DAYS * 864e5);
  const byKey = {};

  ss.getSheets().filter(s => s.getName().indexOf('RAW_') === 0).forEach(sheet => {
    const last = sheet.getLastRow();
    if (last < 2) return;
    sheet.getRange(2, 1, last - 1, CONSOLIDATED_RAW_HEADERS.length).getValues().forEach(r => {
      const item = { 매체: _s(r[0]), 캠페인: _s(r[2]), 그룹: _s(r[3]), 소재이름: _s(r[4]) };
      if (!item.매체 || !item.소재이름) return;
      const key = _creativeKey(item);
      if (registered[key]) return;
      const d = _normDate(r[1]);
      if (d instanceof Date && d < cutoff) return;
      const q = byKey[key] || (byKey[key] = Object.assign(item, { 비용: 0, 노출수: 0, 첫집행일: '', 최근집행일: '' }));
      q.비용 += Number(r[17]) || 0;
      q.노출수 += Number(r[14]) || 0;
      if (d instanceof Date) {
        const ds = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
        if (!q.첫집행일 || ds < q.첫집행일) q.첫집행일 = ds;
        if (!q.최근집행일 || ds > q.최근집행일) q.최근집행일 = ds;
      }
    });
  });

  const items = Object.keys(byKey).map(k => byKey[k]).sort((a, b) => b.비용 - a.비용);
  return { count: items.length, items: items };
}

function _s(v) { return String(v == null ? '' : v).trim(); }
function _creativeKey(o) { return [o.매체, o.캠페인, o.그룹, o.소재이름].map(_s).join('\x00'); }

function _registeredCreativeKeys(ss) {
  const keys = {};
  const master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (master && master.getLastRow() >= 2) {
    master.getRange(2, 3, master.getLastRow() - 1, 4).getValues()
      .forEach(r => { keys[_creativeKey({ 매체: r[0], 캠페인: r[1], 그룹: r[2], 소재이름: r[3] })] = true; });
  }
  const dgpm = ss.getSheetByName(DGPM_SHEET_NAME);
  if (dgpm && dgpm.getLastRow() >= 2) {
    dgpm.getRange(2, 1, dgpm.getLastRow() - 1, DGPM_HEADERS.length).getValues().forEach(r => {
      keys[_creativeKey({ 매체: r[DGPM_COL['매체']], 캠페인: r[DGPM_COL['캠페인']], 그룹: r[DGPM_COL['그룹']], 소재이름: r[DGPM_COL['소재이름']] })] = true;
    });
  }
  return keys;
}

// --------------------------------------------------
// AI 태그 추천 — items: [{ fileId, 소재이름, 캠페인 }]
// 이미지 + 소재이름(이름에 '_특약_', '간병인' 같은 단서가 많다)을 같이 보고, 설정 시트의 선택지 안에서만
// 고르게 한다(responseSchema enum). 기존 전매체 인덱스의 실제 태깅 예시를 같이 줘서 팀의 분류 습관을 따르게 한다.
// 심의필 문구도 함께 읽어 만료일을 뽑는다.
// 반환: items와 같은 순서의 [{ tags, 심의필, 심의만료일 } | { error, message }]
// --------------------------------------------------
function suggestCreativeTags(items, excludeNames) {
  if (!getAiProjectId()) return (items || []).map(() => ({ error: true, message: 'GCP 프로젝트 ID가 설정되지 않았습니다.' }));
  const settings = getSettings();
  const guide = _buildTagGuide(settings, excludeNames || []);
  const schema = _tagResponseSchema(settings);

  const blobs = _fetchDriveFiles(items.map(it => it.fileId));
  const built = items.map((it, i) => {
    try {
      const blob = blobs[i];
      if (!blob) return { error: '이미지 파일을 읽을 수 없습니다.' };
      if (!/^image\//.test(blob.getContentType() || '')) return { error: '이미지 파일만 AI 추천을 지원합니다(동영상은 직접 입력).' };
      return { parts: [
        { text: guide + '\n\n[이번 소재]\n캠페인: ' + _s(it.캠페인) + '\n소재이름: ' + _s(it.소재이름) },
        { inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } }
      ] };
    } catch (e) {
      return { error: '이미지 파일을 읽을 수 없습니다: ' + e.message };
    }
  });

  const out = built.map(b => (b.error ? { error: true, message: b.error } : null));
  let pending = built.map((b, i) => i).filter(i => !out[i]);
  for (let round = 0; round <= VISION_MAX_RETRY && pending.length; round++) {
    if (round > 0) Utilities.sleep(3000 * round);
    const next = [];
    for (let c = 0; c < pending.length; c += VISION_CONCURRENCY) {
      const chunk = pending.slice(c, c + VISION_CONCURRENCY);
      const responses = UrlFetchApp.fetchAll(chunk.map(i => _tagRequestOptions(built[i].parts, schema)));
      chunk.forEach((i, k) => {
        const parsed = _parseVisionResponse(responses[k]);
        if (parsed.error) { out[i] = parsed; next.push(i); return; }
        out[i] = _normalizeTagSuggestion(parsed, settings);
      });
    }
    pending = next;
  }
  return out;
}

// Drive 파일 여러 개를 동시에 내려받는다(DriveApp.getBlob은 한 장씩 순서대로라 8장이면 3~4초가 더 걸렸다).
function _fetchDriveFiles(fileIds) {
  const token = ScriptApp.getOAuthToken();
  return UrlFetchApp.fetchAll(fileIds.map(id => ({
    url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?alt=media&supportsAllDrives=true',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  }))).map(r => (r.getResponseCode() === 200 ? r.getBlob() : null));
}

function _tagRequestOptions(parts, schema) {
  return {
    url: _visionUrl(),
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: '너는 보험 광고 소재를 팀의 분류 체계에 맞춰 태깅하는 담당자다. 반드시 주어진 선택지 안에서만 고르고, 확신이 없으면 기존 예시와 가장 비슷한 쪽을 고른다. 심의필은 이미지에 보이는 그대로 받아 적고 못 읽는 글자는 ?로 적는다.' }] },
      contents: [{ role: 'user', parts: parts }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 4000,
        responseMimeType: 'application/json',
        responseSchema: schema,
        thinkingConfig: { thinkingLevel: 'LOW' }
      }
    }),
    muteHttpExceptions: true
  };
}

function _tagResponseSchema(settings) {
  const enumOf = f => ({ type: 'STRING', enum: (settings[f] || []).length ? settings[f] : [''] });
  return {
    type: 'OBJECT',
    properties: {
      보종: enumOf('보종'),
      소재유형: enumOf('소재유형'),
      소구포인트: enumOf('소구포인트'),
      후킹방식: enumOf('후킹방식'),
      이미지유형: enumOf('이미지유형'),
      모델유형: { type: 'ARRAY', items: enumOf('모델유형') },
      소구상세: { type: 'STRING' },
      심의필: { type: 'STRING' }
    },
    required: ['보종', '소재유형', '소구포인트', '후킹방식', '이미지유형', '모델유형', '소구상세', '심의필']
  };
}

// 설정 시트 선택지 + 전매체 인덱스의 실제 태깅 예시(값마다 소재이름 몇 개)로 모델용 안내문을 만든다.
function _buildTagGuide(settings, excludeNames) {
  const exclude = {};
  excludeNames.forEach(n => { exclude[_s(n)] = true; });
  const examples = {};   // field -> value -> [소재이름]
  const details = [];
  const master = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (master && master.getLastRow() >= 2) {
    // 5=소재이름 6=보종 8=소재유형 9=소구포인트 10=후킹방식 11=소구상세 12=이미지유형 13=모델유형
    const colOf = { 보종: 6, 소재유형: 8, 소구포인트: 9, 후킹방식: 10, 이미지유형: 12, 모델유형: 13 };
    master.getRange(2, 1, master.getLastRow() - 1, 14).getValues().slice().reverse().forEach(r => {
      const name = _s(r[5]);
      if (!name || exclude[name]) return;
      TAG_FIELDS.forEach(f => {
        const v = _s(r[colOf[f]]);
        if (!v) return;
        const byVal = examples[f] || (examples[f] = {});
        const list = byVal[v] || (byVal[v] = []);
        if (list.length < 4 && list.indexOf(name) === -1) list.push(name);
      });
      const d = _s(r[11]);
      if (d && details.length < 15 && details.indexOf(d) === -1) details.push(d);
    });
  }

  const lines = [
    '광고 이미지와 소재이름을 보고 아래 항목을 태깅하라. 소재이름에 들어 있는 단서(예: _특약_, _비갱신_, _브랜딩_, 상품명)를 우선 활용하라.',
    '',
    '[항목 정의]',
    '- 보종: 광고하는 보험 상품. 이미지 속 상품명과 소재이름으로 판단',
    '- 소재유형: 비갱신(보험료가 오르지 않는 비갱신형 강조) / 특약(특정 특약·보장 강조) / 브랜딩(상품·브랜드 전반 인지)',
    '- 소구포인트: 메시지의 핵심 설득 포인트',
    '- 후킹방식: 헤드라인 문장 형태. 질문형(~할까요?, ~하셨나요?) / 권유형(~하세요, ~해보세요) / 단정형(단정적 서술)',
    '- 이미지유형: 3D(입체 렌더링) / 2D(평면 일러스트·그래픽) / 실사(실제 사진)',
    '- 모델유형: 이미지에 등장하는 인물의 연령대·성별(여러 명이면 모두). 실존 광고모델 사진이면 자사모델, 인물이 없으면 오브젝트',
    '- 소구상세: 특정 특약·보장 이름 등 세부 소구 내용(없으면 빈 문자열)',
    '- 심의필: 이미지에 보이는 심의필 문구 전체(예: 한화손해보험 준법감시인 심의필 제2026-1234호(2026.01.01~2026.12.31)). 없으면 빈 문자열',
    '',
    '[팀의 기존 태깅 예시 — 값별 실제 소재이름]'
  ];
  TAG_FIELDS.forEach(f => {
    lines.push(f + ' (선택지: ' + (settings[f] || []).join(', ') + ')');
    const byVal = examples[f] || {};
    Object.keys(byVal).forEach(v => lines.push('  · ' + v + ': ' + byVal[v].join(' | ')));
  });
  if (details.length) lines.push('소구상세 작성 예: ' + details.join(' | '));
  return lines.join('\n');
}

function _normalizeTagSuggestion(p, settings) {
  const pick = (f, v) => ((settings[f] || []).indexOf(_s(v)) !== -1 ? _s(v) : '');
  const tags = {};
  ['보종', '소재유형', '소구포인트', '후킹방식', '이미지유형'].forEach(f => { tags[f] = pick(f, p[f]); });
  const models = (Array.isArray(p.모델유형) ? p.모델유형 : []).map(v => pick('모델유형', v)).filter(Boolean);
  tags.모델유형 = models.filter((v, i) => models.indexOf(v) === i).join(',');
  tags.소구상세 = _s(p.소구상세);
  const review = _s(p.심의필);
  return { tags: tags, 심의필: review, 심의만료일: _parseReviewEndDate(review) };
}

// "…(2026.10.27~2027.10.26)" → '2027-10-26'. 판독 불가 글자(?)가 섞이면 빈 값.
function _parseReviewEndDate(text) {
  const m = String(text || '').match(/(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})\.?\s*[~∼〜\-]\s*(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})/);
  if (!m) return '';
  return m[4] + '-' + ('0' + m[5]).slice(-2) + '-' + ('0' + m[6]).slice(-2);
}

// --------------------------------------------------
// 일괄 저장 — entries: saveCreative와 같은 형태의 data 배열(+ 심의필, 심의만료일, driveFileId)
// 같은 이미지가 여러 소재에 쓰이면 첫 저장 후 파일해시로 같은 이미지코드가 재사용된다.
// --------------------------------------------------
function saveCreativesBulk(entries) {
  const results = [];
  const hashCache = {};
  (entries || []).forEach(data => {
    try {
      if (data.driveFileId && !data.fileHash) {
        data.fileHash = hashCache[data.driveFileId] || (hashCache[data.driveFileId] = _driveFileSha256(data.driveFileId));
      }
      // 처음 보는 이미지면 Drive 파일을 공개로 바꿔 URL을 만든다(이미 등록된 이미지면 saveCreative가 기존 URL을 재사용).
      if (data.driveFileId && !data.directUploadUrl && !checkExistingImage(data.fileHash)) data.directUploadUrl = setFilePublic(data.driveFileId);
      const res = saveCreative(Object.assign({ forceSave: true }, data));
      if (res.success && (data.심의필 || data.심의만료일)) _setMasterReviewInfo(res.imageCode, data);
      results.push(res);
    } catch (e) {
      results.push({ error: true, message: e.message });
    }
  });
  _ensureHierarchyRows(entries || []);
  return results;
}

function _driveFileSha256(fileId) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, DriveApp.getFileById(fileId).getBlob().getBytes());
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

// 등록 폼의 연쇄 드롭다운이 매체 구조 시트를 기준으로 하므로, 일괄 등록한 경로도 거기에 추가해 둔다.
function _ensureHierarchyRows(entries) {
  const sheet = getSpreadsheet().getSheetByName(HIERARCHY_SHEET_NAME);
  if (!sheet) return;
  const have = {};
  getHierarchy().forEach(h => { have[_creativeKey(h)] = true; });
  entries.forEach(e => {
    const k = _creativeKey(e);
    if (have[k] || !e.매체 || !e.소재이름) return;
    have[k] = true;
    sheet.appendRow([e.매체, e.캠페인, e.그룹, e.소재이름]);
  });
}

function _ensureMasterReviewHeaders(sheet) {
  if (sheet.getLastColumn() < MASTER_REVIEW_END_COL || !sheet.getRange(1, MASTER_REVIEW_COL).getValue()) {
    sheet.getRange(1, MASTER_REVIEW_COL, 1, 2).setValues([['심의필', '심의만료일']]).setFontWeight('bold');
  }
}

// 같은 이미지코드의 모든 행(재사용된 지면 포함)에 심의필 정보를 기록한다.
function _setMasterReviewInfo(imageCode, info) {
  const sheet = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  _ensureMasterReviewHeaders(sheet);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach((r, i) => {
    if (_s(r[0]) === imageCode) sheet.getRange(i + 2, MASTER_REVIEW_COL, 1, 2).setValues([[info.심의필 || '', info.심의만료일 || '']]);
  });
}

// --------------------------------------------------
// 심의필 백필 — 이미 등록된 이미지 중 심의필 정보가 없는 것을 한 번에 REVIEW_BACKFILL_BATCH장씩 판독한다.
// 남은 게 있으면 트리거로 이어서 처리한다(대시보드의 '심의필 정보 채우기' 버튼 → 끝날 때까지 자동 진행).
// --------------------------------------------------
const REVIEW_BACKFILL_BATCH = 24;

function backfillReviewExpiry() {
  const sheet = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { done: true, processed: 0, remaining: 0 };
  _ensureMasterReviewHeaders(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MASTER_REVIEW_END_COL).getValues();
  const todo = [];
  const seen = {};
  rows.forEach(r => {
    const code = _s(r[0]);
    if (!code || seen[code] || code.indexOf('VID') === 0) return;
    seen[code] = true;
    if (_s(r[MASTER_REVIEW_COL - 1]) || _s(r[MASTER_REVIEW_END_COL - 1])) return;
    const fileId = _extractDriveFileId(r[14]);
    if (fileId) todo.push({ code: code, fileId: fileId, 소재이름: r[5], 캠페인: r[3] });
  });

  const batch = todo.slice(0, REVIEW_BACKFILL_BATCH);
  const res = suggestCreativeTags(batch);
  batch.forEach((b, i) => {
    const r = res[i];
    // 판독 실패한 이미지는 '판독실패'로 표시해 다음 배치에서 무한 반복하지 않게 한다.
    _setMasterReviewInfo(b.code, r && !r.error ? { 심의필: r.심의필 || '(없음)', 심의만료일: r.심의만료일 } : { 심의필: '판독실패', 심의만료일: '' });
  });
  const remaining = todo.length - batch.length;
  if (remaining > 0) ScriptApp.newTrigger('backfillReviewExpiry').timeBased().after(3000).create();
  return { done: remaining === 0, processed: batch.length, remaining: remaining };
}

// --------------------------------------------------
// AI 태그 추천 정확도 측정 — 이미 사람이 태깅한 이미지 n장을 AI에 다시 태깅시켜 항목별 일치율을 본다.
// (자기 자신의 소재이름은 예시에서 빼고 측정)
// --------------------------------------------------
function evaluateTagSuggestions(n) {
  const sheet = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  const seen = {};
  const sample = [];
  rows.slice().reverse().forEach(r => {
    const code = _s(r[0]);
    const fileId = _extractDriveFileId(r[14]);
    if (sample.length >= (n || 30) || !fileId || seen[code] || code.indexOf('IMG') !== 0) return;
    seen[code] = true;
    sample.push({ fileId: fileId, 소재이름: r[5], 캠페인: r[3], truth: { 보종: r[6], 소재유형: r[8], 소구포인트: r[9], 후킹방식: r[10], 이미지유형: r[12], 모델유형: r[13] } });
  });
  const res = suggestCreativeTags(sample, sample.map(s => s.소재이름));
  const score = {};
  TAG_FIELDS.forEach(f => { score[f] = { match: 0, total: 0 }; });
  const misses = [];
  sample.forEach((s, i) => {
    const r = res[i];
    if (!r || r.error) return;
    TAG_FIELDS.forEach(f => {
      const want = f === '모델유형' ? _s(s.truth[f]).split(',').map(_s).sort().join(',') : _s(s.truth[f]);
      const got = f === '모델유형' ? _s(r.tags[f]).split(',').map(_s).sort().join(',') : _s(r.tags[f]);
      if (!want) return;
      score[f].total++;
      if (want === got) score[f].match++;
      else if (misses.length < 40) misses.push({ 소재이름: s.소재이름, 항목: f, 정답: want, 추천: got });
    });
  });
  return { n: sample.length, failed: res.filter(r => !r || r.error).length, score: score, misses: misses };
}
