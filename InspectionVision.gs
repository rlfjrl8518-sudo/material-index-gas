// ====================================================
// 소재 검수 - InspectionVision.gs
// Gemini(Vertex AI) Vision 이미지 판독 + ZIP 압축 해제
// ====================================================

// Vertex AI의 OpenAI 호환 chat/completions로 Gemini를 호출한다(2026-10-23, NVIDIA NIM 무료 API에서 전환).
// API 키 없이 스크립트 실행 계정의 OAuth 토큰(cloud-platform 스코프)으로 인증하므로, 그 계정이
// GCP_PROJECT_ID 프로젝트에 'Vertex AI 사용자' 권한이 있어야 하고 프로젝트에 Vertex AI API가 켜져 있어야 한다.
// 소재 검수 Vision · AI 인사이트 · 인사이트 챗봇이 모두 AI_MODEL 하나를 쓴다(소재 검수 Vision만 아래의 네이티브
// generateContent 엔드포인트를 쓴다 — responseSchema로 JSON 형태를 강제하기 위해).
// 모델 비교(2026-10-27): 3.5-flash(low)가 3.1-pro와 같은 정확도에 3~4배 빠르고, 3.5-flash-lite는 숫자를 지어내 제외.
const AI_MODEL = 'google/gemini-3.5-flash';
const VISION_CONCURRENCY = 8; // fetchAll 한 번에 동시 발사할 요청 수(이미지당 판독 2건) — 새 프로젝트는 8건 이상 동시 호출 시 429가 났다(2026-10-27 실측)
const VISION_TIME_BUDGET_MS = 4 * 60 * 1000; // 배치 1회 분석에 쓸 최대 시간(GAS 6분 한도 - 시트 기록 여유)
const VISION_MAX_RETRY = 3; // 호출(429·5xx)/파싱 실패 시 재시도 횟수

function _aiChatUrl() {
  return 'https://aiplatform.googleapis.com/v1/projects/' + encodeURIComponent(getAiProjectId()) +
    '/locations/global/endpoints/openapi/chat/completions';
}

// UrlFetchApp.fetch(단건)와 fetchAll(배치)이 같은 요청 옵션을 쓸 수 있게 url을 포함해 만든다
// (fetch(url, options)는 options 안의 url을 무시하므로 두 경로에 그대로 재사용해도 안전하다).
// Gemini는 max_tokens에 thinking 토큰도 포함돼, 넉넉히 주지 않으면 답이 비어서 온다.
function _aiRequestOptions(messages, opts) {
  return {
    url: _aiChatUrl(),
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({
      model: AI_MODEL,
      messages: messages,
      max_tokens: opts.maxTokens,
      temperature: opts.temperature,
      reasoning_effort: opts.reasoningEffort || 'low'
    }),
    muteHttpExceptions: true
  };
}

// { text } 또는 { error:true, message }. Vertex는 오류를 [{error:{...}}] 배열로 돌려주기도 한다.
function _parseAiHttpResponse(res) {
  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    return { error: true, message: 'Gemini 응답 파싱 실패 (HTTP ' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300) };
  }
  if (Array.isArray(json)) json = json[0] || {};
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    const msg = (json.error && json.error.message) || res.getContentText().slice(0, 300);
    return { error: true, message: 'Gemini API 오류 (HTTP ' + code + '): ' + msg };
  }
  const choice = json.choices && json.choices[0];
  const text = choice && choice.message && choice.message.content;
  if (!text) {
    return { error: true, message: '응답이 비어 있습니다' + (choice && choice.finish_reason === 'length' ? '(토큰 한도 초과)' : '') + ' — 다시 시도해보세요.' };
  }
  return { text: text };
}

function _callAiChat(messages, opts) {
  try {
    return _parseAiHttpResponse(UrlFetchApp.fetch(_aiChatUrl(), _aiRequestOptions(messages, opts)));
  } catch (e) {
    return { error: true, message: 'Gemini 호출 실패: ' + e.message };
  }
}

// --------------------------------------------------
// 소재 검수 Vision — Vertex AI 네이티브 generateContent (2026-10-27 개편)
// 실측(합성 광고 이미지 4종, 회당 3~4회 반복)에서 확인된 것들을 그대로 설계에 반영했다:
//  1) OCR 텍스트를 힌트로 같이 주면 모델이 OCR 오독을 그대로 베낀다(복잡한 배경 4/4회 '00812').
//     → Drive OCR 단계를 없애고 이미지만 보고 읽게 한다(이미지당 순차 OCR 병목도 같이 사라짐).
//  2) 기준값을 알려주고 대조시키면 틀린 기준값도 '일치'라고 답한다(9회 중 5회 false pass).
//     → 기준값은 절대 모델에 주지 않고 '보이는 그대로 받아 적기'만 시킨 뒤 비교는 코드가 한다.
//  3) 모델이 스스로 매기는 confidence는 틀린 답에도 0.9~0.95라 판정에 쓸 수 없다.
//     → 대신 못 읽은 글자를 '?'로 적게 하면 숫자를 지어내지 않고 '?????'로 솔직하게 답한다.
//  4) 같은 이미지를 두 번 독립 판독(VISION_READS)해서 두 결과가 다르면 확인 필요로 돌린다
//     (금융권 OCR의 이중 판독 대조 방식). 두 호출은 fetchAll로 동시에 나가 속도는 거의 그대로다.
//  5) responseSchema로 JSON 형태를 API가 강제하므로 JSON 파싱 실패 재시도가 사실상 없어진다.
// --------------------------------------------------
const VISION_READS = [{ temperature: 0.1 }, { temperature: 0.5 }];

const IMAGE_EXT_MIME = {
  jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp'
};

// 문구가 있는 영역 [ymin, xmin, ymax, xmax] (0~1000 정규화, Gemini 기본 좌표 규약) — 결과 화면에서 틀린 위치를 표시하는 데만 쓰고 판정에는 쓰지 않는다
const VISION_BOX_SCHEMA = { type: 'ARRAY', items: { type: 'INTEGER' } };

const VISION_RESPONSE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    productText: { type: 'OBJECT', properties: { text: { type: 'STRING' }, box: VISION_BOX_SCHEMA }, required: ['text'] },
    reviewText: { type: 'OBJECT', properties: { text: { type: 'STRING' }, box: VISION_BOX_SCHEMA }, required: ['text'] },
    allTexts: { type: 'ARRAY', items: { type: 'STRING' } },
    logo: { type: 'OBJECT', properties: { detected: { type: 'BOOLEAN', nullable: true }, description: { type: 'STRING' } }, required: ['detected', 'description'] },
    font: { type: 'OBJECT', properties: { matched: { type: 'BOOLEAN', nullable: true }, description: { type: 'STRING' } }, required: ['matched', 'description'] }
  },
  required: ['productText', 'reviewText', 'allTexts', 'logo', 'font']
};

function _visionUrl() {
  return 'https://aiplatform.googleapis.com/v1/projects/' + encodeURIComponent(getAiProjectId()) +
    '/locations/global/publishers/google/models/' + AI_MODEL.replace(/^google\//, '') + ':generateContent';
}

function _visionRequestOptions(parts, read) {
  return {
    url: _visionUrl(),
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: _visionSystemPrompt() }] },
      contents: [{ role: 'user', parts: parts }],
      generationConfig: {
        temperature: read.temperature,
        maxOutputTokens: 8000,
        responseMimeType: 'application/json',
        responseSchema: VISION_RESPONSE_SCHEMA,
        thinkingConfig: { thinkingLevel: 'LOW' }
      }
    }),
    muteHttpExceptions: true
  };
}

// 분석 결과 객체 또는 { error:true, message }
function _parseVisionResponse(res) {
  let json;
  try {
    json = JSON.parse(res.getContentText());
  } catch (e) {
    return { error: true, message: 'Gemini 응답 파싱 실패 (HTTP ' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300) };
  }
  if (Array.isArray(json)) json = json[0] || {};
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    const msg = (json.error && json.error.message) || res.getContentText().slice(0, 300);
    return { error: true, message: 'Gemini API 오류 (HTTP ' + code + '): ' + msg };
  }
  const cand = json.candidates && json.candidates[0];
  const parts = (cand && cand.content && cand.content.parts) || [];
  const text = parts.filter(p => p.text && !p.thought).map(p => p.text).join('');
  if (!text) {
    return { error: true, message: '응답이 비어 있습니다' + (cand && cand.finishReason ? '(' + cand.finishReason + ')' : '') + ' — 다시 시도해보세요.' };
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    return { error: true, message: 'JSON 파싱 실패: ' + text.slice(0, 200) };
  }
}

// --------------------------------------------------
// 이미지 1장 분석 — README에 문서화된 "편집기에서 직접 실행" 수동 테스트용 진입점
// --------------------------------------------------
function analyzeImageWithOpenAI(fileId, criteria) {
  return analyzeImagesWithOpenAIBatch([{ fileId: fileId }], criteria)[0];
}

// 이미지 여러 장을 한 번에 분석한다. items: [{fileId, fileName}].
// 이미지마다 VISION_READS 수만큼 독립 판독 요청을 만들어 UrlFetchApp.fetchAll로 동시에 보내고,
// 실패한 요청만 다음 라운드에서 재시도한다. 두 판독이 모두 성공하면 _mergeVisionReads로 합친다.
// 반환값: items와 같은 순서의 분석 결과 배열.
function analyzeImagesWithOpenAIBatch(items, criteria) {
  const startedAt = Date.now();
  if (!getAiProjectId()) return items.map(() => ({ error: true, message: 'GCP 프로젝트 ID가 설정되지 않았습니다.' }));

  const logoBlob = (criteria.logo && criteria.logo.use && criteria.logo.value) ? _resolveDriveRefToBlob(criteria.logo.value) : null;
  const built = items.map(item => {
    try {
      return { parts: _buildVisionParts(DriveApp.getFileById(item.fileId).getBlob(), logoBlob, criteria), error: null };
    } catch (e) {
      return { parts: null, error: '이미지 파일을 읽을 수 없습니다: ' + e.message };
    }
  });

  const reads = items.map(() => new Array(VISION_READS.length));
  const jobs = [];
  built.forEach((b, i) => {
    if (b.error) VISION_READS.forEach((_, r) => { reads[i][r] = { error: true, message: b.error }; });
    else VISION_READS.forEach((_, r) => jobs.push({ i: i, r: r }));
  });

  // VISION_CONCURRENCY건씩 나눠 보낸다. 혼잡 시 재시도 대기가 쌓여 GAS 실행 한도(6분)를 넘기면
  // 결과 기록·다음 배치 트리거 생성 전에 실행이 끊겨 검수가 '진행중'에서 멈춘다. 마감 이후엔
  // 재시도/남은 청크를 포기하고 오류로 돌려 '확인 필요'로 기록되게 한다
  // (결과 화면의 '확인 필요 항목 재처리'로 다시 돌릴 수 있음).
  const deadline = startedAt + VISION_TIME_BUDGET_MS;
  for (let c = 0; c < jobs.length; c += VISION_CONCURRENCY) {
    let pending = jobs.slice(c, c + VISION_CONCURRENCY);
    if (Date.now() > deadline) {
      pending.forEach(j => { reads[j.i][j.r] = { error: true, message: 'Gemini 서버 혼잡으로 시간 내 처리하지 못했습니다 — 확인 필요 항목 재처리로 다시 시도하세요.' }; });
      continue;
    }
    let rateLimited = false;
    for (let round = 0; round <= VISION_MAX_RETRY && pending.length && (round === 0 || Date.now() < deadline); round++) {
      // 503(동시 한도)은 몇 초면 풀리지만 429(분당 한도)는 1분 창이 지나야 풀리므로 길게 기다린다.
      if (round > 0) Utilities.sleep(rateLimited ? 15000 : 2000 * round);

      const requests = pending.map(j => _visionRequestOptions(built[j.i].parts, VISION_READS[j.r]));
      let responses;
      try {
        responses = UrlFetchApp.fetchAll(requests);
      } catch (e) {
        pending.forEach(j => { reads[j.i][j.r] = { error: true, message: 'Gemini 동시 호출 실패: ' + e.message }; });
        break;
      }

      const stillPending = [];
      rateLimited = false;
      pending.forEach((j, k) => {
        if (responses[k].getResponseCode() === 429) rateLimited = true;
        const parsed = _parseVisionResponse(responses[k]);
        reads[j.i][j.r] = parsed;
        if (parsed.error) stillPending.push(j);
      });
      pending = stillPending;
    }
  }
  return reads.map(_mergeVisionReads);
}

// 독립 판독 결과들 → judgeField 등이 쓰는 분석 결과 1건.
// 하나라도 실패하면 대조가 불가능하므로 이미지 전체를 오류(=확인 필요)로 돌린다.
// 문구는 각 판독값을 모두 넘겨 비교·합의 판단은 InspectionCompare.gs가 한다.
function _mergeVisionReads(rs) {
  const failed = rs.filter(r => !r || r.error);
  if (failed.length) return { error: true, message: (failed[0] && failed[0].message) || '판독 실패' };
  const texts = key => rs.map(r => String((r[key] && r[key].text) || '').trim());
  const logos = rs.map(r => r.logo || {});
  const logoDetected = logos.every(l => l.detected === logos[0].detected) ? logos[0].detected : null;
  return {
    productText: { text: texts('productText')[0], reads: texts('productText') },
    reviewText: { text: texts('reviewText')[0], reads: texts('reviewText') },
    boxes: { 상품명: _validBox(rs[0].productText && rs[0].productText.box), 심의필: _validBox(rs[0].reviewText && rs[0].reviewText.box) },
    allTexts: Array.isArray(rs[0].allTexts) ? rs[0].allTexts : [],
    allTextsReads: rs.map(r => Array.isArray(r.allTexts) ? r.allTexts : []),
    logo: { detected: logoDetected, description: (logos[0].description || '') + (logoDetected === null ? ' (2회 판독 결과 상이)' : '') },
    font: rs[0].font || {}
  };
}

function _validBox(b) {
  return Array.isArray(b) && b.length === 4 && b.every(function (n) { return typeof n === 'number' && n >= 0 && n <= 1000; }) && b[2] > b[0] && b[3] > b[1] ? b : null;
}

function _buildVisionParts(imageBlob, logoBlob, criteria) {
  const parts = [
    { text: _buildVisionPromptText(criteria) },
    { inlineData: { mimeType: imageBlob.getContentType() || 'image/png', data: Utilities.base64Encode(imageBlob.getBytes()) } }
  ];
  if (logoBlob) {
    parts.push({ text: '위 이미지는 검수 대상, 아래 이미지는 기준 로고입니다. 검수 대상 이미지에 기준 로고와 동일하거나 매우 유사한 로고가 있는지 확인하세요.' });
    parts.push({ inlineData: { mimeType: logoBlob.getContentType() || 'image/png', data: Utilities.base64Encode(logoBlob.getBytes()) } });
  }
  return parts;
}

function _visionSystemPrompt() {
  return [
    '너는 광고 이미지 소재의 문구를 글자 그대로 받아 적는 판독 도구다.',
    '이미지에 실제로 보이는 글자만 옮겨 적고, 보이지 않거나 흐린 글자를 문맥으로 추측해 채우지 마라.',
    '한 글자라도 확실히 읽을 수 없으면 그 글자 자리에 ?를 적어라(예: 제2026-??18호). 숫자는 특히 추측 금지다.',
    '화면 하단의 작은 고지 문구까지 모두 확인하라.'
  ].join(' ');
}

function _buildVisionPromptText(criteria) {
  const lines = ['이 이미지에서 다음을 판독하라.'];
  lines.push('- productText.text: 상품명 또는 보종명으로 보이는 문구');
  lines.push('- reviewText.text: 심의필 문구 전체(예: OO손해보험 준법감시인 심의필 제2026-1234호(2026.01.01~2026.12.31)). 없으면 빈 문자열');
  lines.push('- allTexts: 이미지에 보이는 모든 문구(줄 단위)');
  lines.push('- logo: 이미지 내 로고 유무(detected)와 설명');
  if (criteria.font && criteria.font.use && criteria.font.value) {
    const targetText = (criteria.font.targetText || '').trim();
    if (targetText) {
      lines.push('- font: 이미지에서 "' + targetText + '" 문구를 찾아 그 문구에만 적용된 폰트가 "' + criteria.font.value +
        '"와 유사해 보이는지 matched로 판단하라(다른 문구는 무시). 문구를 찾지 못하면 matched는 null, description은 "대상 문구를 찾을 수 없음".');
    } else {
      lines.push('- font: 이미지 전반의 문구 폰트가 "' + criteria.font.value + '"와 유사해 보이는지 matched로 판단');
    }
  } else {
    lines.push('- font: matched는 null, description은 빈 문자열');
  }
  lines.push('- productText.box / reviewText.box: 그 문구가 있는 영역을 [ymin, xmin, ymax, xmax] (이미지 크기를 0~1000으로 정규화)로. 문구가 없으면 빈 배열');
  lines.push('글자가 전혀 없는 항목은 빈 문자열로 답하라.');

  const custom = _redactCriteriaValues(criteria.customPrompt, criteria);
  if (custom) lines.push('추가 지시사항(검수 담당자가 직접 입력함): ' + custom);
  return lines.join('\n');
}

// 커스텀 지시사항에 상품명·심의필 같은 기준값이 들어 있으면 모델이 그 값을 '읽은 것처럼' 답해
// 틀린 소재를 정상으로 통과시킨다(실측 false pass). 모델에 보내기 전에 기준값을 가린다.
function _redactCriteriaValues(text, criteria) {
  let s = String(text || '').trim();
  if (!s) return '';
  _criteriaValues(criteria).forEach(v => { s = s.split(v).join('[기준값]'); });
  return s;
}

function _criteriaValues(criteria) {
  return [criteria.productName, criteria.reviewNumber, criteria.extra1, criteria.extra2]
    .map(f => String((f && f.value) || '').trim())
    .filter(v => v.length >= 2);
}

// 검수기준 시트의 '로고' 입력값(Drive 공유 URL 또는 파일 ID)을 Blob으로 변환
function _resolveDriveRefToBlob(ref) {
  const id = _extractDriveFileId(ref);
  if (!id) return null;
  try {
    return DriveApp.getFileById(id).getBlob();
  } catch (e) {
    return null;
  }
}

function _extractDriveFileId(ref) {
  if (!ref) return null;
  const s = String(ref).trim();
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{15,}$/.test(s)) return s;
  return null;
}

// --------------------------------------------------
// ZIP 파일 → 이미지만 추출해서 지정 폴더에 저장
// 숨김파일 / __MACOSX / 비이미지 / 손상파일은 건너뛴다
// --------------------------------------------------
function extractImagesFromZip(zipFileId, targetFolder) {
  let zipBlob;
  try {
    zipBlob = DriveApp.getFileById(zipFileId).getBlob();
  } catch (e) {
    return { error: true, message: 'ZIP 파일을 읽을 수 없습니다: ' + e.message };
  }

  let entries;
  try {
    entries = Utilities.unzip(zipBlob);
  } catch (e) {
    return { error: true, message: 'ZIP 압축 해제 실패: ' + e.message };
  }

  const images = [];
  const skipped = [];

  entries.forEach(function (entryBlob) {
    try {
      const name = entryBlob.getName() || '';
      const base = name.split('/').pop();

      if (!base || base.indexOf('__MACOSX') !== -1 || base.charAt(0) === '.') {
        skipped.push({ name: name, reason: '숨김/시스템 파일' });
        return;
      }
      const ext = (base.split('.').pop() || '').toLowerCase();
      const mime = IMAGE_EXT_MIME[ext];
      if (!mime) {
        skipped.push({ name: name, reason: '이미지 파일이 아님' });
        return;
      }
      if (entryBlob.getBytes().length === 0) {
        skipped.push({ name: name, reason: '손상된 파일(빈 파일)' });
        return;
      }

      entryBlob.setContentTypeFromExtension ? entryBlob : null; // no-op, 안전장치
      const file = targetFolder.createFile(entryBlob.setName(base));
      images.push({ fileId: file.getId(), fileName: base });
    } catch (e) {
      skipped.push({ name: entryBlob.getName ? entryBlob.getName() : '(알 수 없음)', reason: '손상된 파일: ' + e.message });
    }
  });

  return { images: images, skipped: skipped };
}
