// ====================================================
// 소재 검수 - InspectionVision.gs
// Gemini(Vertex AI) Vision 이미지 분석 + ZIP 압축 해제
// ====================================================

// Vertex AI의 OpenAI 호환 chat/completions로 Gemini를 호출한다(2026-10-23, NVIDIA NIM 무료 API에서 전환).
// API 키 없이 스크립트 실행 계정의 OAuth 토큰(cloud-platform 스코프)으로 인증하므로, 그 계정이
// GCP_PROJECT_ID 프로젝트에 'Vertex AI 사용자' 권한이 있어야 하고 프로젝트에 Vertex AI API가 켜져 있어야 한다.
// 소재 검수 Vision · AI 인사이트 · 인사이트 챗봇이 모두 AI_MODEL 하나를 쓴다.
const AI_MODEL = 'google/gemini-3.5-flash';
const VISION_CONCURRENCY = 20; // fetchAll 한 번에 동시 발사할 이미지 수(배치 크기와 같게 — 유료 쿼터라 한 번에 보낸다)
const VISION_TIME_BUDGET_MS = 4 * 60 * 1000; // 배치 1회 분석에 쓸 최대 시간(GAS 6분 한도 - 시트 기록 여유)
const VISION_MAX_RETRY = 3; // JSON 파싱/호출(429·5xx) 실패 시 재시도 횟수

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

// 문구 추출은 창작이 아니라 낮은 temperature로 고정.
const VISION_OPTS = { maxTokens: 8000, temperature: 0.1 };

const IMAGE_EXT_MIME = {
  jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp'
};

// --------------------------------------------------
// 이미지 1장 분석 → { productText, reviewText, allTexts, logo, font } 또는 { error:true, message }
// (README에 문서화된 "편집기에서 직접 실행" 수동 테스트용 진입점이라 시그니처를 그대로 둔다)
// --------------------------------------------------
function analyzeImageWithOpenAI(fileId, criteria) {
  if (!getAiProjectId()) return { error: true, message: 'GCP 프로젝트 ID가 설정되지 않았습니다.' };

  let imageBlob;
  try {
    imageBlob = DriveApp.getFileById(fileId).getBlob();
  } catch (e) {
    return { error: true, message: '이미지 파일을 읽을 수 없습니다: ' + e.message };
  }

  const ocrText = _ocrImageText(fileId);
  const messages = _buildVisionMessages(imageBlob, criteria, ocrText);

  let lastError = null;
  for (let attempt = 0; attempt <= VISION_MAX_RETRY; attempt++) {
    if (attempt > 0) Utilities.sleep(2000 * attempt);
    const res = _callAiChat(messages, VISION_OPTS);
    if (res.error) { lastError = res.message; continue; }
    const parsed = _extractJsonObject(res.text);
    if (parsed) return parsed;
    lastError = 'JSON 파싱 실패: 응답에서 JSON 객체를 찾지 못했습니다.';
  }
  return { error: true, message: lastError || '알 수 없는 오류' };
}

// 이미지 여러 장을 한 번에 분석한다. items: [{fileId, fileName}].
// Drive OCR(_ocrImageText)은 Drive Advanced Service 호출이라 fetchAll로 동시에 못 보내서
// 이미지당 순차로 하지만(원래도 순차였으니 이 부분은 속도가 그대로), Vision 호출은
// UrlFetchApp.fetchAll()로 대기 중인 항목을 한 번에 동시 발사한다 — 배치 처리 시간이
// "이미지 수 × 호출시간"에서 "가장 느린 호출 1번" 수준으로 준다. 프롬프트·모델·OCR은
// analyzeImageWithOpenAI와 완전히 동일한 _buildVisionMessages를 그대로 쓰므로 정확도는
// 그대로다. JSON 파싱 실패 등으로 실패한 항목만 다음 라운드에서 다시 동시 발사해
// 재시도한다(최대 VISION_MAX_RETRY+1라운드, analyzeImageWithOpenAI의 재시도 횟수와 동일).
// 반환값: items와 같은 순서의 분석 결과 배열.
function analyzeImagesWithOpenAIBatch(items, criteria) {
  const startedAt = Date.now();
  if (!getAiProjectId()) return items.map(() => ({ error: true, message: 'GCP 프로젝트 ID가 설정되지 않았습니다.' }));

  const built = items.map(item => {
    try {
      const imageBlob = DriveApp.getFileById(item.fileId).getBlob();
      const ocrText = _ocrImageText(item.fileId);
      return { messages: _buildVisionMessages(imageBlob, criteria, ocrText), error: null };
    } catch (e) {
      return { messages: null, error: '이미지 파일을 읽을 수 없습니다: ' + e.message };
    }
  });

  const results = new Array(items.length);
  const ready = [];
  built.forEach((b, i) => {
    if (b.error) results[i] = { error: true, message: b.error };
    else ready.push(i);
  });

  // VISION_CONCURRENCY장씩 나눠 보낸다(NVIDIA 무료 티어 땐 20장 동시 발사가 한도에 걸려 5장씩이었다).
  // 혼잡 시 재시도 대기가 쌓여 GAS 실행 한도(6분)를 넘기면 결과 기록·다음 배치 트리거 생성 전에
  // 실행이 끊겨 검수가 '진행중'에서 멈춘다. 마감 이후엔 재시도/남은 청크를 포기하고 오류로 돌려
  // '확인 필요'로 기록되게 한다(결과 화면의 '확인 필요 항목 재처리'로 다시 돌릴 수 있음).
  const deadline = startedAt + VISION_TIME_BUDGET_MS;
  for (let c = 0; c < ready.length; c += VISION_CONCURRENCY) {
    let pending = ready.slice(c, c + VISION_CONCURRENCY);
    if (Date.now() > deadline) {
      pending.forEach(i => { results[i] = { error: true, message: 'Gemini 서버 혼잡으로 시간 내 처리하지 못했습니다 — 확인 필요 항목 재처리로 다시 시도하세요.' }; });
      continue;
    }
    let rateLimited = false;
    for (let round = 0; round <= VISION_MAX_RETRY && pending.length && (round === 0 || Date.now() < deadline); round++) {
      // 재시도 라운드 사이엔 텀을 둔다 — 503(동시 한도)은 몇 초면 풀리지만 429(분당 한도)는
      // 1분 창이 지나야 풀리므로 길게 기다린다.
      if (round > 0) Utilities.sleep(rateLimited ? 15000 : 2000 * round);

      const requests = pending.map(i => _aiRequestOptions(built[i].messages, VISION_OPTS));
      let responses;
      try {
        responses = UrlFetchApp.fetchAll(requests);
      } catch (e) {
        pending.forEach(i => { results[i] = { error: true, message: 'Gemini 동시 호출 실패: ' + e.message }; });
        break;
      }

      const stillPending = [];
      rateLimited = false;
      pending.forEach((i, k) => {
        if (responses[k].getResponseCode() === 429) rateLimited = true;
        const res = _parseAiHttpResponse(responses[k]);
        if (res.error) { results[i] = res; stillPending.push(i); return; }
        const parsed = _extractJsonObject(res.text);
        if (parsed) { results[i] = parsed; return; }
        results[i] = { error: true, message: 'JSON 파싱 실패: 응답에서 JSON 객체를 찾지 못했습니다.' };
        stillPending.push(i);
      });
      pending = stillPending;
    }
  }
  return results;
}

// analyzeImageWithOpenAI / analyzeImagesWithOpenAIBatch가 공유하는 Vision 메시지 구성.
function _buildVisionMessages(imageBlob, criteria, ocrText) {
  const content = [{ type: 'text', text: _buildVisionPromptText(criteria, ocrText) }];
  content.push({ type: 'image_url', image_url: { url: _blobToDataUri(imageBlob) } });

  if (criteria.logo && criteria.logo.use && criteria.logo.value) {
    const refBlob = _resolveDriveRefToBlob(criteria.logo.value);
    if (refBlob) {
      content.push({ type: 'text', text: '위 이미지는 검수 대상, 아래 이미지는 기준 로고입니다. 검수 대상 이미지에 기준 로고와 동일하거나 매우 유사한 로고가 있는지 확인하세요.' });
      content.push({ type: 'image_url', image_url: { url: _blobToDataUri(refBlob) } });
    }
  }

  return [
    { role: 'system', content: _visionSystemPrompt() },
    { role: 'user', content: content }
  ];
}

// Google Drive 내장 OCR로 이미지에서 문자를 정확히 추출한다(은행권 신분증 검증 수준의 글자 단위 정확도 확보용).
// 실패해도 검수 전체가 멈추면 안 되므로 null을 반환하고 호출부(비전 프롬프트)에서 폴백한다.
function _ocrImageText(fileId) {
  if (!fileId) return null;
  let tempFileId = null;
  let text = null;
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const created = Drive.Files.insert({ title: 'OCR_TEMP_' + fileId, mimeType: MimeType.GOOGLE_DOCS }, blob, { ocr: true, ocrLanguage: 'ko' });
    tempFileId = created.id;
    text = DocumentApp.openById(tempFileId).getBody().getText();
  } catch (e) {
    text = null;
  }
  if (tempFileId) {
    try { Drive.Files.remove(tempFileId); } catch (e2) {}
  }
  return text;
}

// 모델은 response_format 강제 없이도 JSON만 답하도록 프롬프트로 지시하지만,
// 코드블록(```json ... ```)으로 감싸 보내는 경우가 있어 관대하게 파싱한다.
function _extractJsonObject(text) {
  if (!text) return null;
  try { return JSON.parse(text); } catch (e) {}
  const match = String(text).match(/\{[\s\S]*\}/);
  if (match) {
    try { return JSON.parse(match[0]); } catch (e) {}
  }
  return null;
}

function _visionSystemPrompt() {
  return [
    '너는 광고 이미지 소재의 문구를 확인하는 검수 보조 도구다.',
    '이미지에 실제로 보이는 문구만 반환하고, 보이지 않는 문구를 추측하지 마라.',
    '상품명으로 보이는 문구와 심의필로 보이는 문구를 각각 추출하고, 화면 하단의 작은 문구도 최대한 확인하라.',
    '문구가 흐리거나 잘려서 불확실하면 낮은 confidence 값을 반환하라.',
    '문구를 정확히 읽지 못했다면 절대로 그럴듯한 상품명이나 문구를 지어내지 마라. 읽을 수 없으면 text를 빈 문자열로, confidence를 0으로 반환하고, 확실히 읽은 글자만 text에 담아라.',
    '반드시 JSON 형식으로만 답하고 다른 설명은 출력하지 마라.'
  ].join(' ');
}

function _buildVisionPromptText(criteria, ocrText) {
  const lines = [];
  let ocrAvailable = false;
  if (ocrText) {
    ocrAvailable = true;
    lines.push('아래는 Google Drive OCR 엔진이 이 이미지에서 추출한 참고용 텍스트입니다. OCR은 화려한 배경·작은 글씨·장식 폰트가 많은 광고 소재에서 문구를 통째로 놓치는 경우가 있으니 절대적인 정답이 아니라 1차 힌트로만 참고하고, 최종 판단은 반드시 이미지를 직접 보고 하라. OCR에 없는 문구라도 이미지에 실제로 선명하게 보이면 text에 포함해도 된다 — 다만 실제로 읽은 것이 아니라 그럴듯하게 지어낸 문구는 절대 포함하지 말고, 조금이라도 불확실하면 confidence를 낮게 매겨라.');
    lines.push('--- OCR 텍스트 시작 ---');
    lines.push(ocrText);
    lines.push('--- OCR 텍스트 끝 ---');
  } else {
    lines.push('OCR 텍스트를 가져오지 못했다. 이미지에서 직접 읽되, 확신이 없는 글자는 낮은 confidence로 표시하라.');
  }
  lines.push('이 이미지에서 다음을 분석해서 JSON으로 반환하라.');
  lines.push('- productText: 상품명 또는 보종명으로 보이는 문구, confidence(0~1)' + (ocrAvailable ? ', ocrConfirmed(boolean) — 이 text가 위 OCR 텍스트 안에 그대로(또는 사소한 오탈자 수준으로) 있으면 true, 이미지에서는 보이지만 OCR 텍스트엔 전혀 없으면 false. 정답 여부 판단이 아니라 OCR과의 교차검증용이니 솔직하게 표시하라.' : ''));
  lines.push('- reviewText: 심의필 문구(예: 확인필-제OOOO-NNNNNN호 형태)로 보이는 문구, confidence(0~1)' + (ocrAvailable ? ', ocrConfirmed(boolean) — 위와 동일한 기준' : ''));
  if (!ocrAvailable) {
    lines.push('- (OCR 텍스트가 없으므로 productText/reviewText의 ocrConfirmed는 항상 null로 답하라)');
  }
  lines.push('- allTexts: 이미지에서 인식되는 전체 문구 목록(배열) — 마찬가지로 OCR에 없어도 이미지에서 실제로 보이면 포함하라');
  lines.push('- logo: { detected: boolean, description: string, confidence } — 이미지 내 로고 유무와 설명');
  if (criteria.font && criteria.font.use && criteria.font.value) {
    const targetText = (criteria.font.targetText || '').trim();
    if (targetText) {
      lines.push(
        '- font: { matched: boolean|null, description: string, confidence } — ' +
        '이미지에서 "' + targetText + '" 라는 문구를 찾아라. 그 문구가 이미지에 있으면, ' +
        '그 문구에만 적용된 폰트가 "' + criteria.font.value + '"와 유사해 보이는지 판단하라(이미지의 다른 문구는 무시). ' +
        '"' + targetText + '" 문구를 이미지에서 찾지 못하면 matched는 null, description에 "대상 문구를 찾을 수 없음"이라고 답하라.'
      );
    } else {
      lines.push('- font: { matched: boolean|null, description: string, confidence } — 이미지 전반의 문구 폰트가 "' + criteria.font.value + '"와 유사해 보이는지 판단 (대상 문구가 따로 지정되지 않음)');
    }
  } else {
    lines.push('- font: { matched: null, description: "", confidence: 0 }');
  }
  lines.push('응답 예시 형식:');
  lines.push('{"productText":{"text":"","confidence":0.9,"ocrConfirmed":true},"reviewText":{"text":"","confidence":0.9,"ocrConfirmed":true},"allTexts":[""],"logo":{"detected":false,"description":"","confidence":0.5},"font":{"matched":null,"description":"","confidence":0.5}}');
  lines.push('텍스트가 전혀 보이지 않는 항목은 text를 빈 문자열로, confidence를 0으로 반환하라.');

  if (criteria.customPrompt) {
    lines.push('추가 지시사항(검수 담당자가 직접 입력함, 반드시 반영하라): ' + criteria.customPrompt);
  }

  return lines.join('\n');
}

function _blobToDataUri(blob) {
  const mime = blob.getContentType() || 'image/png';
  return 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
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
