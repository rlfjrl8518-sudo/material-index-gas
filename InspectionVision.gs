// ====================================================
// 소재 검수 - InspectionVision.gs
// OpenAI Vision 이미지 분석 + ZIP 압축 해제
// ====================================================

// 이 프로젝트는 기존 AI 인사이트 기능(Code.gs getOpenAIInsight)에서 gpt-5-mini를 쓰고 있고,
// 실측으로 확인된 제약을 그대로 따른다: temperature는 기본값(1) 외엔 거부됨 → 아예 넣지 않음,
// max_tokens가 아니라 max_completion_tokens 사용, reasoning_effort 미지정 시 토큰을 reasoning에
// 다 쓰고 content가 빈 문자열로 오는 경우가 있어 명시적으로 지정한다.
const OPENAI_MODEL = 'gpt-5-mini';
const OPENAI_CHAT_URL = 'https://api.openai.com/v1/chat/completions';
const VISION_MAX_RETRY = 2; // JSON 파싱/호출 실패 시 재시도 횟수

const IMAGE_EXT_MIME = {
  jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp'
};

// --------------------------------------------------
// 이미지 1장 분석 → { productText, reviewText, allTexts, logo, font } 또는 { error:true, message }
// --------------------------------------------------
function analyzeImageWithOpenAI(fileId, criteria) {
  const apiKey = getOpenAIApiKey();
  if (!apiKey) return { error: true, message: 'OpenAI API 키가 설정되지 않았습니다.' };

  let imageBlob;
  try {
    imageBlob = DriveApp.getFileById(fileId).getBlob();
  } catch (e) {
    return { error: true, message: '이미지 파일을 읽을 수 없습니다: ' + e.message };
  }

  const content = [{ type: 'text', text: _buildVisionPromptText(criteria) }];
  content.push({ type: 'image_url', image_url: { url: _blobToDataUri(imageBlob) } });

  if (criteria.logo && criteria.logo.use && criteria.logo.value) {
    const refBlob = _resolveDriveRefToBlob(criteria.logo.value);
    if (refBlob) {
      content.push({ type: 'text', text: '위 이미지는 검수 대상, 아래 이미지는 기준 로고입니다. 검수 대상 이미지에 기준 로고와 동일하거나 매우 유사한 로고가 있는지 확인하세요.' });
      content.push({ type: 'image_url', image_url: { url: _blobToDataUri(refBlob) } });
    }
  }

  const messages = [
    { role: 'system', content: _visionSystemPrompt() },
    { role: 'user', content: content }
  ];

  let lastError = null;
  for (let attempt = 0; attempt <= VISION_MAX_RETRY; attempt++) {
    const res = _callOpenAIChat(apiKey, messages);
    if (res.error) { lastError = res.message; continue; }
    const parsed = _extractJsonObject(res.text);
    if (parsed) return parsed;
    lastError = 'JSON 파싱 실패: 응답에서 JSON 객체를 찾지 못했습니다.';
  }
  return { error: true, message: lastError || '알 수 없는 오류' };
}

// gpt-5 계열은 response_format 강제 없이도 JSON만 답하도록 프롬프트로 지시하지만,
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
    '반드시 JSON 형식으로만 답하고 다른 설명은 출력하지 마라.'
  ].join(' ');
}

function _buildVisionPromptText(criteria) {
  const lines = [];
  lines.push('이 이미지에서 다음을 분석해서 JSON으로 반환하라.');
  lines.push('- productText: 상품명 또는 보종명으로 보이는 문구와 confidence(0~1)');
  lines.push('- reviewText: 심의필 문구(예: 확인필-제OOOO-NNNNNN호 형태)로 보이는 문구와 confidence(0~1)');
  lines.push('- allTexts: 이미지에서 인식되는 전체 문구 목록(배열)');
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
  lines.push('{"productText":{"text":"","confidence":0.9},"reviewText":{"text":"","confidence":0.9},"allTexts":[""],"logo":{"detected":false,"description":"","confidence":0.5},"font":{"matched":null,"description":"","confidence":0.5}}');
  lines.push('텍스트가 전혀 보이지 않는 항목은 text를 빈 문자열로, confidence를 0으로 반환하라.');

  if (criteria.customPrompt) {
    lines.push('추가 지시사항(검수 담당자가 직접 입력함, 반드시 반영하라): ' + criteria.customPrompt);
  }

  return lines.join('\n');
}

function _callOpenAIChat(apiKey, messages) {
  const payload = {
    model: OPENAI_MODEL,
    messages: messages,
    // gpt-5 계열: max_tokens 대신 max_completion_tokens, temperature는 기본값 외 거부됨(둘 다
    // Code.gs의 getOpenAIInsight/getOpenAIChatReply에서 2026-07-28 실측으로 확인된 제약).
    // 이미지 분석 + JSON 조립이라 순수 텍스트 답변보다 reasoning 여유가 더 필요해 'low'로 지정.
    max_completion_tokens: 1500,
    reasoning_effort: 'low'
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const res = UrlFetchApp.fetch(OPENAI_CHAT_URL, options);
    const code = res.getResponseCode();
    const json = JSON.parse(res.getContentText());
    if (code < 200 || code >= 300) {
      const msg = (json.error && json.error.message) || res.getContentText().slice(0, 300);
      return { error: true, message: 'OpenAI API 오류 (HTTP ' + code + '): ' + msg };
    }
    const choice = json.choices && json.choices[0];
    const text = choice && choice.message && choice.message.content;
    if (!text) {
      if (choice && choice.finish_reason === 'length') {
        return { error: true, message: '응답이 비어 있습니다(토큰 한도 초과) — reasoning 토큰 소모로 content 없음' };
      }
      return { error: true, message: 'OpenAI 응답에 content가 없습니다.' };
    }
    return { text: text };
  } catch (e) {
    return { error: true, message: 'OpenAI 호출 실패: ' + e.message };
  }
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
