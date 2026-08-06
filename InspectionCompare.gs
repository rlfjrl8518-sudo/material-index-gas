// ====================================================
// 소재 검수 - InspectionCompare.gs
// 문자열 정규화/비교/판정 — AI가 아니라 GAS 함수로 처리한다
// ====================================================

// --------------------------------------------------
// 정규화
// 심의필(isReviewNumber=true)은 숫자·하이픈이 중요하므로
// specialCharStrict 설정과 무관하게 특수문자를 제거하지 않는다
// --------------------------------------------------
function normalizeText(text, options, isReviewNumber) {
  let s = String(text || '');
  s = s.normalize('NFKC');

  if (options.ignoreLineBreak) s = s.replace(/[\r\n]+/g, ' ');

  if (options.spacingStrict) {
    s = s.replace(/\s+/g, ' ').trim();
  } else {
    s = s.replace(/\s+/g, '');
  }

  if (!options.caseSensitive) s = s.toLowerCase();

  if (!isReviewNumber && !options.specialCharStrict) {
    s = s.replace(/[!"#$%&'()*+,\-./:;<=>?@\[\]^_`{|}~·•""'']/g, '');
  }

  return s;
}

// --------------------------------------------------
// 글자 단위 차이 — 공통 접두/접미를 잘라내고 가운데 다른 부분만 보여준다
// 예: "확인필-제2026-123456호" vs "확인필-제2026-123465호"
//     → "123456 → 123465"
// --------------------------------------------------
function diffText(base, recognized) {
  const a = String(base || '');
  const b = String(recognized || '');
  if (a === b) return '';

  const maxOverlap = Math.min(a.length, b.length);
  let i = 0;
  while (i < maxOverlap && a[i] === b[i]) i++;

  let j = 0;
  while (j < (maxOverlap - i) && a[a.length - 1 - j] === b[b.length - 1 - j]) j++;

  const aMid = a.slice(i, a.length - j);
  const bMid = b.slice(i, b.length - j);

  if (aMid && bMid) return aMid + ' → ' + bMid;
  if (!aMid && bMid) return '(추가됨) ' + bMid;
  if (aMid && !bMid) return '(누락됨) ' + aMid;
  return '기준: ' + a + ' / 인식값: ' + b;
}

// --------------------------------------------------
// 필드 1개 판정 (상품명 / 심의필 / 기타문구1 / 기타문구2 공통)
// use=false면 검사 자체를 안 하므로 null 반환
// --------------------------------------------------
function judgeField(baseValue, useFlag, recognizedObj, options, isReviewNumber) {
  if (!useFlag) return null;

  if (!recognizedObj || recognizedObj.error) {
    return { result: '확인 필요', recognizedText: '', diff: '', reason: 'API 응답 실패' };
  }

  const text = (recognizedObj.text || '').trim();
  const confidence = Number(recognizedObj.confidence);

  if (!text) {
    return { result: '불일치', recognizedText: '', diff: diffText(baseValue, ''), reason: '문구 없음' };
  }

  if (!isNaN(confidence) && confidence < options.ocrConfidence) {
    return { result: '확인 필요', recognizedText: text, diff: diffText(baseValue, text), reason: '낮은 OCR 신뢰도' };
  }

  const normBase = normalizeText(baseValue, options, isReviewNumber);
  const normRec = normalizeText(text, options, isReviewNumber);

  if (normBase === normRec) {
    return { result: '일치', recognizedText: text, diff: '' };
  }
  return { result: '불일치', recognizedText: text, diff: diffText(baseValue, text) };
}

// --------------------------------------------------
// 로고 판정 — 존재 여부 + 기준 로고와의 유사 여부만 확인 (MVP)
// --------------------------------------------------
function judgeLogo(logoCriteria, recognizedLogo, ocrConfidence) {
  if (!logoCriteria || !logoCriteria.use) return null;

  if (!recognizedLogo || recognizedLogo.error) {
    return { result: '확인 필요', detail: 'API 응답 실패' };
  }
  const confidence = Number(recognizedLogo.confidence);
  const detail = recognizedLogo.description || '';

  if (!isNaN(confidence) && confidence < ocrConfidence) {
    return { result: '확인 필요', detail: detail };
  }
  if (recognizedLogo.detected === true) return { result: '일치', detail: detail };
  if (recognizedLogo.detected === false) return { result: '불일치', detail: detail || '로고가 확인되지 않음' };
  return { result: '확인 필요', detail: detail };
}

// --------------------------------------------------
// 폰트 판정 — 보조 항목. 최종 결과 계산에서는 항상 제외한다
// --------------------------------------------------
function judgeFont(fontCriteria, recognizedFont) {
  if (!fontCriteria || !fontCriteria.use) return null;

  if (!recognizedFont || recognizedFont.error) {
    return { result: '확인 필요', detail: 'API 응답 실패' };
  }
  const detail = recognizedFont.description || '';
  if (recognizedFont.matched === true) return { result: '일치 가능성 높음', detail: detail };
  if (recognizedFont.matched === false) return { result: '불일치 가능성 있음', detail: detail };
  return { result: '확인 필요', detail: detail };
}

// --------------------------------------------------
// 최종 결과 — 폰트는 참고 항목이라 제외하고 계산한다 (호출부에서 결과 배열에 폰트를 넣지 않는다)
// --------------------------------------------------
function computeFinalResult(fieldResults) {
  const active = fieldResults.filter(function (r) { return r !== null && r !== undefined; });
  if (active.some(function (r) { return r.result === '불일치'; })) return '불일치';
  if (active.some(function (r) { return r.result === '확인 필요'; })) return '확인 필요';
  return '정상';
}

// --------------------------------------------------
// 수정 요청 문구 — 실제로 발견된 불일치 항목만 포함한다
// items: [{ label, base, recognizedText }]
// --------------------------------------------------
function buildRevisionMessage(fileName, mismatchItems) {
  if (!mismatchItems || mismatchItems.length === 0) return '';

  const lines = [fileName + ' 소재 수정 요청드립니다.', ''];
  mismatchItems.forEach(function (item, idx) {
    lines.push((idx + 1) + '. ' + item.label + '이(가) 입력한 기준값과 다르게 적용되어 있습니다.');
    lines.push('- 기준: ' + item.base);
    lines.push('- 이미지 인식값: ' + (item.recognizedText || '(인식되지 않음)'));
    lines.push('');
  });
  lines.push('확인 후 수정 부탁드립니다.');
  return lines.join('\n');
}
