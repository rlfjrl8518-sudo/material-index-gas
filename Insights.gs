// ====================================================
// 오늘 할 일 / 아침 요약 - Insights.gs
// 사람이 대시보드를 뒤지지 않아도 '지금 손대야 할 것'만 모아 준다.
// 같은 계산(getActionItems)을 대시보드 상단 카드와 매일 아침 메신저 알림(sendDailyDigest)이 같이 쓴다.
// ====================================================

const EXPIRY_ALERT_DAYS = 14;        // 심의 만료 이 기간 안이면 알림
const ACTIVE_LOOKBACK_DAYS = 7;      // '집행 중' 판단 기간
const FATIGUE_MIN_IMP = 3000;        // 피로도 판단 최소 노출(초기 7일·최근 7일 각각)
const FATIGUE_CTR_DROP = 0.7;        // 최근 CTR이 초기의 70% 미만이면 교체 권장
const FATIGUE_MIN_DAYS = 14;         // 집행 일수가 이보다 짧으면 피로도 판단 안 함

// --------------------------------------------------
// 소재_통합RAW → 소재(매체·캠페인·그룹·소재이름)별 일자 데이터
// --------------------------------------------------
function _loadCreativeDaily() {
  const sheet = getSpreadsheet().getSheetByName(CONSOLIDATED_RAW_SHEET_NAME);
  const out = { byKey: {}, lastDate: '' };
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, CONSOLIDATED_RAW_HEADERS.length).getValues().forEach(r => {
    const d = _normDate(r[1]);
    if (!(d instanceof Date)) return;
    const ds = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
    const key = _creativeKey({ 매체: r[0], 캠페인: r[2], 그룹: r[3], 소재이름: r[4] });
    const c = out.byKey[key] || (out.byKey[key] = { 매체: _s(r[0]), 캠페인: _s(r[2]), 그룹: _s(r[3]), 소재이름: _s(r[4]), 이미지URL: _s(r[13]), days: {} });
    const day = c.days[ds] || (c.days[ds] = { imp: 0, clk: 0, cost: 0, conv: 0 });
    day.imp += Number(r[14]) || 0;
    day.clk += Number(r[15]) || 0;
    day.cost += Number(r[17]) || 0;
    day.conv += Number(r[18]) || 0;
    if (ds > out.lastDate) out.lastDate = ds;
  });
  return out;
}

function _sumDays(c, from, to) {
  const s = { imp: 0, clk: 0, cost: 0, conv: 0 };
  Object.keys(c.days).forEach(ds => {
    if (ds < from || ds > to) return;
    const d = c.days[ds];
    s.imp += d.imp; s.clk += d.clk; s.cost += d.cost; s.conv += d.conv;
  });
  return s;
}

function _shiftDate(ds, n) {
  return Utilities.formatDate(new Date(new Date(ds + 'T00:00:00+09:00').getTime() + n * 864e5), 'Asia/Seoul', 'yyyy-MM-dd');
}

function _median(arr) {
  const a = arr.slice().sort((x, y) => x - y);
  return a.length ? (a.length % 2 ? a[(a.length - 1) / 2] : (a[a.length / 2 - 1] + a[a.length / 2]) / 2) : 0;
}

// --------------------------------------------------
// 피로도 — 초기 7일 대비 최근 7일 CTR이 크게 떨어진 집행 중 소재
// --------------------------------------------------
function _findFatigued(daily) {
  if (!daily.lastDate) return [];
  const recentFrom = _shiftDate(daily.lastDate, -(ACTIVE_LOOKBACK_DAYS - 1));
  const out = [];
  Object.keys(daily.byKey).forEach(k => {
    const c = daily.byKey[k];
    const dates = Object.keys(c.days).filter(ds => c.days[ds].imp > 0).sort();
    if (dates.length < FATIGUE_MIN_DAYS || dates[dates.length - 1] < recentFrom) return;
    const first = _sumDays(c, dates[0], _shiftDate(dates[0], 6));
    const recent = _sumDays(c, recentFrom, daily.lastDate);
    if (first.imp < FATIGUE_MIN_IMP || recent.imp < FATIGUE_MIN_IMP || !first.clk) return;
    const ctr0 = first.clk / first.imp, ctr1 = recent.clk / recent.imp;
    if (ctr1 >= ctr0 * FATIGUE_CTR_DROP) return;
    out.push({
      매체: c.매체, 캠페인: c.캠페인, 그룹: c.그룹, 소재이름: c.소재이름, 이미지URL: c.이미지URL,
      집행일수: dates.length, 초기CTR: ctr0, 최근CTR: ctr1, 하락률: 1 - ctr1 / ctr0, 최근비용: recent.cost
    });
  });
  return out.sort((a, b) => b.최근비용 - a.최근비용);
}

// --------------------------------------------------
// 최근 7일 성과 상·하위 — 같은 매체 안의 CPA 중앙값과 비교
//  · 상위: 전환 5건 이상, CPA가 매체 중앙값의 60% 이하
//  · 하위: 전환 3건 이상인데 CPA가 중앙값의 1.8배 이상, 또는 전환 0건인데 중앙값 CPA의 2배 넘게 씀
// --------------------------------------------------
function _findWinnersLosers(daily) {
  if (!daily.lastDate) return { winners: [], losers: [] };
  const from = _shiftDate(daily.lastDate, -(ACTIVE_LOOKBACK_DAYS - 1));
  const byMedia = {};
  Object.keys(daily.byKey).forEach(k => {
    const c = daily.byKey[k];
    const s = _sumDays(c, from, daily.lastDate);
    if (!s.cost) return;
    (byMedia[c.매체] || (byMedia[c.매체] = [])).push({ c: c, s: s });
  });
  const winners = [], losers = [];
  Object.keys(byMedia).forEach(m => {
    const list = byMedia[m];
    const med = _median(list.filter(x => x.s.conv >= 3).map(x => x.s.cost / x.s.conv));
    if (!med) return;
    list.forEach(x => {
      const cpa = x.s.conv ? x.s.cost / x.s.conv : null;
      const row = { 매체: m, 캠페인: x.c.캠페인, 그룹: x.c.그룹, 소재이름: x.c.소재이름, 이미지URL: x.c.이미지URL, 비용: x.s.cost, 전환: x.s.conv, CPA: cpa, 매체중앙CPA: med };
      if (x.s.conv >= 5 && cpa <= med * 0.6) winners.push(row);
      else if ((x.s.conv >= 3 && cpa >= med * 1.8) || (!x.s.conv && x.s.cost >= med * 2)) losers.push(row);
    });
  });
  const bySpend = (a, b) => b.비용 - a.비용;
  return { winners: winners.sort(bySpend).slice(0, 10), losers: losers.sort(bySpend).slice(0, 10) };
}

// --------------------------------------------------
// 심의 만료 임박 — 전매체 인덱스 R열(심의만료일)과 최근 7일 집행 여부를 잇는다
// --------------------------------------------------
function _findExpiring(daily) {
  const sheet = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
  const res = { items: [], unknown: 0 };
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < MASTER_REVIEW_END_COL) return res;
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const limit = _shiftDate(today, EXPIRY_ALERT_DAYS);
  const activeFrom = daily.lastDate ? _shiftDate(daily.lastDate, -(ACTIVE_LOOKBACK_DAYS - 1)) : today;
  const seen = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, MASTER_REVIEW_END_COL).getValues().forEach(r => {
    const key = _creativeKey({ 매체: r[2], 캠페인: r[3], 그룹: r[4], 소재이름: r[5] });
    const c = daily.byKey[key];
    const recent = c ? _sumDays(c, activeFrom, daily.lastDate || today) : null;
    if (!recent || !recent.cost) return; // 집행 중인 소재만
    const end = r[MASTER_REVIEW_END_COL - 1] instanceof Date
      ? Utilities.formatDate(r[MASTER_REVIEW_END_COL - 1], 'Asia/Seoul', 'yyyy-MM-dd') : _s(r[MASTER_REVIEW_END_COL - 1]);
    const dedupe = _s(r[0]) + '\x00' + key;
    if (seen[dedupe]) return;
    seen[dedupe] = true;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(end)) { res.unknown++; return; }
    if (end > limit) return;
    res.items.push({
      이미지코드: _s(r[0]), 매체: _s(r[2]), 캠페인: _s(r[3]), 그룹: _s(r[4]), 소재이름: _s(r[5]), 이미지URL: _s(r[14]),
      심의필: _s(r[MASTER_REVIEW_COL - 1]), 만료일: end,
      남은일수: Math.round((new Date(end + 'T00:00:00+09:00') - new Date(today + 'T00:00:00+09:00')) / 864e5),
      최근비용: recent.cost
    });
  });
  res.items.sort((a, b) => a.남은일수 - b.남은일수);
  return res;
}

function _recentInspectionIssues() {
  try {
    // 최근 7일 안의 검수 중 아직 진행 중이거나 사람이 볼 항목(불일치·확인 필요)이 남은 것만
    const since = Date.now() - 7 * 864e5;
    return getRecentInspections(20).filter(h => new Date(h.date).getTime() >= since &&
      (h.status === '진행중' || Number(h.needCheck) > 0 || Number(h.mismatch) > 0));
  } catch (e) {
    return [];
  }
}

// --------------------------------------------------
// 대시보드 '오늘 할 일' 카드 + 아침 요약이 같이 쓰는 데이터
// --------------------------------------------------
function getActionItems() {
  const daily = _loadCreativeDaily();
  const queue = getRegistrationQueue();
  const wl = _findWinnersLosers(daily);
  return {
    기준일: daily.lastDate,
    queue: { count: queue.count, top: queue.items.slice(0, 5) },
    expiring: _findExpiring(daily),
    fatigued: _findFatigued(daily).slice(0, 10),
    winners: wl.winners,
    losers: wl.losers,
    inspections: _recentInspectionIssues()
  };
}

// --------------------------------------------------
// 아침 요약 — 설정 관리 > API 키의 DIGEST_WEBHOOK_URL(Slack 또는 Google Chat 수신 웹훅)으로 보낸다.
// Vercel Cron(매일 08시 KST)이 호출한다. 웹훅이 없으면 아무것도 안 한다.
// --------------------------------------------------
function sendDailyDigest() {
  const url = PropertiesService.getScriptProperties().getProperty('DIGEST_WEBHOOK_URL');
  if (!url) return { skipped: true, message: 'DIGEST_WEBHOOK_URL 미설정' };
  const text = buildDigestText(getActionItems());
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: text }), muteHttpExceptions: true });
  return { success: res.getResponseCode() < 300, code: res.getResponseCode(), text: text };
}

function buildDigestText(a) {
  const won = n => Math.round(n).toLocaleString('ko-KR') + '원';
  const pct = n => (n * 100).toFixed(2) + '%';
  const name = x => x.매체 + ' · ' + x.소재이름;
  const lines = ['📌 소재 운영 아침 요약 (데이터 기준일 ' + (a.기준일 || '-') + ')'];
  const section = (title, list, fmt) => {
    if (!list.length) return;
    lines.push('', title);
    list.slice(0, 5).forEach(x => lines.push('  • ' + fmt(x)));
    if (list.length > 5) lines.push('  … 외 ' + (list.length - 5) + '건');
  };
  section('⛔ 심의 만료 임박(집행 중)', a.expiring.items, x => name(x) + ' — ' + (x.남은일수 < 0 ? '만료됨' : x.남은일수 + '일 남음') + ' (' + x.만료일 + ')');
  section('📉 피로도 — 교체 검토', a.fatigued, x => name(x) + ' — CTR ' + pct(x.초기CTR) + ' → ' + pct(x.최근CTR) + ' (−' + Math.round(x.하락률 * 100) + '%)');
  section('🔻 최근 7일 성과 하위', a.losers, x => name(x) + ' — 비용 ' + won(x.비용) + ', ' + (x.전환 ? 'CPA ' + won(x.CPA) : '전환 0') + ' (매체 중앙 ' + won(x.매체중앙CPA) + ')');
  section('🏆 최근 7일 성과 상위', a.winners, x => name(x) + ' — CPA ' + won(x.CPA) + ' (매체 중앙 ' + won(x.매체중앙CPA) + '), 전환 ' + x.전환);
  if (a.queue.count) section('📝 미등록 소재 ' + a.queue.count + '건(광고비순)', a.queue.top, x => name(x) + ' — ' + won(x.비용));
  section('🔍 확인이 필요한 검수', a.inspections, x => x.id + ' — 불일치 ' + x.mismatch + ', 확인 필요 ' + x.needCheck + ' (' + x.status + ')');
  if (lines.length === 1) lines.push('', '오늘은 따로 손댈 항목이 없습니다 👍');
  lines.push('', '대시보드: ' + DASHBOARD_URL);
  return lines.join('\n');
}

// --------------------------------------------------
// 다음 소재 제작 브리프 — 대시보드가 계산한 근거(속성별 순효과·조합 성과·미시도 조합·피로도)를 받아
// 제작팀에 그대로 넘길 수 있는 브리프를 쓴다. 숫자는 받은 데이터에 있는 것만 인용하게 한다.
// --------------------------------------------------
function generateCreativeBrief(payload) {
  if (!getAiProjectId()) return { error: true, message: 'GCP 프로젝트 ID가 설정되지 않았습니다.' };
  const prompt = [
    '너는 보험 퍼포먼스 광고의 크리에이티브 전략가다. 아래 데이터만 근거로 다음에 만들 소재 제작 브리프를 한국어 마크다운으로 써라.',
    '규칙: 데이터에 없는 숫자를 지어내지 말 것. 근거를 댈 때는 받은 수치를 그대로 인용할 것. "표본부족"인 값은 근거로 쓰지 말고 테스트 대상으로만 다룰 것.',
    '',
    '형식:',
    '## 1. 지금 통하는 공식',
    '유의한 순효과 위주로 3~5줄. 각 줄에 근거 수치.',
    '## 2. 교체·중단 검토',
    '피로도 소재와 성과가 나쁜 조합. 없으면 "해당 없음".',
    '## 3. 제작안 (3~5개)',
    '각 안마다: **안 이름** / 보종 · 소재유형 · 소구포인트 · 후킹방식 · 이미지유형 · 모델 / 헤드라인 예시 2개(보험 광고 심의를 고려해 과장·단정적 보장 표현 금지) / 근거 / 기대 효과.',
    '최소 1개는 "미시도조합" 중 가능성 높은 조합을 검증하는 실험안으로.',
    '## 4. 테스트 설계',
    '어떤 안끼리 비교할지, 성공 판단 지표와 최소 표본(예: 안별 전환 30건 또는 클릭 1,000회).',
    '',
    '데이터(JSON):',
    JSON.stringify(payload)
  ].join('\n');
  const res = _callAiChat([{ role: 'user', content: prompt }], { maxTokens: 12000, temperature: 0.4 });
  return res.error ? res : { success: true, text: res.text };
}
