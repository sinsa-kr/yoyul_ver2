// 대시보드 핵심 동작 회귀 검증 (jsdom).
// ※ 이 하네스는 스크래치패드에서 두 번 유실된 뒤 리포 안으로 옮겨온 것이다.
const XLSX = require('xlsx-js-style');
const { load, runner, makeWageWorkbook } = require('./harness');

const T = runner('smoke');
const { ok, near, section } = T;

// ────────────────────────────────────────────────────────────
section('[1] 초기 로드');
{
  const A = load();
  ok('스크립트 에러 없음', A.errors.length === 0, A.errors.slice(0, 2).join(' | '));
  ok('탭 패널 2개', !!A.$('tab_calc') && !!A.$('tab_cmp'));
  ok('탭② 기본 숨김', A.$('tab_cmp').style.display === 'none');
  ok('예시 업무 1건 생성', !!A.$('work_1'));
  ok('비교 행 자동 생성', !!A.$('cmpRow_1'));
  ok('Chart 미로딩 가드 통과 (범례는 렌더)', A.$('legendTbl').innerHTML.length > 0);
  const b = A.win.computeAll();
  ok('기본 예시 인원 3명', b.totHead === 3, String(b.totHead));
  ok('월 총액 > 0', b.grandTot > 0);
}

// ────────────────────────────────────────────────────────────
section('[2] 계산 엔진 — 산식 불변');
{
  const A = load();
  const { win } = A;
  A.set('wRate_1', 10000); A.set('wHead_1', 1);
  A.set('wReg_1', 200); A.set('wOT_1', 0); A.set('wNt_1', 0); A.set('wWk_1', 0);
  A.check('wHolFlat_1', false);
  win.calculate();
  const r = win.computeAll().results[0];
  ok('기본급 = 시급 × 소정', r.wBase === 10000 * 200, String(r.wBase));
  ok('연장 0 → 가산 0', r.wOT === 0);

  A.set('wOT_1', 10); A.set('wNt_1', 20); win.calculate();
  const r2 = win.computeAll().results[0];
  ok('연장 가산 = 시급×h×0.5', r2.wOT === 10000 * 10 * 0.5, String(r2.wOT));
  ok('야간 가산 = 시급×h×0.5', r2.wNt === 10000 * 20 * 0.5, String(r2.wNt));

  A.set('wWk_1', 8); win.calculate();
  const r3 = win.computeAll().results[0];
  ok('휴일 법정 = 시급×h×0.5', r3.wHol === 10000 * 8 * 0.5, String(r3.wHol));
  A.check('wHolFlat_1', true); A.set('wHolRate_1', '9,000'); win.calculate();
  const r4 = win.computeAll().results[0];
  ok('휴일 정액 = 정액×h (가산 아님)', r4.wHol === 9000 * 8, String(r4.wHol));
  A.check('wHolFlat_1', false); win.calculate();
  ok('정액 해제 시 법정 계산으로 복귀', win.computeAll().results[0].wHol === 40000);

  // 합계 항등식
  const g = win.computeAll();
  const agg = win.aggCost(g.results);
  const sum = Object.values(agg).reduce((a, b) => a + b, 0);
  ok('★ 비용 7분류 합 = 월 총액', near(sum, g.grandTot, 1e-9), `${Math.round(sum)} vs ${Math.round(g.grandTot)}`);
}

// ────────────────────────────────────────────────────────────
section('[3] 근무시간 프리셋 · 캘린더');
{
  const A = load();
  const { win } = A;
  const DEF = win.eval('JSON.stringify(SHIFT_DEF)');
  const def = JSON.parse(DEF);
  ok('주간고정 209/0/0', def.day.regular === 209 && def.day.overtime === 0 && def.day.night === 0);
  ok('격일제 209/120/60', def.alternate.regular === 209 && def.alternate.overtime === 120 && def.alternate.night === 60);
  ok('4조3교대 야간 61h (오후조 22~24 포함분 반영)', def.shift43.night === 61, String(def.shift43.night));

  A.set('wShift_1', 'day'); win.onShiftChange(1);
  ok('교대형태 변경 → 시간 자동 입력', A.num('wReg_1') === 209 && A.num('wOT_1') === 0);
  ok('주간고정 → 주휴 가산 ON', A.$('wRest_1').checked === true);

  // 캘린더 환산
  const WD = { start: '09:00', end: '17:00' }, SAT = { start: '09:00', end: '13:00' };
  const pat = { weeks: 1, days: [null, WD, WD, WD, WD, WD, SAT] };
  const inc = win.computePatternHours(pat, false);
  const sep = win.computePatternHours(pat, true);
  ok('분리 OFF: 소정에 토요일 포함', inc.hoursRegular > sep.hoursRegular, `${inc.hoursRegular} vs ${sep.hoursRegular}`);
  ok('분리 ON: 소정에서 토요일 제외', inc.hoursRegular - sep.hoursRegular === Math.round(4 * (365 / 7 / 12)));
  ok('휴일근로 시간은 두 경우 동일', inc.holidayHours === sep.holidayHours && inc.holidayHours > 0);

  // 4조3교대 조별 균형
  win.sbApplyPreset && win.openShiftBuilder(1);
  win.sbApplyPreset('shift43');
  const days = JSON.parse(win.eval('JSON.stringify(sbDraft.days)'));
  const cnt = { day: 0, eve: 0, night: 0 };
  for (const d of days) {
    if (!d) continue;
    if (d.start === '08:00') cnt.day++;
    else if (d.start === '16:00') cnt.eve++;
    else if (d.start === '00:00') cnt.night++;
  }
  ok('★ 4조3교대 조별 배치 균등 (7/7/7)', cnt.day === 7 && cnt.eve === 7 && cnt.night === 7,
     JSON.stringify(cnt));
  win.closeShiftBuilder();
}

// ────────────────────────────────────────────────────────────
section('[4] 경고 — 최저임금 · 연장근로 한도');
{
  const A = load();
  const { win } = A;
  ok('격일제 예시: §63 적용제외 기본 ON', A.$('wExempt_1').checked === true);
  ok('격일제 예시: 한도 배너 미표시', !A.vis('otLimitAlert'));

  A.check('wExempt_1', false); win.calculate();
  ok('적용제외 해제 → 배너 표시', A.vis('otLimitAlert'));
  ok('배너에 업무명 포함', A.$('otLimitAlert').innerHTML.includes('경비'));

  A.set('wOT_1', 52); win.calculate();
  ok('경계값 52h → 경고 없음', !A.vis('otLimitAlert'));
  A.set('wOT_1', 53); win.calculate();
  ok('53h → 경고 표시', A.vis('otLimitAlert'));

  A.set('wRate_1', 9000); win.calculate();
  ok('최저임금 미만 → 배너', A.vis('minWageAlert'));
  ok('최저임금 미만 → 카드 인라인 경고', A.vis('wWarn_1'));
  A.set('wRate_1', 10320); win.calculate();
  ok('최저임금 이상 → 경고 해제', !A.vis('minWageAlert') && !A.vis('wWarn_1'));
}

// ────────────────────────────────────────────────────────────
section('[5] 인상률 시뮬레이션');
{
  const A = load();
  const { win } = A;
  const before = win.computeAll().grandTot;
  ok('adj 없으면 기존과 동일', win.computeAll().grandTot === before);
  ok('adj 전부 0 도 동일', near(win.computeAll({ rate:0, wf:0, fx:0, mgmt:0 }).grandTot, before, 1e-9));

  A.set('wRate_1', 10000); win.calculate();
  const sim3 = win.computeAll({ rate:3, wf:0, fx:0, mgmt:0 }).grandTot;
  ok('시뮬은 원본 입력을 안 바꿈', A.num('wRate_1') === 10000);
  A.set('wRate_1', 10300); win.calculate();
  ok('시급 +3% 시뮬 = 10,300 직접입력', near(sim3, win.computeAll().grandTot, 1e-9));
  A.set('wRate_1', 10000); win.calculate();

  A.set('mgmtRate', 10); win.calculate();
  const simM = win.computeAll({ rate:0, wf:0, fx:0, mgmt:2 }).grandTot;
  A.set('mgmtRate', 12); win.calculate();
  ok('관리비율 +2%p = 12% 직접입력', near(simM, win.computeAll().grandTot, 1e-9));
  A.set('mgmtRate', 10); win.calculate();

  A.set('simRate', -3); win.calculate();
  ok('인하(-3%)가 부호 그대로 읽힘', win.readRaiseAdj().rate === -3, String(win.readRaiseAdj().rate));
  A.set('simRate', 0);

  A.set('simRate', 5); A.set('simMgmt', 2.5); win.calculate();
  win.applyRaise();
  ok('반영 후 시급 10,000 → 10,500', A.num('wRate_1') === 10500, A.$('wRate_1').value);
  ok('반영 후 관리비율 12.5', A.num('mgmtRate') === 12.5);
  ok('반영 후 시뮬 입력 리셋', ['simRate','simWf','simFx','simMgmt'].every(id => A.num(id) === 0));
}

// ────────────────────────────────────────────────────────────
section('[6] 산정 정보 · 파일명');
{
  const A = load();
  const { win } = A;
  ok('초기 metaLine 숨김', !A.vis('metaLine'));
  A.set('metaVendor', '남양산업'); A.set('metaFrom', '2026-01-01'); A.set('metaTo', '2026-12-31');
  A.set('metaAuthor', '구매팀 홍길동');
  win.renderMetaLine();
  ok('metaLine 표시', A.vis('metaLine'));
  ok('metaLine 에 계약기간', A.$('metaLine').innerHTML.includes('2026.01.01 ~ 2026.12.31'));

  ok('파일명 태그 — 금지문자·공백 제거', (A.set('metaVendor', '남양 산업/2026'), win.metaFileTag()) === '_남양산업2026', win.metaFileTag());
  A.set('metaVendor', '');
  ok('업체명 비면 태그 없음', win.metaFileTag() === '');
}

// ────────────────────────────────────────────────────────────
section('[7] 저장 / 복원 라운드트립');
{
  const A = load();
  const { win } = A;
  A.set('wName_1', '경비'); A.set('wRate_1', 11000); A.set('wHead_1', 4);
  A.set('metaVendor', '테스트업체');
  win.calculate();
  const g0 = win.computeAll().grandTot;
  const snap = win.collectState();
  ok('collectState 에 meta·exempt 포함',
     snap.meta.vendor === '테스트업체' && typeof snap.works[0].exempt === 'boolean');

  win.applyState(JSON.parse(JSON.stringify(snap)));
  ok('복원 후 금액 동일', near(win.computeAll().grandTot, g0, 1));
  ok('복원 후 업체명 유지', A.$('metaVendor').value === '테스트업체');

  // 구버전 저장본 — meta / exempt / jobTitle / wageTable 없음
  const legacy = JSON.parse(JSON.stringify(snap));
  delete legacy.meta; delete legacy.wageTable;
  delete legacy.works[0].exempt; delete legacy.works[0].jobTitle;
  delete legacy.works[0].holidayHours; legacy.works[0].weekendDays = 2;
  win.applyState(legacy);
  ok('구버전 저장본 로드 시 예외 없음', A.errors.length === 0, A.errors.slice(0, 2).join(' | '));
  ok('구버전 weekendDays 2일 → 16h 환산', A.num('wWk_1') === 16, A.$('wWk_1').value);
  ok('구버전 → 메타 비움', A.$('metaVendor').value === '');

  win.resetAll();
  ok('초기화 후 예시 1건 복구', !!A.$('work_1'));
}

// ────────────────────────────────────────────────────────────
section('[8] 업무 순서 불변식 (맨 위 추가)');
{
  const A = load();
  const { win } = A;
  win.addWork({ name: '두번째', hourlyRate: 12000, headcount: 1 }, true);   // atTop
  const ids = win.workIdsInOrder();
  ok('새 업무가 맨 위', ids[0] === 2, JSON.stringify(ids));
  const res = win.computeAll().results;
  ok('결과 순서 = 화면 순서', res[0].name === '두번째', res.map(r => r.name).join(','));
  const st = win.collectState();
  ok('저장 순서도 화면 순서', st.works[0].name === '두번째');
  ok('비교탭 행도 같은 순서', win.compareAll(win.computeAll()).rows[0].name === '두번째');
}

// ────────────────────────────────────────────────────────────
section('[9] 직종 단가표 — 파싱');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  const P = wb => win.wtParseWorkbook(wb);

  let p = P(makeWageWorkbook(XLSX, [['보통인부', 172068], ['용접공', 282536]]));
  ok('표준 헤더 파싱', p.rows.length === 2 && p.rows[0].job === '보통인부' && p.rows[0].dayWage === 172068);

  p = P(makeWageWorkbook(XLSX, [[172068, '보통인부']], { head: ['일급', '직종명'] }));
  ok('열 순서 뒤바뀜 대응', p.rows[0].job === '보통인부' && p.rows[0].dayWage === 172068);

  p = P(makeWageWorkbook(XLSX, [['보통인부', 172068]], { pre: [['2026 노임단가'], ['출처: 대건협'], []] }));
  ok('헤더가 4행에서 시작해도 찾음', p.rows.length === 1);

  p = P(makeWageWorkbook(XLSX, [['보통인부', '172,068원']]));
  ok('콤마·원 붙은 문자열 숫자 파싱', p.rows[0].dayWage === 172068, String(p.rows[0].dayWage));

  p = P(makeWageWorkbook(XLSX, [['보통인부', 172068], ['', 999], ['용접공', 282536]]));
  ok('직종명 없는 행은 제외', p.rows.length === 2, `rows=${p.rows.length}`);

  // 단가가 0·공란·'-' 인 행은 고를 수 없는 직종이라 목록·검색에 넣지 않는다
  p = P(makeWageWorkbook(XLSX, [['보통인부', 172068], ['연마공', '-'], ['함석공', ''], ['조력공', 0]]));
  ok('★ 단가 없는 직종은 제외', p.rows.length === 1 && p.rows[0].job === '보통인부', `rows=${p.rows.map(r=>r.job).join(',')}`);
  ok('제외 건수를 보고', p.unpriced === 3, `unpriced=${p.unpriced}`);
  ok('단가 있는 행은 그대로', p.rows[0].unpriced === false && p.rows[0].dayWage === 172068);

  // 제외된 직종은 검색·목록 어느 쪽에서도 잡히지 않아야 한다
  win.wtApply(p, { name: 'u', source: 'upload', fileName: 'u.xlsx' });
  ok('제외된 직종은 조회되지 않음', win.wtFind('연마공') === null && win.wtFind('함석공') === null);
  ok('0원 직종도 조회되지 않음', win.wtFind('조력공') === null);
  ok('제외된 직종은 검색 결과에 없음',
     win.jobMatches('공').every(m => m.r.job !== '연마공' && m.r.job !== '함석공'),
     win.jobMatches('공').map(m => m.r.job).join(','));
  win.renderWtList();
  ok('제외된 직종은 목록보기에 없음', !A.$('wtList').textContent.includes('연마공'));
  ok('패널에 제외 건수 표시', A.txt('wtStatus').includes('3건 제외'), A.txt('wtStatus'));

  // 단가 없는 행이 빠지면서 동명이직이 하나만 남으면 구분자도 붙지 않아야 한다
  p = P(makeWageWorkbook(XLSX, [['작업반장', 215907], ['작업반장', '-']]));
  ok('★ 남는 동명이직이 1건이면 구분자 없음', p.rows.length === 1 && p.rows[0].job === '작업반장', p.rows.map(r=>r.job).join('|'));

  // 단가가 하나도 없는 표는 "열을 못 찾았다"가 아니라 진짜 이유를 알려야 한다
  let wtErr = null;
  try { P(makeWageWorkbook(XLSX, [['연마공', '-'], ['함석공', '']])); } catch (e) { wtErr = e; }
  ok('★ 전부 단가 없으면 그 이유로 거절', !!wtErr && wtErr.message.includes('단가가 있는 직종이 하나도'), wtErr?.message);

  // 동명이직 — 분야·적용월만 다른 같은 이름. 첫 행만 남기면 나머지를 고를 수 없다
  p = P(makeWageWorkbook(XLSX, [
    ['작업반장', 215907, '일', '2026년 01월', '건설'],
    ['작업반장', 147122, '일', '2026년 07월', '제조'],
  ], { head: ['직종명', '단가', '단위', '기준', '해설'] }));
  ok('★ 동명이직 전부 보존', p.rows.length === 2 && p.dup === 1, `rows=${p.rows.length} dup=${p.dup}`);
  ok('기준이 다르면 적용월로 구분',
     p.rows.map(r => r.job).join('|') === '작업반장 (2026년 01월)|작업반장 (2026년 07월)',
     p.rows.map(r => r.job).join('|'));

  p = P(makeWageWorkbook(XLSX, [
    ['기술사', 461459, '일', '2026년 01월', '엔지니어링사업_전기'],
    ['기술사', 467217, '일', '2026년 01월', '엔지니어링사업_건설'],
  ], { head: ['직종명', '단가', '단위', '기준', '해설'] }));
  ok('기준이 같으면 해설(분야)로 구분',
     p.rows.map(r => r.job).join('|') === '기술사 (엔지니어링사업_전기)|기술사 (엔지니어링사업_건설)',
     p.rows.map(r => r.job).join('|'));

  p = P(makeWageWorkbook(XLSX, [['같은것', 100], ['같은것', 200]]));
  ok('구분할 정보가 없으면 순번', p.rows.map(r => r.job).join('|') === '같은것 (#1)|같은것 (#2)',
     p.rows.map(r => r.job).join('|'));
  ok('어떤 경우에도 직종명은 유일', new Set(p.rows.map(r => r.job)).size === p.rows.length);

  p = P(makeWageWorkbook(XLSX, [['보통인부', 172068, 1.47]], { head: ['직종', '일급', '산재율'] }));
  ok('산재율 열 인식', p.rows[0].accidentRate === 1.47, String(p.rows[0].accidentRate));

  let threw = null;
  try { P(makeWageWorkbook(XLSX, [['a', 1]], { head: ['이름', '값'] })); } catch (e) { threw = e; }
  ok('필수 열 없으면 안내 에러', !!threw && threw.wtParse === true && /직종명 열/.test(threw.message));
}

// ────────────────────────────────────────────────────────────
section('[9-1] 표준 양식 — 직종명·단가·단위·기준·해설');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  const HEAD = ['직종명', '단가', '단위', '기준', '해설'];
  const p = win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['작업반장', 215907, '일', '2026년 01월', '각 공종별로 인부를 통솔하여 작업을 지휘하는 사람(십장)'],
    ['보통인부', 172068, '일', '2026년 01월', '기능을 요하지 않는 경작업인 일반잡역에 종사하는 사람'],
  ], { head: HEAD }));

  ok('첨부 양식 5열 파싱', p.rows.length === 2 && p.rows[0].job === '작업반장' && p.rows[0].dayWage === 215907);
  ok('단위 열 인식 (일)', p.rows[0].unit === 'day', String(p.rows[0].unit));
  ok('기준 열을 표 대표값으로', p.basis === '2026년 01월', p.basis);
  ok('해설 열 보존', p.rows[1].desc.includes('일반잡역'), p.rows[1].desc);
  ok('「직종명」이 직종 열로 인식됨 (「단위」를 단가로 오인하지 않음)',
     p.rows[0].dayWage === 215907 && p.rows[0].job === '작업반장');

  win.wtApply(p, { name: '시중노임단가', source: 'upload', fileName: 'w.xlsx' });
  ok('표에 기준 저장', win.eval('wageTable.basis') === '2026년 01월');
  ok('패널에 기준 표시', A.$('wtStatus').textContent.includes('2026년 01월 기준'), A.$('wtStatus').textContent);
  ok('일 단가 → ÷8 환산', win.wtRateOf('작업반장') === Math.round(215907 / 8), String(win.wtRateOf('작업반장')));

  // 드롭다운에 해설이 함께 뜬다
  A.set('wJob_1', '작업'); win.openJobList(1);
  const opt = A.$('wJobList_1').querySelector('.job-opt');
  ok('목록에 해설 표시', opt.textContent.includes('십장'), opt.textContent);
  ok('목록에 단위 포함한 단가', opt.textContent.includes('215,907원/일'));

  // ── 단위가 「시간」이면 나누지 않는다 (8로 나누면 크게 틀린다) ──
  const ph = win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['경비원', 12000, '시간', '2026년 01월', ''],
  ], { head: HEAD }));
  ok('단위 시간 인식', ph.rows[0].unit === 'hour', String(ph.rows[0].unit));
  win.wtApply(ph, { name: '시급표', source: 'upload', fileName: 'h.xlsx' });
  ok('★ 시간 단가는 그대로 시급', win.wtRateOf('경비원') === 12000, String(win.wtRateOf('경비원')));
  ok('시간 단가만 있으면 ÷N 표기 숨김', win.wtHasDayRows() === false);
  win.onJobPick(1, '경비원');
  ok('근거 줄에 ÷ 표기 없음', !A.$('wJobNote_1').textContent.includes('÷'), A.$('wJobNote_1').textContent);

  // ── 월 단가는 209h 로 ──
  const pm = win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['사무직', 2508000, '월', '2026년 01월', ''],
  ], { head: HEAD }));
  win.wtApply(pm, { name: '월급표', source: 'upload', fileName: 'm.xlsx' });
  ok('월 단가 → ÷209 환산', win.wtRateOf('사무직') === Math.round(2508000 / 209), String(win.wtRateOf('사무직')));

  // ── 단위 열이 없거나 알 수 없으면 일(日)로 본다 (기존 동작 유지) ──
  const pn = win.wtParseWorkbook(makeWageWorkbook(XLSX, [['보통인부', 172068, '???', '', '']], { head: HEAD }));
  ok('알 수 없는 단위 → 일로 간주', pn.rows[0].unit === null);
  win.wtApply(pn, { name: 'x', source: 'upload', fileName: 'x.xlsx' });
  ok('단위 미지정도 ÷8', win.wtRateOf('보통인부') === 21509, String(win.wtRateOf('보통인부')));
  ok('기준 없으면 빈 문자열', win.eval('wageTable.basis') === '');

  // ── 혼합 단위 표 ──
  const pmix = win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['보통인부', 172068, '일', '2026년 01월', ''],
    ['경비원',    12000, '시간', '2026년 01월', ''],
  ], { head: HEAD }));
  win.wtApply(pmix, { name: '혼합표', source: 'upload', fileName: 'mix.xlsx' });
  ok('혼합표 — 일 단가는 ÷8', win.wtRateOf('보통인부') === 21509);
  ok('혼합표 — 시간 단가는 그대로', win.wtRateOf('경비원') === 12000);
  ok('혼합표에도 일 단가가 있으면 ÷N 표기 유지', win.wtHasDayRows() === true);
}

// ────────────────────────────────────────────────────────────
section('[9-2] 해설 검색 · 하단 미리보기 · 하이라이트');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  const HEAD = ['직종명', '단가', '단위', '기준', '해설'];
  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['작업반장', 215907, '일', '2026년 01월', '각 공종별로 인부를 통솔하여 작업을 지휘하는 사람(십장)'],
    ['보통인부', 172068, '일', '2026년 01월', '기능을 요하지 않는 경작업인 일반잡역에 종사하는 사람'],
    ['건설기계운전사', 283297, '일', '2026년 01월', '지게차 등 건설기계를 운전하는 사람'],
    ['무해설직종', 100000, '일', '2026년 01월', ''],
  ], { head: HEAD })), { name: 't', source: 'upload', fileName: 't.xlsx' });

  const opts = () => [...A.$('wJobList_1').querySelectorAll('.job-opt')];
  const prev = () => A.$('wJobPrev_1');

  // ── 해설로 검색 ──
  A.set('wJob_1', '지게차'); win.openJobList(1);
  ok('★ 해설의 단어로 직종을 찾음 (지게차 → 건설기계운전사)',
     opts().length === 1 && opts()[0].dataset.job === '건설기계운전사',
     opts().map(o => o.dataset.job).join(','));
  ok('해설로 걸린 항목에 「해설」 배지', opts()[0].innerHTML.includes('hitlbl'));

  A.set('wJob_1', '십장'); win.openJobList(1);
  ok('해설 검색 — 십장 → 작업반장', opts().length === 1 && opts()[0].dataset.job === '작업반장');

  // 이름 일치가 해설 일치보다 앞
  A.set('wJob_1', '인부'); win.openJobList(1);
  const order = opts().map(o => o.dataset.job);
  ok('이름 일치를 해설 일치보다 먼저 배치',
     order[0] === '보통인부' && order.includes('작업반장') && order.indexOf('보통인부') < order.indexOf('작업반장'),
     JSON.stringify(order));
  ok('이름으로 걸린 항목엔 배지 없음', !opts()[0].innerHTML.includes('hitlbl'));

  // ── 하단 미리보기 ──
  A.set('wJob_1', ''); win.openJobList(1);
  ok('목록 열면 첫 항목 해설을 미리보기에 표시',
     prev().style.display === 'block' && prev().textContent.includes('십장'), prev().textContent);
  ok('미리보기에 환산 시급도 표시', prev().textContent.includes('26,988'), prev().textContent);  // 215,907 ÷ 8

  win.showJobPrev(1, 2);
  ok('짚은 항목으로 미리보기 갱신', prev().textContent.includes('건설기계를 운전하는'), prev().textContent);

  win.showJobPrev(1, 3);
  ok('해설 없는 직종은 「해설 없음」', prev().textContent.includes('해설 없음'));

  A.set('wJob_1', '없는직종'); win.openJobList(1);
  ok('일치 없으면 미리보기 숨김', prev().style.display === 'none');

  // 키보드 이동이 미리보기를 따라오는지
  A.set('wJob_1', ''); win.openJobList(1);
  const kev = k => ({ key: k, preventDefault(){} });
  win.onJobKey(kev('ArrowDown'), 1);
  win.onJobKey(kev('ArrowDown'), 1);
  ok('↑↓ 이동 시 미리보기 동기화', prev().textContent.includes('일반잡역'), prev().textContent);
  ok('짚은 행에 sel 표시', opts()[1].classList.contains('sel'));

  // ── 하이라이트 · 이스케이프 ──
  A.set('wJob_1', '인부'); win.openJobList(1);
  ok('검색어 하이라이트 <mark>', opts()[0].innerHTML.includes('<mark>인부</mark>'), opts()[0].innerHTML.slice(0, 120));
  ok('미리보기 제목에도 하이라이트', prev().innerHTML.includes('<mark>인부</mark>'), prev().innerHTML.slice(0, 120));

  // ★ 직종명·해설은 엑셀에서 온 외부 데이터 — 태그가 요소로 살아나면 안 된다.
  //   (직렬화된 innerHTML 의 data-* 속성값에는 < 가 원문으로 보이는 게 정상이므로,
  //    문자열이 아니라 실제로 만들어진 DOM 요소로 판정한다)
  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['<img src=x onerror=alert(1)>', 50000, '일', '', '<script>alert(2)<\/script> 위험한 해설'],
  ], { head: HEAD })), { name: 'x', source: 'upload', fileName: 'x.xlsx' });
  A.set('wJob_1', ''); win.openJobList(1);
  const pop = A.$('wJobPop_1');
  ok('★ 주입된 img 요소가 생기지 않음', pop.querySelectorAll('img').length === 0);
  ok('★ 주입된 script 요소가 생기지 않음', pop.querySelectorAll('script').length === 0);
  ok('직종명이 글자 그대로 표시됨',
     pop.querySelector('.job-opt b').textContent === '<img src=x onerror=alert(1)>',
     pop.querySelector('.job-opt b').textContent);
  ok('해설도 글자 그대로 표시됨',
     pop.querySelector('.job-opt .d').textContent.includes('<script>alert(2)</script>'),
     pop.querySelector('.job-opt .d').textContent);
  ok('미리보기에도 요소가 안 생김',
     prev().querySelectorAll('img,script').length === 0 && prev().textContent.includes('<script>'));

  // 검색어에 HTML 특수문자가 들어가도 하이라이트가 이스케이프를 깨지 않는다
  A.set('wJob_1', '<img'); win.openJobList(1);
  ok('특수문자 검색어에서도 요소 미생성', A.$('wJobPop_1').querySelectorAll('img,script').length === 0);
  ok('특수문자 검색어도 하이라이트됨', A.$('wJobList_1').innerHTML.includes('<mark>&lt;img</mark>'),
     A.$('wJobList_1').innerHTML.slice(0, 160));
}

// ────────────────────────────────────────────────────────────
section('[10] 직종 단가표 — 환산 · 선택 · 시급 수정');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['보통인부', 172068], ['인력운반공', 180193], ['일반기계운전사', 173995], ['저임금직', 60000],
  ])), { name: '테스트표', source: 'upload', fileName: 't.xlsx' });

  ok('★ 172,068 ÷ 8 = 21,509 (견적서 참고표와 일치)', win.wtRateOf('보통인부') === 21509, String(win.wtRateOf('보통인부')));
  ok('표 적용 후 직종 칸 노출', A.vis('wJobRow_1'));
  ok('패널에 직종 수 표시', A.$('wtStatus').textContent.includes('4개 직종'));

  win.onJobPick(1, '보통인부');
  ok('직종 선택 → 시급 채워짐', A.num('wRate_1') === 21509, A.$('wRate_1').value);
  ok('근거 표시 노출', A.vis('wJobNote_1') && A.$('wJobNote_1').textContent.includes('172,068원/일'));
  ok('시급 칸은 잠기지 않음', A.$('wRate_1').disabled !== true);

  // 사용자가 직접 수정 — 직종은 유지되고 근거만 사실대로 바뀐다
  A.set('wRate_1', '22,000'); win.renderJobNote(1); win.calculate();
  ok('수동 수정해도 직종 유지', A.$('wJob_1').value === '보통인부');
  ok('근거가 「직접 수정」으로 전환', A.$('wJobNote_1').textContent.includes('직접 수정'));
  ok('되돌리기 링크 노출', A.$('wJobNote_1').innerHTML.includes('revertToTableRate'));
  ok('수정한 시급이 계산에 반영', win.computeAll().results[0].hourlyRate === 22000);

  win.revertToTableRate(1);
  ok('되돌리기 → 표 단가로 복원', A.num('wRate_1') === 21509);
  ok('되돌린 뒤 링크 사라짐', !A.$('wJobNote_1').innerHTML.includes('revertToTableRate'));

  // 빈칸 = 직접입력 — 시급은 건드리지 않는다
  A.set('wRate_1', '25,000');
  A.set('wJob_1', ''); win.renderJobNote(1);
  ok('직종 비우면 시급 불변', A.num('wRate_1') === 25000);
  ok('직종 비우면 근거 숨김', !A.vis('wJobNote_1'));

  // 최저임금 연동
  win.onJobPick(1, '저임금직');
  ok('저임금 직종(60,000÷8=7,500) → 최저임금 경고', A.vis('wWarn_1') && A.num('wRate_1') === 7500);

  // ── 단가 없는 직종은 아예 고를 수 없다 ──
  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['보통인부', 172068], ['연마공', '-'], ['조력공', 0],
  ])), { name: 'u', source: 'upload', fileName: 'u.xlsx' });
  win.onJobPick(1, '보통인부');
  ok('유상 직종 선택 → 시급 채움', A.num('wRate_1') === 21509);

  A.set('wJob_1', ''); win.openJobList(1);
  const jl = A.$('wJobList_1').textContent;
  ok('★ 단가 없는 직종은 목록에 안 뜸', !jl.includes('연마공') && !jl.includes('조력공'), jl);
  ok('단가 있는 직종은 그대로 뜸', jl.includes('보통인부'));
  ok('제외된 직종은 환산 시급 null', win.wtRateOf('연마공') === null);

  // 직접 타이핑해도 표에 없는 이름이라 확정되지 않고 직전 선택으로 되돌아간다
  A.set('wRate_1', '18,000');
  A.$('wJob_1').dataset.lastJob = '보통인부';
  A.set('wJob_1', '연마공'); win.resolveJobInput(1);
  ok('★ 제외된 직종은 타이핑해도 확정 안 됨', A.$('wJob_1').value === '보통인부', A.$('wJob_1').value);
  ok('되돌아가도 수정한 시급은 보존', A.num('wRate_1') === 18000, A.$('wRate_1').value);

  // 환산 기준 변경
  win.onJobPick(1, '보통인부');
  A.set('wtHours', 10); win.onWtHoursChange();
  ok('환산 기준 10h → 표 시급 17,207', win.wtRateOf('보통인부') === Math.round(172068 / 10), String(win.wtRateOf('보통인부')));
  ok('기준 변경이 업무 시급에도 반영', A.num('wRate_1') === 17207, A.$('wRate_1').value);
  ok('근거 줄이 새 기준 표기', A.$('wJobNote_1').textContent.includes('÷ 10h'), A.$('wJobNote_1').textContent);
  // 손으로 고쳐둔 시급은 기준을 바꿔도 유지된다
  A.set('wRate_1', '30,000'); win.renderJobNote(1);
  A.set('wtHours', 8); win.onWtHoursChange();
  ok('수동 수정한 시급은 기준 변경에도 불변', A.num('wRate_1') === 30000, A.$('wRate_1').value);
  win.revertToTableRate(1);
  ok('되돌리면 새 기준(8h) 단가', A.num('wRate_1') === 21509, A.$('wRate_1').value);
}

// ────────────────────────────────────────────────────────────
section('[11] 직종 검색 · 고아 직종 · 일괄 갱신');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [
    ['보통인부', 172068], ['인력운반공', 180193], ['일반기계운전사', 173995],
  ])), { name: '테스트표', source: 'upload', fileName: 't.xlsx' });

  // 검색 필터
  A.set('wJob_1', '인력'); win.openJobList(1);
  let opts = [...A.$('wJobList_1').querySelectorAll('.job-opt')].map(o => o.dataset.job);
  ok('부분 일치 검색 (인력 → 인력운반공)', opts.length === 1 && opts[0] === '인력운반공', JSON.stringify(opts));
  A.set('wJob_1', '운전'); win.openJobList(1);
  opts = [...A.$('wJobList_1').querySelectorAll('.job-opt')].map(o => o.dataset.job);
  ok('중간 일치 검색 (운전 → 일반기계운전사)', opts.length === 1 && opts[0] === '일반기계운전사');
  A.set('wJob_1', '  인 력 '); win.openJobList(1);
  ok('공백 무시 검색', A.$('wJobList_1').querySelectorAll('.job-opt').length === 1);
  A.set('wJob_1', ''); win.openJobList(1);
  ok('빈 검색어 → 전체 표시', A.$('wJobList_1').querySelectorAll('.job-opt').length === 3);
  A.set('wJob_1', '없는직종'); win.openJobList(1);
  ok('일치 없음 안내', A.$('wJobList_1').textContent.includes('일치하는 직종이 없습니다'));
  ok('목록에 일급 함께 표시', (win.openJobList(1), A.set('wJob_1', '보통'), win.openJobList(1),
      A.$('wJobList_1').textContent.includes('172,068원/일')));

  // 목록에 없는 문자열로 blur → 직전 선택 복귀
  win.onJobPick(1, '보통인부');
  A.set('wJob_1', '아무거나');
  win.resolveJobInput(1);          // onJobBlur 가 120ms 뒤 호출하는 확정 로직
  ok('표에 없는 입력은 직전 선택으로 복귀', A.$('wJob_1').value === '보통인부', A.$('wJob_1').value);
  A.set('wJob_1', '인력운반공'); win.resolveJobInput(1);
  ok('표에 있는 입력은 그대로 확정', A.$('wJob_1').value === '인력운반공');

  // 고아 직종 — 저장본에는 있으나 현재 표에 없음
  win.addWork({ name: '옛업무', jobTitle: '사라진직종', hourlyRate: 20000, headcount: 1 });
  const oid = win.workIdsInOrder().slice(-1)[0];
  ok('고아 직종 값 보존', A.$(`wJob_${oid}`).value === '사라진직종');
  ok('고아 직종 안내 표시', A.$(`wJobNote_${oid}`).textContent.includes('단가표에 없는 직종'));

  // 일괄 갱신
  const A2 = load({ xlsx: XLSX });
  A2.win.wtApply(A2.win.wtParseWorkbook(makeWageWorkbook(XLSX, [['보통인부', 160000]])),
                 { name: '구표', source: 'upload', fileName: 'old.xlsx' });
  A2.win.onJobPick(1, '보통인부');
  ok('구표 시급 20,000', A2.num('wRate_1') === 20000, A2.$('wRate_1').value);
  const before = A2.win.wtSnapshotJobs();
  ok('스냅샷이 직종 지정 업무만 담음', before.length === 1 && before[0].job === '보통인부');
  A2.win.wtApply(A2.win.wtParseWorkbook(makeWageWorkbook(XLSX, [['보통인부', 172068]])),
                 { name: '신표', source: 'upload', fileName: 'new.xlsx' });
  A2.win.applyWageTableToAll(before);
  ok('일괄 갱신 → 새 단가 반영', A2.num('wRate_1') === 21509, A2.$('wRate_1').value);
}
// ────────────────────────────────────────────────────────────
section('[12] 단가표 로딩 순서 · 상태 배선');
{
  // 업로드본이 localStorage 에 있으면 그것이 우선 (fetch 는 하네스에서 항상 실패)
  const stored = JSON.stringify({
    name: '저장된표', source: 'upload', fileName: 's.xlsx', hoursPerDay: 8,
    loadedAt: new Date().toISOString(), rows: [{ job: '보관직종', dayWage: 160000, accidentRate: null }],
  });
  const A = load({ xlsx: XLSX, storage: { 'rateCalc_v1_wageTable': stored } });
  ok('저장된 업로드본을 로드', A.win.wtRateOf('보관직종') === 20000, String(A.win.wtRateOf('보관직종')));
  ok('출처 배지 = 직접 올린 표', A.$('wtStatus').textContent.includes('직접 올린 표'));
  ok('기본 표로 복귀 버튼 노출', A.vis('wtResetBtn'));

  // fetch 실패(file://)해도 에러 없이 표 없음으로
  const B = load({ xlsx: XLSX });
  ok('fetch 실패해도 스크립트 에러 없음', B.errors.length === 0, B.errors.slice(0, 2).join(' | '));
  ok('표 없으면 직종 칸 숨김', !B.vis('wJobRow_1'));
  ok('표 없으면 안내 문구', B.$('wtStatus').textContent.includes('단가표가 없습니다'));

  // 상태 라운드트립
  A.win.addWork({ name: '보관업무', jobTitle: '보관직종', hourlyRate: 20000, headcount: 2 });
  const st = A.win.collectState();
  ok('collectState 에 wageTable 포함', st.wageTable?.rows?.length === 1);
  ok('collectState 에 jobTitle 포함', st.works.some(w => w.jobTitle === '보관직종'));

  const C = load({ xlsx: XLSX });
  C.win.applyState(JSON.parse(JSON.stringify(st)));
  ok('저장본 로드로 단가표까지 복원', C.win.wtRateOf('보관직종') === 20000);
  const restored = C.win.workIdsInOrder().find(id => C.$(`wJob_${id}`)?.value === '보관직종');
  ok('저장본 로드로 직종 복원', !!restored, JSON.stringify(C.win.collectState().works.map(w => w.jobTitle)));
  ok('복원 후 직종 칸 노출', C.vis('wJobRow_1'));

  // ── 이 규칙이 생기기 전의 저장본 — 단가 없는 행이 표에 박혀 있다 ──
  const legacy = JSON.parse(JSON.stringify(st));
  legacy.wageTable.rows.push(
    { job: '연마공', dayWage: null, unpriced: true, unit: null, desc: '', basis: '', accidentRate: null },
    { job: '조력공', dayWage: 0,    unpriced: true, unit: null, desc: '', basis: '', accidentRate: null },
  );
  const D = load({ xlsx: XLSX });
  D.win.applyState(legacy);
  ok('★ 옛 저장본의 단가 없는 행도 걸러짐', D.win.wtRows().length === 1, JSON.stringify(D.win.wtRows().map(r => r.job)));
  ok('걸러진 직종은 조회 안 됨', D.win.wtFind('연마공') === null && D.win.wtFind('조력공') === null);
  ok('걸러진 직종은 검색에도 없음', D.win.jobMatches('공').every(m => m.r.job !== '연마공'));
  ok('남은 직종은 정상 동작', D.win.wtRateOf('보관직종') === 20000);
  ok('제외 건수 기록', D.win.collectState().wageTable.excluded === 2, String(D.win.collectState().wageTable.excluded));

  // 브라우저 저장본(localStorage)도 같은 기준으로 걸러진다
  const E = load({ xlsx: XLSX, storage: { rateCalc_v1_wageTable: JSON.stringify(legacy.wageTable) } });
  ok('★ localStorage 저장본도 걸러짐', E.win.wtRows().length === 1, JSON.stringify(E.win.wtRows().map(r => r.job)));

  // resetAll 은 단가표를 남긴다
  A.win.resetAll();
  ok('초기화해도 단가표 유지', A.win.wtRateOf('보관직종') === 20000);
}

// ────────────────────────────────────────────────────────────
section('[13] 클립보드 출력');
{
  const A = load({ xlsx: XLSX });
  const { win } = A;
  let copied = '';
  win.copyText = t => { copied = t; return Promise.resolve(); };

  win.copyResult();
  ok('표 없을 때 단가표 줄 없음', !copied.includes('단가표:'));
  ok('표 없을 때 직종 표기 없음', !copied.includes('| 직종 '));

  win.wtApply(win.wtParseWorkbook(makeWageWorkbook(XLSX, [['보통인부', 172068]])),
              { name: '2026 상반기', source: 'upload', fileName: 't.xlsx' });
  win.onJobPick(1, '보통인부');
  win.copyResult();
  ok('단가표 머리말 기록', copied.includes('단가표: 2026 상반기 (일급÷8h)'), copied.slice(0, 200));
  ok('업무 줄에 직종·일급 기록', copied.includes('직종 보통인부(172,068원/일)'));
}

// ────────────────────────────────────────────────────────────
section('[14] 최종 에러 점검');
{
  const A = load({ xlsx: XLSX });
  A.win.calculate();
  ok('실행 중 스크립트 에러 없음', A.errors.length === 0, A.errors.slice(0, 3).join(' | '));
}

T.done();
