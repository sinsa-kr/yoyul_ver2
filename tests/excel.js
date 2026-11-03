// 엑셀 출력 회귀 검증.
// 핵심 안전망: 모든 수식 셀(.f)을 자체 평가기로 다시 계산해 캐시값(.v)과 일치하는지 확인한다.
// (이 검사가 과거 연산자 우선순위 버그를 실제로 잡아냈다.)
const XLSX = require('xlsx-js-style');
const { load, runner, makeWageWorkbook } = require('./harness');
const { makeEvaluator } = require('./formulaEval');

const T = runner('excel');
const { ok, section, near } = T;
const ref = (r, c) => XLSX.utils.encode_cell({ r, c });

// exportExcel() 이 만드는 워크북과 파일명을 가로챈다
function grab(setup) {
  const A = load({ xlsx: XLSX });
  let wb = null, name = '';
  A.win.XLSX = new Proxy(XLSX, {
    get: (t, k) => (k === 'writeFile' ? (w, f) => { wb = w; name = f; } : t[k]),
  });
  if (setup) setup(A);
  A.win.exportExcel();
  return { A, wb, name };
}

// 컬럼 A 에서 라벨로 행 찾기
function findRow(ws, label) {
  const rng = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rng.s.r; r <= rng.e.r; r++) {
    const c = ws[ref(r, 0)];
    if (c && String(c.v).trim() === label) return r;
  }
  return -1;
}
function findContains(ws, sub) {
  const rng = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rng.s.r; r <= rng.e.r; r++) {
    const c = ws[ref(r, 0)];
    if (c && String(c.v).includes(sub)) return r;
  }
  return -1;
}

function verifyFormulas(wb, label) {
  const evalF = makeEvaluator(wb);
  let checked = 0;
  const bad = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const r of Object.keys(ws)) {
      if (r.startsWith('!')) continue;
      const cell = ws[r];
      if (typeof cell.f !== 'string' || !cell.f) continue;
      checked++;
      try {
        const got = evalF(cell.f, name);
        // 절대오차 — 부동소수 잡음만 허용한다. harness.near 는 상대오차라 0 vs 수백만 원도 통과시켰다.
        const okv = (typeof got === 'number' && typeof cell.v === 'number')
          ? Math.abs(got - cell.v) <= 1e-6 + 1e-12 * Math.abs(cell.v)
          : String(got) === String(cell.v);
        if (!okv) bad.push(`${name}!${r} f=${cell.f} → ${got} vs .v=${cell.v}`);
      } catch (e) {
        bad.push(`${name}!${r} f=${cell.f} → 평가실패: ${e.message}`);
      }
    }
  }
  ok(`${label} — 수식 ${checked}개 전부 캐시값과 일치`, bad.length === 0, bad.slice(0, 3).join(' | '));
  return checked;
}

// ────────────────────────────────────────────────────────────
section('[1] 기본 워크북 구조');
{
  const { wb, name } = grab();
  ok('시트 3장 (공통입력·요약·산정내역)',
     wb.SheetNames.length === 3 && wb.SheetNames.includes('요약') && wb.SheetNames.includes('산정내역'),
     JSON.stringify(wb.SheetNames));
  ok('파일명 기본 형식', /^요율산정_\d{8}\.xlsx$/.test(name), name);

  const sm = wb.Sheets['요약'];
  ok('요약 제목', String(sm[ref(0, 0)].v) === '협력업체 요율산정 결과 요약', String(sm[ref(0,0)]?.v));
  const rKpi = findRow(sm, '전체 인원 (명)');
  ok('KPI 라벨 행 존재', rKpi > 0, String(rKpi));
  ok('KPI 3장이 0·2·4열', String(sm[ref(rKpi, 2)]?.v) === '월 청구 총액 (원)'
                        && String(sm[ref(rKpi, 4)]?.v) === '연간 추정 (원)');
  ok('KPI 값 = 인원 3', sm[ref(rKpi + 1, 0)]?.v === 3, String(sm[ref(rKpi+1,0)]?.v));
  ok('요약 열 너비 지정', Array.isArray(sm['!cols']) && sm['!cols'].length >= 4);
  ok('인쇄영역이 요약 끝까지',
     (wb.Workbook?.Names || []).some(n => n.Name === '_xlnm.Print_Area'
        && n.Ref === `'요약'!$A$1:$F$${XLSX.utils.decode_range(sm['!ref']).e.r + 1}`));

  const det = wb.Sheets['산정내역'];
  ok('산정내역 시급 행', findRow(det, '시급') > 0);
  ok('산정내역 교대형태 행', findRow(det, '교대형태') > 0);
  ok('단가표 미사용 시 직종 행 없음', findRow(det, '직종 (단가표)') === -1);

  verifyFormulas(wb, '기본 워크북');
}

// ────────────────────────────────────────────────────────────
section('[2] 화면 값과의 일치');
{
  const { A, wb } = grab();
  const base = A.win.computeAll();
  const sm = wb.Sheets['요약'];
  const rKpi = findRow(sm, '전체 인원 (명)');
  ok('월 청구 총액이 화면과 일치',
     near(sm[ref(rKpi + 1, 2)].v, Math.round(base.grandTot), 3),
     `${sm[ref(rKpi+1,2)]?.v} vs ${Math.round(base.grandTot)}`);
  ok('연간 = 월 × 12', near(sm[ref(rKpi + 1, 4)].v, sm[ref(rKpi + 1, 2)].v * 12, 12));

  const cin = wb.Sheets['공통입력'];
  const rNP = findRow(cin, '국민연금 (%)');
  ok('공통입력 국민연금 4.75', rNP > 0 && near(cin[ref(rNP, 1)].v, 4.75, 0.001), String(cin[ref(rNP,1)]?.v));
}

// ────────────────────────────────────────────────────────────
section('[3] 산정 정보 · 경고 · 파일명');
{
  const { wb, name } = grab(A => {
    A.set('metaVendor', '남양산업');
    A.set('metaFrom', '2026-01-01'); A.set('metaTo', '2026-12-31');
    A.set('metaAuthor', '구매팀 홍길동'); A.set('metaNote', '야간수당 별도');
    A.check('wExempt_1', false);
    A.win.calculate();
  });
  const sm = wb.Sheets['요약'];
  ok('요약 제목에 업체명 병기',
     String(sm[ref(0, 0)].v) === '협력업체 요율산정 결과 요약 — 남양산업', String(sm[ref(0,0)]?.v));
  const rMeta = findContains(sm, '계약기간 2026.01.01 ~ 2026.12.31');
  ok('산정 정보 밴드 존재', rMeta > 0);
  ok('밴드에 작성자·메모', rMeta > 0 && String(sm[ref(rMeta, 0)].v).includes('작성자 구매팀 홍길동')
                                    && String(sm[ref(rMeta, 0)].v).includes('야간수당 별도'));
  const rWarn = findContains(sm, '연장근로 한도 확인');
  ok('한도 초과 경고 행', rWarn > 0 && String(sm[ref(rWarn, 0)].v).includes('경비(120h)'));
  ok('경고에 근거 조문', rWarn > 0 && String(sm[ref(rWarn, 0)].v).includes('§53①'));
  ok('파일명에 업체명', /^요율산정_남양산업_\d{8}\.xlsx$/.test(name), name);
  verifyFormulas(wb, '메타 포함 워크북');
}

// ────────────────────────────────────────────────────────────
section('[4] 단가표 근거 출력');
{
  const { wb } = grab(A => {
    A.win.wtApply(A.win.wtParseWorkbook(makeWageWorkbook(XLSX, [
      ['보통인부', 172068], ['용접공', 282536],
    ])), { name: '2026 상반기 노임단가', source: 'upload', fileName: 't.xlsx' });
    A.win.onJobPick(1, '보통인부');
    A.win.calculate();
  });
  const sm = wb.Sheets['요약'];
  ok('요약 밴드에 단가표 기록', findContains(sm, '단가표 2026 상반기 노임단가 (일급÷8h)') > 0);

  const det = wb.Sheets['산정내역'];
  const rJob = findRow(det, '직종 (단가표)');
  ok('산정내역에 직종 행 추가', rJob > 0, String(rJob));
  ok('직종 행에 직종명·일급', rJob > 0 && String(det[ref(rJob, 2)]?.v) === '보통인부 · 172,068원/일',
     rJob > 0 ? String(det[ref(rJob, 2)]?.v) : '');
  const rRate = findRow(det, '시급');
  ok('시급 행이 환산값 21,509', det[ref(rRate, 2)]?.v === 21509, String(det[ref(rRate,2)]?.v));
  verifyFormulas(wb, '단가표 포함 워크북');
}

// ────────────────────────────────────────────────────────────
section('[3-1] 휴일 8h초과 (§56②2호)');
{
  const { wb } = grab(A => {
    A.set('wRate_1', '1,000'); A.set('wReg_1', 52); A.set('wOT_1', 0);
    A.set('wNt_1', 0); A.set('wWk_1', 12); A.set('wWkOver_1', 4);
    A.win.calculate();
  });
  const det = wb.Sheets['산정내역'];
  const rOv = findContains(det, '8h초과'), rHol = findRow(det, '휴일수당');
  ok('산정내역에 8h초과 행', rOv > 0, String(rOv));
  ok('8h초과 = 4h', det[ref(rOv, 2)]?.v === 4, String(det[ref(rOv, 2)]?.v));
  ok('★ 휴일수당 = 1000×(12×0.5 + 4×0.5)', det[ref(rHol, 2)]?.v === 1000 * 8,
     String(det[ref(rHol, 2)]?.v));
  // 엑셀에서 손으로 8h초과를 부풀려도 MIN 클램프가 막는다
  verifyFormulas(wb, '휴일 8h초과 워크북');
}

// ────────────────────────────────────────────────────────────
section('[4-1] 기준시급 × 적용률 출력');
{
  // 안 쓰면 시트 구조가 종전과 완전히 같아야 한다
  const plain = grab(A => A.win.calculate());
  ok('안 쓰면 기준시급 행 없음', findRow(plain.wb.Sheets['산정내역'], '기준시급') === -1);
  ok('안 쓰면 적용률 행 없음',   findRow(plain.wb.Sheets['산정내역'], '적용률') === -1);

  const { wb } = grab(A => {
    A.win.wtApply(A.win.wtParseWorkbook(makeWageWorkbook(XLSX, [['보통인부', 172068]])),
                  { name: '표', source: 'upload', fileName: 't.xlsx' });
    A.win.setRateMode(1, true);
    A.set('wPct_1', '85');
    A.win.onBaseChange(1);
    A.win.onJobPick(1, '보통인부');
    A.win.addWork({ name: '직접입력업무', shiftType: 'day', hourlyRate: 30000, headcount: 1,
                    hoursRegular: 209, hoursOvertime: 0, hoursNight: 0, holidayHours: 0 });
    A.win.calculate();
  });
  const det = wb.Sheets['산정내역'];
  const rB = findRow(det, '기준시급'), rP = findRow(det, '적용률'), rR = findRow(det, '시급');
  ok('기준시급 행 추가', rB > 0, String(rB));
  ok('적용률 행 추가',   rP > 0, String(rP));
  ok('기준시급 = 21,509', det[ref(rB, 2)]?.v === 21509, String(det[ref(rB, 2)]?.v));
  ok('적용률 = 85',       det[ref(rP, 2)]?.v === 85,     String(det[ref(rP, 2)]?.v));
  ok('시급 = 18,283',     det[ref(rR, 2)]?.v === 18283,  String(det[ref(rR, 2)]?.v));
  // 가중평균 열이라 문자가 섞이면 합계·수식이 깨진다 — 안 쓰는 업무도 숫자여야 한다
  ok('★ 적용률 안 쓰는 업무도 숫자(100)', det[ref(rP, 3)]?.v === 100, String(det[ref(rP, 3)]?.v));
  ok('★ 그 업무 기준시급은 시급과 동일', det[ref(rB, 3)]?.v === 30000, String(det[ref(rB, 3)]?.v));
  ok('적용률 행 서식이 %', String(det[ref(rP, 2)]?.z || '').includes('%'), String(det[ref(rP, 2)]?.z));
  verifyFormulas(wb, '적용률 포함 워크북');
}

// ────────────────────────────────────────────────────────────
section('[5] 업체 비교 시트');
{
  const { A, wb } = grab(A => {
    const unit = A.win.computeAll().results[0].finPer;
    A.set('v_unit_1', Math.round(unit * 1.2));
    A.win.onCmp();
    A.win.calculate();
  });
  ok('비교 시트 추가됨', wb.SheetNames.includes('산정안 비교'), JSON.stringify(wb.SheetNames));
  const cs = wb.Sheets['산정안 비교'];
  ok('비교 시트에 산정일', findRow(cs, '산정일') >= 0);
  verifyFormulas(wb, '비교 포함 워크북');
}

// ────────────────────────────────────────────────────────────
section('[6] 스타일·서식');
{
  const { wb } = grab();
  const sm = wb.Sheets['요약'];
  ok('제목 셀에 스타일', !!sm[ref(0, 0)].s && !!sm[ref(0, 0)].s.font);
  ok('제목 병합', (sm['!merges'] || []).some(m => m.s.r === 0 && m.e.c >= 5));
  ok('행 높이 지정', Array.isArray(sm['!rows']) && sm['!rows'].some(r => r && r.hpt > 0));
  const det = wb.Sheets['산정내역'];
  const rRate = findRow(det, '시급');
  ok('금액 셀에 숫자 서식', !!det[ref(rRate, 2)]?.z, String(det[ref(rRate,2)]?.z));
  ok('산정내역 자동필터', !!det['!autofilter']);
  ok('한글 폰트 지정', JSON.stringify(sm[ref(0, 0)].s).includes('맑은 고딕'));
}

// ────────────────────────────────────────────────────────────
section('[7] 기타고정비 분리 (T1103-02) — 별도 월 정액 · 화면과 일치');
{
  // 원 단위 절대오차 — harness 의 near 는 상대오차라 0 과 수백만 원도 통과시킬 수 있다
  const nearWon = (a, b, tol = 1) => typeof a === 'number' && Math.abs(a - b) <= tol;
  const setup = A => {
    A.set('wName_1', '경비'); A.set('wHead_1', 3);
    A.win.addWork({ name: '미화', shiftType: 'day', hourlyRate: 12000, headcount: 2,
                    hoursRegular: 209, hoursOvertime: 0, hoursNight: 0, holidayHours: 0 });
    A.win.addFixed(); A.win.addFixed();
    const ids = [...A.doc.querySelectorAll('[id^="fxAmt_"]')].map(e => e.id.replace('fxAmt_', ''));
    ids.forEach((i, k) => { A.set(`fxName_${i}`, `고정${k}`); A.set(`fxAmt_${i}`, k ? '120,000' : '330,000'); });
    A.set('mgmtRate', 10);
    A.win.calculate();
    // 비교 입력 — 업무1 간편, 업무2 상세 · 부가세 포함
    const ws = A.win.workIdsInOrder();
    A.check('cmpVatOn', true); A.set('cmpVatRate', 10);
    A.set(`v_unit_${ws[0]}`, '4,000,000'); A.set(`v_head_${ws[0]}`, 3); A.set(`v_vat_${ws[0]}`, '400,000');
    A.win.toggleCmpMode(ws[1], 'detail');
    A.set(`v_head_${ws[1]}`, 2);
    [['base', '2,600,000'], ['ins', '250,000'], ['ret', '210,000'], ['wf', '200,000'],
     ['fx', '100,000'], ['mgmt', '200,000'], ['profit', '100,000'], ['vat', '366,000']]
      .forEach(([k, v]) => A.set(`v_${k}_${ws[1]}`, v));
    A.win.onCmp(); A.win.calculate();
  };
  const { A, wb } = grab(setup);
  const base = A.win.computeAll();
  const cmp  = A.win.compareAll(base);
  ok('전제: 고정비 > 0 · 업무 2건 · 비교 입력', base.fxTotal > 0 && base.results.length === 2 && cmp.has);

  // 산정내역
  const det = wb.Sheets['산정내역'];
  const CT = 2 + base.results.length;
  const rFin = findRow(det, '1인 청구단가'), rH = findRow(det, '시간당 단가');
  ok('순비용 라벨 = ①+②+③', findRow(det, '1인 순비용 (①+②+③)') > 0);
  ok('★ 1인 청구단가 = 화면 finPer', base.results.every((r, i) => near(det[ref(rFin, 2 + i)].v, r.finPer, 1e-9)));
  ok('★ 시간당 단가 = 화면 hBill',  base.results.every((r, i) => near(det[ref(rH, 2 + i)].v, r.hBill, 1e-9)));
  const rFxT = findRow(det, '기타고정비 계'), rFxB = findRow(det, '기타고정비 청구액'), rG = findRow(det, '월 총액');
  ok('고정비 블록 행 존재', rFxT > 0 && rFxB > rFxT && rG > rFxB, [rFxT, rFxB, rG].join(','));
  ok('고정비 계 = fxTotal (합계열)', near(det[ref(rFxT, CT)].v, base.fxTotal, 1e-9));
  ok('★ 고정비 청구액 = fxBill', near(det[ref(rFxB, CT)].v, base.fxBill, 1e-9));
  ok('★ 월 총액 = grandTot',      near(det[ref(rG, CT)].v, base.grandTot, 1e-9));
  ok('고정비 행 업무별 열은 비움', [rFxT, rFxB, rG].every(r => det[ref(r, 2)]?.v === '' && !det[ref(r, 2)]?.f));
  ok('고정비 행은 1인 단가 위에 없음 (산식 순서)', rFxT > findRow(det, '월 소계'));

  // 요약
  const sm = wb.Sheets['요약'];
  const rKpi = findRow(sm, '전체 인원 (명)');
  ok('★ KPI 월 총액 = grandTot', nearWon(sm[ref(rKpi + 1, 2)].v, Math.round(base.grandTot)));
  const rFx = findRow(sm, '기타고정비 (월 정액 · 관리비 포함)'), rTot = findRow(sm, '전체 합계');
  ok('요약 고정비 행 존재', rFx > 0 && rFx < rTot);
  ok('★ 요약 고정비 행 = fxBill', nearWon(sm[ref(rFx, 4)].v, Math.round(base.fxBill)));
  ok('요약 고정비 행 인원·1인단가 "-"', sm[ref(rFx, 2)].v === '-' && sm[ref(rFx, 3)].v === '-');
  ok('★ 요약 전체 합계 = grandTot', nearWon(sm[ref(rTot, 4)].v, Math.round(base.grandTot)));
  const rAggT = findRow(sm, '합계');
  // 비용 구성은 분류마다 원 단위로 반올림한 뒤 더하므로 분류 수 × 0.5원까지 어긋날 수 있다 (7분류 → 3.5원)
  const aggTol = A.win.AGG_LBL.length * 0.5;
  ok('★ 비용 구성 합계 = grandTot', nearWon(sm[ref(rAggT, 3)].v, base.grandTot, aggTol), `${sm[ref(rAggT, 3)].v} vs ${base.grandTot}`);
  const rAggFx = findRow(sm, '기타고정비');
  ok('비용 구성 기타고정비 = fxTotal', nearWon(sm[ref(rAggFx, 3)].v, Math.round(base.fxTotal)));

  // 비교
  const cs = wb.Sheets['산정안 비교'];
  const rows = cmp.rows.map(x => findRow(cs, x.name));
  ok('★ 비교 내 1인단가 = 화면', cmp.rows.every((x, i) => nearWon(cs[ref(rows[i], 4)].v, Math.round(x.mineUnit))));
  ok('★ 비교 내 월소계 = 화면',  cmp.rows.every((x, i) => nearWon(cs[ref(rows[i], 7)].v, Math.round(x.mineTot))));
  const rCT = findRow(cs, '전체 합계');
  ok('★ 비교 내 월 합계 = 화면', nearWon(cs[ref(rCT, 7)].v, Math.round(cmp.tot.mineMonthly)));
  const itemOf = key => cmp.tot.items.find(it => it.key === key);
  [['기타고정비', 'fx'], ['일반관리비+이윤', 'mgmt'], ['부가세', 'vat']].forEach(([lbl, key]) => {
    const r = findRow(cs, lbl);
    ok(`★ 비교 항목 ${lbl} = 화면`, r > 0 && nearWon(cs[ref(r, 1)].v, Math.round(itemOf(key).mine)),
       r > 0 ? `${cs[ref(r, 1)].v} vs ${itemOf(key).mine}` : 'row 없음');
  });

  verifyFormulas(wb, '고정비 분리 워크북');

  // 고정비 0 → 요약 고정비 행 생략, 산정내역 청구액 0
  const z = grab(A => A.win.calculate());
  ok('고정비 0이면 요약 고정비 행 생략', findRow(z.wb.Sheets['요약'], '기타고정비 (월 정액 · 관리비 포함)') === -1);
  verifyFormulas(z.wb, '고정비 0 워크북');
}

T.done();
