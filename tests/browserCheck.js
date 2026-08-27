// 실제 브라우저에서만 확인 가능한 것들:
//   · https 로 서빙했을 때 리포 단가표 자동 로드 (jsdom 은 fetch 를 흉내낼 뿐)
//   · file:// 에서 fetch 가 막혔을 때의 폴백
//   · 드래그앤드롭 (HTML5 DnD 는 jsdom 이 effectAllowed 를 강제하지 않아 회귀를 놓친다)
//   · 3개 테마 렌더
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..');
const DASH = '요율산정_대시보드_v2.html';
const MIME = { '.html': 'text/html; charset=utf-8', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.json': 'application/json' };

let pass = 0, fail = 0;
const failures = [];
const ok = (l, c, x) => { c ? (pass++, console.log('  ✓ ' + l)) : (fail++, failures.push(l), console.log('  ✗ ' + l + (x ? ' → ' + x : ''))); };

function serve() {
  return new Promise(res => {
    const srv = http.createServer((req, r) => {
      const rel = decodeURIComponent(req.url.split('?')[0]).replace(/^\//, '') || DASH;
      const f = path.join(ROOT, rel);
      if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.writeHead(404); r.end(); return; }
      r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' });
      fs.createReadStream(f).pipe(r);
    });
    srv.listen(0, '127.0.0.1', () => res({ srv, port: srv.address().port }));
  });
}

async function drag(page, a, b, { ctrl = false } = {}) {
  const ax = a.x + a.width / 2, ay = a.y + a.height / 2;
  const bx = b.x + b.width / 2, by = b.y + b.height / 2;
  if (ctrl) await page.keyboard.down('Control');
  await page.mouse.move(ax, ay); await page.waitForTimeout(60);
  await page.mouse.down(); await page.waitForTimeout(60);
  for (let i = 1; i <= 20; i++) {
    await page.mouse.move(ax + (bx - ax) * i / 20, ay + (by - ay) * i / 20);
    await page.waitForTimeout(18);
  }
  await page.waitForTimeout(90);
  await page.mouse.up(); await page.waitForTimeout(220);
  if (ctrl) await page.keyboard.up('Control');
}

(async () => {
  const { srv, port } = await serve();
  const browser = await chromium.launch();
  const errs = [];

  // ── https(로컬 서버) — 리포 단가표 자동 로드 ──
  console.log('\n===== 서버 서빙 (GitHub Pages 모사) =====');
  {
    const p = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    p.on('pageerror', e => errs.push('SERVE: ' + e.message));
    await p.goto(`http://127.0.0.1:${port}/${encodeURIComponent(DASH)}`);
    await p.waitForTimeout(1500);
    // 표 내용은 사용자가 반기마다 갈아끼우므로, 특정 직종명을 박아두지 않고
    // 로드된 표에서 조건에 맞는 표본을 골라 검증한다.
    const meta = await p.evaluate(() => {
      const rows = wtRows();
      const priced = rows.find(r => !r.unpriced);
      // 해설에만 있고 이름에는 없는 단어 하나를 찾아 해설검색 표본으로 쓴다
      let descProbe = null;
      for (const r of rows) {
        const w = (r.desc || '').match(/[가-힣]{3,6}/g) || [];
        const cand = w.find(x => !r.job.includes(x) && rows.filter(o => o.job.includes(x)).length === 0);
        if (cand) { descProbe = { job: r.job, term: cand }; break; }
      }
      return { n: rows.length, unpriced: rows.filter(r => r.unpriced).length,
               noRate: rows.filter(r => !(r.dayWage > 0)).length,
               excluded: wageTable?.excluded || 0,
               priced: priced && { job: priced.job, rate: wtRateOf(priced.job) }, descProbe };
    });
    ok('업로드 없이 리포 단가표 자동 로드', meta.n > 0, `${meta.n}개 직종`);
    ok('출처 배지 = 기본 표', (await p.locator('#wtStatus').innerText()).includes('기본 표'));
    ok('직종 칸 노출', await p.locator('#wJobRow_1').isVisible());
    ok('환산 시급이 계산됨', meta.priced && meta.priced.rate > 0,
       meta.priced ? `${meta.priced.job} → ${meta.priced.rate}` : '유상 직종 없음');

    // 검색 → 선택 (표본 직종명의 앞 두 글자로 검색)
    const probe = meta.priced.job.slice(0, 2);
    await p.click('#wJob_1');
    await p.fill('#wJob_1', probe);
    await p.waitForTimeout(250);
    const opts = await p.locator('#wJobList_1 .job-opt').allTextContents();
    ok('검색 목록이 뜨고 걸러짐', opts.length > 0 && opts.some(o => o.includes(probe)),
       `"${probe}" → ${opts.length}건`);
    ok('기준(적용월) 패널 표시', /\d{4}년 \d{2}월 기준/.test(await p.locator('#wtStatus').innerText()),
       await p.locator('#wtStatus').innerText());
    await p.locator(`#wJobList_1 .job-opt[data-job="${meta.priced.job.replace(/"/g, '&quot;')}"]`).first().click();
    await p.waitForTimeout(250);
    ok('선택 → 시급 자동 입력',
       (await p.inputValue('#wRate_1')).replace(/,/g, '') === String(meta.priced.rate),
       await p.inputValue('#wRate_1'));
    ok('근거 줄 표시', (await p.locator('#wJobNote_1').innerText()).includes('원/'));

    // 해설 검색 + 하단 미리보기
    if (meta.descProbe) {
      await p.fill('#wJob_1', meta.descProbe.term); await p.waitForTimeout(250);
      const byDesc = await p.locator('#wJobList_1 .job-opt').allTextContents();
      ok(`해설 단어로 검색 ("${meta.descProbe.term}" → 이름엔 없는 직종)`,
         byDesc.length > 0, `${byDesc.length}건`);
      ok('해설 일치 배지 표시', await p.locator('#wJobList_1 .hitlbl').count() > 0);
    } else {
      ok('해설 검색 표본을 찾지 못함 (표에 해설이 없음)', false, '표 확인 필요');
    }
    ok('미리보기가 보임', await p.locator('#wJobPrev_1').isVisible());
    // 실제 표에 단가 없는 행이 섞여 있어도 로드된 목록에는 한 건도 없어야 한다
    ok('★ 단가 없는 직종은 목록에 없음', meta.noRate === 0 && meta.unpriced === 0,
       `단가없음 ${meta.noRate} · unpriced ${meta.unpriced} / ${meta.n}`);
    ok('제외 건수를 패널에 밝힘',
       !meta.excluded || (await p.locator('#wtStatus').innerText()).includes(`${meta.excluded}건 제외`),
       `excluded=${meta.excluded} · ${await p.locator('#wtStatus').innerText()}`);

    await p.fill('#wJob_1', ''); await p.waitForTimeout(250);
    const prevBefore = await p.locator('#wJobPrev_1').innerText();
    const listH1 = (await p.locator('#wJobList_1').boundingBox()).height;
    await p.locator('#wJobList_1 .job-opt').nth(3).hover(); await p.waitForTimeout(220);
    const prevAfter = await p.locator('#wJobPrev_1').innerText();
    ok('마우스로 짚으면 미리보기 갱신', prevBefore !== prevAfter, prevAfter.slice(0, 40));
    const listH2 = (await p.locator('#wJobList_1').boundingBox()).height;
    ok('짚어도 목록 높이 불변 (출렁임 없음)', Math.abs(listH1 - listH2) < 1, `${listH1} → ${listH2}`);
    ok('미리보기에 해설 전문(잘림 없음)',
       !prevAfter.includes('…') && prevAfter.length > 20, prevAfter.slice(0, 60));

    // 상세칸이 목록의 또 한 줄처럼 보이면 안 된다 — 배경·경계선이 실제로 다른지 잰다.
    // 눈으로만 확인하면 토큰이 바뀔 때 조용히 되돌아간다.
    const sep = await p.evaluate(() => {
      const rgb = s => (s.match(/\d+/g) || [0, 0, 0]).slice(0, 3).map(Number);
      const lum = c => { const [r, g, b] = rgb(c).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; });
                         return .2126 * r + .7152 * g + .0722 * b; };
      const prev = getComputedStyle(document.querySelector('.job-prev'));
      const opt  = getComputedStyle(document.querySelector('.job-opt'));
      // 목록 행은 대개 투명이라 팝업 배경이 실제 바탕이 된다
      const listBg = opt.backgroundColor === 'rgba(0, 0, 0, 0)'
        ? getComputedStyle(document.querySelector('.job-pop')).backgroundColor : opt.backgroundColor;
      const L1 = lum(prev.backgroundColor), L2 = lum(listBg);
      return { ratio: (Math.max(L1, L2) + .05) / (Math.min(L1, L2) + .05),
               border: parseFloat(prev.borderTopWidth), shadow: prev.boxShadow !== 'none',
               label: !!document.querySelector('.job-prev .pv') };
    });
    ok('★ 상세칸 배경이 목록과 구분됨', sep.ratio >= 1.06, `대비 ${sep.ratio.toFixed(3)}:1`);
    ok('★ 상세칸 경계선이 2px 이상', sep.border >= 2, `${sep.border}px`);
    ok('상세칸에 그림자로 단차 표현', sep.shadow);
    ok('상세칸에 「상세」 표식', sep.label);

    // 키보드 조작
    await p.fill('#wJob_1', '');
    await p.click('#wJob_1');
    await p.waitForTimeout(150);
    await p.keyboard.press('ArrowDown');
    await p.keyboard.press('ArrowDown');
    await p.keyboard.press('Enter');
    await p.waitForTimeout(250);
    ok('↑↓ + Enter 로 선택', (await p.inputValue('#wJob_1')).length > 0, await p.inputValue('#wJob_1'));

    // 시급 수동 수정 후 되돌리기
    await p.fill('#wRate_1', '30000');
    await p.waitForTimeout(250);
    ok('시급 수동 수정 가능', (await p.inputValue('#wRate_1')).replace(/,/g, '') === '30000');
    ok('근거가 직접 수정으로 전환', (await p.locator('#wJobNote_1').innerText()).includes('직접 수정'));
    await p.locator('#wJobNote_1 button').click();
    await p.waitForTimeout(250);
    ok('되돌리기 동작', (await p.locator('#wJobNote_1').innerText()).includes('📋'));

    // ── 기준시급 × 적용률 — 좁은 좌측 컬럼(380px)에 실제로 들어가는지 ──
    ok('꺼진 상태가 기본', !(await p.locator('#wBaseRow_1').isVisible()));
    ok('켜기 링크 노출', await p.locator('#wPctBtn_1').isVisible());
    const rateBefore = await p.inputValue('#wRate_1');
    await p.locator('#wPctBtn_1').click();
    await p.waitForTimeout(250);
    ok('★ 켜도 시급 불변', (await p.inputValue('#wRate_1')) === rateBefore,
       `${rateBefore} → ${await p.inputValue('#wRate_1')}`);
    ok('입력행 노출 · 켜기 링크 숨김',
       (await p.locator('#wBaseRow_1').isVisible()) && !(await p.locator('#wPctBtn_1').isVisible()));

    await p.fill('#wPct_1', '85');
    await p.waitForTimeout(250);
    const fit = await p.evaluate(() => {
      const row = document.querySelector('#wBaseRow_1');
      const lab = document.querySelector('#wPctBtn_1').parentElement;
      return { rowOver: row.scrollWidth > row.clientWidth + 1,
               labOver: lab.scrollWidth > lab.clientWidth + 1,
               page: document.documentElement.scrollWidth > document.documentElement.clientWidth,
               base: +document.querySelector('#wBase_1').value.replace(/,/g, ''),
               rate: +document.querySelector('#wRate_1').value.replace(/,/g, '') };
    });
    ok('★ 입력행이 카드 폭을 넘지 않음', !fit.rowOver);
    ok('★ 시급 라벨이 줄바꿈되지 않음', !fit.labOver);
    ok('페이지 가로 스크롤 없음', !fit.page);
    ok('★ 시급 = 기준 × 85%', fit.rate === Math.round(fit.base * 0.85), `${fit.base} → ${fit.rate}`);
    ok('근거 줄에 계산식', (await p.locator('#wJobNote_1').innerText()).includes('× 85% ='),
       await p.locator('#wJobNote_1').innerText());

    // 다크 테마에서도 새 입력행이 배경에 묻히지 않아야 한다
    await p.evaluate(() => setTheme('dark'));
    await p.waitForTimeout(300);
    const dk = await p.evaluate(() => {
      const rgb = s => (s.match(/\d+/g) || [0, 0, 0]).slice(0, 3).map(Number);
      const lum = c => { const [r, g, b] = rgb(c).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; });
                         return .2126 * r + .7152 * g + .0722 * b; };
      const row = getComputedStyle(document.querySelector('#wBaseRow_1')).backgroundColor;
      const card = getComputedStyle(document.querySelector('#work_1')).backgroundColor;
      const L1 = lum(row), L2 = lum(card);
      return (Math.max(L1, L2) + .05) / (Math.min(L1, L2) + .05);
    });
    ok('다크에서 입력행이 카드와 구분됨', dk >= 1.05, `대비 ${dk.toFixed(3)}:1`);
    await p.evaluate(() => setTheme('warm'));

    await p.close();
  }

  // ── file:// — fetch 차단 폴백 ──
  console.log('\n===== file:// (로컬 파일) =====');
  {
    const p = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    p.on('pageerror', e => errs.push('FILE: ' + e.message));
    await p.goto('file:///' + encodeURI(path.join(ROOT, DASH).split(path.sep).join('/')));
    await p.waitForTimeout(1400);
    ok('fetch 차단돼도 페이지 정상', await p.locator('#work_1').isVisible());
    ok('표 없음 → 직종 칸 숨김', !(await p.locator('#wJobRow_1').isVisible()));
    ok('안내 문구 표시', (await p.locator('#wtStatus').innerText()).includes('단가표가 없습니다'));
    ok('업로드 버튼은 노출', await p.locator('#wtImport').count() === 1);
    await p.close();
  }

  // ── 드래그앤드롭 · 테마 ──
  console.log('\n===== 드래그앤드롭 · 테마 =====');
  {
    const p = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    p.on('pageerror', e => errs.push('DND: ' + e.message));
    await p.goto(`http://127.0.0.1:${port}/${encodeURIComponent(DASH)}`);
    await p.waitForTimeout(1400);

    await p.click('#work_1 button[onclick^="openShiftBuilder"]');
    await p.waitForTimeout(400);
    await p.click('button[onclick="sbApplyPreset(\'day\')"]');
    await p.waitForTimeout(300);

    const cell = i => p.locator(`#sbGrid [onclick="sbToggleCell(${i})"]`);
    const chip = p.locator('button[onclick="sbApplyBlock(\'night\')"]');
    await drag(p, await chip.boundingBox(), await cell(6).boundingBox());
    ok('칩 → 빈 칸 드롭', await p.evaluate(() => JSON.stringify(sbDraft.days[6])) === JSON.stringify({ start: '00:00', end: '08:00' }),
       await p.evaluate(() => JSON.stringify(sbDraft.days[6])));

    await drag(p, await cell(1).boundingBox(), await cell(0).boundingBox());
    ok('칸 → 칸 이동: 대상에 값', await p.evaluate(() => !!sbDraft.days[0]));
    ok('칸 → 칸 이동: 원본 비움', await p.evaluate(() => sbDraft.days[1] === null));

    await drag(p, await cell(2).boundingBox(), await cell(1).boundingBox(), { ctrl: true });
    ok('Ctrl 복사: 대상에 값', await p.evaluate(() => !!sbDraft.days[1]));
    ok('Ctrl 복사: 원본 유지', await p.evaluate(() => !!sbDraft.days[2]));
    await p.click('#shiftBuilderModal button[onclick="closeShiftBuilder()"]');
    await p.waitForTimeout(250);

    for (const t of ['light', 'dark', 'warm']) {
      await p.click('#tw_' + t); await p.waitForTimeout(400);
      const themed = await p.evaluate(() => document.documentElement.dataset.theme);
      const chartOk = await p.evaluate(() => typeof Chart !== 'undefined' && Chart.defaults.color === pal().tick);
      ok(`${t} 테마 적용 + 차트 색 동기화`, themed === t && chartOk, `${themed}/${chartOk}`);
    }
    // 다크에서 직종 목록이 테마를 따르는지
    await p.click('#tw_dark'); await p.waitForTimeout(400);
    await p.click('#wJob_1'); await p.waitForTimeout(250);
    const listBg = await p.locator('#wJobPop_1').evaluate(e => getComputedStyle(e).backgroundColor);
    ok('다크에서 직종 목록 배경도 어둡게', listBg === 'rgb(29, 35, 42)', listBg);
    const anyJob = await p.evaluate(() => wtRows()[0].job.slice(0, 2));
    await p.fill('#wJob_1', anyJob); await p.waitForTimeout(250);
    const markBg = await p.locator('#wJobList_1 mark').first().evaluate(e => getComputedStyle(e).backgroundColor);
    ok('다크에서 하이라이트도 테마색', markBg === 'rgb(51, 43, 28)', markBg);
    await p.close();
  }

  ok('페이지 에러 없음', errs.length === 0, errs.slice(0, 3).join(' | '));
  await browser.close();
  srv.close();
  console.log(`\n${'='.repeat(46)}\n  browser — 통과 ${pass} · 실패 ${fail}\n${'='.repeat(46)}\n`);
  if (fail) { console.log('실패 목록:'); failures.forEach(f => console.log('  · ' + f)); }
  process.exit(fail ? 1 : 0);
})();
