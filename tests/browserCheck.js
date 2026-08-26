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
    ok('미공표 직종도 목록에 남음', meta.unpriced === 0 || meta.n > meta.unpriced,
       `미공표 ${meta.unpriced}/${meta.n}`);

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
