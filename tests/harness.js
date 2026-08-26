// 대시보드를 jsdom 에 띄우고 내부 함수를 그대로 호출할 수 있게 해주는 공용 하네스.
// CDN(Chart.js / XLSX)은 제거해 오프라인 가드가 동작하는지도 함께 확인하고,
// 엑셀이 필요한 검사에서는 win.XLSX 로 직접 주입한다.
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const DASH = path.resolve(__dirname, '..', '요율산정_대시보드_v2.html');

function load({ xlsx = null, storage = {} } = {}) {
  const html = fs.readFileSync(DASH, 'utf8').replace(/<script src="https:[^"]*"><\/script>/g, '');
  const errors = [];
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    pretendToBeVisual: true,
    // localStorage 는 불투명 origin(about:blank)에서 막힌다 — 실제 URL 을 줘야 동작한다
    url: 'https://example.test/요율산정_대시보드_v2.html',
    beforeParse(win) {
      // 스크립트가 돌기 전에 localStorage 를 미리 채워둔다 (단가표 로딩 순서 검사용)
      for (const [k, v] of Object.entries(storage)) {
        try { win.localStorage.setItem(k, v); } catch (e) {}
      }
      if (xlsx) win.XLSX = xlsx;
      win.alert = () => {};
      win.confirm = () => true;
      win.fetch = () => Promise.reject(new Error('file:// 차단 모사'));
      // jsdom 에 없는 브라우저 API — 코드 결함이 아니라 환경의 공백이라 여기서 메운다
      win.Element.prototype.scrollIntoView = function () {};
    },
    virtualConsole: new VirtualConsole()
      .on('jsdomError', e => errors.push('JSDOM: ' + e.message))
      .on('error', e => errors.push('CONSOLE: ' + e)),
  });
  const win = dom.window, doc = win.document;
  win.alert = () => {};
  win.confirm = () => true;

  const $ = id => doc.getElementById(id);
  return {
    dom, win, doc, errors, $,
    set:    (id, v) => { const e = $(id); if (!e) throw new Error('no element ' + id); e.value = String(v); },
    check:  (id, v) => { const e = $(id); if (!e) throw new Error('no element ' + id); e.checked = !!v; },
    txt:    id => ($(id) ? $(id).textContent : '<missing:' + id + '>'),
    // 금액 입력칸은 천단위 콤마로 표시된다 — 숫자 비교 시 콤마를 걷어낸다
    num:    id => Number(String($(id)?.value ?? '').replace(/,/g, '')),
    vis:    id => $(id) && $(id).style.display !== 'none',
  };
}

function runner(title) {
  let pass = 0, fail = 0;
  const failures = [];
  return {
    section: name => console.log('\n' + name),
    ok(name, cond, extra) {
      if (cond) { pass++; console.log('  ✓ ' + name); }
      else { fail++; failures.push(name); console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
    },
    near: (a, b, tol = 1e-6) => Math.abs(a - b) <= tol * Math.max(1, Math.abs(a), Math.abs(b)),
    done() {
      console.log(`\n${'='.repeat(46)}\n  ${title} — 통과 ${pass} · 실패 ${fail}\n${'='.repeat(46)}\n`);
      if (fail) { console.log('실패 목록:'); failures.forEach(f => console.log('  · ' + f)); }
      process.exitCode = fail ? 1 : 0;
      return fail;
    },
  };
}

// 테스트용 단가표 워크북을 그때그때 만들어 쓴다 (고정 픽스처 파일 불필요)
function makeWageWorkbook(XLSX, rows, opt = {}) {
  const head = opt.head || ['직종', '일급(원)'];
  const pre  = opt.pre  || [];                       // 헤더 앞에 끼워 넣을 행들
  const aoa  = [...pre, head, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, opt.sheet || '단가표');
  return wb;
}

module.exports = { load, runner, makeWageWorkbook, DASH };
