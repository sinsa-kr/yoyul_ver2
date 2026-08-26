// exportExcel() 이 직접 만들어내는 제한된 수식 문법(사칙연산·SUM·SUMPRODUCT·IF·비교연산자·
// 배열(범위) 산술·크로스시트 참조)만 다루는 미니 평가기. 범용 엑셀 엔진이 아니라
// "이 셀의 .f 문자열을 실제로 계산하면 .v(캐시값)와 같은가"를 자동 검증해
// 잘못된 셀 주소·참조 실수를 잡아내는 것이 목적이다.
//
// 값 모델: 모든 하위 파서는 스칼라(number|string|boolean) 또는 그 배열을 반환한다.
// 산술·비교 연산자는 피연산자 중 하나라도 배열이면 원소별로 브로드캐스트한다 —
// SUMPRODUCT((A1:A5>0)*(A1:A5-B1:B5)*C1:C5) 같은 표준 배열 수식 패턴을 지원하기 위함.
const XLSX = require('xlsx-js-style');

function makeEvaluator(wb) {
  function cellVal(sheet, ref) {
    const ws = wb.Sheets[sheet];
    if (!ws) throw new Error(`시트 없음: ${sheet}`);
    const cell = ws[ref];
    if (!cell) throw new Error(`셀 없음: ${sheet}!${ref}`);
    return cell.v;   // 텍스트 셀 비교를 위해 강제 형변환하지 않는다
  }

  function broadcast(a, b, fn) {
    if (Array.isArray(a) && Array.isArray(b)) {
      if (a.length !== b.length) throw new Error(`배열 길이 불일치: ${a.length} vs ${b.length}`);
      return a.map((v, k) => fn(v, b[k]));
    }
    if (Array.isArray(a)) return a.map(v => fn(v, b));
    if (Array.isArray(b)) return b.map(v => fn(a, v));
    return fn(a, b);
  }

  return function evalFormula(formula, currentSheet) {
    const s = formula;
    let i = 0;
    const skipWs = () => { while (s[i] === ' ') i++; };

    function parseCellToken() {
      let sheet = currentSheet, m;
      if ((m = /^'([^']+)'!/.exec(s.slice(i)))) { sheet = m[1]; i += m[0].length; }
      else if ((m = /^([A-Za-z0-9가-힣_]+)!/.exec(s.slice(i)))) { sheet = m[1]; i += m[0].length; }
      m = /^\$?[A-Z]+\$?[0-9]+/.exec(s.slice(i));
      if (!m) throw new Error(`셀 참조 파싱 실패 @${i}: ...${s.slice(i, i + 20)}`);
      i += m[0].length;
      return { sheet, ref: m[0].replace(/\$/g, '') };
    }

    function tryRange() {
      const save = i;
      try {
        const a = parseCellToken();
        skipWs();
        if (s[i] !== ':') { i = save; return null; }
        i++; skipWs();
        const b = parseCellToken();
        return { sheet: a.sheet, from: a.ref, to: b.ref };
      } catch (e) { i = save; return null; }
    }

    function rangeValues(rg) {
      const from = XLSX.utils.decode_cell(rg.from), to = XLSX.utils.decode_cell(rg.to);
      const vals = [];
      for (let r = from.r; r <= to.r; r++)
        for (let c = from.c; c <= to.c; c++)
          vals.push(cellVal(rg.sheet, XLSX.utils.encode_cell({ r, c })));
      return vals;
    }

    function parseArgs() {
      const args = [];
      skipWs();
      if (s[i] === ')') return args;
      while (true) {
        skipWs();
        args.push(parseCompare());
        skipWs();
        if (s[i] === ',') { i++; continue; }
        break;
      }
      return args;
    }

    const asArr = v => (Array.isArray(v) ? v : [v]);

    // 우리가 생성하는 SUMIF/COUNTIF 조건은 항상 ">0" 형태의 문자열 리터럴이다
    function matchCriteria(v, crit) {
      const m = /^(>=|<=|<>|>|<|=)?(.+)$/.exec(String(crit));
      const op = m[1] || '=', num = Number(m[2]);
      const val = Number(v) || 0;
      if (op === '>')  return val >  num;
      if (op === '<')  return val <  num;
      if (op === '>=') return val >= num;
      if (op === '<=') return val <= num;
      if (op === '<>') return val !== num;
      return val === num;
    }

    function callFunc(fname, args) {
      if (fname === 'SUM') {
        let sum = 0;
        for (const v of args) for (const x of asArr(v)) sum += Number(x) || 0;
        return sum;
      }
      if (fname === 'SUMPRODUCT') {
        const arrs = args.map(asArr);
        const len = Math.max(...arrs.map(a => a.length));
        let sum = 0;
        for (let k = 0; k < len; k++) {
          let p = 1;
          for (const arr of arrs) p *= Number(arr.length === 1 ? arr[0] : arr[k]) || 0;
          sum += p;
        }
        return sum;
      }
      if (fname === 'ROUND') {
        const p = Math.pow(10, args[1]);
        return Math.round(args[0] * p) / p;
      }
      if (fname === 'COUNTIF') return asArr(args[0]).filter(v => matchCriteria(v, args[1])).length;
      if (fname === 'SUMIF') {
        const rr = asArr(args[0]), ss = asArr(args[2] ?? args[0]);
        let sum = 0;
        rr.forEach((v, k) => { if (matchCriteria(v, args[1])) sum += Number(ss[k]) || 0; });
        return sum;
      }
      throw new Error('알 수 없는 함수: ' + fname);
    }

    function parseIf() {
      i += 3;                                   // 'IF('
      const cond = parseCompare();
      skipWs(); if (s[i] !== ',') throw new Error('IF 콤마 없음'); i++;
      const a = parseCompare();
      skipWs(); if (s[i] !== ',') throw new Error('IF 콤마2 없음'); i++;
      const b = parseCompare();
      skipWs(); if (s[i] !== ')') throw new Error('IF 괄호 안닫힘'); i++;
      return cond ? a : b;
    }

    function parseFactor() {
      skipWs();
      if (s[i] === '-') { i++; return broadcast(parseFactor(), null, v => -v); }
      if (s[i] === '"') {
        const end = s.indexOf('"', i + 1);
        if (end < 0) throw new Error('문자열 리터럴이 닫히지 않음');
        const str = s.slice(i + 1, end);
        i = end + 1;
        return str;
      }
      if (s[i] === '(') { i++; const v = parseCompare(); skipWs(); if (s[i] !== ')') throw new Error('괄호 안맞음'); i++; return v; }
      if (/^IF\(/.test(s.slice(i))) return parseIf();
      let m = /^[A-Z]+(?=\()/.exec(s.slice(i));
      if (m) {
        const fname = m[0]; i += m[0].length; i++;   // '(' 소비
        const args = parseArgs();
        skipWs(); if (s[i] !== ')') throw new Error('함수 괄호 안맞음'); i++;
        return callFunc(fname, args);
      }
      m = /^[0-9]+(\.[0-9]+)?/.exec(s.slice(i));
      if (m) { i += m[0].length; return Number(m[0]); }
      const rg = tryRange();
      if (rg) return rangeValues(rg);
      const tok = parseCellToken();
      return cellVal(tok.sheet, tok.ref);
    }

    function parseMulDiv() {
      let v = parseFactor();
      skipWs();
      while (s[i] === '*' || s[i] === '/') {
        const op = s[i]; i++; skipWs();
        v = broadcast(v, parseFactor(), op === '*' ? (x, y) => x * y : (x, y) => x / y);
        skipWs();
      }
      return v;
    }

    function parseAddSub() {
      let v = parseMulDiv();
      skipWs();
      while (s[i] === '+' || s[i] === '-') {
        const op = s[i]; i++; skipWs();
        v = broadcast(v, parseMulDiv(), op === '+' ? (x, y) => x + y : (x, y) => x - y);
        skipWs();
      }
      return v;
    }

    // 비교 연산자는 IF() 조건뿐 아니라 SUMPRODUCT((범위>0)*...) 처럼 괄호 안 어디서든
    // 나타나므로 최상위 문법으로 둔다 (엑셀에서 비교는 체이닝되지 않는다)
    function parseCompare() {
      let v = parseAddSub();
      skipWs();
      let op = null;
      if (s.slice(i, i + 2) === '>=') { op = '>='; i += 2; }
      else if (s.slice(i, i + 2) === '<=') { op = '<='; i += 2; }
      else if (s[i] === '>') { op = '>'; i++; }
      else if (s[i] === '<') { op = '<'; i++; }
      else if (s[i] === '=') { op = '='; i++; }
      if (op) {
        skipWs();
        const fn = op === '>' ? (x, y) => x > y : op === '<' ? (x, y) => x < y
                 : op === '>=' ? (x, y) => x >= y : op === '<=' ? (x, y) => x <= y
                 : (x, y) => x === y;
        v = broadcast(v, parseAddSub(), fn);
      }
      return v;
    }

    const result = parseCompare();
    skipWs();
    if (i !== s.length) throw new Error('수식 끝까지 못 읽음, 남은부분=' + s.slice(i));
    return result;
  };
}

module.exports = { makeEvaluator };
