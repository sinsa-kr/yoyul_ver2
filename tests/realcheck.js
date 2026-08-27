const XLSX = require('xlsx-js-style');
const fs = require('fs');
const path = require('path');
const { load } = require('./harness');
const A = load({ xlsx: XLSX });
// 실행 위치와 무관하게 리포 루트의 표를 읽는다
const wb = XLSX.read(fs.readFileSync(path.resolve(__dirname, '..', '단가표.xlsx')), { type: 'buffer' });
const p = A.win.wtParseWorkbook(wb);
console.log('파싱 직종 수 :', p.rows.length);
console.log('제외(단가없음):', p.unpriced, '· 동명이직 구분', p.dup);
console.log('기준         :', p.basis);
const units = p.rows.reduce((m, r) => (m[r.unit || 'null'] = (m[r.unit || 'null'] || 0) + 1, m), {});
console.log('단위 분포    :', JSON.stringify(units));
console.log('해설 있는 행 :', p.rows.filter(r => r.desc).length);
console.log('동명이직 예시:', p.rows.filter(r => /^(작업반장|기술사)\s*\(/.test(r.job)).slice(0,4).map(r=>`${r.job}=${r.dayWage}`).join(' / '));
A.win.wtApply(p, { name: '시중노임단가', source: 'repo', fileName: '단가표.xlsx' });
const low = p.rows.filter(r => !r.unpriced && A.win.wtHourlyOf(r) < 10320);
console.log('최저임금 미만:', low.length, low.slice(0,3).map(r=>`${r.job} ${r.dayWage}→${A.win.wtHourlyOf(r)}`).join(' / '));
const hi = p.rows.filter(r=>!r.unpriced).sort((a,b)=>b.dayWage-a.dayWage)[0];
console.log('최고 단가    :', hi.job, hi.dayWage, '→ 시급', A.win.wtHourlyOf(hi));
console.log('중복 직종명  :', (() => {
  const seen = new Set(), dup = [];
  for (const r of p.rows) { if (seen.has(r.job)) dup.push(r.job); seen.add(r.job); }
  return dup.length ? dup.join(',') : '없음';
})());
