---
name: dashboard
description: 협력업체 요율산정 대시보드(요율산정_대시보드.html) 스타일로 단일 HTML 대시보드를 만들거나 수정할 때 사용. "대시보드 만들어줘", "요율산정 스타일로", "dashboard 추가/수정/업데이트" 같은 요청에 활성화. 기존 대시보드 파일 수정 작업에도 자동 적용.
version: 1.0.0
---

# 요율산정 대시보드 스킬

이 스킬은 `요율산정_대시보드.html` 에서 검증된 모든 패턴을 담고 있다.
새 대시보드를 만들거나 기존 파일을 수정할 때 반드시 이 패턴을 따른다.

---

## 기술 스택 (고정)

```html
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
```

- 빌드 없음, 단일 `.html` 파일로 완결
- 폰트: `font-family:'Malgun Gothic','맑은 고딕',sans-serif`
- 서버 불필요, 로컬 브라우저에서 바로 실행

---

## 색상 팔레트 (따뜻한 베이지 계열)

| 용도 | HEX |
|------|-----|
| 배경(body) | `#f5ede0` |
| 패널 배경 | `#fdf9f3` |
| 헤더 배경 | `#7a5c3e` |
| 기본 텍스트(진한) | `#4a3828` |
| 기본 텍스트(중간) | `#7a6a58` |
| 테두리 | `#ddd0bc` |
| 강조(amber) | `#a07040` |
| 강조(green) | `#4d7a5a` |
| 강조(blue) | `#5a7aaa` |
| 입력 포커스 테두리 | `#b07d4e` |
| 섹션 구분선 | `#f0e8d8` |

업무 색상 순환 배열 (8색):
```js
const WK_COLORS = ['#7a9fcc','#70b898','#d4a060','#d07878','#9880cc','#d87aaa','#5ab8cc','#90c070'];
```

---

## CSS 공통 클래스 (항상 `<style>` 태그에 포함)

```css
body { font-family:'Malgun Gothic','맑은 고딕',sans-serif; background:#f5ede0; }
.panel { background:#fdf9f3; border-radius:10px; box-shadow:0 1px 4px rgba(100,70,30,.09); padding:14px; }
.sec-hdr { font-size:12px; font-weight:700; color:#4a3828; margin-bottom:8px; }
.inp { padding:5px 8px; border:1px solid #ddd0bc; border-radius:6px; font-size:12px; outline:none; width:100%; box-sizing:border-box; }
.inp:focus { border-color:#b07d4e; }
.inp-r { padding:3px 6px; border:1px solid #ddd0bc; border-radius:4px; font-size:12px; text-align:right; outline:none; width:70px; }
.inp-r:focus { border-color:#b07d4e; }
.rate-row { display:flex; align-items:center; justify-content:space-between; padding:3px 0; }
.res-row { display:flex; justify-content:space-between; align-items:center; padding:4px 0; font-size:13px; border-bottom:1px solid #f0e8d8; }
.rstbtn { color:#a09080; background:none; border:none; cursor:pointer; font-size:12px; padding:1px 3px; border-radius:3px; }
.rstbtn:hover { color:#4a3828; background:#f0e8d8; }
.sum-card { background:#fdf9f3; border-radius:10px; box-shadow:0 1px 4px rgba(100,70,30,.09); padding:14px; text-align:center; }
```

---

## 레이아웃 구조

```html
<!-- 헤더 -->
<div style="background:#7a5c3e;color:white;padding:12px 20px;...">
  제목 + 액션 버튼들 (엑셀, 복사, 저장/불러오기, 초기화)
</div>

<!-- 경고 배너 (조건부 표시) -->
<div id="alertBanner" style="display:none;..."></div>

<!-- 3컬럼 메인 그리드 -->
<div style="display:grid;grid-template-columns:380px 1fr 280px;gap:14px;padding:14px;align-items:start;">
  <!-- 좌측: 입력 패널들 (스크롤 가능) -->
  <div style="max-height:calc(100vh - 76px);overflow-y:auto;">...</div>
  <!-- 중앙: 요약카드 + 테이블 + 상세 아코디언 -->
  <div>...</div>
  <!-- 우측: 차트 + 범례 -->
  <div>...</div>
</div>
```

좌측 패널 너비 `380px`, 우측 `280px`는 고정값. 중앙이 나머지를 채움.

---

## JS 유틸 함수 (항상 포함)

```js
const $   = id => document.getElementById(id);
const gv  = id => parseFloat($(id)?.value) || 0;
const gc  = id => $(id)?.checked || false;
const fmt  = n => Math.round(n).toLocaleString('ko-KR') + '원';
const fmtn = n => Math.round(n).toLocaleString('ko-KR');

// 입력값을 기본값으로 복원하고 재계산
function rv(id, def) { $(id).value = def; calculate(); }

// 섹션 접기/펼치기
function toggleSec(secId, btnId) {
  const s = $(secId), b = $(btnId);
  if (s.style.display === 'none') { s.style.display='block'; b.textContent='접기'; }
  else { s.style.display='none'; b.textContent='펼치기'; }
}
```

---

## 동적 DOM 패턴 (항목 추가/삭제)

항목마다 전역 카운터로 고유 ID를 부여한다. 삭제해도 카운터는 증가만 함.

```js
let itemCnt = 0;

function addItem(p = {}) {
  const id = ++itemCnt;
  const div = document.createElement('div');
  div.id = `item_${id}`;
  div.innerHTML = `
    <input id="iName_${id}" oninput="calculate()">
    <button onclick="removeItem(${id})">✕</button>
  `;
  $('itemList').appendChild(div);
  calculate();
}

function removeItem(id) {
  $(`item_${id}`)?.remove();
  calculate();
}
```

입력값 수집:
```js
function collectItems() {
  const items = [];
  for (let i = 1; i <= itemCnt; i++) {
    if (!$(`item_${i}`)) continue; // 삭제된 항목 건너뜀
    items.push({ id: i, name: $(`iName_${i}`)?.value || '' });
  }
  return items;
}
```

---

## 계산 엔진 파이프라인

```
collectItems() ──▶ calcOne(item, rates) ──▶ calculate() ──▶ render*()
```

- 모든 입력 필드에 `oninput="calculate()"` 달기
- `calculate()`는 전체를 다시 계산하고 모든 렌더 함수를 호출
- 렌더 함수는 DOM 직접 조작 (innerHTML 교체 방식)

```js
function calculate() {
  const items = collectItems();
  const rates = { /* gv() 로 읽기 */ };
  const results = items.map(item => calcOne(item, rates));
  renderTable(results);
  renderCharts(results);
}
```

---

## 요율 입력 행 패턴 (↺ 복원 버튼 포함)

```html
<div class="rate-row">
  <label style="font-size:12px;color:#4a3828;width:80px;">항목명</label>
  <div style="display:flex;align-items:center;gap:3px;">
    <input type="number" id="rXX" value="4.5" step="0.01" oninput="calculate()" class="inp-r">
    <span style="font-size:11px;color:#a09080;">%</span>
    <button onclick="rv('rXX',4.5)" class="rstbtn" title="기본값 복원">↺</button>
  </div>
</div>
```

---

## 아코디언 상세 패널 패턴

```html
<!-- 헤더 (클릭으로 토글) -->
<div onclick="toggleDetail(${id})" style="cursor:pointer;...">
  <span>항목명</span>
  <button id="detail_btn_${id}">▼ 상세 보기</button>
</div>
<!-- 내용 (기본 숨김) -->
<div id="detail_body_${id}" style="display:none;">
  상세 내용
</div>
```

```js
function toggleDetail(id) {
  const el = $('detail_body_' + id);
  const btn = $('detail_btn_' + id);
  if (!el) return;
  if (el.style.display === 'none') { el.style.display='block'; btn.textContent='▲ 접기'; }
  else { el.style.display='none'; btn.textContent='▼ 상세 보기'; }
}
```

---

## Chart.js 싱글톤 패턴

차트를 처음 만든 후에는 `update()`만 호출. 재생성하지 않는다.

```js
let myChart = null;

function renderCharts(data) {
  const labels = data.map(d => d.name);
  const values = data.map(d => d.value);

  if (myChart) {
    myChart.data.labels = labels;
    myChart.data.datasets[0].data = values;
    myChart.update();
  } else {
    myChart = new Chart($('myChart').getContext('2d'), {
      type: 'bar',
      data: { labels, datasets: [{ data: values, backgroundColor: WK_COLORS }] },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        scales: { y: { ticks: { callback: v => fmtn(v/10000)+'만', font:{size:10} } } }
      }
    });
  }
}
```

도넛 차트:
```js
type: 'doughnut',
options: {
  cutout: '58%',
  plugins: {
    legend: { display: false },
    tooltip: { callbacks: { label: ctx => {
      const t = ctx.dataset.data.reduce((a,b)=>a+b,0);
      return `${ctx.label}: ${fmtn(ctx.raw)}원 (${t>0?(ctx.raw/t*100).toFixed(1):0}%)`;
    }}}
  }
}
```

---

## 모달 패턴 (저장/불러오기)

```html
<!-- 배경 클릭으로 닫기 -->
<div id="myModal" onclick="if(event.target===this)closeModal()"
  style="display:none;position:fixed;inset:0;background:rgba(60,40,20,0.55);z-index:2000;align-items:center;justify-content:center;">
  <div style="background:#fdf9f3;border-radius:14px;width:520px;...">
    내용
  </div>
</div>
```

```js
function openModal()  { $('myModal').style.display = 'flex'; }
function closeModal() { $('myModal').style.display = 'none'; }
```

---

## localStorage 저장/불러오기

```js
const SAVE_KEY = 'myApp_v1_saves';

function getSaves() {
  try { return JSON.parse(localStorage.getItem(SAVE_KEY) || '[]'); }
  catch(e) { return []; }
}
function setSaves(arr) { localStorage.setItem(SAVE_KEY, JSON.stringify(arr)); }

function doSave(name) {
  const saves = getSaves();
  saves.unshift({ id: Date.now(), name, savedAt: new Date().toISOString(), data: collectState() });
  if (saves.length > 30) saves.pop();
  setSaves(saves);
}
```

---

## 엑셀 내보내기 (xlsx-js-style)

핵심 패턴: AOA(배열of배열) + 메타 배열로 스타일 분리 적용

```js
function exportExcel() {
  if (typeof XLSX === 'undefined') { alert('엑셀 라이브러리 로딩 중. 잠시 후 재시도.'); return; }

  const aoa = [];    // 셀 데이터
  const meta = [];   // { row, col, s(스타일), z(숫자포맷) }
  const merges = []; // 병합 범위
  let r = 0;

  function st(row, col, s, z) { meta.push({ row, col, s, z }); }
  function mg(r1,c1,r2,c2)   { merges.push({ s:{r:r1,c:c1}, e:{r:r2,c:c2} }); }

  // 데이터 추가
  aoa.push(['제목']); mg(r,0,r,5); r++;
  aoa.push(['항목', '값']); r++;

  // 시트 생성
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const m of meta) {
    const ref = XLSX.utils.encode_cell({ r:m.row, c:m.col });
    if (!ws[ref]) ws[ref] = { t:'z', v:'' };
    ws[ref].s = m.s;
    if (m.z) ws[ref].z = m.z;
  }
  ws['!merges'] = merges;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '시트명');
  const today = new Date().toISOString().slice(0,10).replace(/-/g,'');
  XLSX.writeFile(wb, `결과_${today}.xlsx`);
}
```

엑셀 스타일 팔레트 (재사용):
```js
const bt = { style:'thin',   color:{ rgb:'B0CCD8' } };
const bm = { style:'medium', color:{ rgb:'7A9AB8' } };
const font = (sz, bold, rgb) => ({ sz, bold:!!bold, color:rgb?{rgb}:undefined });
const fill = rgb => ({ patternType:'solid', fgColor:{ rgb } });
const al   = (h, wrap) => ({ horizontal:h, vertical:'center', wrapText:!!wrap });
```

---

## 클립보드 복사

```js
function copyResult() {
  const lines = ['=== 결과 ===', `날짜: ${new Date().toLocaleDateString('ko-KR')}`, ''];
  // lines에 내용 추가
  navigator.clipboard.writeText(lines.join('\n'))
    .then(() => alert('클립보드에 복사되었습니다.'))
    .catch(() => alert('복사 실패. 브라우저 권한을 확인해주세요.'));
}
```

---

## 요약 카드 (3열 그리드)

```html
<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
  <div class="sum-card">
    <div style="font-size:11px;color:#7a6a58;">항목명</div>
    <div id="s_value" style="font-size:26px;font-weight:800;color:#5a7aaa;margin:4px 0;">-</div>
    <div style="font-size:10px;color:#a09080;">단위</div>
  </div>
</div>
```

---

## 결과 행 렌더링 패턴 (짝수/홀수 배경)

```js
tbody.innerHTML = results.map((r, i) => `
  <div style="...;${i%2?'background:#faf5ee':''}">
    <span>${r.name}</span>
    <span style="font-weight:700;color:#4d7a5a;">${fmt(r.value)}</span>
  </div>
`).join('');
```

---

## 헤더 버튼 스타일

```html
<button onclick="exportExcel()" style="background:#4d7a5a;color:white;border:none;padding:7px 14px;border-radius:6px;font-size:13px;cursor:pointer;">📥 엑셀 다운로드</button>
<button onclick="copyResult()"  style="background:#5a8a6a;color:white;border:none;padding:7px 14px;border-radius:6px;font-size:13px;cursor:pointer;">결과 복사</button>
<button onclick="openModal()"   style="background:#5a6a9a;color:white;border:none;padding:7px 14px;border-radius:6px;font-size:13px;cursor:pointer;">📂 저장/불러오기</button>
<button onclick="resetAll()"    style="background:#8a7a6a;color:white;border:none;padding:7px 14px;border-radius:6px;font-size:13px;cursor:pointer;">초기화</button>
```

---

## 핵심 설계 원칙

1. **단일 파일**: 외부 파일 의존 없이 HTML 1개로 완결
2. **즉시 반응**: 모든 입력에 `oninput="calculate()"` — 실시간 갱신
3. **복원 버튼**: 수정 가능한 모든 비율 필드에 ↺ 버튼 필수
4. **기존 파일 수정 시**: 같은 색상 팔레트, 클래스명, 유틸 함수 유지
5. **항목 추가/삭제**: 카운터 증가만, 감소 없음. `if (!$(...)) continue` 패턴으로 삭제 항목 건너뜀
6. **차트**: 싱글톤 패턴 — 재생성 금지, `.update()` 사용

---

## 기존 대시보드 파일 위치

`C:\Users\josgo\Desktop\claudestudy\요율산정_대시보드.html`
