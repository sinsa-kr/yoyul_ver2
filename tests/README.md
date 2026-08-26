# 회귀 검증 하네스

대시보드(`../요율산정_대시보드_v2.html`)의 동작을 고정해두는 테스트다.
전에는 임시 폴더에 두었다가 두 번 통째로 유실돼, 리포 안으로 옮겼다.

## 실행

```bash
cd tests
npm install          # 최초 1회 (jsdom · xlsx-js-style · playwright)
npx playwright install chromium

node smoke.js        # 계산·경고·저장복원·단가표  (jsdom)
node excel.js        # 엑셀 출력 + 전 수식 재계산 검증
node browserCheck.js # 실제 브라우저 — 자동로드·드래그앤드롭·테마

node realcheck.js    # ../단가표.xlsx 를 갈아끼운 뒤 잘 읽히는지 점검
```

`realcheck.js` 는 **반기마다 단가표를 교체한 뒤** 한 번 돌려보면 된다.
파싱된 직종 수·미공표 건수·동명이직 구분 결과·단위 분포·최저임금 미만 직종을 출력한다.

`node_modules/` 와 `package-lock.json` 은 `.gitignore` 로 제외된다.

## 파일

| 파일 | 역할 |
|---|---|
| `harness.js` | 대시보드를 jsdom 에 띄우는 공용 로더 + 결과 집계 |
| `smoke.js` | 화면 동작 전반 |
| `excel.js` | 엑셀 워크북 구조·서식·값 |
| `formulaEval.js` | 엑셀 수식 미니 평가기 — 모든 `.f` 를 다시 계산해 `.v` 와 대조 |
| `browserCheck.js` | Playwright — jsdom 이 못 잡는 것들 |
| `realcheck.js` | 실제 `단가표.xlsx` 파싱 결과 점검 (교체 후 확인용) |

## 왜 이렇게 나눠 두었나

- **`formulaEval.js`** — 엑셀 셀에 수식(`.f`)과 캐시값(`.v`)을 함께 쓰는데, 둘이 어긋나면
  파일을 열기 전까지 아무도 모른다. 실제로 연산자 우선순위 버그를 이 검사가 잡았다.
- **`browserCheck.js`** — HTML5 드래그앤드롭은 jsdom 이 `effectAllowed` 를 강제하지 않아
  통과하지만 실제 브라우저에서는 드롭이 취소되는 회귀가 있었다. 단가표 자동 로드(`fetch`)도
  실제 서버가 있어야 확인된다.

## 단가표 테스트는 내용에 의존하지 않는다

`단가표.xlsx` 는 반기마다 통째로 교체된다. 그래서 브라우저 검사는 특정 직종명을 박아두지 않고,
**로드된 표에서 조건에 맞는 표본을 골라** 검증한다(유상 직종 하나, 해설에만 있는 단어 하나).
표를 갈아끼워도 테스트를 고칠 필요가 없다.
