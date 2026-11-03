# 협업 프로토콜 (지휘 Claude · 작업 Claude · 검토 Codex)

Herdr 한 탭에서 세 에이전트가 협업한다. **이 문서를 먼저 읽고 자기 역할만 수행한다.**

| 역할 | Herdr 이름 | 위치 | 하는 일 | 하지 않는 일 |
|---|---|---|---|---|
| 지휘 (orchestrator) | — | 왼쪽 pane (Claude) | 요구 해석·계획·작업 분해, 작업지시서 작성, 결과 1차 확인, 검토 지적 채택/기각 판정, 사용자 보고 | 직접 구현(원칙), 검토 횟수 상한 초과 |
| 작업자 (worker) | `worker` | 가운데 pane (Claude) | 지시서대로 구현, 테스트 실행, 보고서 작성 | 지시서 범위 밖 변경, **커밋/푸시**, 리뷰 파일 수정 |
| 검토자 (reviewer) | `reviewer` | 오른쪽 pane (Codex) | 읽기 전용 코드 검토, 검토서 작성 | **코드 수정**, 커밋, 지시서·보고서 수정 |

> 사용자는 언제든 어느 pane에 직접 지시할 수 있다. 사용자의 직접 지시가 이 문서보다 우선한다.

---

## 1. 작업 폴더 구조

```
.collab/
├─ PROTOCOL.md            ← 이 문서 (git 추적)
├─ templates/             ← 양식 (git 추적)
├─ BOARD.md               ← 작업 현황판 (지휘만 수정, gitignore)
└─ tasks/<ID>/            ← 작업별 폴더 (gitignore — 실데이터가 섞일 수 있음)
   ├─ brief.md            ← 지휘 작성: 작업지시서 + 판정 기록
   ├─ report.md           ← worker 작성: 구현 보고 (수정 라운드는 아래에 추가)
   ├─ diff.patch          ← 지휘 작성: 검토 대상 diff 스냅샷
   ├─ review-R1.md        ← reviewer 작성: 1차 검토
   └─ review-R2.md        ← reviewer 작성: 2차(최종) 검토
```

- 작업 ID: `T<MMDD>-<NN>` (예: `T1103-01`)
- **파일당 작성자는 한 명**이다. 남의 파일은 읽기만 한다 → 편집 충돌 없음.
- 저장소는 GitHub Pages로 공개된다. `.collab/tasks/`, `BOARD.md`는 gitignore 대상이며, 추적되는 파일에 실제 업체명·임금·원가를 쓰지 않는다.

## 2. 작업 흐름

```
사용자 요청
  → [지휘] 계획·분해 → tasks/<ID>/brief.md 작성 → worker에 지시
  → [worker] 구현 → 테스트 → report.md 작성 → "완료: <ID>" 응답
  → [지휘] report + diff 확인 → diff.patch 저장 → reviewer에 R1 요청
  → [reviewer] review-R1.md 작성 → "검토완료: <ID> R1 <판정>" 응답
  → [지휘] 지적별 채택/기각 판정(brief.md에 기록)
       ├ PASS 또는 채택 0건 → 종료
       └ 채택 있음 → worker에 수정 지시 → report.md에 "수정 R1" 추가
  → [reviewer] review-R2.md (최종) → [지휘] 판정
  → [지휘] 사용자에게 완료 보고 (요약 · 검증 명령 · 커밋 메시지 제안)
```

## 3. 검토 횟수 규칙 (중요)

1. **작업 1건당 Codex 검토는 최대 2회(R1, R2).** 검토 요청 1회 전송 = 1회 차감.
2. R1이 PASS면 R2는 생략한다.
3. R2는 **최종 검토**다. 범위: R1 채택 지적의 해결 여부 + 수정 diff에서 생긴 새 문제.
4. R2 이후 잔여 이슈 처리:
   - P2/P3 → 지휘가 확인 후 종료하거나 후속 작업으로 등록
   - P0/P1 → 지휘가 직접 검증해 worker에게 수정시키고 테스트로 확인. 그래도 **추가 검토가 필요하다고 판단되면 사용자 승인을 받은 뒤에만 R3** 진행 (사유·잔여 이슈·검토 범위를 제시)
5. 작업을 쪼개거나 ID를 바꿔 검토 횟수를 우회하지 않는다.
6. 검토 횟수는 `BOARD.md`의 `검토` 칸에 `1/2`, `2/2`, `3/2(승인)` 형식으로 기록한다.

## 4. 심각도 기준

| 등급 | 의미 | 처리 |
|---|---|---|
| P0 | 데이터 손상·계산 오류·실데이터 노출·크래시 | 반드시 수정 |
| P1 | 요구사항 미충족, 회귀, 명확한 버그 | 원칙적으로 수정 |
| P2 | 엣지 케이스, 유지보수성, 테스트 누락 | 지휘 판단 |
| P3 | 스타일·취향 | 기본 기각 |

검토 판정: `PASS` (P0/P1 없음) / `CHANGES_REQUESTED` (P0/P1 있음)

## 5. 역할별 세부 규칙

### worker (Claude)
- `brief.md`의 **범위·완료조건**만 수행. 범위 밖 문제는 고치지 말고 report의 "발견 사항"에 적는다.
- 불명확하면 추측하지 말고 report에 질문을 적고 `질문: <ID>`로 응답한 뒤 멈춘다.
- 테스트: brief에 지정된 명령을 실행하고 결과(통과/실패 수)를 report에 붙인다.
- 커밋·푸시·브랜치 생성 금지. 작업 완료 응답은 한 줄: `완료: <ID>` (또는 `질문: <ID>`, `실패: <ID>`).

### reviewer (Codex)
- **코드를 수정하지 않는다.** 쓰기는 `.collab/tasks/<ID>/review-R<n>.md` 하나만.
- 검토 기준: brief의 완료조건 충족 여부, 정확성(계산·수식·저장복원), 회귀, 실데이터 노출, 테스트 적절성.
- 지적마다 `파일:라인`, 근거, 재현 방법 또는 실패 시나리오, 수정 제안을 쓴다. 확신 없는 지적은 `[추정]` 표시.
- 취향성 리팩터링 제안은 P3로만. 응답은 한 줄: `검토완료: <ID> R<n> PASS|CHANGES_REQUESTED`.

### 지휘 (Claude)
- 지시는 파일로, 프롬프트는 짧게 ("`<경로>` 읽고 수행").
- 검토 요청 직전 `git diff -- <변경 파일>`을 `diff.patch`로 저장해 범위를 고정한다.
- 판정 기록: brief.md 하단 "판정" 표에 지적별 채택/기각과 사유.
- 커밋은 사용자 요청 시에만, conventional commit 형식으로.

## 6. 효율 운영

- **파이프라이닝**: 변경 파일이 겹치지 않으면 reviewer가 작업 A를 검토하는 동안 worker는 작업 B를 진행한다 (diff.patch로 범위가 고정되므로 가능). 겹치면 순차.
- 작은 연관 수정은 한 작업으로 묶어 검토 횟수를 아낀다. 큰 작업은 검토 가능한 단위(대략 diff 300줄 이하)로 나눈다.
- 막힘(`blocked`, 권한 요청 등)은 지휘가 확인 후, 판단이 필요하면 사용자에게 묻는다.

## 7. 지휘용 Herdr 명령

```bash
herdr agent prompt worker   ".collab/tasks/<ID>/brief.md 읽고 수행" --wait --timeout 600000
herdr agent prompt reviewer ".collab/tasks/<ID>/ R1 검토 수행 (.collab/PROTOCOL.md 기준)" --wait --timeout 600000
herdr agent wait   worker                     # 긴 작업은 백그라운드로 대기
herdr agent read   worker --source recent-unwrapped --lines 80
herdr agent get    reviewer                   # 상태 확인 (idle/working/blocked)
```
