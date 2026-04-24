# 프로젝트 규칙: 모바일뉴스리더 신규 가입자 보고 자동화

## 1. 개요
- 연합뉴스 콘텐츠사업부가 매일 하던 수작업(관리자에서 전체 사용자 엑셀 다운 → 오늘 신규 집계 → 일간 보고 문장 작성)을 자동화한 웹 도구.
- 여러 작업자가 공유 URL로 사용. 배포: https://mnlnewuser.pages.dev
- 저장소: github.com/gimeon/mnlnewuser (main 브랜치 푸시 → Cloudflare Pages 자동 배포)

## 2. 기술 스택 & 아키텍처 결정
- **순수 정적 웹앱**: HTML/CSS/JS. 빌드 시스템 없음
- **Cloudflare Pages**: 정적 호스팅 + Pages Functions
- **D1 (mnlnewuser-db)**: 보고 기록 공유 저장소. 바인딩 변수명 `DB`
- **엑셀 파싱**: SheetJS (CDN)
- **Web Worker**: `XLSX.read`를 백그라운드 스레드에서 수행 (UI 블로킹 방지)
- **날짜 피커**: flatpickr (ko locale)

### 번복 금지 (Do-not-reopen)
- **자동 다운로드 기능은 구현 불가능**. 관리자 서버(mng.yna.co.kr)가 내부망 IP + Referer 체크라 외부 호스팅에서 호출 시 차단. 과거에 시도 후 확정됨. 업로드 방식만 사용.
- **업로드 카드는 파일 업로드 전용**. 자동 다운로드 UI는 완전 제거했으니 되돌리지 말 것.

## 3. 도메인 규칙

### 엑셀 컬럼 매핑 (`COLUMN_MAP` in app.js)
| 논리 필드 | 실제 컬럼명 |
|---|---|
| organization | 소속기관 |
| department | 부서 |
| name | 이름 |
| title | 직책 |
| phone | 전화번호 |
| registrationTime | 전화번호 등록 일시 |

### 대표자 선정 로직
1. 직책이 기재된 사람 우선 후보
2. 사전(rank-config.js의 `RANK_PRIORITY`)과 매칭 → 가장 높은 직급
3. 매칭 실패 시 가장 늦게 등록된 사람

### 보고 문장 포맷 (고정)
```
▲ 모바일 뉴스리더 제공(N개처)
- {기관} {부서} {이름} {직책} 등 N명(before명 -> after명)
```
- 한 기관에서 1명이면 "등" 생략, 인원수만
- 대표자 정보(부서/이름/직책)가 전부 비어 있으면 "등"도 생략 (예: "질병관리청 2명")
- 기관 정렬: 총 인원(after) 내림차순 → 신규 수 내림차순 → 기관명 (한글 로캘)

### 기간 계산
- **집계 시작점 옵션**: `lastReport` | `endZero` | `last24h` | `custom`
- **집계 종료점 옵션**: `fileModified` (기본) | `now` | `custom`
- 파일 업로드 시 endOption이 'now'면 'fileModified'로 자동 전환

## 4. UI/UX 관습

### 날짜 표기
모든 날짜/시각 표기는 `yyyy-mm-dd(요일) HH:MM` 형식.
예: `2026-04-23(목) 15:37`

### 언어
모든 UI 문구는 한국어.

### 색상
- **주 강조**: `#6366f1` (인디고) — 버튼, 링크, 강조 테두리
- **서브 강조**: `#f59e0b` (앰버) — "마지막/최신"을 표현할 때
- 결과 카드 강조 테두리는 **보고 생성 후에만** 적용 (빈 상태에선 연한 테두리)

### 상태 메시지
- `is-info`: 파란색. 진행 중/안내
- `is-success`: 초록색. 완료
- `is-error`: 빨간색. 오류
- 긴 작업 중엔 우측에 120px 프로그레스바 (determinate/indeterminate)

## 5. 유지보수 포인트

### 자주 바뀌는 상수 (app.js 상단)
- `COLUMN_MAP`: 엑셀 헤더 변경 시 후보 배열만 교체
- `HISTORY_MAX_ITEMS = 10`: 보고 기록 보관 개수. `functions/api/reports.js`의 `MAX_ITEMS`와 동기화 필요
- `MAX_FILE_SIZE_MB = 15`

### 직급 사전 (rank-config.js)
- `RANK_PRIORITY`: 기관 유형별 직급 배열
- `ORG_CATEGORY`: 기관명 키워드 → 유형 매핑. `includes()` 매칭

### 파일 구성
| 파일 | 역할 |
|---|---|
| `index.html` | 단일 페이지 UI |
| `app.js` | 전체 로직 (파싱/렌더/이벤트/팝오버) |
| `styles.css` | 스타일 |
| `rank-config.js` | 직급/기관 사전 |
| `functions/api/reports.js` | GET/POST/DELETE-all |
| `functions/api/reports/[id].js` | DELETE 개별 |

### 버전 표기
- 푸터 `<span class="footer-version">vX.Y</span>` 에서 수정
- 현재: v1.1
- 주요 기능 추가 시 minor 증가, UI/버그 픽스는 patch

## 6. 배포 운영

### 흐름
```
local edit → git add . → git commit -m "..." → git push
→ Cloudflare Pages 자동 빌드 (~30초~1분)
→ https://mnlnewuser.pages.dev 반영
```

### 커밋 메시지 스타일
- 제목 70자 이내, 동사형
- 본문에 변경 이유·사용자 요청 맥락 포함
- `Co-Authored-By` 유지

### D1 복구
- Cloudflare 대시보드 → D1 → mnlnewuser-db → Time Travel (일정 기간 내 복구)
- Console 탭에서 `SELECT * FROM reports;` 직접 조회 가능

## 7. 알려진 보안·제약
- Pages Functions API는 **인증 없음**. URL 아는 누구나 보고 기록 읽기/쓰기/삭제 가능. 현재는 "URL 비공개" 수준 보안. 필요시 Cloudflare Access(무료 50인) 추가 고려
- localStorage 키:
  - `dailyReportHistory_v1` (API 실패 시 폴백)
  - `dailyReportLastAuthor_v1` (작성자 이름 기억)
  - `dailyReportParseTimeMs_v1` (파싱 시간 캐시, 진행률 추정용)

## 8. `_thinking/` 폴더 운영 규칙

대화 중 프로젝트에 중요한 내용을 마크다운 문서로 `_thinking/` 폴더에 누적 저장한다.

### 저장 규칙
- **저장 트리거**: 사용자가 명시적으로 `"NNN 문서로 저장"`이라고 지시할 때만 저장한다. (예: `"001 문서로 저장"`)
- **파일명 형식**: `NNN-짧은제목.md` (예: `001-프로젝트에 필요한 파일.md`)
- **번호 부여**: 001부터 순차 증가
- **제목**: 짧고 내용을 잘 나타내도록

### 누적 원칙
- 기존 문서는 **수정하지 않는다**. 새 문서를 계속 쌓는 방식으로 운영.
- 내용이 겹치거나 업데이트되더라도 새 번호의 문서로 추가.

### 읽기 규칙
- `_thinking/`은 **명시적 요청이 있을 때만** 전체/부분 읽기. 자동 선제 로드 금지.
