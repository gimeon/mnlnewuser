# 일간 신규가입자 보고 자동화 도구

관리자 페이지에서 받은 사용자 엑셀(~10MB)을 업로드하면 지정된 일간 보고 문장을 자동 생성하는 순수 정적 웹앱.
서버·DB·로그인 없음. 데이터는 사용자의 브라우저를 벗어나지 않음.

## 빠른 시작 (로컬)

1. 이 폴더의 `index.html`을 브라우저로 직접 열기 (`file://` 로도 동작)
2. `sample_users.xlsx` 를 드롭존에 드래그 → 예시 보고 문장이 나오면 OK

## 사용 흐름

1. 작업자가 관리자 페이지에서 사용자 엑셀을 다운로드 (기존 절차와 동일)
2. 도구 URL 접속 → 엑셀 파일을 드롭
3. 기준 날짜 = 기본값 "오늘". 필요하면 datepicker로 변경
4. 결과 문장이 표시되면 **[복사]** 버튼으로 클립보드로 옮겨 일간 보고에 붙여넣기

## 엑셀 컬럼 매핑 변경

실제 엑셀의 헤더가 예상과 다르면 [app.js](app.js) 상단의 `COLUMN_MAP`만 수정하면 된다.

```js
const COLUMN_MAP = {
  organization: ['기관', '기관명', '소속', 'organization'],   // 후보를 배열로 나열
  department:   ['부서', '부서명', 'department'],
  name:         ['이름', '성명', 'name'],
  title:        ['직책', '직급', '직위'],
  signupDate:   ['가입일', '가입일자', '등록일'],
};
```

- 후보를 여러 개 넣어두면 순서대로 시도
- 실제 헤더 문자열을 그대로 넣으면 됨

## 직급 사전 보강

매칭이 안 되는 직책은 콘솔에 경고가 남고 원본 순서 첫 번째로 폴백한다.
필요한 직급을 [rank-config.js](rank-config.js)의 해당 카테고리에 추가하면 된다.

- 기관 → 카테고리 매핑도 `ORG_CATEGORY`에서 관리
- 기관명은 `includes()`로 매칭되므로 `경찰청`만 등록해두면 `경찰청 서울청 …` 모두 대응

## 배포 (여러 명이 웹 URL로 공유)

### Cloudflare Pages (추천)
1. GitHub/GitLab 저장소에 이 폴더 푸시
2. Cloudflare 대시보드 → Pages → Create project → 저장소 선택
3. 빌드 명령어 비워두고 Deploy
4. `https://<프로젝트>.pages.dev` URL 공유

### GitHub Pages
1. GitHub 저장소 Settings → Pages
2. Source: `main` / root
3. `https://<계정>.github.io/<저장소>` URL 공유

두 방법 모두 무료, 서버 관리 없음, 5분 이내 배포 완료.

## 검증

- `node test_logic.js` — 순수 로직 헤드리스 테스트 (기대 출력 일치 + 엣지 케이스)
- `python3 generate_sample.py` — 테스트용 `sample_users.xlsx` 재생성

## 파일 구성

| 파일 | 역할 |
|---|---|
| `index.html`      | UI (드롭존, 결과 영역, 복사 버튼) |
| `styles.css`      | 스타일 |
| `app.js`          | 엑셀 파싱 → 필터 → 그룹핑 → 문장 생성 |
| `rank-config.js`  | 직급 우선순위·기관 카테고리 사전 |
| `test_logic.js`   | Node 단독 실행 로직 테스트 |
| `generate_sample.py` | 테스트 xlsx 생성 (openpyxl 필요) |
| `sample_users.xlsx` | 테스트용 샘플 데이터 |
