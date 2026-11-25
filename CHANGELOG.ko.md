# 변경 내역

[English](CHANGELOG.md) | **한국어**

## 2026 — 전면 정리

### 2023 버전에서 업그레이드하기

1. Apps Script 편집기에서 `Code.gs`, `MailMerge.html`을 교체하고 `i18n.gs`, `Lang_en.gs`, `Lang_ko.gs`, `Manual_en.html`, `Manual_ko.html`을 추가합니다.
2. `UserManual.html`, `createDraft.html`, `createSlide.html`은 삭제합니다 (더 이상 쓰지 않음).
3. `appsscript.json`을 교체합니다 (또는 서비스(+)에서 **Gmail API**, **Google Slides API** 추가).
4. 스프레드시트를 새로고침하고 테스트 메일로 확인합니다.

기존 시트의 `EMAIL_SENT` 기록과 `{{Name}}` 형식 필드는 그대로 호환됩니다.

### 이모지 깨짐 (🎄 → ��) 원인과 해결

- 2023년에 보내는 주소 문제로 바꾼 `GmailApp.sendEmail`은 4바이트 문자(대부분의 이모지)를 제대로 인코딩하지 못합니다. 이모지 1개가 깨진 글자 2개(��)로 바뀌고, ➕ 같은 3바이트 기호만 살아남습니다.
- 그렇다고 `MailApp`으로 되돌리면 보내는 주소(별칭)가 무시됩니다.
- 그래서 메일 원문(MIME)을 UTF-8로 직접 만들어 **Gmail API**로 보냅니다. 별칭과 이모지가 둘 다 됩니다.

### 고친 버그

| # | 기존 문제 | 수정 |
|---|---|---|
| 1 | 빈 행이 있으면 `EMAIL_SENT`가 **다른 사람 행**에 기록됨 | 실제 시트 행 번호로 기록 |
| 2 | 머리행이 한글이거나 빈 칸이면 필드가 **한 칸씩 밀림** | 열 단위로 짝지음, `{{이름}}` 사용 가능 |
| 3 | `in:draft`/`in:drafts`를 섞어 순번으로 선택 → **다른 초안 발송** 가능, 답장 초안은 남의 메일을 템플릿으로 씀 | 초안 ID로 선택 |
| 4 | 한 행 오류로 **전체 중단**, 오류 창도 깨짐 | 행별로 `ERROR: 사유` 기록 후 계속 |
| 5 | 6분 제한에 걸리면 어디까지 갔는지 모름 | 약 5분에 멈추고 다시 RUN하면 이어서 |
| 6 | 슬라이드 이미지 사용 시 원본에 **장이 계속 쌓임** | 복사본 삭제 |
| 7 | 날짜가 `Mon Dec 22 2026 …`으로, `0`은 빈칸으로 들어감 | 시트에 보이는 그대로 |
| 8 | `&`, `<`가 HTML을 깨뜨림, `$`가 이상하게 치환 | 이스케이프, 그대로 삽입 |
| 9 | `{{<span>Name</span>}}`처럼 태그가 끼면 치환 안 됨 | 태그 제거 후 인식 |
| 10 | `Email` 헤더가 없으면 기존 머리행을 덮어씀 | `Email`/`Email Address`/`이메일` 자동 인식 |
| 11 | `Mail Merge Status`가 끝 열이 아니면 열 중복, 마지막 열에 덮어씀 | 머리행에서 위치를 찾음 |
| 12 | RUN 연타 시 중복 발송 가능 | 실행 잠금 + 2단계 확인 |
| 13 | 1행뿐인 시트에서 '시트 내용 지우기' 오류 | 행 수 확인 |
| 14 | 초안 제목을 자를 때 이모지가 반으로 잘림 | 문자 단위로 자름 |

### 추가

- 한/영 화면 (`i18n.gs`, `Lang_*.gs`). Google 계정 언어로 자동 선택, 사이드바 🌐에서 변경
- 발송 전 대상 인원 확인
- `EMAIL_SENT` 인식 시 대소문자·띄어쓰기·`_`/`-` 무시 (`Email_Sent`, `email sent`도 건너뜀)
- CC 입력칸, CC/BCC 주소 검사
- Gmail '보내는 주소' 설정의 이름 자동 입력, 마지막 입력값 기억
- 시트에 없는 필드 경고
- 텍스트 버전 본문 함께 발송
- 일시적 Gmail API 오류 자동 재시도
- Gmail/Slides API가 꺼져 있으면 켜는 링크와 함께 안내

### 삭제

- 이름을 알아내려고 드라이브에 `Mail Merge.txt`를 만들던 `getMyName()`
- `getResult()`, `rangeToObjects()`, `onInstall()`, 예전 행 읽기 함수들 (→ `readTable_()`)
- 사이드바의 개별 호출 4번 (→ `getSidebarData()` 1번), 불필요한 JSON 왕복, jQuery, 주석 처리된 코드
- `createDraft.html`, `createSlide.html`, `UserManual.html` (사이드바·`Manual_*.html`로 대체)
- "오늘 보낼 수 있는 메일 N통" 표시 (MailApp 할당량이라 실제 발송과 맞지 않음)

## 2023-11-15

- 보내는 주소(별칭)가 적용되도록 `MailApp.sendEmail` → `GmailApp.sendEmail`
- Reply-To를 보내는 주소로 설정

## 2019

- 첫 공개: Gmail 초안 템플릿, 본문 삽입 이미지, 첨부파일, 테스트 메일, BCC, 슬라이드 이미지
