# Overview 저장 — PC·휴대폰 공유 (최초 1회만)

GitHub Pages 주소(`https://kriskang02-max.github.io/...`)에서는 **브라우저만으로는 파일을 저장할 수 없습니다.**  
그래서 **Firebase(무료 클라우드)** 에 Overview 내용을 올리도록 했습니다. **한 번만 설정**하면, 이후에는 **저장 버튼**만 누르면 휴대폰에서도 같은 내용이 보입니다.

## 1. Firebase 프로젝트 만들기 (약 5분)

1. [Firebase 콘솔](https://console.firebase.google.com/) 접속 → Google 로그인  
2. **프로젝트 추가** → 이름 예: `market-db-dashboard` → 만들기  
3. 왼쪽 **빌드 → Realtime Database** → **데이터베이스 만들기**  
   - 지역: `asia-southeast1` (또는 가까운 곳)  
   - 보안 규칙: **테스트 모드로 시작** (개인용; 나중에 규칙 잠글 수 있음)  
4. ⚙ **프로젝트 설정** → 아래로 스크롤 → **앱 추가** → **</> 웹**  
   - 앱 닉네임 아무거나 → 앱 등록  
5. `firebaseConfig` 안에서 아래 네 가지를 복사해 둡니다.  
   - `apiKey`  
   - `authDomain`  
   - `databaseURL` (Realtime Database URL — `https://....firebasedatabase.app`)  
   - `projectId`

## 2. 대시보드에 붙여넣기

`firebase_overview_config.js` 파일을 열고 복사한 값을 넣습니다.

```javascript
window.FIREBASE_OVERVIEW_CONFIG = {
  apiKey: "여기에",
  authDomain: "여기에",
  databaseURL: "여기에",
  projectId: "여기에",
};
```

## 3. GitHub에 올리기

PowerShell:

```powershell
cd C:\Users\infomax\Documents\market_db_dashboard
git add firebase_overview_config.js
git commit -m "chore: Firebase Overview 동기화 설정"
git push
```

1~2분 뒤 Pages가 갱신되면 끝입니다.

## 4. 사용 방법

1. PC·휴대폰 모두 **같은 주소**로 접속  
   `https://kriskang02-max.github.io/market_db_dashboard/dashboard.html`  
2. **Overview** 탭 → 상단에 **「클라우드 연결됨」** 이 보이면 OK  
3. **저장** 클릭 → 메시지 **「저장됨 · PC·휴대폰에 반영됩니다」**  
4. 휴대폰에서 Overview를 열거나 새로고침하면 같은 내용이 보입니다.

## 문제 해결

| 증상 | 해결 |
|------|------|
| 상단에 「클라우드 미연결」 | `firebase_overview_config.js` 값·push 여부 확인 |
| 저장해도 휴대폰에 안 보임 | 휴대폰에서 **Overview** 탭 + **새로고침** |
| Firebase 규칙 오류 | Realtime Database → **규칙** 탭에서 테스트용으로 읽기·쓰기 허용 |

테스트용 규칙 예 (본인만 쓰는 비공개 용도):

```json
{
  "rules": {
    "overview_state_v1": {
      ".read": true,
      ".write": true
    }
  }
}
```
