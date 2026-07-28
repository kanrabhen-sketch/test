# 푸드의비서 Google 로그인 장애 — 원인, 경위, 이전 결과

대상 앱: https://foobierp.ai.studio
조사·이전일: 2026-07-27 ~ 2026-07-28

## 결론

기존 Firebase 프로젝트 `radiant-badge-xfs6l`에서는 로그인을 고칠 수 없어,
정동진님 소유의 새 프로젝트 `foobierp-vy1v5`를 만들어 이전했다.

새 프로젝트에서는 `init.json`이 정상 생성되어 로그인 차단 요인이 해소됐다.

## 원인

Google 로그인 팝업은 `authDomain`의 인증 핸들러 페이지를 연다.
그 페이지의 `handler.js`가 `/__/firebase/init.json`을 읽어 초기화하는데,
Firebase Hosting에 한 번도 배포된 적이 없어 이 파일이 없었다.
초기화에 실패하니 로그인 결과가 앱으로 전달되지 못하고 팝업이 즉시 닫혔다.

`handler.js`(280KB)에서 해당 경로 참조를 확인했고,
응답 본문이 Hosting의 "Site Not Found" 페이지인 것도 확인했다.
`/__/auth/handler`, `/__/auth/iframe`, `/__/auth/experiments.js`는 모두 200으로 정상이었다.

## 기존 프로젝트에서 막힌 이유

`radiant-badge-xfs6l`은 AI Studio가 생성해 Google이 소유·관리하는 managed project이고,
`kanrabhen@gmail.com`에게는 읽기 전용 역할만 부여되어 있었다.

```
POST https://firebasehosting.googleapis.com/v1beta1/projects/-/sites/radiant-badge-xfs6l/versions
→ 403 {"error":{"code":403,"message":"The caller does not have permission","status":"PERMISSION_DENIED"}}

POST cloudresourcemanager.googleapis.com/v1/projects/radiant-badge-xfs6l:testIamPermissions
  {"permissions":["firebase.projects.get","firebasehosting.sites.update"]}
→ 403 USER_PROJECT_DENIED
```

부여된 역할은 로그 뷰어, 모니터링 뷰어, Cloud Datastore 사용자, Cloud Run 뷰어,
Cloud SQL User, Firebase 뷰어, Firebase User, Managed Projects Deleter/Upgrader/Viewer,
Run SaaS Namespace User 뿐이었다. 소유자·편집자·Firebase Hosting 관리자가 없고,
IAM 목록에 사람 소유자가 아예 없어 권한을 요청할 대상도 없었다.

검토했으나 모두 막힌 경로:

| 시도 | 결과 |
|---|---|
| Hosting 직접 배포 | 403 PERMISSION_DENIED |
| IAM 자가 역할 부여 | `setIamPolicy` 권한 없음 |
| authDomain을 `foobierp.ai.studio`로 교체 | 해당 도메인은 SPA 폴백만 반환, 인증 핸들러 없음 |
| GIS + signInWithCredential 우회 | OAuth 클라이언트에 출처 미등록(403), 등록 권한 없음 |
| AI Studio 플랫폼에 배포 요청 | 플랫폼에 Hosting 프로비저닝 권한·기능 자체가 없다고 공식 회신 |

## 이전 내용

신규 프로젝트: **`foobierp-vy1v5`** (Foobi ERP, 프로젝트 번호 793652092378)

수행한 작업:

1. Firebase 프로젝트 생성 (생성자가 소유자)
2. Firebase Hosting 배포 → `init.json` 생성 확인
3. 웹 앱 등록 → appId 발급
4. Firestore API·Identity Toolkit API 활성화
5. Firestore 데이터베이스 생성 — 기존과 **동일한 이름**을 사용해
   앱의 `firestoreDatabaseId` 상수는 수정 불필요
   (`ai-studio-2ebc80c3-0c85-4e16-bcac-c4329e982059`, us-west1)
6. 문서 19건 복사 (하위 컬렉션 없음)
7. 보안 규칙 배포

검증 결과:

```
GET https://foobierp-vy1v5.firebaseapp.com/__/firebase/init.json  → 200 (JSON 정상)
GET https://foobierp-vy1v5.firebaseapp.com/__/auth/handler        → 200
```

복사된 데이터:

| 컬렉션 | 문서 수 |
|---|---|
| activity_logs | 4 |
| clients | 2 |
| projects | 5 |
| resources | 1 |
| tasks | 1 |
| teams | 5 |
| users | 1 |

`users`의 유일한 문서(`kanrabhen@gmail.com`, `team_admin`, 활성)가
문서 ID까지 그대로 이전됐다.

기존 프로젝트는 삭제하지 않았고 읽기만 했다. 언제든 되돌릴 수 있다.

## 보안 규칙 변경 (확인 필요)

기존 프로젝트의 규칙은 다음과 같았다.

```
match /{document=**} {
  allow read, write: if true; // Internal business app authenticated users rule
}
```

주석은 "인증된 사용자"라고 되어 있으나 실제로는 **누구나 읽고 쓸 수 있는 상태**였다.
API 키만 알면 사내 데이터 전체가 노출·변조 가능하다.

새 프로젝트에는 주석의 의도대로 다음 규칙을 배포했다.

```
match /{document=**} {
  allow read, write: if request.auth != null;
}
```

앱의 모든 Firestore 접근이 로그인 이후에 일어나므로 동작에 문제가 없어야 한다.
만약 이전 후 데이터가 안 보이는 증상이 생기면 `firestore.rules`의 조건을
`if true`로 되돌려 원인을 분리할 수 있다. 다만 그 상태를 유지하는 것은 권장하지 않는다.

## UID 변경 영향 — 없음

앱은 로그인 후 사용자를 UID가 아니라 **이메일로 조회**한다.

```js
const email = user.email?.toLowerCase().trim() || "";
const q = query(collection(db,"users"), where("email","==",email), limit(1));
```

따라서 새 프로젝트에서 UID가 새로 발급돼도 이메일이 같으면
기존 문서를 찾아 역할·권한이 그대로 유지된다.
관리자가 직원을 추가할 때도 UID가 아닌 자체 생성 ID를 쓰고 있어,
문서 ID와 UID가 다른 상황을 앱이 이미 정상 처리한다.

Firebase Storage는 설정 문자열만 있고 실제 사용처가 없어 이전 대상이 아니다.

## 남은 작업 (콘솔·AI Studio에서 수동)

### 1. Google 로그인 공급자 사용 설정

https://console.firebase.google.com/project/foobierp-vy1v5/authentication/providers

Authentication → 시작하기 → Google → 사용 설정 → 저장.

API로는 OAuth 클라이언트를 자동 생성할 수 없어(`client_id cannot be empty`)
콘솔에서 한 번 눌러야 한다. 이 동작이 Auth 초기화와 OAuth 클라이언트 생성을 함께 처리한다.

### 2. 승인된 도메인에 앱 도메인 추가

Authentication → Settings → 승인된 도메인 → `foobierp.ai.studio` 추가.

### 3. 앱의 Firebase 설정 교체 (AI Studio)

| 항목 | 기존 | 신규 |
|---|---|---|
| projectId | `radiant-badge-xfs6l` | `foobierp-vy1v5` |
| appId | `1:99186442110:web:c7526be378cdfafb597d2f` | `1:793652092378:web:4c701082aa004f5d2ef4e3` |
| apiKey | `AIzaSyDmcdmcV02iKY-lhnCco65f453LkCg_lhs` | `AIzaSyADfeeTDvvYQW_fy4pZQYwSQtEEAIueD4c` |
| authDomain | `radiant-badge-xfs6l.firebaseapp.com` | `foobierp-vy1v5.firebaseapp.com` |
| storageBucket | `radiant-badge-xfs6l.firebasestorage.app` | `foobierp-vy1v5.firebasestorage.app` |
| messagingSenderId | `99186442110` | `793652092378` |
| firestoreDatabaseId | `ai-studio-2ebc80c3-...` | **변경 없음** |

## 이 디렉터리

`firebase.json`과 `public/index.html`은 Hosting 배포에 사용한 최소 구성이다.
재배포가 필요하면 다음을 실행한다.

```bash
firebase deploy --only hosting --project foobierp-vy1v5
```
