# 푸드의비서 Google 로그인 장애 — 조사 결과와 배포 구성

대상 앱: https://foobierp.ai.studio
Firebase 프로젝트: `radiant-badge-xfs6l` (프로젝트 번호 99186442110)
조사일: 2026-07-27

## 요약

로그인이 깨지는 원인은 확인했으나, **현재 계정 권한으로는 수정할 수 없다.**
프로젝트가 Google AI Studio가 생성·소유한 managed project이고,
`kanrabhen@gmail.com`에게는 읽기 전용 역할만 부여되어 있다.

수정 권한을 가진 주체는 AI Studio 쪽이므로, 해결은 AI Studio를 통해 요청해야 한다.

## 원인 (확인 완료)

Google 로그인 팝업은 `authDomain`인 `radiant-badge-xfs6l.firebaseapp.com`의
인증 핸들러 페이지를 연다. 그 페이지가 로드하는 `handler.js`가
`/__/firebase/init.json`을 읽어 초기화하는데, 이 파일이 없다.

Firebase Hosting에 한 번도 배포된 적이 없어서 해당 경로가
Hosting의 "Site Not Found" 페이지를 반환한다. 초기화에 실패하니
로그인 결과가 앱으로 전달되지 못하고 팝업이 즉시 닫힌다.

검증 내역:

| 확인 항목 | 결과 |
|---|---|
| `handler.js`(280KB) 내 `/__/firebase/init.json` 참조 | 존재함 |
| `firebaseapp.com/__/firebase/init.json` | 404, 본문은 Hosting "Site Not Found" 페이지 |
| `firebaseapp.com/__/auth/handler` | 200 (정상) |
| `firebaseapp.com/__/auth/iframe` | 200 (정상) |
| `firebaseapp.com/__/auth/experiments.js` | 200 (정상) |
| `radiant-badge-xfs6l.web.app/` | 404 (배포 이력 없음) |
| 승인된 도메인에 `foobierp.ai.studio` | 등록되어 있음 |
| 앱 번들의 Firebase 설정 | projectId·authDomain 모두 `radiant-badge-xfs6l` |

즉 Hosting에 아무 파일이나 한 번 배포하면 `init.json`이 생성되어 해결된다.
문제는 그 배포를 할 권한이 없다는 것이다.

## 막힌 경로들 (모두 실측)

### 1. Firebase Hosting 배포 — 권한 없음

```
POST https://firebasehosting.googleapis.com/v1beta1/projects/-/sites/radiant-badge-xfs6l/versions
→ 403 {"error":{"code":403,"message":"The caller does not have permission","status":"PERMISSION_DENIED"}}

POST https://cloudresourcemanager.googleapis.com/v1/projects/radiant-badge-xfs6l:testIamPermissions
  body: {"permissions":["firebase.projects.get","firebasehosting.sites.update"]}
→ 403 USER_PROJECT_DENIED
  "Caller does not have required permission to use project radiant-badge-xfs6l.
   Grant the caller the roles/serviceusage.serviceUsageConsumer role, or a custom role
   with the serviceusage.services.use permission"
```

조회성 호출(프로젝트 조회, 사이트 조회, `projects:list`, `hosting:sites:list`)은
전부 200으로 성공한다. 쓰기만 막힌다.

### 2. IAM 자가 수정 — 권한 없음

`kanrabhen@gmail.com`에게 부여된 역할:
로그 뷰어, 모니터링 뷰어, Cloud Datastore 사용자, Cloud Run 뷰어,
Cloud SQL User (Free Tier), Firebase 뷰어, Firebase User (Free Tier),
Managed Projects Deleter / Upgrader / Viewer, Run SaaS Namespace User (Free Tier).

소유자(Owner)·편집자(Editor)·Firebase Hosting 관리자 없음.
`resourcemanager.projects.setIamPolicy`가 없어 스스로 역할을 추가할 수 없다.

IAM 목록의 나머지 주 구성원은 전부 Google 내부 서비스 계정이다
(`alkali-makersuite@prod.google.com` — makersuite는 AI Studio의 옛 이름,
`ais-sa@...`, `firebase-adminsdk-fbsvc@...`).
**사람 소유자가 한 명도 없어 권한을 부탁할 대상 자체가 없다.**

### 3. authDomain 교체 — 불가능

`foobierp.ai.studio/__/firebase/init.json`은 200을 반환하지만
본문이 앱의 `index.html`(SPA 폴백)이라 인증 핸들러로 쓸 수 없다.
`web.app` 도메인도 같은 프로젝트라 동일하게 404.

### 4. 팝업 우회 (Google Identity Services + signInWithCredential) — 막힘

Google 공급자 설정은 조회에 성공했다.

```
GET https://identitytoolkit.googleapis.com/admin/v2/projects/radiant-badge-xfs6l/defaultSupportedIdpConfigs/google.com
→ 200 { "enabled": true,
        "clientId": "99186442110-lmja4k6cf8ooetj020qpede57uj4n8cm.apps.googleusercontent.com" }
```

그러나 GIS는 페이지 출처가 해당 OAuth 클라이언트의
"승인된 JavaScript 원본"에 등록되어 있어야 동작한다. 확인 결과 미등록이다.

```
GET https://accounts.google.com/gsi/status?client_id=<위 clientId>&origin=https://foobierp.ai.studio
→ 403

GET https://accounts.google.com/gsi/status?client_id=<위 clientId>&origin=https://radiant-badge-xfs6l.firebaseapp.com
→ 403
```

원본을 추가하려면 OAuth 클라이언트 편집 권한(`clientauthconfig.clients.update`)이
필요한데, 이 역시 없다.

## 권장 해결 순서

### 1순위 — AI Studio에 배포를 요청한다

이 프로젝트에 쓰기 권한을 가진 주체는 AI Studio다.
AI Studio 채팅에 아래를 그대로 붙여넣어 요청한다.

> 이 앱의 Firebase 프로젝트(radiant-badge-xfs6l)에 Firebase Hosting이 한 번도 배포되지 않아서
> https://radiant-badge-xfs6l.firebaseapp.com/__/firebase/init.json 이 404를 반환하고,
> 그 결과 Google 로그인 팝업이 초기화에 실패해 즉시 닫힌다.
> Firebase Hosting에 최소 구성(index.html 한 개)으로 한 번 배포해서 init.json이 생성되게 해달라.
> Firestore와 Authentication 설정·데이터는 건드리지 말고 Hosting만 배포해달라.

배포 후 아래가 200이고 JSON이 나오면 해결된 것이다.

```
curl -i https://radiant-badge-xfs6l.firebaseapp.com/__/firebase/init.json
```

### 2순위 — 프로젝트 쓰기 권한을 확보한다

AI Studio에서 이 프로젝트를 본인 소유로 전환하는 옵션이 있는지 확인한다
(`Managed Projects Upgrader` 역할이 있는 것으로 보아 관련 기능이 존재할 가능성이 있다).
불가능하면 Firebase 지원(https://firebase.google.com/support/troubleshooter/contact)에
managed project의 Hosting 배포 권한을 요청한다.

권한을 얻은 뒤에는 이 디렉터리에서 바로 배포하면 된다.

```bash
npm install -g firebase-tools
firebase login
firebase deploy --only hosting --project radiant-badge-xfs6l
```

`firebase.json`과 `public/index.html`은 이미 준비되어 있다.
Hosting만 배포하며 Firestore·Authentication에는 영향이 없다.

### 3순위 — 직접 소유하는 Firebase 프로젝트로 이전

위가 모두 막힐 때의 최후 수단이다. 새 프로젝트를 만들면 소유자가 되므로
Hosting 배포도 OAuth 설정도 자유롭다. 다만 Firestore 데이터 이전이 필요하고,
기존 사용자 계정은 내보낼 권한이 없어 UID가 바뀐다.
데이터가 UID를 키로 쓰고 있다면 영향이 크므로 사전 검토가 필요하다.

## 참고

- 이번 조사 중 Firestore·Authentication 설정과 데이터는 일절 변경하지 않았다.
  수행한 쓰기 시도는 Hosting 배포 하나뿐이고 403으로 거부되어 기록된 변경이 없다.
- 앱 번들에서 확인한 `apiKey`, `clientId`는 클라이언트에 노출되는 공개 값이다.
  OAuth 클라이언트 시크릿은 조회하지 않았고 이 문서에도 포함하지 않았다.
