# 인명사전 읽기+수정 시스템 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 읽기 전용 인명사전 사이트에 Apps Script 웹앱 기반의 본인 PIN 인증 수정 기능을 추가한다.

**Architecture:** GitHub Pages 정적 사이트(index.html)가 Apps Script 웹앱(구글 시트 Master에 바인딩)을 JSON API로 호출한다. 읽기는 `doGet`(PIN 열 제외), 쓰기는 `doPost`(PIN 검증 → 해당 행만 수정 + 최종수정일 자동 기록). Apps Script의 핵심 로직은 GAS 전역 객체에 의존하지 않는 순수 함수로 작성해 Node 내장 테스트 러너로 검증하고, 같은 로직을 재사용하는 로컬 mock 서버로 프런트엔드를 검증한다.

**Tech Stack:** Vanilla JS + Tailwind CDN (기존 유지), Google Apps Script, Node.js 내장 `node:test` (devDependency 없음)

**스펙:** `docs/superpowers/specs/2026-06-12-member-db-edit-system-design.md`

---

## 파일 구조

| 파일 | 역할 |
|------|------|
| `apps-script/Code.gs` (신규) | API 전체. 순수 로직(handleGet/handlePost 등) + GAS 글루(doGet/doPost). 시트 편집기에 복붙해 배포 |
| `tests/code-gs.test.mjs` (신규) | Code.gs 순수 로직 단위 테스트 (`node --test`) |
| `tests/mock-server.mjs` (신규) | Code.gs 로직을 그대로 재사용하는 로컬 API mock — 프런트 수동 검증용 |
| `config.js` (수정) | `SHEET_CSV_URL` → `API_URL` 교체, `SURVEY_URL` 유지 |
| `index.html` (수정) | JSON 읽기 전환, 역대 소속·최종수정일 표시, 수정 모달 |
| `readme.md` (전면 갱신) | 비개발자용 운영·배포·인수인계 매뉴얼 |

**API 필드명 ↔ 시트 헤더 매핑 (모든 태스크 공통):**
`name`=이름, `cohort`=기수, `role`=역할, `org`=현재 소속, `history`=역대 소속, `phone`=전화, `email`=이메일, `interests`=관심사, `pin`=PIN, `updatedAt`=최종수정일

**API 응답 형식 (모든 태스크 공통):**
- `GET` → `{ok:true, members:[{name,cohort,role,org,history,phone,email,interests,updatedAt}, ...]}` (PIN 없음)
- `POST` body `{name, cohort, pin, fields:{...}}` → 성공 `{ok:true, member:{...}}` / 실패 `{ok:false, error:"bad_request"|"not_found"|"wrong_pin"|"field_not_editable"}`
- `fields`가 `{}`이면 PIN 검증만 수행 (시트 수정 없음)
- 수정 가능 필드: `org, history, phone, email, interests, pin`

---

### Task 1: Code.gs 읽기 로직 (handleGet) — TDD

**Files:**
- Create: `apps-script/Code.gs`
- Create: `tests/code-gs.test.mjs`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/code-gs.test.mjs` 생성:

```js
import { readFileSync } from "node:fs";
import test from "node:test";
import assert from "node:assert/strict";

// Code.gs는 GAS용 plain JS — var/function 선언만 쓰므로 new Function으로 로드 가능.
// GAS 전역(SpreadsheetApp 등)은 글루 함수 내부에서만 참조하므로 정의 시점에는 안전하다.
const src = readFileSync(new URL("../apps-script/Code.gs", import.meta.url), "utf8");
const gas = new Function(
  src + "; return { buildColumnIndex, rowToMember, findRowIndex, handleGet, handlePost };"
)();

const HEADER = ["이름", "기수", "역할", "현재 소속", "역대 소속", "전화", "이메일", "관심사", "PIN", "최종수정일"];

function sampleSheet() {
  return [
    HEADER.slice(),
    ["홍길동", "33", "부원", "KAIST", "충북과학고 → KAIST", "010-1111-2222", "hong@example.com", "로봇/제어", "5678", "2026-01-01"],
    ["김철수", "35", "기장", "충북과학고", "충북과학고", "010-2222-3333", "kim@example.com", "임베디드", "3333", "2026-01-02"],
    ["", "", "", "", "", "", "", "", "", ""]
  ];
}

test("handleGet: 멤버 목록을 반환하고 빈 행은 거른다", () => {
  const result = gas.handleGet(sampleSheet());
  assert.equal(result.ok, true);
  assert.equal(result.members.length, 2);
  assert.equal(result.members[0].name, "홍길동");
  assert.equal(result.members[0].history, "충북과학고 → KAIST");
  assert.equal(result.members[1].cohort, "35");
});

test("handleGet: PIN은 어떤 멤버 객체에도 포함되지 않는다", () => {
  const result = gas.handleGet(sampleSheet());
  for (const m of result.members) {
    assert.equal("pin" in m, false);
    assert.equal(JSON.stringify(m).includes("5678"), false);
  }
});

test("handleGet: Date 객체 셀은 yyyy-MM-dd 문자열로 변환한다", () => {
  const sheet = sampleSheet();
  sheet[1][9] = new Date(2026, 0, 15); // 2026-01-15 (월은 0부터)
  const result = gas.handleGet(sheet);
  assert.equal(result.members[0].updatedAt, "2026-01-15");
});

test("handleGet: 열 순서가 바뀌어도 헤더명 기준으로 매핑한다", () => {
  const sheet = [
    ["기수", "이름", "PIN", "관심사"],
    ["33", "홍길동", "5678", "로봇"]
  ];
  const result = gas.handleGet(sheet);
  assert.equal(result.members[0].name, "홍길동");
  assert.equal(result.members[0].cohort, "33");
  assert.equal(result.members[0].interests, "로봇");
  assert.equal(result.members[0].org, ""); // 없는 열은 빈 문자열
});
```

- [ ] **Step 2: 테스트 실행 — 실패 확인**

Run: `node --test`
Expected: FAIL — `ENOENT ... apps-script/Code.gs` (파일 없음)

- [ ] **Step 3: Code.gs 읽기 로직 구현**

`apps-script/Code.gs` 생성:

```js
/**
 * 충북과학고 공학동아리 인명사전 API (Google Apps Script 웹앱)
 *
 * - Master 시트가 데이터 원천(SSOT)입니다.
 * - 배포/재배포 방법은 저장소 readme.md를 참고하세요.
 * - "순수 로직" 구역은 Node 테스트(tests/code-gs.test.mjs)로 검증됩니다.
 *   GAS 전역 객체(SpreadsheetApp 등)는 "GAS 글루" 구역에서만 사용하세요.
 */

var SHEET_NAME = "Master";

// 시트 헤더명 <-> API 필드명 매핑 (시트 열 이름을 바꾸면 여기도 바꿔야 합니다)
var HEADERS = {
  name: "이름",
  cohort: "기수",
  role: "역할",
  org: "현재 소속",
  history: "역대 소속",
  phone: "전화",
  email: "이메일",
  interests: "관심사",
  pin: "PIN",
  updatedAt: "최종수정일"
};

// 웹 응답에 포함하는 필드 — PIN은 절대 추가 금지
var PUBLIC_FIELDS = ["name", "cohort", "role", "org", "history", "phone", "email", "interests", "updatedAt"];

// 본인이 웹에서 수정할 수 있는 필드
var EDITABLE_FIELDS = ["org", "history", "phone", "email", "interests", "pin"];

// ===== 순수 로직 (Node 테스트 대상) =====

function formatDateValue(d) {
  var y = d.getFullYear();
  var m = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  return y + "-" + m + "-" + day;
}

function buildColumnIndex(headerRow) {
  var index = {};
  Object.keys(HEADERS).forEach(function (field) {
    var col = headerRow.indexOf(HEADERS[field]);
    if (col !== -1) index[field] = col;
  });
  return index;
}

function cellToString(value) {
  if (value instanceof Date) return formatDateValue(value);
  return String(value == null ? "" : value).trim();
}

function rowToMember(row, colIndex) {
  var member = {};
  PUBLIC_FIELDS.forEach(function (field) {
    var col = colIndex[field];
    member[field] = col === undefined ? "" : cellToString(row[col]);
  });
  return member;
}

function handleGet(sheetData) {
  var colIndex = buildColumnIndex(sheetData[0]);
  var members = sheetData.slice(1).filter(function (row) {
    return colIndex.name !== undefined && cellToString(row[colIndex.name]) !== "";
  }).map(function (row) {
    return rowToMember(row, colIndex);
  });
  return { ok: true, members: members };
}

function findRowIndex(sheetData, colIndex, name, cohort) {
  for (var i = 1; i < sheetData.length; i++) {
    if (cellToString(sheetData[i][colIndex.name]) === String(name).trim() &&
        cellToString(sheetData[i][colIndex.cohort]) === String(cohort).trim()) {
      return i;
    }
  }
  return -1;
}

function handlePost(body, sheetData, now) {
  // Task 2에서 구현
  return { response: { ok: false, error: "bad_request" } };
}
```

- [ ] **Step 4: 테스트 실행 — 통과 확인**

Run: `node --test`
Expected: PASS (4 tests)

- [ ] **Step 5: 커밋**

```bash
git add apps-script/Code.gs tests/code-gs.test.mjs
git commit -m "Add Apps Script read logic (handleGet) with Node tests"
```

---

### Task 2: Code.gs 쓰기 로직 (handlePost) — TDD

**Files:**
- Modify: `apps-script/Code.gs` (handlePost 본문 교체)
- Modify: `tests/code-gs.test.mjs` (테스트 추가)

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/code-gs.test.mjs` 끝에 추가:

```js
test("handlePost: 올바른 PIN + 필드 수정 → 행 갱신 + 최종수정일 기록", () => {
  const sheet = sampleSheet();
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { org: "삼성전자", interests: "반도체" } },
    sheet, "2026-06-12"
  );
  assert.equal(result.response.ok, true);
  assert.equal(result.response.member.org, "삼성전자");
  assert.equal(result.response.member.updatedAt, "2026-06-12");
  assert.equal(result.rowIndex, 1);
  assert.equal(result.updatedRow[3], "삼성전자");  // 현재 소속 열
  assert.equal(result.updatedRow[9], "2026-06-12"); // 최종수정일 열
  assert.equal(result.updatedRow[0], "홍길동");     // 이름은 그대로
  // 원본 sheetData는 변경하지 않는다 (쓰기는 글루 코드 책임)
  assert.equal(sheet[1][3], "KAIST");
});

test("handlePost: PIN 불일치 → wrong_pin, 행 미반환", () => {
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "0000", fields: { org: "X" } },
    sampleSheet(), "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "wrong_pin" });
  assert.equal(result.updatedRow, undefined);
});

test("handlePost: 시트의 PIN이 빈 값이면 어떤 PIN도 거부한다", () => {
  const sheet = sampleSheet();
  sheet[1][8] = "";
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "", fields: {} },
    sheet, "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "wrong_pin" });
});

test("handlePost: 없는 멤버 → not_found", () => {
  const result = gas.handlePost(
    { name: "없는사람", cohort: "99", pin: "1234", fields: {} },
    sampleSheet(), "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "not_found" });
});

test("handlePost: fields가 비면 PIN 검증만 수행 (시트 수정 없음)", () => {
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: {} },
    sampleSheet(), "2026-06-12"
  );
  assert.equal(result.response.ok, true);
  assert.equal(result.response.member.name, "홍길동");
  assert.equal("pin" in result.response.member, false);
  assert.equal(result.updatedRow, undefined);
});

test("handlePost: 수정 불가 필드(name 등) 시도 → field_not_editable", () => {
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { name: "변조" } },
    sampleSheet(), "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "field_not_editable" });
});

test("handlePost: PIN 변경 가능, 빈 PIN으로 변경은 거부", () => {
  const ok = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { pin: "9999" } },
    sampleSheet(), "2026-06-12"
  );
  assert.equal(ok.response.ok, true);
  assert.equal(ok.updatedRow[8], "9999");

  const bad = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { pin: "  " } },
    sampleSheet(), "2026-06-12"
  );
  assert.deepEqual(bad.response, { ok: false, error: "bad_request" });
});

test("handlePost: 필수 키 누락 → bad_request", () => {
  for (const body of [null, {}, { name: "홍길동" }, { name: "홍길동", cohort: "33" }]) {
    const result = gas.handlePost(body, sampleSheet(), "2026-06-12");
    assert.deepEqual(result.response, { ok: false, error: "bad_request" });
  }
});

test("handlePost: 동명이인은 기수로 구분한다", () => {
  const sheet = sampleSheet();
  sheet.push(["홍길동", "36", "부원", "충북과학고", "충북과학고", "010-9999-8888", "", "AI", "8888", ""]);
  const result = gas.handlePost(
    { name: "홍길동", cohort: "36", pin: "8888", fields: { interests: "AI/ML" } },
    sheet, "2026-06-12"
  );
  assert.equal(result.response.ok, true);
  assert.equal(result.rowIndex, 4);
});
```

- [ ] **Step 2: 테스트 실행 — 새 테스트 실패 확인**

Run: `node --test`
Expected: Task 1의 4개 PASS, 새 테스트 다수 FAIL (handlePost가 스텁이므로)

- [ ] **Step 3: handlePost 구현**

`apps-script/Code.gs`의 스텁을 교체:

```js
function handlePost(body, sheetData, now) {
  if (!body || !body.name || !body.cohort || body.pin == null) {
    return { response: { ok: false, error: "bad_request" } };
  }

  var colIndex = buildColumnIndex(sheetData[0]);
  var rowIndex = findRowIndex(sheetData, colIndex, body.name, body.cohort);
  if (rowIndex === -1) {
    return { response: { ok: false, error: "not_found" } };
  }

  var row = sheetData[rowIndex].slice();
  var storedPin = cellToString(row[colIndex.pin]);
  if (storedPin === "" || storedPin !== String(body.pin).trim()) {
    return { response: { ok: false, error: "wrong_pin" } };
  }

  var fields = body.fields || {};
  var fieldKeys = Object.keys(fields);

  for (var i = 0; i < fieldKeys.length; i++) {
    if (EDITABLE_FIELDS.indexOf(fieldKeys[i]) === -1) {
      return { response: { ok: false, error: "field_not_editable" } };
    }
  }
  // PIN을 빈 값으로 바꾸면 영영 수정 불가가 되므로 거부
  if (fields.pin != null && String(fields.pin).trim() === "") {
    return { response: { ok: false, error: "bad_request" } };
  }

  if (fieldKeys.length === 0) {
    return { response: { ok: true, member: rowToMember(row, colIndex) } };
  }

  fieldKeys.forEach(function (field) {
    row[colIndex[field]] = String(fields[field]).trim();
  });
  row[colIndex.updatedAt] = now;

  return {
    response: { ok: true, member: rowToMember(row, colIndex) },
    updatedRow: row,
    rowIndex: rowIndex
  };
}
```

- [ ] **Step 4: 테스트 실행 — 전체 통과 확인**

Run: `node --test`
Expected: PASS (13 tests)

- [ ] **Step 5: 커밋**

```bash
git add apps-script/Code.gs tests/code-gs.test.mjs
git commit -m "Add Apps Script write logic (handlePost) with PIN auth"
```

---

### Task 3: Code.gs GAS 글루 코드 (doGet/doPost)

**Files:**
- Modify: `apps-script/Code.gs` (파일 끝에 추가)

글루는 시트 I/O와 락만 담당 — 로직은 전부 Task 1·2의 순수 함수로 검증 완료. GAS 환경에서만 도는 코드라 로컬 단위 테스트는 없고, 구문 검사 + (배포 후) curl 검증.

- [ ] **Step 1: 글루 코드 추가**

`apps-script/Code.gs` 끝에 추가:

```js
// ===== GAS 글루 코드 (Apps Script 환경 전용 — Node 테스트 제외 구역) =====

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function doGet() {
  return jsonOutput_(handleGet(readSheet_().getDataRange().getValues()));
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: "bad_request" });
  }

  // 동시 쓰기 방지 (읽기-수정-쓰기 사이 끼어들기 차단)
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = readSheet_();
    var result = handlePost(body, sheet.getDataRange().getValues(), formatDateValue(new Date()));
    if (result.updatedRow) {
      sheet.getRange(result.rowIndex + 1, 1, 1, result.updatedRow.length)
        .setValues([result.updatedRow]);
    }
    return jsonOutput_(result.response);
  } finally {
    lock.releaseLock();
  }
}
```

- [ ] **Step 2: 구문 검사 + 기존 테스트 회귀 확인**

Run: `node -e "new Function(require('fs').readFileSync('apps-script/Code.gs','utf8'))" && node --test`
Expected: 구문 에러 없음, 13 tests PASS

- [ ] **Step 3: 커밋**

```bash
git add apps-script/Code.gs
git commit -m "Add GAS glue code (doGet/doPost) with script lock"
```

---

### Task 4: mock 서버 + config.js 전환

**Files:**
- Create: `tests/mock-server.mjs`
- Modify: `config.js`

- [ ] **Step 1: mock 서버 작성**

`tests/mock-server.mjs` 생성 — Code.gs의 순수 로직을 그대로 재사용해 로컬에서 실제 API와 동일하게 동작:

```js
// 로컬 개발용 mock API 서버.
// 사용법: node tests/mock-server.mjs  →  config.js의 API_URL을 http://localhost:8788 로 임시 변경
import { createServer } from "node:http";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("../apps-script/Code.gs", import.meta.url), "utf8");
const gas = new Function(src + "; return { handleGet, handlePost };")();

const sheetData = [
  ["이름", "기수", "역할", "현재 소속", "역대 소속", "전화", "이메일", "관심사", "PIN", "최종수정일"],
  ["홍길동", "33", "부원", "KAIST", "충북과학고 → KAIST", "010-1111-2222", "hong@example.com", "로봇/제어", "5678", "2026-01-01"],
  ["김철수", "35", "기장", "충북과학고", "충북과학고", "010-2222-3333", "kim@example.com", "임베디드/키오스크", "3333", "2026-01-02"],
  ["이영희", "31", "부원", "서울대학교/네이버", "충북과학고 → 서울대 → 네이버", "010-5555-6666", "lee@example.com", "백엔드, 인프라", "6666", "2024-03-01"]
];

const today = () => new Date().toISOString().slice(0, 10);

createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Content-Type", "application/json; charset=utf-8");

  if (req.method === "GET") {
    res.end(JSON.stringify(gas.handleGet(sheetData)));
    return;
  }

  let buf = "";
  req.on("data", (c) => (buf += c));
  req.on("end", () => {
    let body = null;
    try { body = JSON.parse(buf); } catch {}
    const result = gas.handlePost(body, sheetData, today());
    if (result.updatedRow) sheetData[result.rowIndex] = result.updatedRow;
    res.end(JSON.stringify(result.response));
  });
}).listen(8788, () => console.log("mock API: http://localhost:8788"));
```

- [ ] **Step 2: mock 서버 동작 확인**

Run (백그라운드로 띄운 뒤):
```bash
node tests/mock-server.mjs &
sleep 1
curl -s http://localhost:8788 | head -c 200
curl -s -X POST http://localhost:8788 -d '{"name":"홍길동","cohort":"33","pin":"5678","fields":{}}'
curl -s -X POST http://localhost:8788 -d '{"name":"홍길동","cohort":"33","pin":"0000","fields":{}}'
kill %1
```
Expected: 첫 응답에 `"ok":true,"members":[...`(pin 미포함), 둘째 `{"ok":true,"member":...}`, 셋째 `{"ok":false,"error":"wrong_pin"}`

- [ ] **Step 3: config.js 교체**

`config.js` 전체를 다음으로 교체:

```js
// Apps Script 웹앱 배포 URL (재배포 방법은 readme.md 참고)
// 로컬 테스트 시에는 "http://localhost:8788" 로 임시 변경 (커밋 금지!)
const API_URL = "https://script.google.com/macros/s/배포후_여기에_붙여넣기/exec";

// 신규 부원 등록용 구글 폼
const SURVEY_URL = "https://forms.gle/BkaXfCUPz3hYhc9F7";
```

- [ ] **Step 4: 커밋**

```bash
git add tests/mock-server.mjs config.js
git commit -m "Add mock API server; switch config to Apps Script API_URL"
```

---

### Task 5: index.html 읽기 경로 JSON 전환 + 신규 필드 표시

**Files:**
- Modify: `index.html`

- [ ] **Step 1: 캐시 키 버전업 (init 함수)**

`init()` 안의 `localStorage.getItem("memberData")`를 `localStorage.getItem("memberData_v2")`로 변경 (옛 CSV 포맷 캐시와 충돌 방지).

- [ ] **Step 2: loadMembers를 JSON API 호출로 교체**

기존 `loadMembers` 함수 전체를 교체:

```js
async function loadMembers() {
    try {
        const response = await fetch(API_URL + "?t=" + Date.now());
        const data = await response.json();
        if (!data.ok) throw new Error(data.error || "api_error");
        const newMembers = data.members;

        if (JSON.stringify(newMembers) !== JSON.stringify(allMembers)) {
            allMembers = newMembers;
            localStorage.setItem("memberData_v2", JSON.stringify(allMembers));
            renderMembers(allMembers);
        }
    } catch (error) {
        console.error(error);
        if (allMembers.length > 0) {
            showToast("새 데이터를 불러오지 못해 저장된 목록을 표시합니다.");
        } else {
            showToast("데이터를 불러오지 못했습니다. 잠시 후 다시 시도해주세요.");
        }
    }
}
```

- [ ] **Step 3: 이스케이프 헬퍼 추가**

`getBadgeStyle` 함수 위에 추가 (이제 사용자 입력이 화면에 직접 들어가므로):

```js
function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, c =>
        ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// onclick="..." 속성 안의 JS 문자열 인자용
function jsArg(s) {
    return String(s ?? "").replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/"/g, "&quot;");
}
```

- [ ] **Step 4: renderMembers에 역대 소속·최종수정일·수정 버튼 추가**

`renderMembers`의 카드 템플릿(`return \`...\``) 직전에 추가:

```js
const historyHtml = m.history
    ? `<div class="text-xs text-gray-400 mt-1.5 flex items-start gap-1.5">
         <i class="fas fa-route mt-0.5"></i><span>${escapeHtml(m.history)}</span>
       </div>`
    : '';

const updatedHtml = m.updatedAt
    ? `<span class="text-[11px] text-gray-300">최종수정 ${escapeHtml(m.updatedAt)}</span>`
    : '<span></span>';

const editBtn = `<button onclick="openEditModal('${jsArg(m.name)}', '${jsArg(m.cohort)}')"
    class="flex-1 bg-white border border-gray-200 text-gray-700 py-2.5 rounded-xl text-center hover:bg-gray-50 transition font-medium text-sm flex items-center justify-center gap-2 cursor-pointer active:bg-gray-100">
    <i class="fas fa-pen text-gray-500"></i> 수정
   </button>`;
```

카드 템플릿에서:
1. `<div class="flex flex-wrap items-center">${orgHtml}</div>` 바로 아래에 `${historyHtml}` 삽입
2. 버튼 줄 `${phoneBtn} ${emailBtn}`을 `${phoneBtn} ${emailBtn} ${editBtn}`으로
3. 버튼 줄 div 바로 아래(카드 닫는 태그 직전)에 `<div class="flex justify-end mt-2">${updatedHtml}</div>` 추가

- [ ] **Step 5: 검색에 역대 소속 포함**

검색 이벤트 리스너의 filter 조건에 추가:

```js
(m.history && m.history.toLowerCase().includes(term)) ||
```

- [ ] **Step 6: mock 서버로 수동 검증**

```bash
node tests/mock-server.mjs &
# config.js의 API_URL을 "http://localhost:8788" 로 임시 변경
python3 -m http.server 8080 &
```
브라우저(또는 curl)로 `http://localhost:8080` 확인:
- 멤버 3명 카드 표시, 역대 소속 줄·최종수정일 표시
- 검색에 "서울대" 입력 → 이영희만 표시 (역대 소속 검색)
- 확인 후: **config.js를 원복**, 서버 종료 (`kill %1 %2`)

- [ ] **Step 7: 커밋**

```bash
git add index.html
git commit -m "Switch read path to JSON API; show career history and last-updated"
```

---

### Task 6: index.html 수정 모달

**Files:**
- Modify: `index.html`

- [ ] **Step 1: 모달 HTML 추가**

`<script src="config.js"></script>` 바로 위에 추가:

```html
<!-- 정보 수정 모달 (z-[60]: FAB(z-50)보다 위. 토스트는 z-index:100으로 모달 위에 보이게 함 — 아래 Step 1b 참고) -->
<div id="edit-modal" class="hidden fixed inset-0 z-[60] items-end sm:items-center justify-center bg-black/40 px-4 pb-4 sm:pb-0">
    <div class="bg-white w-full max-w-md rounded-2xl p-6 shadow-2xl max-h-[85vh] overflow-y-auto">
        <div class="flex justify-between items-center mb-4">
            <h2 id="edit-modal-title" class="text-lg font-bold text-gray-900"></h2>
            <button onclick="closeEditModal()" class="text-gray-400 hover:text-gray-600 p-1">
                <i class="fas fa-times text-lg"></i>
            </button>
        </div>

        <!-- 1단계: PIN 확인 -->
        <div id="edit-step-pin">
            <p class="text-sm text-gray-500 mb-3">본인 확인을 위해 PIN을 입력해주세요.<br>(기본값: 전화번호 뒤 4자리)</p>
            <input id="edit-pin-input" type="password" inputmode="numeric" maxlength="8" autocomplete="off"
                   class="w-full border border-gray-200 rounded-xl px-4 py-3 mb-4 focus:outline-none focus:ring-2 focus:ring-blue-500"
                   placeholder="PIN">
            <button id="edit-pin-submit" onclick="verifyPin()"
                    class="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 disabled:opacity-50">
                확인
            </button>
        </div>

        <!-- 2단계: 정보 수정 폼 -->
        <div id="edit-step-form" class="hidden space-y-3">
            <label class="block text-sm font-medium text-gray-600">현재 소속
                <input id="edit-org" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="예: KAIST / 삼성전자">
            </label>
            <label class="block text-sm font-medium text-gray-600">역대 소속
                <input id="edit-history" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="예: 충북과학고 → KAIST">
            </label>
            <label class="block text-sm font-medium text-gray-600">전화번호
                <input id="edit-phone" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="010-0000-0000">
            </label>
            <label class="block text-sm font-medium text-gray-600">이메일
                <input id="edit-email" type="email" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500">
            </label>
            <label class="block text-sm font-medium text-gray-600">관심사 <span class="font-normal text-gray-400">(쉼표나 / 로 구분)</span>
                <input id="edit-interests" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="예: 로봇/제어, 임베디드">
            </label>
            <label class="block text-sm font-medium text-gray-600">새 PIN <span class="font-normal text-gray-400">(바꿀 때만 입력)</span>
                <input id="edit-newpin" type="password" inputmode="numeric" maxlength="8" autocomplete="off" class="mt-1 w-full border border-gray-200 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500">
            </label>
            <button id="edit-save" onclick="submitEdit()"
                    class="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 disabled:opacity-50 mt-2">
                저장
            </button>
        </div>
    </div>
</div>
```

- [ ] **Step 1b: 토스트 z-index 상향**

`index.html`의 `<style>` 안 `#toast` 규칙에서 `z-index: 50;`을 `z-index: 100;`으로 변경 (모달이 열린 상태에서 PIN 오류 토스트가 모달 뒤에 가려지지 않도록).

- [ ] **Step 2: 모달 JS 추가**

`init();` 호출 직전에 추가:

```js
// ===== 정보 수정 모달 =====
let editTarget = null; // { name, cohort, pin }

async function postApi(payload) {
    const response = await fetch(API_URL, {
        method: "POST",
        // Content-Type을 text/plain으로 보내면 CORS preflight 없이 Apps Script로 전송 가능
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload)
    });
    return response.json();
}

window.openEditModal = function(name, cohort) {
    editTarget = { name, cohort, pin: null };
    document.getElementById("edit-modal-title").innerText = `${name} (${cohort}기) 정보 수정`;
    document.getElementById("edit-pin-input").value = "";
    document.getElementById("edit-step-pin").classList.remove("hidden");
    document.getElementById("edit-step-form").classList.add("hidden");
    const modal = document.getElementById("edit-modal");
    modal.classList.remove("hidden");
    modal.classList.add("flex");
    document.getElementById("edit-pin-input").focus();
}

window.closeEditModal = function() {
    const modal = document.getElementById("edit-modal");
    modal.classList.add("hidden");
    modal.classList.remove("flex");
    editTarget = null;
}

window.verifyPin = async function() {
    const pin = document.getElementById("edit-pin-input").value.trim();
    if (!pin) { showToast("PIN을 입력해주세요."); return; }
    const btn = document.getElementById("edit-pin-submit");
    btn.disabled = true;
    try {
        const data = await postApi({ name: editTarget.name, cohort: editTarget.cohort, pin, fields: {} });
        if (!data.ok) {
            showToast(data.error === "wrong_pin" ? "PIN이 일치하지 않습니다." : "확인에 실패했습니다. 기장에게 문의해주세요.");
            return;
        }
        editTarget.pin = pin;
        document.getElementById("edit-org").value = data.member.org || "";
        document.getElementById("edit-history").value = data.member.history || "";
        document.getElementById("edit-phone").value = data.member.phone || "";
        document.getElementById("edit-email").value = data.member.email || "";
        document.getElementById("edit-interests").value = data.member.interests || "";
        document.getElementById("edit-newpin").value = "";
        document.getElementById("edit-step-pin").classList.add("hidden");
        document.getElementById("edit-step-form").classList.remove("hidden");
    } catch (err) {
        console.error(err);
        showToast("네트워크 오류가 발생했습니다. 다시 시도해주세요.");
    } finally {
        btn.disabled = false;
    }
}

window.submitEdit = async function() {
    const fields = {
        org: document.getElementById("edit-org").value.trim(),
        history: document.getElementById("edit-history").value.trim(),
        phone: document.getElementById("edit-phone").value.trim(),
        email: document.getElementById("edit-email").value.trim(),
        interests: document.getElementById("edit-interests").value.trim()
    };
    const newPin = document.getElementById("edit-newpin").value.trim();
    if (newPin) fields.pin = newPin;

    const btn = document.getElementById("edit-save");
    btn.disabled = true;
    try {
        const data = await postApi({ name: editTarget.name, cohort: editTarget.cohort, pin: editTarget.pin, fields });
        if (!data.ok) {
            showToast("저장에 실패했습니다. 다시 시도해주세요.");
            return;
        }
        const idx = allMembers.findIndex(m => m.name === editTarget.name && m.cohort === editTarget.cohort);
        if (idx !== -1) allMembers[idx] = data.member;
        localStorage.setItem("memberData_v2", JSON.stringify(allMembers));
        renderMembers(allMembers);
        closeEditModal();
        showToast("정보가 수정되었습니다!");
    } catch (err) {
        console.error(err);
        showToast("네트워크 오류가 발생했습니다. 다시 시도해주세요.");
    } finally {
        btn.disabled = false;
    }
}

// 모달 바깥 클릭 시 닫기
document.getElementById("edit-modal").addEventListener("click", (e) => {
    if (e.target.id === "edit-modal") closeEditModal();
});
```

- [ ] **Step 3: mock 서버로 수동 검증**

Task 5 Step 6과 동일하게 mock 서버 + 로컬 서버 가동, config.js 임시 변경 후:
- 카드 "수정" 클릭 → 모달 → 틀린 PIN(0000) → "PIN이 일치하지 않습니다" 토스트
- 올바른 PIN(홍길동: 5678) → 폼에 기존 값 채워짐 → 소속을 "삼성전자"로 변경 → 저장 → 카드 즉시 갱신 + 최종수정일 오늘 날짜
- 새로고침 → 변경값 유지 (mock 서버 메모리에 반영됨)
- 확인 후 config.js 원복, 서버 종료

- [ ] **Step 4: 커밋**

```bash
git add index.html
git commit -m "Add self-service edit modal with PIN verification"
```

---

### Task 7: readme.md 전면 갱신 (인수인계 매뉴얼)

**Files:**
- Modify: `readme.md` (전체 교체)

- [ ] **Step 1: readme.md 전체를 다음 내용으로 교체**

````markdown
# 📂 충북과학고 공학동아리 인명사전 (Digital Handbook)

동아리 선후배 네트워킹을 위한 웹 인명사전입니다.
서버 비용 0원 — **구글 스프레드시트**(데이터) + **Apps Script**(API) + **GitHub Pages**(화면)로 운영됩니다.

## 🚀 바로가기
- **웹사이트:** https://cbsh-genie.github.io/CBSH_genie_DB/
- **관리자용 시트 (Master):** [link](https://docs.google.com/spreadsheets/d/1_RrJyaNhBlk7uDH0w_-FWXv-AZnmWPcnYyXqwmZMEkM/edit)
- 관리 계정: CBSHgenie2020 구글 계정 (비밀번호는 네이버 카페 기장단 게시판 참고)

---

## 🛠️ 운영 원리 (3줄 요약)
1. 모든 데이터는 구글 시트 `Master` 한 곳에만 있습니다. (기장단은 시트를 직접 고쳐도 됩니다)
2. 시트에 붙은 **Apps Script**가 무료 API 역할을 해서, 웹사이트가 즉시 읽고 씁니다.
3. 부원 각자는 웹사이트에서 **본인 PIN**(기본: 전화번호 뒤 4자리)으로 자기 정보를 직접 수정합니다.

## 📋 시트 열 구성 (열 이름을 바꾸면 안 됩니다!)

| 이름 | 기수 | 역할 | 현재 소속 | 역대 소속 | 전화 | 이메일 | 관심사 | PIN | 최종수정일 |

- **PIN**: 웹에는 절대 노출되지 않습니다. 본인 수정 시 본인 확인용.
- **최종수정일**: 웹에서 수정하면 자동 기록. 시트에서 직접 고쳤다면 **수동으로 오늘 날짜를 적어주세요.**
- 열 이름을 바꾸고 싶다면 `apps-script/Code.gs`의 `HEADERS`도 같이 바꾸고 재배포해야 합니다.

---

## 📅 매년 해야 할 일 (기장/총무 필독)

### 1. 신입 기수 추가
1. 구글 폼(웹사이트 우하단 "정보 등록/수정" 버튼과 같은 폼)으로 신입생 정보를 걷습니다.
2. `Master` 시트에 행으로 붙여넣습니다.
3. **PIN 열에 각자 전화번호 뒤 4자리를 넣어줍니다.** (전화번호가 없으면 아무 4자리를 넣고 개인톡으로 알려주기)
4. 끝. 웹사이트에 바로 반영됩니다.

### 2. 졸업생 정보 업데이트
- 각자 웹사이트에서 카드의 **[수정]** 버튼 → PIN 입력 → 직접 수정하라고 전체방에 공지하면 됩니다.
- 최종수정일이 2년 이상 지난 사람에게는 업데이트 부탁 연락을 돌립니다. (기장 톡방 합의 사항)

### 3. PIN을 까먹었다는 연락이 오면
- `Master` 시트에서 그 사람 PIN 열을 직접 확인하거나 새 값으로 바꿔서 알려주면 됩니다.

---

## 🔧 Apps Script 배포/재배포 방법 (처음 또는 코드 변경 시)

> 한 번 배포해두면 평소에는 손댈 일이 없습니다. `Code.gs`를 고쳤을 때만 필요합니다.

1. CBSHgenie2020 계정으로 `Master` 시트를 엽니다.
2. 메뉴 **확장 프로그램 → Apps Script** 클릭.
3. 편집기에 이 저장소의 `apps-script/Code.gs` 내용을 전부 복사해 붙여넣고 저장(💾).
4. 우상단 **배포 → 새 배포** 클릭.
   - 유형: **웹 앱**
   - 실행 계정: **나(CBSHgenie2020)**
   - 액세스 권한: **모든 사용자**
5. **배포** 버튼 → 나오는 `https://script.google.com/macros/s/.../exec` 주소를 복사.
6. 이 저장소 `config.js`의 `API_URL` 값을 그 주소로 바꿔서 커밋/푸시.
7. 1~2분 뒤 웹사이트 새로고침해서 명단이 뜨면 성공!

> ⚠️ **이미 배포된 적이 있다면**: "새 배포" 대신 **배포 → 배포 관리 → ✏️ 수정 → 버전: 새 버전**으로 올리세요. 이러면 주소가 안 바뀌어서 `config.js`를 안 고쳐도 됩니다.

---

## ⚠️ 문제 해결 (Troubleshooting)

### Q. 웹사이트에 명단이 안 떠요!
1. `Master` 시트 1행의 열 이름이 위 표와 정확히 같은지 확인.
2. Apps Script 배포가 살아있는지 확인: `config.js`의 `API_URL`을 브라우저 주소창에 직접 입력했을 때 `{"ok":true,...}` 글자가 보여야 정상.
3. 안 보인다면 위의 **재배포 방법**대로 다시 배포.

### Q. 수정하려는데 "PIN이 일치하지 않습니다"래요.
- `Master` 시트에서 그 사람의 PIN 열 값을 확인해서 알려주세요. PIN 칸이 비어 있으면 수정이 막히니 값을 채워주세요.

### Q. '정보 등록' 버튼(신규 등록 폼) 링크를 바꾸고 싶어요.
- `config.js`의 `SURVEY_URL` 값을 새 구글 폼 주소로 바꾸면 됩니다.

---

## 💻 개발자용 정보

- **`index.html`**: 화면 전체 (검색, 카드, 수정 모달). Vanilla JS + Tailwind CDN.
- **`config.js`**: `API_URL`(Apps Script 배포 주소), `SURVEY_URL`(신규 등록 폼).
- **`apps-script/Code.gs`**: API 코드. 윗부분(순수 로직)은 Node 테스트로 검증, 아랫부분(GAS 글루)은 배포해야 동작.
- **테스트:** `node --test`
- **로컬에서 화면 확인:**
  ```bash
  node tests/mock-server.mjs &          # 가짜 API (포트 8788)
  python3 -m http.server 8080 &         # 정적 서버
  # config.js의 API_URL을 http://localhost:8788 로 임시 변경 (커밋 금지)
  ```

---
Made with by 안연수 (33기)
````

- [ ] **Step 2: 테스트 회귀 확인 + 커밋**

Run: `node --test`
Expected: PASS

```bash
git add readme.md
git commit -m "Rewrite readme as handover manual for Apps Script architecture"
```

---

### Task 8: 최종 검증 + 푸시

- [ ] **Step 1: 전체 테스트**

Run: `node --test && node -e "new Function(require('fs').readFileSync('apps-script/Code.gs','utf8'))"`
Expected: 13 tests PASS, 구문 에러 없음

- [ ] **Step 2: 풀스택 스모크 (mock)**

mock 서버 + 정적 서버를 띄우고 config.js를 localhost로 임시 변경한 뒤 curl 또는 브라우저로:
- 목록 로드 / 검색(역대 소속 포함) / PIN 오류 / 정상 수정 / 수정 후 카드 갱신
확인 후 config.js 원복 (git diff로 원복 확인), 서버 종료.

- [ ] **Step 3: 푸시**

```bash
git push origin main
```

> 푸시하면 GitHub Pages가 갱신되지만, `config.js`의 API_URL이 아직 placeholder이므로
> 사이트는 "데이터를 불러오지 못했습니다"가 뜨는 상태가 된다. 이는 의도된 중간 상태 —
> 최종 사용자(연수)가 아래 "사람이 해야 할 일"을 마치면 완성된다.

## 사람이 해야 할 일 (코드 외 — 연수 담당)

1. `Master` 시트에 **역대 소속 / PIN / 최종수정일** 열 추가, PIN 일괄 세팅 (전화 뒤 4자리)
2. readme의 "Apps Script 배포 방법"대로 배포 → `config.js`에 실제 URL 커밋/푸시
3. 동작 확인 후 기존 "파일 → 공유 → 웹에 게시"(CSV) 중지, `Web` 시트 삭제
