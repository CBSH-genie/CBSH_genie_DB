import { readFileSync } from "node:fs";
import test from "node:test";
import assert from "node:assert/strict";

// Code.gs는 GAS용 plain JS — var/function 선언만 쓰므로 new Function으로 로드 가능.
// GAS 전역(SpreadsheetApp 등)은 글루 함수 내부에서만 참조하므로 정의 시점에는 안전하다.
// 실행: node --test   (Node 18+ 필요, npm 의존성 없음)
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
  assert.equal("pin" in result.members[0], false);
  assert.equal(JSON.stringify(result.members[0]).includes("5678"), false);
});

test("handleGet: 빈 시트 데이터는 빈 목록을 반환한다", () => {
  assert.deepEqual(gas.handleGet([]), { ok: true, members: [] });
});

test("handleGet: 이름 열이 없으면 에러를 반환한다", () => {
  const result = gas.handleGet([["기수", "관심사"], ["33", "로봇"]]);
  assert.deepEqual(result, { ok: false, error: "missing_name_column" });
});

test("handleGet: 숫자 셀(기수 등)은 문자열로 변환된다", () => {
  const sheet = sampleSheet();
  sheet[1][1] = 33; // GAS getValues()는 숫자 셀을 number로 반환
  const result = gas.handleGet(sheet);
  assert.equal(result.members[0].cohort, "33");
});

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
  assert.equal(ok.updatedRow[9], "2026-06-12"); // PIN만 바꿔도 최종수정일 갱신

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

test("handlePost: pin: null 로 변경 시도 → bad_request (문자열 'null' 기록 방지)", () => {
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { pin: null } },
    sampleSheet(), "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "bad_request" });
  assert.equal(result.updatedRow, undefined);
});

test("handlePost: 시트 PIN이 숫자 서식이어도(앞 0 소실) 인증된다", () => {
  const sheet = sampleSheet();
  sheet[1][8] = 421; // 시트 셀이 숫자 서식이면 "0421"이 421(number)로 저장됨
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "0421", fields: {} },
    sheet, "2026-06-12"
  );
  assert.equal(result.response.ok, true);
});

test("handlePost: 빈 sheetData → bad_request (크래시 방지)", () => {
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: {} },
    [], "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "bad_request" });
});

test("handlePost: 수정 대상 열이 시트에 없으면 missing_column", () => {
  const sheet = [
    ["이름", "기수", "PIN"],
    ["홍길동", "33", "5678"]
  ];
  const result = gas.handlePost(
    { name: "홍길동", cohort: "33", pin: "5678", fields: { org: "X" } },
    sheet, "2026-06-12"
  );
  assert.deepEqual(result.response, { ok: false, error: "missing_column" });
});
