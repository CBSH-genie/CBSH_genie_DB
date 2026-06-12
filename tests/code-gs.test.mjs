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
