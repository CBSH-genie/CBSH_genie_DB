// 로컬 개발용 mock API 서버.
// 사용법: node tests/mock-server.mjs  →  config.js의 API_URL을 http://localhost:8788 로 임시 변경
import { createServer } from "node:http";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("../apps-script/Code.gs", import.meta.url), "utf8");
const gas = new Function(src + "; return { handleGet, handlePost };")();

const sheetData = [
  ["이름", "기수", "역할", "현재 소속", "역대 소속", "전화", "이메일", "관심사", "PIN", "최종수정일"],
  ["홍길동", "33", "부원", "KAIST", "충북과학고 → KAIST", "010-1111-2222", "hong@example.com", "로봇/제어", "5678", "2026-01-01"],
  ["김철수", "35", "기장", "충북과학고", "충북과학고", "010-2222-4444", "kim@example.com", "임베디드/키오스크", "3333", "2026-01-02"],
  ["이영희", "31", "부원", "서울대학교/네이버", "충북과학고 → 서울대 → 네이버", "010-5555-7777", "lee@example.com", "백엔드, 인프라", "6666", "2024-03-01"]
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
