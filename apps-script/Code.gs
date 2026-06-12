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
  if (!sheetData || sheetData.length === 0) return { ok: true, members: [] };
  var colIndex = buildColumnIndex(sheetData[0]);
  if (colIndex.name === undefined) return { ok: false, error: "missing_name_column" };
  var members = sheetData.slice(1).filter(function (row) {
    return cellToString(row[colIndex.name]) !== "";
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

function pinsMatch(storedPin, inputPin) {
  if (storedPin === "" || inputPin === "") return false;
  if (storedPin === inputPin) return true;
  // 시트 셀이 숫자 서식이면 앞자리 0이 사라진다(예: 0421 → 421).
  // 둘 다 숫자로만 구성된 경우 수치로도 비교해 잠금 사고를 막는다.
  return /^[0-9]+$/.test(storedPin) && /^[0-9]+$/.test(inputPin) &&
         Number(storedPin) === Number(inputPin);
}

// handleGet과 달리 {response, updatedRow, rowIndex} 형태를 반환한다 —
// 시트 쓰기는 GAS 글루(doPost)의 책임이고, 순수 로직은 결과만 계산하기 때문.
function handlePost(body, sheetData, now) {
  if (!body || !body.name || !body.cohort || body.pin == null) {
    return { response: { ok: false, error: "bad_request" } };
  }
  if (!sheetData || sheetData.length === 0) {
    return { response: { ok: false, error: "bad_request" } };
  }

  var colIndex = buildColumnIndex(sheetData[0]);
  var rowIndex = findRowIndex(sheetData, colIndex, body.name, body.cohort);
  if (rowIndex === -1) {
    return { response: { ok: false, error: "not_found" } };
  }

  var row = sheetData[rowIndex].slice();
  var storedPin = cellToString(row[colIndex.pin]);
  if (!pinsMatch(storedPin, String(body.pin).trim())) {
    return { response: { ok: false, error: "wrong_pin" } };
  }

  var fields = body.fields || {};
  var fieldKeys = Object.keys(fields);

  for (var i = 0; i < fieldKeys.length; i++) {
    if (EDITABLE_FIELDS.indexOf(fieldKeys[i]) === -1) {
      return { response: { ok: false, error: "field_not_editable" } };
    }
  }
  // PIN을 빈 값(null 포함)으로 바꾸면 영영 수정 불가가 되므로 거부
  if (fieldKeys.indexOf("pin") !== -1 &&
      (fields.pin == null || String(fields.pin).trim() === "")) {
    return { response: { ok: false, error: "bad_request" } };
  }

  if (fieldKeys.length === 0) {
    return { response: { ok: true, member: rowToMember(row, colIndex) } };
  }

  // 대상 열이 시트에 없으면 row[undefined]에 쓰여 조용히 유실되므로 명시적으로 거부
  for (var j = 0; j < fieldKeys.length; j++) {
    if (colIndex[fieldKeys[j]] === undefined) {
      return { response: { ok: false, error: "missing_column" } };
    }
  }
  if (colIndex.updatedAt === undefined) {
    return { response: { ok: false, error: "missing_column" } };
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
