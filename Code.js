// ═══════════════════════════════════════════════════════════════
// 빌리지 유입로그 — GAS 코드
// 새 스프레드시트의 Apps Script에 이 코드 전체를 붙여넣기
// ═══════════════════════════════════════════════════════════════

var INFLOW_SHEET = "유입로그";
var ANALYSIS_SHEET = "유입분석";

// ─── API 엔드포인트 ─────────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";
  var callback = (e && e.parameter && e.parameter.callback) || "";
  var result;

  // action 없으면 웹폼 페이지 서빙
  if (!action) {
    return HtmlService.createHtmlOutputFromFile("inflow-page")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0, user-scalable=no");
  }



  try {
    switch (action) {
      case "list":
        result = doList_(e.parameter);
        break;
      case "get":
        result = doGetRow_(e.parameter);
        break;
      case "create":
        result = doCreate_(e.parameter);
        break;
      case "update":
        result = doUpdate_(e.parameter);
        break;

      default:
        result = { success: false, error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var callback = "";
  var result;

  try {
    var params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter || {};
    }
    callback = params.callback || "";
    var action = params.action || "";

    switch (action) {
      case "create":
        result = doCreate_(params);
        break;
      case "update":
        result = doUpdate_(params);
        break;

      default:
        result = { success: false, error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── 목록 조회 ──────────────────────────────────────────────

function doList_(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFLOW_SHEET);
  if (!sheet) return { success: false, error: "유입로그 시트 없음" };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      row: i + 2,
      일시: data[i][0],
      유입경로: data[i][1],
      고객유형: data[i][2],
      문의장비: data[i][3],
      예약여부: data[i][4],
      예약금액: data[i][5],
      미예약사유: data[i][6],
      메모: data[i][7]
    });
  }

  // 최신순 정렬
  rows.reverse();

  // 필터
  var filter = params.filter || "";
  if (filter === "pending") {
    rows = rows.filter(function(r) { return r.예약여부 === "N" && !r.미예약사유; });
  } else if (filter === "booked") {
    rows = rows.filter(function(r) { return r.예약여부 === "Y"; });
  } else if (filter === "lost") {
    rows = rows.filter(function(r) { return r.예약여부 === "N" && r.미예약사유; });
  }

  // 검색
  var keyword = params.keyword || "";
  if (keyword) {
    keyword = keyword.toLowerCase();
    rows = rows.filter(function(r) {
      return r.문의장비.toLowerCase().indexOf(keyword) >= 0 ||
             r.메모.toLowerCase().indexOf(keyword) >= 0;
    });
  }

  return { success: true, data: rows, total: rows.length };
}

// ─── 단건 조회 ──────────────────────────────────────────────

function doGetRow_(params) {
  var rowNum = parseInt(params.row, 10);
  if (!rowNum || rowNum < 2) return { success: false, error: "잘못된 행 번호" };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFLOW_SHEET);
  if (!sheet) return { success: false, error: "유입로그 시트 없음" };

  var data = sheet.getRange(rowNum, 1, 1, 8).getDisplayValues()[0];
  return {
    success: true,
    data: {
      row: rowNum,
      일시: data[0],
      유입경로: data[1],
      고객유형: data[2],
      문의장비: data[3],
      예약여부: data[4],
      예약금액: data[5],
      미예약사유: data[6],
      메모: data[7]
    }
  };
}

// ─── 새 건 등록 ─────────────────────────────────────────────

function doCreate_(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFLOW_SHEET);
  if (!sheet) return { success: false, error: "유입로그 시트 없음" };

  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
  var channel = params.유입경로 || "";
  var custType = params.고객유형 || "";
  var equipment = params.문의장비 || "";
  var booked = params.예약여부 || "N";
  var amount = params.예약금액 || "";
  var reason = params.미예약사유 || "";
  var memo = params.메모 || "";

  if (!channel || !custType || !equipment) {
    return { success: false, error: "필수값 누락 (유입경로, 고객유형, 문의장비)" };
  }

  sheet.appendRow([now, channel, custType, equipment, booked, amount ? Number(amount) : "", reason, memo]);

  return { success: true, message: "등록 완료" };
}

// ─── 기존 건 수정 ───────────────────────────────────────────

function doUpdate_(params) {
  var rowNum = parseInt(params.row, 10);
  if (!rowNum || rowNum < 2) return { success: false, error: "잘못된 행 번호" };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFLOW_SHEET);
  if (!sheet) return { success: false, error: "유입로그 시트 없음" };

  // 수정 가능한 필드만 업데이트
  var range = sheet.getRange(rowNum, 1, 1, 8);
  var current = range.getValues()[0];

  if (params.예약여부 !== undefined) current[4] = params.예약여부;
  if (params.예약금액 !== undefined) current[5] = params.예약금액 ? Number(params.예약금액) : "";
  if (params.미예약사유 !== undefined) current[6] = params.미예약사유;
  if (params.메모 !== undefined) current[7] = params.메모;
  if (params.유입경로 !== undefined) current[1] = params.유입경로;
  if (params.고객유형 !== undefined) current[2] = params.고객유형;
  if (params.문의장비 !== undefined) current[3] = params.문의장비;

  range.setValues([current]);

  return { success: true, message: "수정 완료" };
}





// ─── google.script.run 용 래퍼 (HTML에서 호출) ──────────────

function apiCreate(params) { return doCreate_(params); }
function apiList(params)   { return doList_(params); }
function apiGetRow(params) { return doGetRow_(params); }
function apiUpdate(params) { return doUpdate_(params); }


// ═══════════════════════════════════════════════════════════════
// 시트 초기 설정 (최초 1회 실행)
// ═══════════════════════════════════════════════════════════════

function setupInflowLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName(INFLOW_SHEET);
  if (existing) {
    SpreadsheetApp.getUi().alert("'" + INFLOW_SHEET + "' 시트가 이미 존재합니다.");
    return;
  }

  var sheet = ss.getSheets()[0];
  sheet.setName(INFLOW_SHEET);

  var headers = ["일시", "유입경로", "고객유형", "문의장비", "예약여부", "예약금액", "미예약사유", "메모"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1a73e8").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 200);

  var maxRows = 1000;

  sheet.getRange(2, 2, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["네이버검색", "인스타그램", "당근마켓", "지인소개", "기타"], true)
      .setAllowInvalid(false).build());

  sheet.getRange(2, 3, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["신규", "재방문"], true)
      .setAllowInvalid(false).build());

  sheet.getRange(2, 5, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Y", "N"], true)
      .setAllowInvalid(false).build());

  sheet.getRange(2, 7, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["가격", "장비없음", "일정안맞음", "타업체이용", "단순문의"], true)
      .setAllowInvalid(false).build());

  sheet.getRange(2, 6, maxRows, 1).setNumberFormat("#,##0");
  sheet.getRange(2, 1, maxRows, 1).setNumberFormat("yyyy-MM-dd HH:mm");

  var yRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Y"').setBackground("#e6f4ea")
    .setRanges([sheet.getRange(2, 1, maxRows, headers.length)]).build();
  var nRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2="N",$E2<>"")').setBackground("#fce8e6")
    .setRanges([sheet.getRange(2, 1, maxRows, headers.length)]).build();
  sheet.setConditionalFormatRules([yRule, nRule]);
  sheet.setTabColor("#1a73e8");

  setupAnalysis_(ss);

  SpreadsheetApp.getUi().alert("유입로그 + 유입분석 시트가 생성되었습니다.\n웹앱을 새 버전으로 배포하세요.");
}

function setupAnalysis_(ss) {
  var existing = ss.getSheetByName(ANALYSIS_SHEET);
  if (existing) ss.deleteSheet(existing);

  var as = ss.insertSheet(ANALYSIS_SHEET);
  var src = INFLOW_SHEET;

  function styleHeader(range, color) {
    range.setBackground(color).setFontColor("#ffffff")
         .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  }

  as.getRange("A1").setValue("채널별 분석").setFontSize(13).setFontWeight("bold");
  as.getRange("A2:E2").setValues([["유입경로", "유입건수", "예약건수", "전환율", "매출합계"]]);
  styleHeader(as.getRange("A2:E2"), "#1a73e8");

  var channels = ["네이버검색", "인스타그램", "당근마켓", "지인소개", "기타"];
  for (var i = 0; i < channels.length; i++) {
    var row = i + 3;
    as.getRange(row, 1).setValue(channels[i]);
    as.getRange(row, 2).setFormula('=COUNTIF(\'' + src + '\'!B:B,"' + channels[i] + '")');
    as.getRange(row, 3).setFormula('=COUNTIFS(\'' + src + '\'!B:B,"' + channels[i] + '",\'' + src + '\'!E:E,"Y")');
    as.getRange(row, 4).setFormula('=IFERROR(C' + row + '/B' + row + ',0)');
    as.getRange(row, 5).setFormula('=SUMIFS(\'' + src + '\'!F:F,\'' + src + '\'!B:B,"' + channels[i] + '",\'' + src + '\'!E:E,"Y")');
  }
  var totalRow = channels.length + 3;
  as.getRange(totalRow, 1).setValue("합계").setFontWeight("bold");
  as.getRange(totalRow, 2).setFormula("=SUM(B3:B" + (totalRow - 1) + ")");
  as.getRange(totalRow, 3).setFormula("=SUM(C3:C" + (totalRow - 1) + ")");
  as.getRange(totalRow, 4).setFormula("=IFERROR(C" + totalRow + "/B" + totalRow + ",0)");
  as.getRange(totalRow, 5).setFormula("=SUM(E3:E" + (totalRow - 1) + ")");
  as.getRange("D3:D" + totalRow).setNumberFormat("0.0%");
  as.getRange("E3:E" + totalRow).setNumberFormat("#,##0");
  as.getRange("B3:C" + totalRow).setNumberFormat("#,##0");

  var sec2Start = totalRow + 2;
  as.getRange(sec2Start, 1).setValue("미예약사유 분포").setFontSize(13).setFontWeight("bold");
  as.getRange(sec2Start + 1, 1, 1, 2).setValues([["미예약사유", "건수"]]);
  styleHeader(as.getRange(sec2Start + 1, 1, 1, 2), "#e8710a");

  var reasons = ["가격", "장비없음", "일정안맞음", "타업체이용", "단순문의"];
  for (var j = 0; j < reasons.length; j++) {
    var rRow = sec2Start + 2 + j;
    as.getRange(rRow, 1).setValue(reasons[j]);
    as.getRange(rRow, 2).setFormula('=COUNTIF(\'' + src + '\'!G:G,"' + reasons[j] + '")');
  }

  as.getRange(sec2Start, 4).setValue("인기 문의장비 TOP5").setFontSize(13).setFontWeight("bold");
  as.getRange(sec2Start + 1, 4, 1, 2).setValues([["장비명", "문의건수"]]);
  styleHeader(as.getRange(sec2Start + 1, 4, 1, 2), "#34a853");
  as.getRange(sec2Start + 2, 4).setFormula(
    '=IFERROR(QUERY(\'' + src + '\'!D:D,"SELECT D, COUNT(D) WHERE D IS NOT NULL GROUP BY D ORDER BY COUNT(D) DESC LIMIT 5 LABEL COUNT(D) \'문의건수\'"),"데이터 없음")');

  var sec4Start = sec2Start + 2 + reasons.length + 1;
  as.getRange(sec4Start, 1).setValue("고객유형 분석").setFontSize(13).setFontWeight("bold");
  as.getRange(sec4Start + 1, 1, 1, 4).setValues([["고객유형", "유입건수", "예약건수", "전환율"]]);
  styleHeader(as.getRange(sec4Start + 1, 1, 1, 4), "#9334e6");

  var custTypes = ["신규", "재방문"];
  for (var k = 0; k < custTypes.length; k++) {
    var cRow = sec4Start + 2 + k;
    as.getRange(cRow, 1).setValue(custTypes[k]);
    as.getRange(cRow, 2).setFormula('=COUNTIF(\'' + src + '\'!C:C,"' + custTypes[k] + '")');
    as.getRange(cRow, 3).setFormula('=COUNTIFS(\'' + src + '\'!C:C,"' + custTypes[k] + '",\'' + src + '\'!E:E,"Y")');
    as.getRange(cRow, 4).setFormula('=IFERROR(C' + cRow + '/B' + cRow + ',0)');
  }
  as.getRange("D" + (sec4Start + 2) + ":D" + (sec4Start + 3)).setNumberFormat("0.0%");

  as.setColumnWidth(1, 130);
  as.setColumnWidth(2, 100);
  as.setColumnWidth(3, 100);
  as.setColumnWidth(4, 130);
  as.setColumnWidth(5, 120);
  as.setTabColor("#e8710a");
}
