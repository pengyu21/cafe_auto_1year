/**
 * 설정값
 */
const SHEET_NAME = "작업로그";
const ERROR_SHEET_NAME = "카페작업리스트"; // ⭐ 에러 체크용 시트명
const RECIPIENT = "ysg@byuls.com";
const LINK = "https://docs.google.com/spreadsheets/d/1LNB7mhszGpWRPrIIh7YZz0Rcdmx9EgFp7SRRFB2A87o/edit?gid=1275153677#gid=1275153677";

const START_ROW = 3;       // 작업로그 시작행
const START_ROW_ERR = 3;   // 카페작업리스트 처리 시작행
const TIMESTAMP_COL = 8;   // 작업로그 H열 (1-indexed: A=1, B=2, ... H=8)
const CAPTCHA_CELL = "O1"; // ⭐ 프로그램이 캡챠 실패 목록을 기록하는 칸 (카페작업리스트)

/**
 * 메인 실행 함수 (트리거가 이 함수를 호출합니다)
 */
function runAllChecks() {
  checkNewRowsAndSendEmail();
  checkErrorsAndSendEmail();
  checkCaptchaAndSendEmail();
}

/**
 * 1. 작업로그 시트: H열 타임스탬프 기준 신규 행 감지 (행 삭제에 안전)
 */
function checkNewRowsAndSendEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return;

  const props = PropertiesService.getScriptProperties();
  const key = `lastSentTimestamp_${ss.getId()}_${SHEET_NAME}`;
  const lastSentStr = props.getProperty(key);

  // 첫 실행이면 "현재 시각"을 기준점으로 저장하고 종료 (폭주 방지)
  if (!lastSentStr) {
    props.setProperty(key, String(Date.now()));
    Logger.log("🔰 첫 실행: 현재 시각을 기준점으로 저장. 이후 H열 타임스탬프가 이 시각보다 미래인 행만 발송 대상.");
    return;
  }

  const lastSentMs = Number(lastSentStr);

  // 전체 데이터 한 번에 읽기 (A~H, 8열)
  const numRows = lastRow - START_ROW + 1;
  const values = sheet.getRange(START_ROW, 1, numRows, 8).getValues();

  // 신규 행 수집 (시간순으로 정렬해서 발송하기 위해)
  const newRows = [];

  values.forEach((row, i) => {
    const rowNumber = START_ROW + i;
    const colB = safeText(row[1]);
    const colD = safeText(row[3]);
    const colE = safeText(row[4]);
    const colG = safeText(row[6]);
    const colH = row[7]; // H열 = 타임스탬프 (Date 객체 또는 문자열)

    // 빈 행 스킵
    if (!colB && !colD && !colE && !colG && !colH) return;

    // H열을 Date로 파싱
    const ts = parseTimestamp(colH);
    if (!ts) {
      Logger.log(`⚠️ ${rowNumber}행: H열 타임스탬프 파싱 실패 (${colH}) - 스킵`);
      return;
    }

    const tsMs = ts.getTime();
    // 마지막 발송 시각보다 이후에 올라온 행만 대상
    if (tsMs > lastSentMs) {
      newRows.push({ rowNumber, colB, colD, colE, colG, colH, tsMs });
    }
  });

  if (newRows.length === 0) return;

  // 시간 오름차순 정렬 (오래된 것부터 발송)
  newRows.sort((a, b) => a.tsMs - b.tsMs);

  let maxTsMs = lastSentMs;

  newRows.forEach(item => {
    const subject = `카페후기 ${item.colB} ${item.colD},${item.colE} 업로드 완료`;
    const tsText = Utilities.formatDate(new Date(item.tsMs), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const body = `${item.colG} / ${tsText}\n\n${LINK}`;

    try {
      GmailApp.sendEmail(RECIPIENT, subject, body);
      Logger.log(`📧 [발송] ${item.rowNumber}행 / ${tsText} / ${item.colB}`);
      if (item.tsMs > maxTsMs) maxTsMs = item.tsMs;
    } catch (e) {
      Logger.log(`❌ ${item.rowNumber}행 발송 실패: ${e.message}`);
      // 실패해도 다음 행은 계속 시도. maxTsMs는 갱신 안 함 → 다음 실행 때 재시도
    }
  });

  // 가장 늦은 발송 성공 시각으로 갱신
  props.setProperty(key, String(maxTsMs));
}

/**
 * 타임스탬프 파싱 (Date 객체 또는 "2026-03-10 19:06:14" 형태의 문자열 처리)
 */
function parseTimestamp(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  const s = String(v).trim();
  if (!s) return null;

  // "2026-03-10 19:06:14" → "2026-03-10T19:06:14"
  let normalized = s.replace(" ", "T");
  let d = new Date(normalized);
  if (!isNaN(d.getTime())) return d;

  // 폴백: 원본 그대로 시도
  d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 2. 카페작업리스트 시트: 전수조사 방식 (기존 유지)
 */
function checkErrorsAndSendEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW_ERR) return;

  const props = PropertiesService.getScriptProperties();
  const errorLogKey = `sentErrors_${ss.getId()}`;
  let sentErrors = JSON.parse(props.getProperty(errorLogKey) || "{}");

  const headers = sheet.getRange(2, 10, 1, 4).getValues()[0];

  const numRows = lastRow - START_ROW_ERR + 1;
  const values = sheet.getRange(START_ROW_ERR, 1, numRows, 13).getValues();

  let currentErrors = {};

  values.forEach((row, i) => {
    const rowNumber = START_ROW_ERR + i;
    const colB = safeText(row[1]);

    [9, 10, 11, 12].forEach((colIdx, hIdx) => {
      const cellValue = safeText(row[colIdx]);
      const headerTitle = safeText(headers[hIdx]) || `${colIdx + 1}열`;

      // 캡챠 에러는 checkCaptchaAndSendEmail(O1)에서 따로 '캡챠' 메일로 보냄 (중복 발송 방지)
      if (cellValue.includes("캡챠")) return;

      if (cellValue.includes("에러") || cellValue.toLowerCase().includes("error") || cellValue.startsWith("#")) {
        const errorKey = `${rowNumber}_${headerTitle}_${cellValue}`;
        currentErrors[errorKey] = true;

        if (!sentErrors[errorKey]) {
          const subject = `${colB} ${headerTitle} 에러`;
          const body = `'${ERROR_SHEET_NAME}' 시트의 ${rowNumber}행에서 에러가 확인되었습니다.\n\n대상: ${colB}\n항목: ${headerTitle}\n상세내용: ${cellValue}\n\n${LINK}`;

          try {
            GmailApp.sendEmail(RECIPIENT, subject, body);
            Logger.log(`🚨 [메일발송] ${rowNumber}행 [${headerTitle}] 새로운 에러 감지됨`);
          } catch (e) {
            Logger.log(`❌ ${rowNumber}행 에러메일 발송 실패: ${e.message}`);
          }
        }
      }
    });
  });

  props.setProperty(errorLogKey, JSON.stringify(currentErrors));
}

/**
 * 3. 카페작업리스트 O1: 프로그램이 기록한 캡챠 실패 목록에서 새로 생긴 항목만 메일 발송
 *    O1 형식: "캡챠 실패 (N)\nNo.150 김가연(대현) 캡챠실패 (사전 로그인 10-22 14:30)\n..."
 *    로그인이 확인되어 목록에서 빠졌다가 다시 캡챠가 뜨면 다시 발송
 */
function checkCaptchaAndSendEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) return;

  const props = PropertiesService.getScriptProperties();
  const captchaKey = `sentCaptcha_${ss.getId()}`;
  const sentCaptcha = JSON.parse(props.getProperty(captchaKey) || "{}");

  const lines = safeText(sheet.getRange(CAPTCHA_CELL).getValue())
    .split("\n").map(safeText).filter(l => l.includes("캡챠실패"));

  const current = {};
  lines.forEach(line => {
    // 시각/단계 표시가 바뀌어도 같은 사람은 한 번만 발송 ("No.150 김가연(대현)" 기준)
    const who = safeText(line.split(" 캡챠실패")[0]);
    current[who] = true;
    if (sentCaptcha[who]) return;

    const subject = `${who} 캡챠실패 - 수동로그인 필요`;
    const body = `${line}\n\n자동 로그인 중 캡챠(자동입력 방지)가 떠서 로그인하지 못했습니다.\n` +
                 `프로그램의 '사전 로그인 → 브라우저 열기'로 직접 로그인('로그인 상태 유지' 체크)해주세요.\n\n${LINK}`;
    try {
      GmailApp.sendEmail(RECIPIENT, subject, body);
      Logger.log(`🔒 [메일발송] ${who} 캡챠 실패`);
    } catch (e) {
      Logger.log(`❌ ${who} 캡챠메일 발송 실패: ${e.message}`);
      delete current[who]; // 다음 실행 때 재시도
    }
  });

  props.setProperty(captchaKey, JSON.stringify(current));
}

/**
 * 트리거 설정 함수 (한 번만 수동 실행)
 */
function setupTimeTriggerEveryMinute() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "checkNewRowsAndSendEmail" || t.getHandlerFunction() === "runAllChecks") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("runAllChecks")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("✅ 1분마다 runAllChecks 실행 트리거 등록 완료");
}

/**
 * 발송 내역 초기화 (필요할 때 수동 실행)
 */
function resetProcessedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const key1 = `lastSentTimestamp_${ss.getId()}_${SHEET_NAME}`;
  const key2 = `sentErrors_${ss.getId()}`;
  const key3 = `sentCaptcha_${ss.getId()}`;

  // 구버전 키도 함께 삭제 (행번호 기반)
  const oldKey = `lastProcessedRow_${ss.getId()}_${SHEET_NAME}`;

  props.deleteProperty(key1);
  props.deleteProperty(key2);
  props.deleteProperty(key3);
  props.deleteProperty(oldKey);

  Logger.log("✅ 모든 발송 기록이 초기화되었습니다. 다음 실행 시 '현재 시각' 이후의 H열 타임스탬프를 가진 행부터 발송됩니다.");
}

function safeText(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
