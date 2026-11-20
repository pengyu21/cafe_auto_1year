/**
 * 작업로그 시트 G3:G의 네이버 카페 URL 삭제 여부를 I3:I에 기록
 * - 정상: "O"
 * - 삭제: "삭제"
 * - 판정불가/에러: "에러"
 *
 * 방식: 네이버 카페 내부 API (v4)
 *   https://article.cafe.naver.com/gw/v4/cafes/{clubId}/articles/{articleId}?useCafeId=true
 *   - HTTP 200 + result.article → 정상
 *   - errorCode "4003"/"4004" 또는 메시지에 "삭제/존재하지 않" → 삭제
 *   - HTTP 404 → 삭제
 *   - 그 외 → 에러 (응답 본문을 로그에 기록)
 *
 * I1: 작업중에는 "작업중: N행", 완료 후에는 "업데이트: yyyy-MM-dd HH:mm:ss"
 *     5분 타임아웃 시 "중단 (N행) - 5분 뒤 자동 재개" → 5분 후 트리거로 자동 재개
 *
 * 자동 재개: 타임아웃 시 5분 뒤 실행되는 일회성 트리거 생성.
 *            I열에 결과가 비어있는 행부터 이어서 처리하므로 별도 인덱스 저장 불필요.
 */

const SHEET_NAME_CHK = '작업로그';
const START_ROW_CHK = 3;
const URL_COL_CHK = 7;    // G
const RESULT_COL_CHK = 9; // I

// 자동 재개 트리거 핸들러 (수동 트리거와 구분하기 위한 전용 이름)
const RESUME_HANDLER_CHK = 'resumeCheckCafePostStatus';

// 5분 후 재개
const RESUME_DELAY_MS = 5 * 60 * 1000;

// 공통 요청 헤더
const REQUEST_HEADERS_CHK = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  'Referer': 'https://cafe.naver.com/',
  'Accept': 'application/json'
};

/**
 * 수동 실행 및 매일 자동 실행 진입점.
 * I3:I 결과를 초기화한 뒤 처음부터 새로 검사.
 */
function checkCafePostStatus() {
  // 시작 시 잔여 자동 재개 트리거 정리 (새로 시작하는 경우)
  removeResumeTriggers_();

  // I3:I 결과 초기화 (매일 새로 검사)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CHK);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= START_ROW_CHK) {
      sheet.getRange(START_ROW_CHK, RESULT_COL_CHK, lastRow - START_ROW_CHK + 1, 1).clearContent();
      SpreadsheetApp.flush();
      Logger.log(`I${START_ROW_CHK}:I${lastRow} 초기화 완료`);
    }
  }

  runCheck_();
}

/**
 * 자동 재개 트리거가 호출하는 진입점.
 * 초기화 없이 I열 비어있는 행부터 이어서 검사.
 */
function resumeCheckCafePostStatus() {
  runCheck_();
}

function runCheck_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CHK);
  if (!sheet) throw new Error(`시트 "${SHEET_NAME_CHK}"를 찾을 수 없습니다.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW_CHK) return;

  const numRows = lastRow - START_ROW_CHK + 1;
  const urls = sheet.getRange(START_ROW_CHK, URL_COL_CHK, numRows, 1).getValues();
  const results = sheet.getRange(START_ROW_CHK, RESULT_COL_CHK, numRows, 1).getValues();

  const statusCell = sheet.getRange(1, RESULT_COL_CHK); // I1
  const tz = ss.getSpreadsheetTimeZone();
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;

  for (let i = 0; i < urls.length; i++) {
    const currentRow = START_ROW_CHK + i;

    // 시간 초과 → 5분 후 자동 재개 트리거 예약 후 종료
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      scheduleResume_();
      statusCell.setValue(`중단 (${currentRow}행) - 5분 뒤 자동 재개 (${Utilities.formatDate(new Date(), tz, 'HH:mm:ss')})`);
      SpreadsheetApp.flush();
      Logger.log(`실행시간 초과: ${currentRow}행에서 중단, 5분 뒤 자동 재개 예약`);
      return;
    }

    const url = (urls[i][0] || '').toString().trim();
    if (!url) continue;
    if (url.startsWith('실패')) continue; // 업로드 실패 기록(URL 대신 실패 이유)은 검사 대상 아님
    if (results[i][0]) continue; // 이미 결과 있으면 스킵 (= 재개 시 이어서 처리)

    statusCell.setValue(`작업중: ${currentRow}행`);
    SpreadsheetApp.flush();

    const status = checkSinglePost_(url);
    sheet.getRange(currentRow, RESULT_COL_CHK).setValue(status);
    SpreadsheetApp.flush();

    Utilities.sleep(400 + Math.floor(Math.random() * 300));
  }

  // 정상 완료 → 자동 재개 트리거 정리 + 완료 시각 기록
  removeResumeTriggers_();
  statusCell.setValue(`업데이트: ${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')}`);
  SpreadsheetApp.flush();
}

/** 5분 후 자동 재개 트리거 예약 (기존 자동 재개 트리거는 먼저 정리) */
function scheduleResume_() {
  removeResumeTriggers_();
  ScriptApp.newTrigger(RESUME_HANDLER_CHK)
    .timeBased()
    .after(RESUME_DELAY_MS)
    .create();
}

/** 자동 재개 트리거만 삭제 (사용자가 수동으로 설정한 다른 트리거는 보호) */
function removeResumeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === RESUME_HANDLER_CHK) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** 디버깅용: 현재 예약된 자동 재개 트리거 확인 */
function viewResumeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === RESUME_HANDLER_CHK);
  Logger.log(`자동 재개 트리거 ${triggers.length}개`);
  triggers.forEach(t => Logger.log(`  - ${t.getHandlerFunction()} (uid=${t.getUniqueId()})`));
}

/**
 * 단일 URL 삭제 여부 판정 (카페 내부 API v4).
 */
function checkSinglePost_(url) {
  try {
    const m = url.match(/cafe\.naver\.com\/([^\/\?#]+)\/(\d+)/);
    if (!m) {
      Logger.log(`[${url}] 카페 URL 패턴 아님`);
      return '에러';
    }
    const cafeName = m[1];
    const articleId = m[2];

    const clubId = getClubId_(cafeName);
    if (!clubId) {
      Logger.log(`[${url}] clubId 추출 실패`);
      return '에러';
    }

    const apiUrl = `https://article.cafe.naver.com/gw/v4/cafes/${clubId}/articles/${articleId}?useCafeId=true`;
    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: REQUEST_HEADERS_CHK
    });

    const code = res.getResponseCode();
    const body = res.getContentText('UTF-8');

    let json = null;
    try { json = JSON.parse(body); } catch (e) {}

    // 정상
    if (code === 200 && json && json.result && json.result.article) {
      return 'O';
    }

    // 삭제 판정: errorCode / reason / errorMessage 어디에 있든 확인
    const r = (json && json.result) || json || {};
    const errorCode = String(r.errorCode || '');
    const msg = String(r.reason || r.errorMessage || r.message || '');

    if (errorCode === '4003' || errorCode === '4004' || /삭제|존재하지\s*않/.test(msg)) {
      Logger.log(`[${url}] 삭제됨 (HTTP ${code}, errorCode=${errorCode}, msg=${msg})`);
      return '삭제';
    }
    if (code === 404) {
      Logger.log(`[${url}] HTTP 404 → 삭제 처리 / body=${body.slice(0, 200)}`);
      return '삭제';
    }

    // 그 외는 에러로 두고 본문을 남겨서 나중에 패턴 확인
    Logger.log(`[${url}] HTTP ${code} 판정불가 / body=${body.slice(0, 300)}`);
    return '에러';

  } catch (e) {
    Logger.log(`[${url}] 예외: ${e.message}`);
    return '에러';
  }
}

/**
 * 카페 URL 이름(예: cosmania)으로 숫자 clubId 조회 + 영구 캐시.
 */
function getClubId_(cafeName) {
  const props = PropertiesService.getScriptProperties();
  const cacheKey = 'clubId_' + cafeName;
  const cached = props.getProperty(cacheKey);
  if (cached) return cached;

  try {
    const res = UrlFetchApp.fetch('https://cafe.naver.com/' + cafeName, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': REQUEST_HEADERS_CHK['User-Agent']
      }
    });
    if (res.getResponseCode() !== 200) return null;

    let html = res.getContentText('EUC-KR');
    let match = html.match(/g_sClubId\s*=\s*["'](\d+)["']/);
    if (!match) {
      html = res.getContentText('UTF-8');
      match = html.match(/g_sClubId\s*=\s*["'](\d+)["']/);
    }
    if (match) {
      const clubId = match[1];
      props.setProperty(cacheKey, clubId);
      Logger.log(`[clubId캐시] ${cafeName} → ${clubId}`);
      return clubId;
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * 디버깅용: URL 하나의 v4 응답 원문 확인.
 * 삭제된 글 URL을 넣고 실행하면 삭제 시 응답 형태를 확인할 수 있음.
 */
function debugOneUrl() {
  const url = 'https://cafe.naver.com/feko/10436067'; // ← 확인할 URL로 교체
  const m = url.match(/cafe\.naver\.com\/([^\/\?#]+)\/(\d+)/);
  if (!m) { Logger.log('카페 URL 패턴 아님'); return; }

  const clubId = getClubId_(m[1]);
  Logger.log('clubId: ' + clubId);

  const apiUrl = `https://article.cafe.naver.com/gw/v4/cafes/${clubId}/articles/${m[2]}?useCafeId=true`;
  const res = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true, headers: REQUEST_HEADERS_CHK });
  Logger.log(`HTTP ${res.getResponseCode()}\n${res.getContentText().slice(0, 500)}`);
  Logger.log('판정 결과: ' + checkSinglePost_(url));
}

/** 스프레드시트 열 때 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('카페 검사')
    .addItem('전체 새로 검사', 'checkCafePostStatus')
    .addItem('이어서 검사 (빈 칸만)', 'resumeCheckCafePostStatus')
    .addSeparator()
    .addItem('자동 재개 취소', 'cancelResume')
    .addToUi();
}

/** 메뉴용: 예약된 자동 재개 트리거 취소 */
function cancelResume() {
  removeResumeTriggers_();
  SpreadsheetApp.getActiveSpreadsheet().toast('자동 재개 예약을 취소했습니다.', '카페 검사', 3);
}