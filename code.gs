/**
 * =========================================================
 * 장원급제반 Backend Script (Google Apps Script)
 * [엽전 경제 시스템 + 저잣거리 + 자동화 트리거 버전]
 * =========================================================
 *
 * [Member 시트 컬럼 구조]
 *  A(0): 학번  B(1): 이름  C(2): 역할  D(3): 엽전잔액  E(4): 사면장보유수
 *  F(5): 기숙사(Y/N)   G(6): 반(공무원/공기업)   H(7): alert
 *
 * [PointLog 시트 컬럼 구조]
 *  A(0): 날짜  B(1): 이름  C(2): 변동량(+N/-N 정수)  D(3): 사유  E(4): 부여자  F(5): 상태
 *
 * [AbsenceReason 시트]
 *  A(0): 날짜  B(1): 학생명  C(2): 사유(병결/공결/기타)  D(3): 등록자
 *
 * [MarketInventory 시트]
 *  A(0): 상품명  B(1): 엽전가격  C(2): 초기재고  D(3): 현재재고  E(4): 비고
 *
 * [MarketLog 시트]
 *  A(0): 날짜  B(1): 학생명  C(2): 상품명  D(3): 수량  E(4): 사용엽전  F(5): 처리자  G(6): 상태(완료/취소)
 *
 * [Teachers 시트]
 *  A(0): 이름  B(1): 전화번호 뒷4자리  C(2): 비고(과목 등)
 *
 * ★ 출결 정책:
 *   일반생:   ~08:10 출석 / 08:10~08:29 지각(-1냥) / 08:30 이후 버튼 차단 → 트리거가 결석 처리
 *   기숙사생: ~08:29 출석 / 08:30 이후 버튼 차단 → 트리거가 결석 처리 (지각 없음)
 *
 * ★ 자동화 트리거 (GAS 트리거 메뉴에서 직접 등록):
 *   - dailyAbsencePenalty  : 매일 오전 9시 (결석 처리 -2엽전)
 *   - dailyPlannerPenalty  : 매일 새벽 5시 (플래너미작성 -1엽전)
 *   - weeklyAttendanceBonus: 매주 월요일 새벽 5시 (주간개근 +5엽전)
 *   - weeklyBankInterest   : 매주 월요일 새벽 5~6시 (은행 저축 이자)      ← 신규
 *   - weeklyTeacherReport  : 매주 월요일 아침 7~8시 (주간 리포트 생성+푸시) ← 신규
 *   - pushCheckInReminder        : 매일 오전 7~8시 (미인증 학생 등교 독촉 푸시)
 *   - pushAfterSchoolAttendanceCheck : 매일 오전 8~9시 (전날 방과후 수업 있으면 교사에게 출석체크 푸시) ← 신규
 *   ※ dailyAbsencePenalty는 결석 확정 후 교사에게 오늘 출결 요약 푸시도 함께 발송(신규)
 *
 * [Bank 시트] A(0): 이름  B(1): 저축잔액  C(2): 최근변동
 * [WeeklyReport 시트] A(0): 주차시작  B(1): 생성시각  C(2): 내용
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const CACHE_TTL = {
  PLANNER:    600,
  ATTENDANCE: 300,
  MEMBER:     1800,
  ROLE_DATA:  300,
  STATS:      600,
  POINT_LOG:  300,
  EXAM:       1800,
  LEGACY:     3600,
  MARKET:     300,
};

// 지각 기준 시간 (일반생만 해당)
const CHECKIN_LIMIT = { hour: 8, min: 10 };

// 등교 인증 마감 시간 (일반생/기숙사생 공통 — 이후 차단)
const CHECKIN_CUTOFF = { hour: 8, min: 30 };

// 엽전 변동량 상수
const COIN = {
  PRESENT:      +1,   // 출석 (제시간 등교)
  LATE:         -1,   // 지각 (일반생만)
  ABSENT:       -2,   // 결석
  NO_PLANNER:   -1,   // 플래너 미작성
  PLANNER_DONE: +2,   // 플래너 100% 달성
  WEEKLY_FULL:  +5,   // 주간개근 (월~금)
};

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  let result   = {};

  // ★ 쓰기 작업 직렬화 (안정화)
  //   행 삭제(deleteRow)와 엽전 갱신이 동시에 겹치면 행 밀림·잔액 꼬임이 생길 수 있음
  //   조회(get*)와 읽기성 작업(로그인 등)은 잠금 없이 병렬 처리
  //   checkIn·savePlanner는 엽전 잔액을 갱신(applyCoins)하므로 반드시 잠금 필요:
  //   교사 지급/승인과 동시에 겹치면 잔액 read-modify-write가 서로를 덮어씀
  // notifyAfterSchoolChange: 시트를 전혀 건드리지 않는 순수 푸시 발송이라 잠금 불필요.
  //   (잠금을 걸면 여러 명에게 FCM 전송하는 동안 다른 교사 작업이 막힘)
  const NO_LOCK = { login: 1, verifyTeacher: 1, instructorLogin: 1,
                    clearAlert: 1, savePushToken: 1, notifyAfterSchoolChange: 1 };
  const needsLock = action.indexOf('get') !== 0 && !NO_LOCK[action];
  let lock = null;
  if (needsLock) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: '다른 작업을 처리 중입니다. 잠시 후 다시 시도해주세요.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  try {
    // ── 기존 기능 ──────────────────────────────────────────
    if      (action === 'login')              result = login(params);
    else if (action === 'checkIn')            result = checkIn(params);
    else if (action === 'getPlanner')         result = getPlanner(params);
    else if (action === 'savePlanner')        result = savePlanner(params);
    else if (action === 'getRoleData')        result = getRoleData(params);
    else if (action === 'roleAction')         result = roleAction(params);
    else if (action === 'getAdminData')       result = getAdminData();
    else if (action === 'adminAction')        result = adminAction(params);
    else if (action === 'grantPointBatch')    result = grantPointBatch(params);
    else if (action === 'updateAttendance')   result = updateAttendance(params);
    else if (action === 'addMember')          result = addMember(params);
    else if (action === 'clearAlert')         result = clearAlert(params);
    else if (action === 'getAdminPlanner')    result = getAdminPlanner(params);
    else if (action === 'getAdminAttendance') result = getAdminAttendance(params);
    // ── 범위 조회 (여러 날짜를 1회 호출로 — 프론트 속도 개선) ──
    else if (action === 'getAdminAttendanceRange') result = getAdminAttendanceRange(params);
    else if (action === 'getAdminPlannerRange')    result = getAdminPlannerRange(params);
    else if (action === 'getCounselingRange')      result = getCounselingRange(params);
    else if (action === 'getMyHistory')       result = getMyHistory(params);
    else if (action === 'getMyPointLog')      result = getMyPointLog(params);
    else if (action === 'getClassStats')      result = getClassStats();
    else if (action === 'saveExamScores')     result = saveExamScores(params);
    else if (action === 'getMyExamScores')    result = getMyExamScores(params);
    else if (action === 'getExamList')        result = getExamList();
    else if (action === 'saveLegacyScores')   result = saveLegacyScores(params);
    else if (action === 'getLegacyStats')     result = getLegacyStats(params);
    else if (action === 'getAllMemberPoints') result = getAllMemberPoints();
    else if (action === 'saveCounseling')     result = saveCounseling(params);
    else if (action === 'getCounselingByDate')    result = getCounselingByDate(params);
    else if (action === 'getCounselingByStudent') result = getCounselingByStudent(params);
    // ── 엽전 시스템 신규 ───────────────────────────────────
    else if (action === 'getAbsenceReasons')   result = getAbsenceReasons(params);
    else if (action === 'saveAbsenceReason')   result = saveAbsenceReason(params);
    else if (action === 'deleteAbsenceReason') result = deleteAbsenceReason(params);
    // ── 저잣거리 ───────────────────────────────────────────
    else if (action === 'getMarketInventory')  result = getMarketInventory();
    else if (action === 'saveMarketItem')      result = saveMarketItem(params);
    else if (action === 'deleteMarketItem')    result = deleteMarketItem(params);
    else if (action === 'marketPurchase')      result = marketPurchase(params);
    else if (action === 'marketPurchaseMulti') result = marketPurchaseMulti(params);
    else if (action === 'getMarketLog')        result = getMarketLog(params);
    else if (action === 'cancelMarketLog')     result = cancelMarketLog(params);
    // ── 교과 선생님 QR ────────────────────────────────────
    else if (action === 'verifyTeacher')       result = verifyTeacher(params);
    // ── 상담 통계 ─────────────────────────────────────────
    else if (action === 'getCounselingAll')    result = getCounselingAll();
    // ── 트리거 일시정지 ───────────────────────────────────
    else if (action === 'getTriggerPauses')    result = getTriggerPauses();
    else if (action === 'saveTriggerPause')    result = saveTriggerPause(params);
    else if (action === 'deleteTriggerPause')  result = deleteTriggerPause(params);
    // ── PointLog 관리 ─────────────────────────────────────
    else if (action === 'getAdminPointLog')    result = getAdminPointLog(params);
    else if (action === 'editPointLog')        result = editPointLog(params);
    else if (action === 'deletePointLog')      result = deletePointLog(params);
    else if (action === 'deletePointLogBatch') result = deletePointLogBatch(params);
    // ── 학생 엽전내역 수정요청 ────────────────────────────
    else if (action === 'requestPointEdit')        result = requestPointEdit(params);
    else if (action === 'getPointEditRequests')    result = getPointEditRequests();
    else if (action === 'getMyEditRequests')       result = getMyEditRequests(params);
    else if (action === 'resolvePointEditRequest') result = resolvePointEditRequest(params);
    // ── 방과후수업 일정표 ─────────────────────────────────
    else if (action === 'instructorLogin')            result = instructorLogin(params);
    else if (action === 'getAfterSchool')             result = getAfterSchool();
    else if (action === 'saveAfterSchoolClass')       result = saveAfterSchoolClass(params);
    else if (action === 'deleteAfterSchoolClass')     result = deleteAfterSchoolClass(params);
    else if (action === 'saveAfterSchoolException')   result = saveAfterSchoolException(params);
    else if (action === 'deleteAfterSchoolException') result = deleteAfterSchoolException(params);
    else if (action === 'notifyAfterSchoolChange')    result = notifyAfterSchoolChange(params);
    else if (action === 'getAfterSchoolAttendance')   result = getAfterSchoolAttendance(params);
    else if (action === 'saveAfterSchoolAttendance')  result = saveAfterSchoolAttendance(params);
    // ── 푸시 알림 ─────────────────────────────────────────
    else if (action === 'savePushToken')              result = savePushToken(params);
    // ── 교사일지 / 졸업생 / 학년도 전환 ───────────────────
    else if (action === 'getTeacherLog')              result = getTeacherLog();
    else if (action === 'saveTeacherLog')             result = saveTeacherLog(params);
    else if (action === 'deleteTeacherLog')           result = deleteTeacherLog(params);
    else if (action === 'getGraduateScores')          result = getGraduateScores(params);
    else if (action === 'promoteSchoolYear')          result = promoteSchoolYear(params);
    // ── 엽전 사유 프리셋 / 시험 D-day / 엽전 은행 / 주간 리포트 ──
    else if (action === 'getCoinPresets')             result = getCoinPresets();
    else if (action === 'saveCoinPresets')            result = saveCoinPresets(params);
    else if (action === 'getExamDdays')               result = getExamDdays();
    else if (action === 'saveExamDday')               result = saveExamDday(params);
    else if (action === 'deleteExamDday')             result = deleteExamDday(params);
    else if (action === 'getBank')                    result = getBank(params);
    else if (action === 'getBankAll')                 result = getBankAll();
    else if (action === 'bankDeposit')                result = bankDeposit(params);
    else if (action === 'bankWithdraw')               result = bankWithdraw(params);
    else if (action === 'getWeeklyReports')           result = getWeeklyReports();
    else if (action === 'runWeeklyReport')            result = runWeeklyReport();
    else if (action === 'getMyStreak')                result = getMyStreak(params);
    else throw new Error('알 수 없는 action: ' + action);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock) lock.releaseLock();
  }
}


/* =========================================================
   [유틸] 공통
   ========================================================= */
function getValidData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

function toDateStr(date) {
  date = date || new Date();
  const koreaTimeStr = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const koreaHour    = parseInt(koreaTimeStr.slice(11, 13), 10);
  if (koreaHour < 6) {
    const prev = new Date(date.getTime() - 24 * 60 * 60 * 1000);
    return Utilities.formatDate(prev, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return koreaTimeStr.slice(0, 10);
}

function normalizeDateCell(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(val).replace(/^'/, '').trim();
}

function normalizeTimeCell(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    const s = String(val.getSeconds()).padStart(2, '0');
    return h + ':' + m + ':' + s;
  }
  return String(val).trim();
}

function isDormStudent(row) {
  return String(row[5] || '').trim().toUpperCase() === 'Y';
}

/**
 * 블록 쓰기(setValues) 전에 시트 그리드 용량을 확보한다.
 * appendRow는 그리드가 차면 스스로 늘리지만 getRange().setValues()는
 * 범위가 시트 밖이면 예외를 던짐 — 잔액은 갱신됐는데 로그만 빠지는 사고 방지.
 */
function _ensureRows(sheet, extra) {
  const need = sheet.getLastRow() + extra - sheet.getMaxRows();
  if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need);
}

/**
 * 트리거(시간 기반 실행)용 스크립트 락 래퍼.
 * doPost의 쓰기 잠금과 같은 락을 잡아, 트리거의 잔액 일괄 갱신이
 * 교사/학생의 실시간 쓰기와 겹쳐 서로를 덮어쓰는 것을 막는다.
 */
function _withScriptLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/**
 * 같은 잡이 같은 날짜에 두 번 '완주'하지 않게 하는 멱등성 가드.
 * - _jobAlreadyRan: 조회만 한다 (잡 시작 시).
 * - _markJobRan: 잡이 끝까지 성공한 뒤에만 호출한다.
 *   → 잡이 중간에 죽으면 마킹이 없어 재실행으로 복구 가능.
 *     (각 잡은 행 단위 멱등 가드가 따로 있어 재실행돼도 이중 차감 없음)
 * - 잡당 키 1개(값=마지막 완주 날짜)라 저장소가 누적되지 않음.
 * - 동시 중복 실행은 _withScriptLock이 직렬화하므로 시작 마킹이 필요 없음.
 */
function _jobAlreadyRan(jobName, dateStr) {
  return PropertiesService.getScriptProperties().getProperty('job_' + jobName) === String(dateStr);
}
function _markJobRan(jobName, dateStr) {
  PropertiesService.getScriptProperties().setProperty('job_' + jobName, String(dateStr));
}

/**
 * 한국 공휴일 여부 — 구글 공휴일 캘린더의 공개 iCal을 UrlFetch로 조회.
 * ★ 캘린더 서비스를 쓰지 않는 이유: 새 OAuth 스코프가 추가되면 소유자가 재승인하기
 *   전까지 모든 시간 트리거가 실행 자체를 실패한다. UrlFetch는 이미 승인된 스코프.
 * ★ 이름 화이트리스트: 캘린더에는 어버이날·식목일 같은 '기념일(수업일)'도 들어있어
 *   법정공휴일 이름만 휴일로 인정. 조회 실패 시 false(기존 동작 유지)로 폴백.
 */
function isKoreanHoliday(dateStr) {
  try {
    const key = 'kr_holiday_' + dateStr;
    const cached = cacheGet(key);
    if (cached !== null) return cached === 1;

    const ymd = String(dateStr).replace(/-/g, '');
    const ics = UrlFetchApp.fetch(
      'https://calendar.google.com/calendar/ical/ko.south_korea%23holiday%40group.v.calendar.google.com/public/basic.ics',
      { muteHttpExceptions: true }).getContentText();

    const NAMES = ['신정', '새해', '설날', '삼일절', '어린이날', '부처님', '석가탄신일',
                   '현충일', '광복절', '추석', '개천절', '한글날', '성탄절', '크리스마스',
                   '대체', '선거', '임시공휴일'];
    let holiday = false;
    const blocks = ics.split('BEGIN:VEVENT');
    for (let i = 1; i < blocks.length; i++) {
      if (blocks[i].indexOf('DTSTART;VALUE=DATE:' + ymd) < 0) continue;
      const m = blocks[i].match(/SUMMARY:(.+)/);
      const name = m ? m[1].trim() : '';
      if (NAMES.some(function(n) { return name.indexOf(n) >= 0; })) { holiday = true; break; }
    }
    cacheSet(key, holiday ? 1 : 0, 21600);   // 6시간 캐시 (하루 트리거 2~3회면 충분)
    return holiday;
  } catch(e) { return false; }
}

function cacheGet(key) {
  const val = CacheService.getScriptCache().get(key);
  if (!val) return null;
  try { return JSON.parse(val); } catch(e) { return null; }
}

function cacheSet(key, value, ttl) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(value), ttl || 600); }
  catch(e) { Logger.log('Cache set failed [' + key + ']: ' + e.toString()); }
}

function cacheRemove(keys) {
  CacheService.getScriptCache().removeAll(keys);
}


/* =========================================================
   [엽전 유틸] Member 시트 엽전 잔액 변경
   ========================================================= */
function applyCoins(name, delta, reason, actor, pre) {
  actor = actor || '시스템';
  // pre = {memSheet, memData}: 호출자가 방금 읽은 Member 데이터 재사용 (같은 잠금 안이라 안전)
  const memSheet = (pre && pre.memSheet) || SS.getSheetByName('Member');
  const memData  = (pre && pre.memData)  || getValidData(memSheet);

  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][1]).trim() === name) {
      const current = Number(memData[i][3]) || 0;
      memSheet.getRange(i + 1, 4).setValue(current + delta);
      memData[i][3] = current + delta;
      SS.getSheetByName('PointLog').appendRow([toDateStr(), name, delta, reason, actor, '완료']);
      cacheRemove(['member_all', 'admin_data', 'pointlog_' + name, 'market_log']);
      _fsPushMemberRow(memData[i]);   // Firestore 즉시 반영 (실시간 잔액)
      return current + delta;
    }
  }
  throw new Error('Member에서 관원을 찾지 못했소: ' + name);
}


/* =========================================================
   1. 로그인
   ========================================================= */
function login(p) {
  if (p.id === '청양고' && p.name === '교사') {
    return { name: '교사', role: 'teacher', coins: 0, amnesty: 0,
             dorm: false, checkedIn: false, todayPlanner: { content: '', progress: '' } };
  }

  const MEMBER_KEY = 'member_all';
  let memberData   = cacheGet(MEMBER_KEY);
  if (!memberData) {
    memberData = getValidData(SS.getSheetByName('Member'));
    cacheSet(MEMBER_KEY, memberData, CACHE_TTL.MEMBER);
  }

  for (let i = 1; i < memberData.length; i++) {
    // ★ trim 비교: 시트 셀에 앞뒤 공백이 섞여도 로그인 실패하지 않도록
    if (String(memberData[i][0]).trim() == String(p.id).trim() &&
        String(memberData[i][1]).trim() == String(p.name).trim()) {
      const user = {
        id:      memberData[i][0],
        name:    memberData[i][1],
        role:    memberData[i][2],
        coins:   Number(memberData[i][3]) || 0,
        amnesty: Number(memberData[i][4]) || 0,
        dorm:    isDormStudent(memberData[i]),
        track:   String(memberData[i][6] || '').trim() || '공무원',
        alert:   memberData[i][7]
      };
      const today       = toDateStr();
      user.checkedIn    = getTodayCheckIn(user.name, today);
      user.todayPlanner = getPlannerByNameDate(user.name, today);
      return user;
    }
  }
  throw new Error('관원 명부에 없는 정보입니다.');
}

function getTodayCheckIn(name, today) {
  const CACHE_KEY = 'checkin_' + name + '_' + today;
  const cached    = cacheGet(CACHE_KEY);
  if (cached !== null) return cached;

  // ★ 하루치 출결 맵을 통째로 캐시 — 아침 로그인 러시 때
  //   학생 수만큼 Attendance 전체 읽기가 반복되는 것을 1회로 축소
  const MAP_KEY = 'att_all_' + today;
  let map = cacheGet(MAP_KEY);
  if (!map) {
    map = {};
    const data = getValidData(SS.getSheetByName('Attendance'));
    for (let i = 1; i < data.length; i++) {
      if (normalizeDateCell(data[i][0]) === today)
        map[String(data[i][1]).trim()] = data[i][3];
    }
    cacheSet(MAP_KEY, map, 120);
  }
  const st = map[String(name).trim()];
  if (st !== undefined && st !== null && st !== '') {
    cacheSet(CACHE_KEY, st, CACHE_TTL.ATTENDANCE);
    return st;
  }
  cacheSet(CACHE_KEY, false, 60);
  return false;
}

function getPlannerByNameDate(name, date) {
  const CACHE_KEY = 'planner_' + name + '_' + date;
  const cached    = cacheGet(CACHE_KEY);
  if (cached !== null) return cached;

  const sheet  = SS.getSheetByName('Planner');
  const data   = getValidData(sheet);
  let   result = { content: '', progress: '' };
  for (let i = data.length - 1; i > 0; i--) {
    if (normalizeDateCell(data[i][0]) === date && data[i][1] === name) {
      result = { content: data[i][2], progress: data[i][3] };
      break;
    }
  }
  cacheSet(CACHE_KEY, result, CACHE_TTL.PLANNER);
  return result;
}


/* =========================================================
   2. 등교 인증
   ★ 출결 정책:
     - 08:30 이후: 일반생/기숙사생 모두 차단 (트리거가 결석 처리)
     - 기숙사생: 08:30 이전이면 무조건 출석 (지각 없음)
     - 일반생:   08:10 이전 출석 / 08:10~08:29 지각(-1냥)
   ========================================================= */
function checkIn(p) {
  const today     = toDateStr();
  const CACHE_KEY = 'checkin_' + p.name + '_' + today;

  const cachedStatus = cacheGet(CACHE_KEY);
  if (cachedStatus) throw new Error('이미 오늘의 등교 인증을 마쳤소. (' + cachedStatus + ')');

  const sheet = SS.getSheetByName('Attendance');
  const data  = getValidData(sheet);
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === today && data[i][1] === p.name) {
      throw new Error('이미 오늘의 등교 인증을 마쳤소. (' + data[i][3] + ')');
    }
  }

  // 기숙사 여부 및 사면장 확인
  const memberSheet  = SS.getSheetByName('Member');
  const memberData   = getValidData(memberSheet);
  let   dormFlag     = false;
  let   amnestyCount = 0;
  for (let i = 1; i < memberData.length; i++) {
    if (memberData[i][1] === p.name) {
      dormFlag     = isDormStudent(memberData[i]);
      amnestyCount = Number(memberData[i][4]) || 0;
      break;
    }
  }

  // ★ 한국 시간 기준으로 비교
  const now       = new Date();
  const koreaTime = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  const koreaH    = parseInt(koreaTime.slice(0, 2), 10);
  const koreaM    = parseInt(koreaTime.slice(3, 5), 10);

  // ★ 06시 이전 차단 — toDateStr의 6시 경계 때문에 00~05시 인증은
  //   '전날 출석'으로 기록돼 심야에 어제 출석 +1을 만드는 구멍이 됨
  if (koreaH < 6) {
    throw new Error('등교 인증은 06:00부터 가능하오.');
  }

  // 08:30 이후 → 일반생/기숙사생 모두 차단
  const isCutoff = koreaH > CHECKIN_CUTOFF.hour ||
                   (koreaH === CHECKIN_CUTOFF.hour && koreaM >= CHECKIN_CUTOFF.min);
  if (isCutoff) {
    throw new Error('08:30 이후에는 등교 인증이 불가합니다. 결석 처리됩니다.');
  }

  // 지각 판정: 일반생만 08:10 기준 적용 (기숙사생은 08:30 이전이면 무조건 출석)
  const isLate = !dormFlag &&
                 (koreaH > CHECKIN_LIMIT.hour ||
                  (koreaH === CHECKIN_LIMIT.hour && koreaM >= CHECKIN_LIMIT.min));
  const status = isLate ? '지각' : '출석';

  // 한국 시간으로 시각 기록
  const timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm:ss');

  sheet.appendRow([today, p.name, timeStr, status, '']);
  cacheRemove([CACHE_KEY, 'role_암행어사_' + today, 'admin_att_' + today, 'att_all_' + today]);

  // 엽전 자동 처리: 출석 +1냥 / 지각 -1냥
  // (위에서 읽은 Member 데이터를 재사용해 전체 시트 이중 읽기 제거)
  const pre = { memSheet: memberSheet, memData: memberData };
  let newCoins = null;
  if (status === '지각') {
    newCoins = applyCoins(p.name, COIN.LATE, '지각', '시스템', pre);
  } else if (status === '출석') {
    newCoins = applyCoins(p.name, COIN.PRESENT, '출석', '시스템', pre);
  }

  return { status, time: timeStr, dorm: dormFlag, newCoins, amnesty: amnestyCount };
}


/* =========================================================
   3. 플래너
   ★ progress=100 달성 시 COIN.PLANNER_DONE(+2) 자동 지급
   ========================================================= */
function getPlanner(p) {
  // ★ 버그 수정: 프론트가 date를 보내 과거 기록을 조회하는데 기존엔 무시하고 항상 오늘만 반환했음
  return getPlannerByNameDate(p.name, p.date || toDateStr());
}

function savePlanner(p) {
  const sheet  = SS.getSheetByName('Planner');
  const today  = toDateStr();
  const data   = getValidData(sheet);

  let progress = 0;
  try {
    const plans = JSON.parse(p.content);
    const total = plans.length;
    if (total > 0) {
      const done = plans.filter(function(t) { return t.status === 3; }).length;
      progress = Math.round((done / total) * 100);
    }
  } catch(e) { progress = 0; }

  let saved        = false;
  let prevProgress = 0;
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === today && data[i][1] === p.name) {
      prevProgress = Number(data[i][3]) || 0;
      sheet.getRange(i + 1, 3).setValue(p.content);
      sheet.getRange(i + 1, 4).setValue(progress);
      saved = true;
      break;
    }
  }
  if (!saved) {
    sheet.appendRow([today, p.name, p.content, progress]);
  }

  let coinMsg = null;
  if (progress === 100 && prevProgress < 100) {
    // ★ 하루 1회만 지급 — 100%↔미만을 반복 저장해 +2를 무한 파밍하는 구멍 차단
    const plogData = getValidData(SS.getSheetByName('PointLog'));
    let alreadyPaid = false;
    for (let i = plogData.length - 1; i >= 1; i--) {
      if (normalizeDateCell(plogData[i][0]) === today &&
          String(plogData[i][1]).trim() === String(p.name).trim() &&
          String(plogData[i][3]).trim() === '플래너100%달성') {
        alreadyPaid = true;
        break;
      }
    }
    if (!alreadyPaid) {
      applyCoins(p.name, COIN.PLANNER_DONE, '플래너100%달성', '시스템');
      coinMsg = '🪙 +' + COIN.PLANNER_DONE + '엽전 (플래너 100% 달성)';
    }
  }

  cacheRemove([
    'planner_' + p.name + '_' + today,
    'role_집현전학사_' + today,
    'stats_class',
    'history_' + p.name,
    'admin_planner_' + today
  ]);

  return { success: true, message: saved ? '수정되었소.' : '등록되었소.', progress, coinMsg };
}


/* =========================================================
   4. 알림 초기화
   ========================================================= */
function clearAlert(p) {
  const sheet = SS.getSheetByName('Member');
  const data  = getValidData(sheet);

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === p.name) {
      if (p.clearedMsg) {
        const existing  = String(data[i][7] || '').trim();
        const datedTypes = ['플래너독촉:', '지각소명:'];
        const prefix    = datedTypes.find(function(pfx) { return p.clearedMsg + ':' === pfx; })
                          ? p.clearedMsg + ':' : null;
        const remaining = existing.split('|')
          .filter(function(a) {
            if (!a) return false;
            if (prefix) return !a.startsWith(prefix);
            return a !== p.clearedMsg;
          }).join('|');
        sheet.getRange(i + 1, 8).setValue(remaining);
      } else {
        sheet.getRange(i + 1, 8).setValue('');
      }

      if (p.reason) {
        const attSheet = SS.getSheetByName('Attendance');
        const attData  = getValidData(attSheet);
        const today    = toDateStr();
        for (let j = attData.length - 1; j > 0; j--) {
          if (normalizeDateCell(attData[j][0]) === today && attData[j][1] === p.name) {
            attSheet.getRange(j + 1, 5).setValue(p.reason);
            break;
          }
        }
      }
      cacheRemove(['member_all']);
      return '확인 완료';
    }
  }
  throw new Error('관원을 찾을 수 없소: ' + p.name);
}


/* =========================================================
   5. 관원 역할 데이터 조회
   ========================================================= */
function getRoleData(p) {
  const queryDate = p.date ? p.date : toDateStr();

  if (p.role === '암행어사') {
    const CACHE_KEY = 'role_암행어사_' + queryDate;
    const cached    = cacheGet(CACHE_KEY);
    if (cached) return cached;

    const memSheet   = SS.getSheetByName('Member');
    const memRows    = getValidData(memSheet).slice(1)
      .filter(function(r) {
        const name = String(r[1] || '').trim();
        return name !== '' && name !== '교사' && r[2] !== 'teacher';
      });
    const allMembers = memRows.map(function(r) { return String(r[1]).trim(); });
    const dormSet    = new Set(memRows.filter(isDormStudent).map(function(r) { return String(r[1]).trim(); }));

    const attSheet = SS.getSheetByName('Attendance');
    const attMap   = {};
    getValidData(attSheet).slice(1).forEach(function(r) {
      if (normalizeDateCell(r[0]) === queryDate) {
        attMap[String(r[1]).trim()] = { time: normalizeTimeCell(r[2]), status: r[3] };
      }
    });

    const alertSentSet = new Set();
    getValidData(memSheet).slice(1).forEach(function(r) {
      const name   = String(r[1] || '').trim();
      const alerts = String(r[7] || '').trim();
      if (name && alerts.split('|').some(function(a) { return a === '지각소명:' + queryDate; })) {
        alertSentSet.add(name);
      }
    });

    const result = allMembers.map(function(name) {
      const att = attMap[name];
      return { name, time: att ? att.time : '-', status: att ? att.status : '미출석',
               dorm: dormSet.has(name), alertSent: alertSentSet.has(name) };
    });
    cacheSet(CACHE_KEY, result, CACHE_TTL.ATTENDANCE);
    return result;
  }

  if (p.role === '집현전학사') {
    const CACHE_KEY = 'role_집현전학사_' + queryDate;
    const cached    = cacheGet(CACHE_KEY);
    if (cached) return cached;

    const memSheet   = SS.getSheetByName('Member');
    const allMembers = getValidData(memSheet).slice(1)
      .filter(function(r) {
        const name = String(r[1] || '').trim();
        return name !== '' && name !== '교사' && r[2] !== 'teacher';
      }).map(function(r) { return String(r[1]).trim(); });

    const planSheet = SS.getSheetByName('Planner');
    const planMap   = {};
    getValidData(planSheet).slice(1).forEach(function(r) {
      if (normalizeDateCell(r[0]) === queryDate) {
        planMap[r[1]] = { submitted: true, progress: Number(r[3]) || 0 };
      }
    });

    const today      = toDateStr();
    const dueSentSet = new Set();
    getValidData(memSheet).slice(1).forEach(function(r) {
      const name   = String(r[1] || '').trim();
      const alerts = String(r[7] || '').trim();
      if (name && alerts.split('|').includes('플래너독촉:' + today)) dueSentSet.add(name);
    });

    const result = allMembers.map(function(name) {
      return { name, hasPlan: !!planMap[name],
               progress: planMap[name] ? planMap[name].progress : 0,
               alertSent: dueSentSet.has(name) };
    });
    cacheSet(CACHE_KEY, result, CACHE_TTL.ROLE_DATA);
    return result;
  }

  if (p.role === '포도대장') {
    const CACHE_KEY = 'role_포도대장_v2';
    const cached    = cacheGet(CACHE_KEY);
    if (cached) return cached;
    const sheet  = SS.getSheetByName('Member');
    const result = getValidData(sheet).slice(1)
      .filter(function(r) {
        const name = String(r[1] || '').trim();
        return name !== '' && name !== '교사' && r[2] !== 'teacher';
      }).map(function(r) { return String(r[1]).trim(); });
    cacheSet(CACHE_KEY, result, CACHE_TTL.MEMBER);
    return result;
  }

  throw new Error('알 수 없는 역할: ' + p.role);
}


/* =========================================================
   6. 관원 역할 수행
   ========================================================= */
function roleAction(p) {
  const today = toDateStr();

  if (p.type === 'penaltyRequest') {
    if (!p.targetName || !p.reason) throw new Error('대상 또는 사유가 없소.');
    SS.getSheetByName('PointLog').appendRow([today, p.targetName, p.amount || COIN.LATE, p.reason, p.requestor, '대기']);
    cacheRemove(['admin_data', 'pointlog_' + p.targetName]);
    try { _sendPushToStudent('교사', '🔔 엽전 차감 요청', (p.requestor || '') + ' → ' + p.targetName + ' (' + p.reason + ')'); } catch(e) {}  // [C]
    return '상소문(차감요청)을 올렸습니다.';
  }

  if (p.type === 'sendAlert') {
    if (!p.targetName || !p.alertMsg) throw new Error('대상 또는 메시지가 없소.');
    const sheet = SS.getSheetByName('Member');
    const data  = getValidData(sheet);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === p.targetName) {
        const existing = String(data[i][7] || '').trim();
        const alerts   = existing ? existing.split('|') : [];
        const storeMsg = (p.alertMsg === '플래너독촉' || p.alertMsg === '지각소명')
                         ? p.alertMsg + ':' + today : p.alertMsg;
        const prefix   = (p.alertMsg === '플래너독촉' || p.alertMsg === '지각소명')
                         ? p.alertMsg + ':' : p.alertMsg;
        const alreadyHas = alerts.some(function(a) { return a === storeMsg || a.startsWith(prefix); });
        if (!alreadyHas) {
          const filtered = alerts.filter(function(a) { return !a.startsWith(prefix); });
          filtered.push(storeMsg);
          sheet.getRange(i + 1, 8).setValue(filtered.join('|'));
        }
        cacheRemove(['member_all']);
        return '해당 관원에게 전갈을 보냈습니다.';
      }
    }
    throw new Error('관원을 찾을 수 없소: ' + p.targetName);
  }

  throw new Error('알 수 없는 roleAction type: ' + p.type);
}


/* =========================================================
   7. 관리자 조회 & 수행
   ========================================================= */
function getAdminData() {
  const CACHE_KEY = 'admin_data';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;

  const members = getValidData(SS.getSheetByName('Member')).slice(1)
    .filter(function(r) { return r[1] && String(r[1]).trim() !== ''; })
    .map(function(r) {
      return { name: r[1], role: r[2], dorm: isDormStudent(r),
               track: String(r[6] || '').trim() || '공무원' };
    });

  const pendingLogs = getValidData(SS.getSheetByName('PointLog')).slice(1)
    .filter(function(r) { return r[5] === '대기'; });

  const result = { members, pendingLogs };
  cacheSet(CACHE_KEY, result, CACHE_TTL.POINT_LOG);
  return result;
}

function adminAction(p) {
  if (p.type === 'approvePenalty' || p.type === 'approveDeduct') {
    if (!p.targetName) throw new Error('대상 이름이 없소.');
    const logSheet = SS.getSheetByName('PointLog');
    const logData  = getValidData(logSheet);
    let   delta    = 0;
    let   approved = false;
    for (let i = 1; i < logData.length; i++) {
      // ★ String().trim() 비교: 사유 셀이 숫자/날짜로 자동 형변환돼도 매칭되도록
      if (String(logData[i][1]).trim() === String(p.targetName).trim() &&
          String(logData[i][3]).trim() === String(p.reason).trim() &&
          logData[i][5] === '대기') {
        delta = Number(logData[i][2]) || -1;
        logSheet.getRange(i + 1, 6).setValue('완료');
        approved = true;
        break;
      }
    }
    if (!approved) throw new Error('승인할 대기 항목을 찾지 못했소.');
    const memSheet = SS.getSheetByName('Member');
    const memData  = getValidData(memSheet);
    for (let i = 1; i < memData.length; i++) {
      if (memData[i][1] === p.targetName) {
        memData[i][3] = (Number(memData[i][3]) || 0) + delta;
        memSheet.getRange(i + 1, 4).setValue(memData[i][3]);
        cacheRemove(['member_all', 'admin_data', 'pointlog_' + p.targetName]);
        _fsPushMemberRow(memData[i]);   // Firestore 즉시 반영
        try { _sendPushToStudent(p.targetName, '🪙 벌점 ' + delta + '냥', String(p.reason || '')); } catch(e) {}  // [A]
        return '엽전 처리 완료.';
      }
    }
    throw new Error('Member에서 관원을 찾지 못했소: ' + p.targetName);
  }

  if (p.type === 'rejectPenalty') {
    if (!p.targetName) throw new Error('대상 이름이 없소.');
    const logSheet = SS.getSheetByName('PointLog');
    const logData  = getValidData(logSheet);
    for (let i = 1; i < logData.length; i++) {
      if (String(logData[i][1]).trim() === String(p.targetName).trim() &&
          String(logData[i][3]).trim() === String(p.reason).trim() &&
          logData[i][5] === '대기') {
        logSheet.getRange(i + 1, 6).setValue('미승인');
        cacheRemove(['member_all', 'admin_data', 'pointlog_' + p.targetName]);
        return '미승인 처리 완료.';
      }
    }
    throw new Error('미승인할 대기 항목을 찾지 못했소.');
  }

  if (p.type === 'setRole') {
    if (!p.targetName || !p.role) throw new Error('대상 또는 역할이 없소.');
    const sheet = SS.getSheetByName('Member');
    const data  = getValidData(sheet);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === p.targetName) {
        sheet.getRange(i + 1, 3).setValue(p.role);
        data[i][2] = p.role;
        cacheRemove(['member_all', 'admin_data', 'role_포도대장_v2']);
        _fsPushMemberRow(data[i]);   // Firestore 즉시 반영 (역할)
        return p.targetName + '을(를) ' + p.role + '로 임명했소.';
      }
    }
    throw new Error('관원을 찾지 못했소: ' + p.targetName);
  }

  if (p.type === 'setDorm') {
    if (!p.targetName) throw new Error('대상 이름이 없소.');
    const sheet = SS.getSheetByName('Member');
    const data  = getValidData(sheet);
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === p.targetName) {
        sheet.getRange(i + 1, 6).setValue(p.dorm ? 'Y' : '');
        data[i][5] = p.dorm ? 'Y' : '';
        cacheRemove(['member_all', 'admin_data']);
        _fsPushMemberRow(data[i]);   // Firestore 즉시 반영 (기숙사)
        return p.targetName + ' 기숙사 ' + (p.dorm ? '등록' : '해제') + ' 완료.';
      }
    }
    throw new Error('관원을 찾지 못했소: ' + p.targetName);
  }

  if (p.type === 'grantPoint') {
    if (!p.targetName || !p.amount || !p.reason) throw new Error('대상, 변동량, 사유가 필요하오.');
    const delta = Number(p.amount);
    if (isNaN(delta)) throw new Error('변동량이 올바르지 않소.');
    const newCoins = applyCoins(p.targetName, delta, p.reason, '교사');
    try {  // [A] 상점/벌점 즉시 알림
      const t = delta >= 0 ? '🪙 상점! +' + delta + '냥' : '🪙 벌점 ' + delta + '냥';
      _sendPushToStudent(p.targetName, t, String(p.reason || ''));
    } catch(e) {}
    return { msg: '엽전 처리 완료.', newCoins };
  }

  throw new Error('알 수 없는 adminAction type: ' + p.type);
}


/* =========================================================
   7-0. 엽전 일괄 지급/차감 (여러 학생 동시)
   ========================================================= */
function grantPointBatch(p) {
  if (!Array.isArray(p.names) || p.names.length === 0) throw new Error('대상 학생이 없소.');
  const amount = Number(p.amount);
  if (isNaN(amount) || amount === 0) throw new Error('변동량이 올바르지 않소.');
  const reason = String(p.reason || '').trim();
  if (!reason) throw new Error('사유가 필요하오.');

  // ★ 배치 처리: 학생 N명에 대해 (시트 전체읽기 + appendRow) × N 반복 대신
  //   Member 1회 읽기 + 잔액 개별 갱신 + PointLog는 한 번의 블록 쓰기.
  //   전역 잠금 점유 시간이 크게 줄어 다른 요청이 덜 막힘.
  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  const today    = toDateStr();
  const results  = [];
  const logRows  = [];
  const applied  = [];   // 잔액 반영할 [행인덱스, 새잔액]
  const cacheKeys = ['member_all', 'admin_data'];

  p.names.forEach(function(rawName) {
    const name = String(rawName || '').trim();
    if (!name) return;
    let found = false;
    for (let i = 1; i < memData.length; i++) {
      if (String(memData[i][1]).trim() === name) {
        const cur = Number(memData[i][3]) || 0;
        applied.push({ idx: i, newCoins: cur + amount });
        logRows.push([today, name, amount, reason, '교사', '완료']);
        cacheKeys.push('pointlog_' + name);
        results.push({ name: name, ok: true, newCoins: cur + amount });
        found = true;
        break;
      }
    }
    if (!found) results.push({ name: name, ok: false, error: 'Member에서 관원을 찾지 못했소: ' + name });
  });

  // ★ 로그를 먼저 기록(용량 확보 포함) — 여기서 실패하면 잔액은 아직 그대로라
  //   교사가 재시도해도 이중 지급이 없다. (반대 순서였다면 로그 실패 시 잔액만 올라감)
  if (logRows.length) {
    const logSheet = SS.getSheetByName('PointLog');
    _ensureRows(logSheet, logRows.length);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 6).setValues(logRows);
  }

  // 잔액 반영 + Firestore 즉시 반영
  applied.forEach(function(a) {
    memSheet.getRange(a.idx + 1, 4).setValue(a.newCoins);
    memData[a.idx][3] = a.newCoins;
    _fsPushMemberRow(memData[a.idx]);
  });
  cacheRemove(cacheKeys);

  // 푸시 알림은 시트 반영 뒤에 몰아서 발송
  results.forEach(function(r) {
    if (!r.ok) return;
    try {
      const t = amount >= 0 ? '🪙 상점! +' + amount + '냥' : '🪙 벌점 ' + amount + '냥';
      _sendPushToStudent(r.name, t, reason);
    } catch(e) {}
  });

  const okCount = results.filter(function(r) { return r.ok; }).length;
  return { count: okCount, total: results.length, results: results };
}


/* =========================================================
   7-0b. 출결 직접 정정 (교사) — 상태 변경 + 엽전 자동 보정
   상태별 엽전: 출석 +1 / 지각 -1 / 결석 -2 / 공결·미출석 0
   변동량 = (새 상태 엽전) - (기존 상태 엽전) 만큼만 보정 기록
   ========================================================= */
function updateAttendance(p) {
  if (!p.date || !p.name || !p.newStatus) throw new Error('날짜, 이름, 새 상태가 필요하오.');
  const VALID = ['출석', '지각', '결석', '공결', '미출석'];
  if (VALID.indexOf(p.newStatus) < 0) throw new Error('알 수 없는 출결 상태: ' + p.newStatus);
  const name = String(p.name).trim();

  const attSheet = SS.getSheetByName('Attendance');
  if (!attSheet) throw new Error("'Attendance' 시트가 없소.");
  const data = getValidData(attSheet);
  let oldStatus = '미출석', rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === p.date && String(data[i][1]).trim() === name) {
      oldStatus = String(data[i][3] || '').trim() || '미출석';
      rowIdx = i;
      break;
    }
  }
  if (oldStatus === p.newStatus) return { success: true, changed: false, msg: '변경 없음' };

  function coinOf(s) {
    if (s === '출석') return COIN.PRESENT;
    if (s === '지각') return COIN.LATE;
    if (s === '결석') return COIN.ABSENT;
    return 0;   // 공결·미출석
  }
  const delta = coinOf(p.newStatus) - coinOf(oldStatus);

  // Attendance 시트 갱신
  if (p.newStatus === '미출석') {
    if (rowIdx >= 0) attSheet.deleteRow(rowIdx + 1);
  } else if (rowIdx >= 0) {
    attSheet.getRange(rowIdx + 1, 4).setValue(p.newStatus);
  } else {
    attSheet.appendRow([p.date, name, '', p.newStatus, '']);
  }

  // ★ '결석'에서 다른 상태로 정정 시, 트리거가 남긴 '결석' 감점 로그를 '사유인정'으로 마킹.
  //   정정의 delta 보정이 이미 -2를 보상하므로, 이후 결석사유를 등록해도
  //   _refundAbsenceIfPenalized가 같은 -2를 한 번 더 환불(+2 이중 지급)하지 않게 한다.
  if (oldStatus === '결석' && p.newStatus !== '결석') {
    const markSheet = SS.getSheetByName('PointLog');
    const markData  = getValidData(markSheet);
    for (let i = markData.length - 1; i >= 1; i--) {
      if (normalizeDateCell(markData[i][0]) === p.date &&
          String(markData[i][1]).trim() === name &&
          String(markData[i][3]).trim() === '결석' &&
          String(markData[i][5]).trim() === '완료') {
        markSheet.getRange(i + 1, 6).setValue('사유인정');
        break;
      }
    }
  }
  // ★ 반대로 다시 '결석'으로 되돌리면 마킹도 원복 — 이후 결석사유를 등록했을 때
  //   정당한 환불이 막히지 않도록 (정정 delta가 -2를 다시 적용했으므로 상태 일관)
  if (p.newStatus === '결석' && oldStatus !== '결석') {
    const markSheet = SS.getSheetByName('PointLog');
    const markData  = getValidData(markSheet);
    for (let i = markData.length - 1; i >= 1; i--) {
      if (normalizeDateCell(markData[i][0]) === p.date &&
          String(markData[i][1]).trim() === name &&
          String(markData[i][3]).trim() === '결석' &&
          String(markData[i][5]).trim() === '사유인정') {
        markSheet.getRange(i + 1, 6).setValue('완료');
        break;
      }
    }
  }

  let newCoins = null;
  if (delta !== 0) {
    newCoins = applyCoins(name, delta, '출결정정: ' + oldStatus + '→' + p.newStatus + ' (' + p.date + ')', '교사');
  }
  cacheRemove(['admin_att_' + p.date, 'checkin_' + name + '_' + p.date, 'role_암행어사_' + p.date, 'att_all_' + p.date]);
  return { success: true, changed: true, oldStatus: oldStatus, newStatus: p.newStatus, delta: delta, newCoins: newCoins };
}


/* =========================================================
   7-1. 관원 명부 등록 (학생 추가)
   Member 시트: A=학번 B=이름 C=역할 D=엽전 E=사면장 F=기숙사 G=반 H=alert
   ========================================================= */
function addMember(p) {
  const id   = String(p.id   || '').trim();
  const name = String(p.name || '').trim();
  if (!id)   throw new Error('학번을 입력하시오.');
  if (!name) throw new Error('이름을 입력하시오.');
  const track = (p.track === '공기업') ? '공기업' : '공무원';
  const dorm  = p.dorm ? 'Y' : '';

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);

  // 중복 검사: 학번 또는 이름이 이미 있으면 거부
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][0]).trim() === id)   throw new Error('이미 등록된 학번이오: ' + id);
    if (String(memData[i][1]).trim() === name) throw new Error('이미 등록된 이름이오: ' + name);
  }

  // 신규 관원: 역할 관원, 엽전 0, 사면장 0
  memSheet.appendRow([id, name, '관원', 0, 0, dorm, track, '']);
  cacheRemove(['member_all', 'admin_data', 'role_포도대장_v2']);

  // Firestore 즉시 반영 (appendRow는 onMemberSheetEdit(onEdit) 트리거를 발생시키지 않음)
  try {
    _fsSet('members', id, {
      id: id, name: name, role: '관원', coins: 0, amnesty: 0,
      dorm: dorm === 'Y', track: track, alert: ''
    });
  } catch(e) { Logger.log('[FS] addMember sync 오류: ' + e.message); }

  return { success: true, id: id, name: name };
}


/* =========================================================
   8. 조회 함수들
   ========================================================= */
function getAdminPlanner(p) {
  if (!p.date) throw new Error('날짜가 없소.');
  const CACHE_KEY = 'admin_planner_' + p.date;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const data   = getValidData(SS.getSheetByName('Planner'));
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === p.date) {
      result.push({ name: data[i][1], content: data[i][2], progress: data[i][3] });
    }
  }
  cacheSet(CACHE_KEY, result, CACHE_TTL.PLANNER);
  return result;
}

function getAdminAttendance(p) {
  if (!p.date) throw new Error('날짜가 없소.');
  const CACHE_KEY = 'admin_att_' + p.date;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;

  const memSheet   = SS.getSheetByName('Member');
  const memRows    = getValidData(memSheet).slice(1)
    .filter(function(r) {
      const name = String(r[1] || '').trim();
      return name !== '' && name !== '교사' && r[2] !== 'teacher';
    });
  const allMembers = memRows.map(function(r) { return String(r[1]).trim(); });
  const dormSet    = new Set(memRows.filter(isDormStudent).map(function(r) { return String(r[1]).trim(); }));

  const attSheet = SS.getSheetByName('Attendance');
  const attMap   = {};
  getValidData(attSheet).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === p.date) {
      attMap[String(r[1]).trim()] = {
        time:   normalizeTimeCell(r[2]),
        status: String(r[3] || '').trim(),
        reason: String(r[4] || '').trim()
      };
    }
  });

  const result = allMembers.map(function(name) {
    const att = attMap[name];
    return { name, time: att ? att.time : '-', status: att ? att.status : '미출석',
             reason: att ? att.reason : '', dorm: dormSet.has(name) };
  });
  cacheSet(CACHE_KEY, result, 60);
  return result;
}

/* =========================================================
   8-1. 범위 조회 (여러 날짜 → 1회 호출)
   프론트가 최근 5일/7일/한 달치를 날짜별로 수십 번 호출하던 것을
   시트 1회 읽기로 처리. 반환 형태: { 'yyyy-mm-dd': [일별 API와 동일한 배열] }
   ========================================================= */
function _listDates(start, end) {
  if (!start || !end) throw new Error('start/end 날짜가 필요하오.');
  const out = [];
  let d = new Date(start + 'T00:00:00+09:00');
  const stop = new Date(end + 'T00:00:00+09:00');
  let guard = 0;
  while (d <= stop && guard++ < 62) {   // 최대 62일 (약 두 달) 안전장치
    out.push(Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd'));
    d = new Date(d.getTime() + 86400000);
  }
  return out;
}

/** getAdminAttendance와 동일한 행 형태를 날짜별로 반환 */
function getAdminAttendanceRange(p) {
  const dates = _listDates(p.start, p.end);
  const memSheet = SS.getSheetByName('Member');
  const memRows  = getValidData(memSheet).slice(1)
    .filter(function(r) {
      const name = String(r[1] || '').trim();
      return name !== '' && name !== '교사' && r[2] !== 'teacher';
    });
  const allMembers = memRows.map(function(r) { return String(r[1]).trim(); });
  const dormSet    = new Set(memRows.filter(isDormStudent).map(function(r) { return String(r[1]).trim(); }));

  const dateSet  = new Set(dates);
  const attByDate = {};   // { date: { name: {time,status,reason} } }
  getValidData(SS.getSheetByName('Attendance')).slice(1).forEach(function(r) {
    const d = normalizeDateCell(r[0]);
    if (!dateSet.has(d)) return;
    if (!attByDate[d]) attByDate[d] = {};
    attByDate[d][String(r[1]).trim()] = {
      time:   normalizeTimeCell(r[2]),
      status: String(r[3] || '').trim(),
      reason: String(r[4] || '').trim()
    };
  });

  const result = {};
  dates.forEach(function(d) {
    const attMap = attByDate[d] || {};
    result[d] = allMembers.map(function(name) {
      const att = attMap[name];
      return { name: name, time: att ? att.time : '-', status: att ? att.status : '미출석',
               reason: att ? att.reason : '', dorm: dormSet.has(name) };
    });
  });
  return result;
}

/** getAdminPlanner와 동일한 행 형태를 날짜별로 반환 (기록 있는 학생만) */
function getAdminPlannerRange(p) {
  const dates   = _listDates(p.start, p.end);
  const dateSet = new Set(dates);
  const result  = {};
  dates.forEach(function(d) { result[d] = []; });
  getValidData(SS.getSheetByName('Planner')).slice(1).forEach(function(r) {
    const d = normalizeDateCell(r[0]);
    if (!dateSet.has(d)) return;
    result[d].push({ name: r[1], content: r[2], progress: r[3] });
  });
  return result;
}

/** getCounselingByDate와 동일한 행 형태를 날짜별로 반환 */
function getCounselingRange(p) {
  const dates = _listDates(p.start, p.end);
  const allMembers = getValidData(SS.getSheetByName('Member')).slice(1)
    .filter(function(r) { const name = String(r[1] || '').trim(); return name && name !== '교사' && r[2] !== 'teacher'; })
    .map(function(r) { return String(r[1]).trim(); });

  const dateSet  = new Set(dates);
  const cByDate  = {};   // { date: { name: {done,memo} } }
  const sheet    = SS.getSheetByName('Counseling');
  if (sheet) {
    getValidData(sheet).slice(1).forEach(function(r) {
      const d = normalizeDateCell(r[0]);
      if (!dateSet.has(d)) return;
      if (!cByDate[d]) cByDate[d] = {};
      cByDate[d][String(r[1]).trim()] = {
        done: String(r[2]).trim().toUpperCase() === 'Y',
        memo: String(r[3] || '').trim()
      };
    });
  }

  const result = {};
  dates.forEach(function(d) {
    const cMap = cByDate[d] || {};
    result[d] = allMembers.map(function(name) {
      const c = cMap[name] || { done: false, memo: '' };
      return { name: name, done: c.done, memo: c.memo };
    });
  });
  return result;
}

function getMyHistory(p) {
  if (!p.name) throw new Error('이름이 없소.');
  const CACHE_KEY = 'history_' + p.name;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const data   = getValidData(SS.getSheetByName('Planner'));
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === p.name) {
      result.push({ date: normalizeDateCell(data[i][0]), content: data[i][2], progress: Number(data[i][3]) || 0 });
    }
  }
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  cacheSet(CACHE_KEY, result, CACHE_TTL.PLANNER);
  return result;
}

function getMyPointLog(p) {
  if (!p.name) throw new Error('이름이 없소.');
  const CACHE_KEY = 'pointlog_' + p.name;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const data   = getValidData(SS.getSheetByName('PointLog'));
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === p.name) {
      result.push({
        date:   normalizeDateCell(data[i][0]),
        delta:  Number(data[i][2]) || 0,
        reason: data[i][3],
        actor:  data[i][4],
        status: data[i][5]
      });
    }
  }
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  cacheSet(CACHE_KEY, result, CACHE_TTL.POINT_LOG);
  return result;
}

function getClassStats() {
  const CACHE_KEY = 'stats_class';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const data  = getValidData(SS.getSheetByName('Planner'));
  const stats = {};
  for (let i = 1; i < data.length; i++) {
    const date     = normalizeDateCell(data[i][0]);
    const progress = Number(data[i][3]) || 0;
    if (!date) continue;
    if (!stats[date]) stats[date] = { sum: 0, count: 0 };
    stats[date].sum   += progress;
    stats[date].count += 1;
  }
  const result = Object.keys(stats).map(function(date) {
    return { date, avg: Math.round(stats[date].sum / stats[date].count) };
  });
  result.sort(function(a, b) { return a.date.localeCompare(b.date); });
  cacheSet(CACHE_KEY, result, CACHE_TTL.STATS);
  return result;
}

function getAllMemberPoints() {
  const CACHE_KEY = 'member_all';
  let memberData  = cacheGet(CACHE_KEY);
  if (!memberData) {
    memberData = getValidData(SS.getSheetByName('Member'));
    cacheSet(CACHE_KEY, memberData, CACHE_TTL.MEMBER);
  }
  return memberData.slice(1)
    .filter(function(r) {
      const name = String(r[1] || '').trim();
      return name !== '' && name !== '교사' && r[2] !== 'teacher';
    })
    .map(function(r) {
      return {
        name:    String(r[1]).trim(),
        role:    String(r[2] || '관원').trim(),
        coins:   Number(r[3]) || 0,
        amnesty: Number(r[4]) || 0,
        dorm:    isDormStudent(r),
        track:   String(r[6] || '').trim() || '공무원'
      };
    });
}




/* =========================================================
   10. 결석사유 관리
   ========================================================= */
function getAbsenceReasons(p) {
  const sheet = SS.getSheetByName('AbsenceReason');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const date = normalizeDateCell(data[i][0]);
    if (p && p.date && date !== p.date) continue;
    result.push({ row: i + 1, date, name: String(data[i][1]).trim(),
                  reason: String(data[i][2]).trim(), actor: String(data[i][3]).trim() });
  }
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return result;
}

function saveAbsenceReason(p) {
  if (!p.date || !p.name || !p.reason) throw new Error('날짜, 이름, 사유가 필요하오.');
  const sheet = SS.getSheetByName('AbsenceReason');
  if (!sheet) throw new Error("'AbsenceReason' 시트가 없소.");
  const data = getValidData(sheet);
  let saved = false;
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === p.date && String(data[i][1]).trim() === p.name) {
      sheet.getRange(i + 1, 3).setValue(p.reason);
      sheet.getRange(i + 1, 4).setValue(p.actor || '교사');
      saved = true;
      break;
    }
  }
  if (!saved) sheet.appendRow([p.date, p.name, p.reason, p.actor || '교사']);

  // ★ 이미 결석 차감(-2)이 적용됐다면 자동 환불 + 출결 '공결' 처리
  //   (교사는 보통 차감 후에야 사유를 알게 되므로, 수동 +2를 없앤다)
  const refundInfo = _refundAbsenceIfPenalized(p.date, p.name, p.actor || '교사');

  return { success: true, msg: saved ? '수정 완료' : '등록 완료',
           refunded: refundInfo.refunded, refund: refundInfo.refund || 0, newCoins: refundInfo.newCoins };
}

/**
 * 해당 날짜에 결석 차감(-2)이 이미 '완료'로 적용돼 있으면 자동 환불한다.
 * - 원래 결석 PointLog 행의 상태를 '사유인정'으로 바꿔 중복 환불을 막는다.
 * - applyCoins로 +|ABSENT|냥 보정(잔액·PointLog·Firestore 일괄) + Attendance를 '공결'로.
 */
function _refundAbsenceIfPenalized(date, name, actor) {
  const logSheet = SS.getSheetByName('PointLog');
  const logData  = getValidData(logSheet);
  for (let i = logData.length - 1; i >= 1; i--) {
    if (normalizeDateCell(logData[i][0]) === date &&
        String(logData[i][1]).trim() === name &&
        String(logData[i][3]).trim() === '결석' &&
        String(logData[i][5]).trim() === '완료') {
      logSheet.getRange(i + 1, 6).setValue('사유인정');   // 중복 환불 방지 마커
      const refund   = Math.abs(COIN.ABSENT);
      const newCoins = applyCoins(name, refund, '결석사유 인정(환불): ' + date, actor || '교사');
      _setAttendanceStatus(date, name, '공결');
      try { _sendPushToStudent(name, '🪙 결석 사유 인정', '+' + refund + '냥 환불되었어요. (' + date + ')'); } catch(e) {}
      return { refunded: true, refund: refund, newCoins: newCoins };
    }
  }
  return { refunded: false };
}

/** Attendance 시트의 (날짜,이름) 상태를 변경 (해당 행이 없으면 무시) */
function _setAttendanceStatus(date, name, status) {
  const attSheet = SS.getSheetByName('Attendance');
  if (!attSheet) return;
  const data = getValidData(attSheet);
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === date && String(data[i][1]).trim() === name) {
      attSheet.getRange(i + 1, 4).setValue(status);
      cacheRemove(['admin_att_' + date, 'checkin_' + name + '_' + date, 'att_all_' + date]);
      return;
    }
  }
}

function deleteAbsenceReason(p) {
  if (!p.date || !p.name) throw new Error('날짜와 이름이 필요하오.');
  const sheet = SS.getSheetByName('AbsenceReason');
  if (!sheet) throw new Error("'AbsenceReason' 시트가 없소.");
  const data = getValidData(sheet);
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizeDateCell(data[i][0]) === p.date && String(data[i][1]).trim() === p.name) {
      sheet.deleteRow(i + 1);
      // ★ 등록 때 자동환불(+2)·공결 처리가 됐다면 삭제 시 되돌린다.
      //   현재 출결이 '공결'(환불 흐름의 결과 상태)일 때만 — 교사가 그 사이
      //   출결을 다른 상태로 직접 바꿨다면 그 결정을 존중해 건드리지 않는다.
      const reclaimed = _reclaimAbsenceRefund(p.date, p.name, p.actor || '교사');
      return { success: true, reclaimed: reclaimed };
    }
  }
  throw new Error('해당 결석사유 기록을 찾지 못했소.');
}

/** 결석사유 삭제 시: '사유인정' 마킹을 '완료'로 되돌리고 환불(+2)을 회수, 출결을 '결석'으로 복귀 */
function _reclaimAbsenceRefund(date, name, actor) {
  // 현재 출결 상태 확인 — '공결'이 아니면 회수하지 않음
  const attData = getValidData(SS.getSheetByName('Attendance'));
  let curStatus = '';
  for (let i = 1; i < attData.length; i++) {
    if (normalizeDateCell(attData[i][0]) === date && String(attData[i][1]).trim() === name) {
      curStatus = String(attData[i][3] || '').trim();
      break;
    }
  }
  if (curStatus !== '공결') return false;

  const logSheet = SS.getSheetByName('PointLog');
  const logData  = getValidData(logSheet);
  for (let i = logData.length - 1; i >= 1; i--) {
    if (normalizeDateCell(logData[i][0]) === date &&
        String(logData[i][1]).trim() === name &&
        String(logData[i][3]).trim() === '결석' &&
        String(logData[i][5]).trim() === '사유인정') {
      logSheet.getRange(i + 1, 6).setValue('완료');   // 마킹 원복 → 이후 사유 재등록 시 환불 가능
      applyCoins(name, -Math.abs(COIN.ABSENT), '결석사유 삭제(환불 회수): ' + date, actor);
      _setAttendanceStatus(date, name, '결석');
      return true;
    }
  }
  return false;
}


/* =========================================================
   11. 저잣거리 - 상품 관리
   ========================================================= */
function getMarketInventory() {
  const CACHE_KEY = 'market_inventory';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheet = SS.getSheetByName('MarketInventory');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (!name) continue;
    result.push({ row: i + 1, name, price: Number(data[i][1]) || 0,
                  initQty: Number(data[i][2]) || 0, qty: Number(data[i][3]) || 0,
                  note: String(data[i][4] || '').trim() });
  }
  cacheSet(CACHE_KEY, result, CACHE_TTL.MARKET);
  return result;
}

function saveMarketItem(p) {
  if (!p.name || p.price === undefined) throw new Error('상품명과 가격이 필요하오.');
  const sheet = SS.getSheetByName('MarketInventory');
  if (!sheet) throw new Error("'MarketInventory' 시트가 없소.");
  const data = getValidData(sheet);
  if (p.row) {
    const rowNum = Number(p.row);
    // ★ 안전장치: 프론트가 본 상품과 시트 실제 행이 다르면(행 밀림) 오수정 방지
    if (p.expectName !== undefined) {
      const curName = String(sheet.getRange(rowNum, 1).getValue()).trim();
      if (curName !== String(p.expectName).trim())
        throw new Error('목록이 바뀌어 안전하게 수정할 수 없소. 새로고침 후 다시 시도하시오.');
    }
    sheet.getRange(rowNum, 1).setValue(p.name);
    sheet.getRange(rowNum, 2).setValue(p.price);
    sheet.getRange(rowNum, 3).setValue(p.initQty || 0);
    sheet.getRange(rowNum, 4).setValue(p.qty !== undefined ? p.qty : p.initQty || 0);
    sheet.getRange(rowNum, 5).setValue(p.note || '');
    cacheRemove(['market_inventory']);
    return { success: true, msg: '수정 완료' };
  }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === p.name) {
      sheet.getRange(i + 1, 2).setValue(p.price);
      sheet.getRange(i + 1, 3).setValue(p.initQty || data[i][2]);
      sheet.getRange(i + 1, 4).setValue(p.qty !== undefined ? p.qty : data[i][3]);
      sheet.getRange(i + 1, 5).setValue(p.note || data[i][4]);
      cacheRemove(['market_inventory']);
      return { success: true, msg: '수정 완료' };
    }
  }
  const qty = p.qty !== undefined ? p.qty : (p.initQty || 0);
  sheet.appendRow([p.name, p.price, p.initQty || qty, qty, p.note || '']);
  cacheRemove(['market_inventory']);
  return { success: true, msg: '등록 완료' };
}

function deleteMarketItem(p) {
  if (!p.row) throw new Error('행 번호가 필요하오.');
  const sheet = SS.getSheetByName('MarketInventory');
  if (!sheet) throw new Error("'MarketInventory' 시트가 없소.");
  // ★ 안전장치: 프론트가 본 상품과 시트 실제 행이 다르면(행 밀림) 오삭제 방지
  if (p.expectName !== undefined) {
    const curName = String(sheet.getRange(Number(p.row), 1).getValue()).trim();
    if (curName !== String(p.expectName).trim())
      throw new Error('목록이 바뀌어 안전하게 삭제할 수 없소. 새로고침 후 다시 시도하시오.');
  }
  sheet.deleteRow(Number(p.row));
  cacheRemove(['market_inventory']);
  return { success: true };
}


/* =========================================================
   12. 저잣거리 - 구매 처리
   ========================================================= */
function marketPurchase(p) {
  if (!p.studentName || !p.itemName || !p.qty) throw new Error('학생명, 상품명, 수량이 필요하오.');
  const qty = Number(p.qty);
  if (qty <= 0 || !Number.isInteger(qty)) throw new Error('수량이 올바르지 않소.');

  const invSheet = SS.getSheetByName('MarketInventory');
  if (!invSheet) throw new Error("'MarketInventory' 시트가 없소.");
  const invData = getValidData(invSheet);

  let itemRow = -1, itemPrice = 0, itemQty = 0;
  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === p.itemName) {
      itemRow = i + 1; itemPrice = Number(invData[i][1]) || 0; itemQty = Number(invData[i][3]) || 0;
      break;
    }
  }
  if (itemRow < 0) throw new Error('상품을 찾지 못했소: ' + p.itemName);
  if (itemQty < qty) throw new Error('재고가 부족하오. (현재 재고: ' + itemQty + '개)');

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   memRow = -1, coins = 0;
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][1]).trim() === p.studentName) {
      memRow = i + 1; coins = Number(memData[i][3]) || 0; break;
    }
  }
  if (memRow < 0) throw new Error('관원을 찾지 못했소: ' + p.studentName);

  const totalCost = itemPrice * qty;
  if (coins < totalCost) throw new Error('엽전이 부족하오. (보유: ' + coins + '냥, 필요: ' + totalCost + '냥)');

  invSheet.getRange(itemRow, 4).setValue(itemQty - qty);
  memSheet.getRange(memRow, 4).setValue(coins - totalCost);
  memData[memRow - 1][3] = coins - totalCost;

  const today = toDateStr();
  SS.getSheetByName('PointLog').appendRow([today, p.studentName, -totalCost, '저잣거리:' + p.itemName + '×' + qty, p.actor || '장원', '완료']);
  SS.getSheetByName('MarketLog').appendRow([today, p.studentName, p.itemName, qty, totalCost, p.actor || '장원', '완료']);

  cacheRemove(['market_inventory', 'market_log', 'member_all', 'admin_data', 'pointlog_' + p.studentName]);
  _fsPushMemberRow(memData[memRow - 1]);   // Firestore 즉시 반영 (구매 후 잔액)
  return { success: true, newCoins: coins - totalCost, newQty: itemQty - qty, totalCost };
}

/* =========================================================
   12-2. 저잣거리 - 여러 상품 한 번에 구매 (한 학생)
   p.items = [{ itemName, qty }, ...]
   재고·총액을 먼저 모두 검증 → 통과 시에만 일괄 처리 (부분 실패 방지)
   ========================================================= */
function marketPurchaseMulti(p) {
  if (!p.studentName) throw new Error('학생명이 필요하오.');
  if (!p.items || !p.items.length) throw new Error('상품을 1개 이상 선택하시오.');

  const invSheet = SS.getSheetByName('MarketInventory');
  if (!invSheet) throw new Error("'MarketInventory' 시트가 없소.");
  const invData = getValidData(invSheet);

  // 1) 요청 검증 + 주문 목록 구성 (재고 확인)
  const orders = [];   // { i(invData 인덱스), name, price, qty }
  let totalCost = 0;
  p.items.forEach(function(it) {
    const itemName = String(it.itemName || '').trim();
    const qty      = Number(it.qty);
    if (!itemName) return;
    if (!(qty > 0) || !Number.isInteger(qty)) throw new Error('수량이 올바르지 않소: ' + itemName);
    let found = -1, price = 0, stock = 0;
    for (let i = 1; i < invData.length; i++) {
      if (String(invData[i][0]).trim() === itemName) {
        found = i; price = Number(invData[i][1]) || 0; stock = Number(invData[i][3]) || 0; break;
      }
    }
    if (found < 0)     throw new Error('상품을 찾지 못했소: ' + itemName);
    if (stock < qty)   throw new Error(itemName + ' 재고가 부족하오. (현재 ' + stock + '개)');
    orders.push({ i: found, name: itemName, price: price, qty: qty });
    totalCost += price * qty;
  });
  if (!orders.length) throw new Error('구매할 상품이 없소.');

  // 2) 학생 엽전 확인
  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   memRow = -1, coins = 0;
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][1]).trim() === p.studentName) {
      memRow = i; coins = Number(memData[i][3]) || 0; break;
    }
  }
  if (memRow < 0)         throw new Error('관원을 찾지 못했소: ' + p.studentName);
  if (coins < totalCost)  throw new Error('엽전이 부족하오. (보유: ' + coins + '냥, 필요: ' + totalCost + '냥)');

  // 3) 일괄 처리: 로그 먼저 기록 → 엽전 차감 → 재고 차감 순.
  //    로그는 상품별 appendRow 반복 대신 블록 한 번에 기록 (속도).
  //    ★ 순서 이유: 그리드 용량 등으로 실패할 수 있는 블록 쓰기를 맨 앞에 둬서,
  //      실패 시 재고·잔액이 하나도 변하지 않게 (재시도 시 재고만 증발하는 사고 방지)
  const today = toDateStr();
  const plog  = SS.getSheetByName('PointLog');
  const mlog  = SS.getSheetByName('MarketLog');
  const plogRows = [], mlogRows = [];
  orders.forEach(function(o) {
    const cost = o.price * o.qty;
    plogRows.push([today, p.studentName, -cost, '저잣거리:' + o.name + '×' + o.qty, p.actor || '장원', '완료']);
    mlogRows.push([today, p.studentName, o.name, o.qty, cost, p.actor || '장원', '완료']);
  });
  if (plogRows.length) {
    _ensureRows(plog, plogRows.length);
    plog.getRange(plog.getLastRow() + 1, 1, plogRows.length, 6).setValues(plogRows);
  }
  if (mlogRows.length) {
    _ensureRows(mlog, mlogRows.length);
    mlog.getRange(mlog.getLastRow() + 1, 1, mlogRows.length, 7).setValues(mlogRows);
  }
  orders.forEach(function(o) {
    const newStock = (Number(invData[o.i][3]) || 0) - o.qty;
    invSheet.getRange(o.i + 1, 4).setValue(newStock);
    invData[o.i][3] = newStock;
  });

  memSheet.getRange(memRow + 1, 4).setValue(coins - totalCost);
  memData[memRow][3] = coins - totalCost;

  cacheRemove(['market_inventory', 'market_log', 'member_all', 'admin_data', 'pointlog_' + p.studentName]);
  _fsPushMemberRow(memData[memRow]);   // Firestore 즉시 반영
  return { success: true, newCoins: coins - totalCost, totalCost: totalCost, count: orders.length };
}


/* =========================================================
   13. 저잣거리 - 거래 내역 조회 & 취소
   ========================================================= */
function getMarketLog(p) {
  const CACHE_KEY = 'market_log';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheet = SS.getSheetByName('MarketLog');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ row: i + 1, date: normalizeDateCell(data[i][0]),
                  name: String(data[i][1]).trim(), item: String(data[i][2]).trim(),
                  qty: Number(data[i][3]) || 0, cost: Number(data[i][4]) || 0,
                  actor: String(data[i][5]).trim(), status: String(data[i][6]).trim() });
  }
  result.sort(function(a, b) { return b.date.localeCompare(a.date) || b.row - a.row; });
  cacheSet(CACHE_KEY, result, CACHE_TTL.MARKET);
  return result;
}

function cancelMarketLog(p) {
  if (!p.row) throw new Error('행 번호가 필요하오.');
  const rowNum   = Number(p.row);
  const logSheet = SS.getSheetByName('MarketLog');
  const logData  = getValidData(logSheet);
  const row      = logData[rowNum - 1];
  if (!row) throw new Error('해당 거래 기록을 찾지 못했소.');
  if (String(row[6]).trim() === '취소') throw new Error('이미 취소된 거래요.');
  // ★ 안전장치: 프론트가 본 거래와 시트 실제 행이 다르면(행 밀림) 오취소 방지
  if (p.expectName !== undefined && String(row[1]).trim() !== String(p.expectName).trim())
    throw new Error('목록이 바뀌어 안전하게 취소할 수 없소. 새로고침 후 다시 시도하시오.');

  const studentName = String(row[1]).trim();
  const itemName    = String(row[2]).trim();
  const qty         = Number(row[3]) || 0;
  const cost        = Number(row[4]) || 0;

  logSheet.getRange(rowNum, 7).setValue('취소');

  const invSheet = SS.getSheetByName('MarketInventory');
  const invData  = getValidData(invSheet);
  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]).trim() === itemName) {
      invSheet.getRange(i + 1, 4).setValue((Number(invData[i][3]) || 0) + qty);
      break;
    }
  }

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   refundRow = -1;
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][1]).trim() === studentName) {
      memData[i][3] = (Number(memData[i][3]) || 0) + cost;
      memSheet.getRange(i + 1, 4).setValue(memData[i][3]);
      refundRow = i;
      break;
    }
  }

  SS.getSheetByName('PointLog').appendRow([toDateStr(), studentName, cost, '저잣거리취소:' + itemName + '×' + qty, '교사', '완료']);
  cacheRemove(['market_inventory', 'market_log', 'member_all', 'admin_data', 'pointlog_' + studentName]);
  if (refundRow >= 0) _fsPushMemberRow(memData[refundRow]);   // Firestore 즉시 반영 (환불 후 잔액)
  return { success: true };
}


/* =========================================================
   14. 모의고사 성적
   ========================================================= */
function saveExamScores(p) {
  if (!p.examName) throw new Error('시험 이름이 없소.');
  if (!p.examDate) throw new Error('시험 날짜가 없소.');
  if (!p.scores || p.scores.length === 0) throw new Error('점수 데이터가 없소.');
  const sheet = SS.getSheetByName('ExamScores');
  if (!sheet) throw new Error("'ExamScores' 시트가 없소.");
  const existing     = getValidData(sheet);
  const rowsToDelete = [];
  for (let i = existing.length - 1; i >= 1; i--) {
    if (String(existing[i][0]).trim() === p.examName && normalizeDateCell(existing[i][1]) === p.examDate)
      rowsToDelete.push(i + 1);
  }
  rowsToDelete.forEach(function(rowNum) { sheet.deleteRow(rowNum); });
  p.scores.forEach(function(s) {
    sheet.appendRow([p.examName, p.examDate, s.name, s.korean, s.english, s.history, s.total, s.rank]);
  });
  if (p.classAvg) {
    sheet.appendRow([p.examName, p.examDate, '__avg__',
      p.classAvg.avgKorean, p.classAvg.avgEnglish, p.classAvg.avgHistory, p.classAvg.avgTotal, '']);
  }
  // ★ 공기업 전용 평균 행 (NCS 기준) — 공기업 학생 화면의 '반평균'이
  //   공무원 평균을 가리키던 문제 해결. 구버전 프론트는 안 보내므로 조건부.
  if (p.classAvgG) {
    sheet.appendRow([p.examName, p.examDate, '__avgG__',
      p.classAvgG.avgKorean, p.classAvgG.avgEnglish, p.classAvgG.avgHistory, p.classAvgG.avgTotal, '']);
  }
  const keysToRemove = ['exam_list'];
  p.scores.forEach(function(s) { keysToRemove.push('exam_' + s.name); });
  cacheRemove(keysToRemove);
  syncExamScoresAfterSave(p.examName, p.examDate, p.scores, p.classAvg, p.classAvgG);
  return { success: true, count: p.scores.length };
}

function getMyExamScores(p) {
  if (!p.name) throw new Error('이름이 없소.');
  const CACHE_KEY = 'exam_' + p.name;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheet = SS.getSheetByName('ExamScores');
  if (!sheet) return { exams: [], classAvg: [] };
  const data = getValidData(sheet);
  const myExams = [], avgMap = {}, avgGMap = {};
  for (let i = 1; i < data.length; i++) {
    const examName = String(data[i][0] || '').trim();
    const examDate = normalizeDateCell(data[i][1]);
    const who      = String(data[i][2] || '').trim();
    if (!examName || !examDate) continue;
    if (who === '__avg__' || who === '__avgG__') {
      const entry = { name: examName, date: examDate,
        avgKorean: Number(data[i][3])||0, avgEnglish: Number(data[i][4])||0,
        avgHistory: Number(data[i][5])||0, avgTotal: Number(data[i][6])||0 };
      if (who === '__avg__') avgMap[examName] = entry;
      else                   avgGMap[examName] = entry;   // 공기업(NCS) 평균
    } else if (who === p.name) {
      myExams.push({ name: examName, date: examDate,
        korean: Number(data[i][3])||0, english: Number(data[i][4])||0,
        history: Number(data[i][5])||0, total: Number(data[i][6])||0,
        rank: data[i][7] ? Number(data[i][7]) : null });
    }
  }
  myExams.sort(function(a, b) { return a.date.localeCompare(b.date); });
  const byDate    = function(a, b) { return a.date.localeCompare(b.date); };
  const classAvg  = Object.values(avgMap).sort(byDate);
  const classAvgG = Object.values(avgGMap).sort(byDate);
  const result    = { exams: myExams, classAvg, classAvgG };
  cacheSet(CACHE_KEY, result, CACHE_TTL.EXAM);
  return result;
}

/**
 * [수동 실행 유틸] 기존 시험들에 __avgG__(공기업 NCS 평균) 행 백필.
 * GAS 편집기에서 1회 실행 — 이미 __avgG__가 있는 시험은 건너뜀.
 * 공기업 학생 판별은 Member 시트 G열(반) 기준(현재 재학생만 정확).
 */
function backfillAvgG() {
  const sheet = SS.getSheetByName('ExamScores');
  if (!sheet) return Logger.log('ExamScores 시트 없음');
  const memData = getValidData(SS.getSheetByName('Member'));
  const gongiupSet = new Set();
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][6] || '').trim() === '공기업')
      gongiupSet.add(String(memData[i][1] || '').trim());
  }
  if (!gongiupSet.size) return Logger.log('공기업 학생 없음 → 백필 불필요');

  const data = getValidData(sheet);
  const exams = {};   // examName -> { date, sum, count, hasAvgG }
  for (let i = 1; i < data.length; i++) {
    const examName = String(data[i][0] || '').trim();
    const who      = String(data[i][2] || '').trim();
    if (!examName) continue;
    if (!exams[examName]) exams[examName] = { date: normalizeDateCell(data[i][1]), sum: 0, count: 0, hasAvgG: false };
    if (who === '__avgG__') exams[examName].hasAvgG = true;
    else if (gongiupSet.has(who)) { exams[examName].sum += Number(data[i][3]) || 0; exams[examName].count++; }
  }

  let added = 0;
  Object.keys(exams).forEach(function(name) {
    const e = exams[name];
    if (e.hasAvgG || e.count === 0) return;
    const ncsAvg = Math.round(e.sum / e.count * 10) / 10;
    sheet.appendRow([name, e.date, '__avgG__', ncsAvg, 0, 0, ncsAvg, '']);
    syncExamScoresAfterSave(name, e.date, [], null, { avgKorean: ncsAvg, avgEnglish: 0, avgHistory: 0, avgTotal: ncsAvg });
    added++;
    Logger.log('  ↳ __avgG__ 추가: ' + name + ' (NCS 평균 ' + ncsAvg + ', ' + e.count + '명)');
  });
  cacheRemove(['exam_list']);
  // 학생별 성적 캐시 무효화 (classAvgG 반영)
  const keys = [];
  for (let i = 1; i < memData.length; i++) {
    const nm = String(memData[i][1] || '').trim();
    if (nm) keys.push('exam_' + nm);
  }
  if (keys.length) cacheRemove(keys);
  Logger.log('[backfillAvgG] 완료: ' + added + '개 시험에 공기업 평균 추가');
}

function getExamList() {
  const CACHE_KEY = 'exam_list';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheet = SS.getSheetByName('ExamScores');
  if (!sheet) return [];
  const data = getValidData(sheet), examMap = {};
  for (let i = 1; i < data.length; i++) {
    const examName = String(data[i][0] || '').trim();
    const examDate = normalizeDateCell(data[i][1]);
    const who      = String(data[i][2] || '').trim();
    if (!examName || who === '__avg__' || who === '__avgG__') continue;
    if (!examMap[examName]) examMap[examName] = { name: examName, date: examDate, totalSum: 0, count: 0 };
    examMap[examName].totalSum += Number(data[i][6]) || 0;
    examMap[examName].count    += 1;
  }
  const result = Object.values(examMap).map(function(e) {
    return { name: e.name, date: e.date, count: e.count,
             avgTotal: e.count > 0 ? Math.round(e.totalSum / e.count * 10) / 10 : 0 };
  });
  result.sort(function(a, b) { return a.date.localeCompare(b.date); });
  cacheSet(CACHE_KEY, result, CACHE_TTL.EXAM);
  return result;
}


/* =========================================================
   15. 역대 선배 성적
   ========================================================= */
function saveLegacyScores(p) {
  if (!p.rows || p.rows.length === 0) throw new Error('업로드할 데이터가 없소.');
  const track     = (p.track === '공기업') ? '공기업' : '공무원';
  const sheetName = 'ExamHistory_' + track;
  const sheet     = SS.getSheetByName(sheetName);
  if (!sheet) throw new Error("'" + sheetName + "' 시트가 없소.");
  const existing = getValidData(sheet), rowsToDelete = [];
  for (var i = existing.length - 1; i >= 1; i--) {
    if (p.rows.some(function(r) {
      return r.cohort === String(existing[i][0]) && r.examName === String(existing[i][1]) && r.name === String(existing[i][2]);
    })) rowsToDelete.push(i + 1);
  }
  rowsToDelete.forEach(function(rowNum) { sheet.deleteRow(rowNum); });
  p.rows.forEach(function(r) {
    sheet.appendRow([r.cohort, r.examName, r.name, r.korean, r.english, r.history,
                     r.total || (r.korean + r.english + r.history), r.passed ? 'Y' : 'N']);
  });
  CacheService.getScriptCache().remove('legacy_stats_' + track);
  try { syncExamHistoryToFirestore(track); } catch(e) { Logger.log('[FS] legacy sync 오류: ' + e.message); }
  return { success: true, count: p.rows.length };
}

function getLegacyStats(p) {
  const track     = (p && p.track === '공기업') ? '공기업' : '공무원';
  const CACHE_KEY = 'legacy_stats_' + track;
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheetName = 'ExamHistory_' + track;
  const sheet     = SS.getSheetByName(sheetName);
  if (!sheet) return { examNames: [], dist: {}, cohorts: {} };
  const data = getValidData(sheet);
  if (data.length < 2) return { examNames: [], dist: {}, cohorts: {} };

  var rawMap = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var cohort = String(row[0]||'').trim(), examName = String(row[1]||'').trim(), name = String(row[2]||'').trim();
    if (!cohort || !examName || !name) continue;
    var korean = Number(row[3])||0, english = Number(row[4])||0, history = Number(row[5])||0;
    var total  = Number(row[6])||(korean+english+history);
    var passed = String(row[7]||'').trim().toUpperCase() === 'Y';
    if (!rawMap[examName])         rawMap[examName] = {};
    if (!rawMap[examName][cohort]) rawMap[examName][cohort] = [];
    rawMap[examName][cohort].push({ korean, english, history, total, passed });
  }
  var examNames = Object.keys(rawMap).sort();
  var dist = {}, cohorts = {};
  examNames.forEach(function(examName) {
    var cohortMap = rawMap[examName], allRows = [], passRows = [];
    Object.values(cohortMap).forEach(function(rows) {
      rows.forEach(function(r) { allRows.push(r); if (r.passed) passRows.push(r); });
    });
    function calcStat(arr, key) {
      if (!arr.length) return null;
      var vals = arr.map(function(r){return r[key];});
      var sum  = vals.reduce(function(a,b){return a+b;},0);
      return { min:Math.min.apply(null,vals), max:Math.max.apply(null,vals),
               avg:Math.round(sum/vals.length*10)/10, count:vals.length, allScores:vals };
    }
    function calcPassAvg(arr, key) {
      if (!arr.length) return null;
      return Math.round(arr.reduce(function(a,r){return a+r[key];},0)/arr.length*10)/10;
    }
    dist[examName] = {};
    ['korean','english','history','total'].forEach(function(key) {
      var stat = calcStat(allRows, key);
      if (stat) { stat.passAvg = calcPassAvg(passRows, key); stat.passCount = passRows.length; }
      dist[examName][key] = stat;
    });
    cohorts[examName] = Object.keys(cohortMap).sort().map(function(cohort) {
      var rows = cohortMap[cohort];
      var avg  = function(k){return Math.round(rows.reduce(function(s,r){return s+r[k];},0)/rows.length*10)/10;};
      return { cohort, avgKorean:avg('korean'), avgEnglish:avg('english'), avgHistory:avg('history'),
               avgTotal:avg('total'), count:rows.length, passCount:rows.filter(function(r){return r.passed;}).length };
    });
  });
  var result = { examNames, dist, cohorts };
  cacheSet(CACHE_KEY, result, 120);
  return result;
}


/* =========================================================
   16. 상담
   ========================================================= */
function saveCounseling(p) {
  if (!p.date || !p.studentName) throw new Error('날짜 또는 학생 이름이 없소.');
  var sheet = SS.getSheetByName('Counseling');
  if (!sheet) throw new Error("'Counseling' 시트가 없소.");
  var data = getValidData(sheet), found = false;
  for (var i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === p.date && String(data[i][1]).trim() === p.studentName) {
      sheet.getRange(i+1,3).setValue(p.done?'Y':'N');
      sheet.getRange(i+1,4).setValue(p.memo||'');
      sheet.getRange(i+1,5).setValue(new Date());
      found = true; break;
    }
  }
  if (!found) sheet.appendRow([p.date, p.studentName, p.done?'Y':'N', p.memo||'', new Date()]);
  // ★ 학생별 달력 캐시는 'counsel_student_<이름>_<연월>' 키 — 연월까지 포함해야 실제로 지워짐
  cacheRemove(['counsel_date_'+p.date,
               'counsel_student_'+p.studentName+'_'+String(p.date).slice(0,7),
               'counsel_all']);
  return { success: true };
}

function getCounselingByDate(p) {
  if (!p.date) throw new Error('날짜가 없소.');
  var CACHE_KEY = 'counsel_date_' + p.date;
  var cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  var memSheet   = SS.getSheetByName('Member');
  var allMembers = getValidData(memSheet).slice(1)
    .filter(function(r){var name=String(r[1]||'').trim();return name&&name!=='교사'&&r[2]!=='teacher';})
    .map(function(r){return String(r[1]).trim();});
  var sheet = SS.getSheetByName('Counseling'), cMap = {};
  if (sheet) {
    getValidData(sheet).slice(1).forEach(function(r) {
      if (normalizeDateCell(r[0]) === p.date)
        cMap[String(r[1]).trim()] = { done:String(r[2]).trim().toUpperCase()==='Y', memo:String(r[3]||'').trim() };
    });
  }
  var result = allMembers.map(function(name) {
    var c = cMap[name] || { done:false, memo:'' };
    return { name, done:c.done, memo:c.memo };
  });
  cacheSet(CACHE_KEY, result, 120);
  return result;
}

function getCounselingByStudent(p) {
  if (!p.studentName || !p.yearMonth) throw new Error('학생명 또는 연월이 없소.');
  var CACHE_KEY = 'counsel_student_' + p.studentName + '_' + p.yearMonth;
  var cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  var sheet = SS.getSheetByName('Counseling');
  if (!sheet) return [];
  var result = [];
  getValidData(sheet).slice(1).forEach(function(r) {
    var date = normalizeDateCell(r[0]);
    if (String(r[1]).trim() === p.studentName && date.startsWith(p.yearMonth))
      result.push({ date, done:String(r[2]).trim().toUpperCase()==='Y', memo:String(r[3]||'').trim() });
  });
  cacheSet(CACHE_KEY, result, 120);
  return result;
}

function getCounselingAll() {
  const CACHE_KEY = 'counsel_all';
  const cached    = cacheGet(CACHE_KEY);
  if (cached) return cached;
  const sheet = SS.getSheetByName('Counseling');
  if (!sheet) return [];
  const result = [];
  getValidData(sheet).slice(1).forEach(function(r) {
    const date = normalizeDateCell(r[0]);
    const name = String(r[1] || '').trim();
    if (!date || !name) return;
    result.push({ date, name,
      done:   String(r[2] || '').trim().toUpperCase() === 'Y',
      memo:   String(r[3] || '').trim(),
      status: String(r[2] || '').trim().toUpperCase() === 'Y' ? '완료' : '미완료'
    });
  });
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  cacheSet(CACHE_KEY, result, 300);
  return result;
}


/* =========================================================
   ★★★ 자동화 트리거 함수 3개 ★★★
   GAS 편집기 → 트리거 메뉴에서 직접 등록:
   1) dailyAbsencePenalty  → 매일 오전 9시 (결석 처리 -2엽전)
   2) dailyPlannerPenalty  → 매일 새벽 5시 (플래너미작성 -1엽전)
   3) weeklyAttendanceBonus → 매주 월요일 새벽 5시 (주간개근 +5엽전)
   ========================================================= */

/**
 * [트리거 1] 매일 09:00
 * 08:30까지 출석/지각 기록 없는 학생 → 결석 처리
 * AbsenceReason에 등록된 학생은 제외
 */
function dailyAbsencePenalty() { _withScriptLock(_dailyAbsencePenaltyCore); }
function _dailyAbsencePenaltyCore() {
  const today = toDateStr();
  Logger.log('[dailyAbsencePenalty] 실행: ' + today);

  // ★ 주말(토=6, 일=0)이면 실행하지 않음
  const dayOfWeek = new Date(today + 'T00:00:00+09:00').getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log('[dailyAbsencePenalty] 주말 → 건너뜀');
    return;
  }

  // ★ 트리거 일시정지 기간이면 실행하지 않음
  if (isTriggerPaused(today)) {
    Logger.log('[dailyAbsencePenalty] 일시정지 기간 → 건너뜀');
    return;
  }

  // ★ 공휴일이면 실행하지 않음 (일시정지 등록을 잊어도 전원 결석 -2 사고 방지)
  if (isKoreanHoliday(today)) {
    Logger.log('[dailyAbsencePenalty] 공휴일 → 건너뜀');
    return;
  }

  // ★ 같은 날 중복 실행 방지 (트리거 중복 등록·수동 재실행 대비)
  if (_jobAlreadyRan('absence', today)) {
    Logger.log('[dailyAbsencePenalty] 이미 오늘 실행됨 → 건너뜀');
    return;
  }

  // 결석사유 등록된 학생
  const absenceSheet = SS.getSheetByName('AbsenceReason');
  const excusedSet   = new Set();
  if (absenceSheet) {
    getValidData(absenceSheet).slice(1).forEach(function(r) {
      if (normalizeDateCell(r[0]) === today && String(r[1]).trim())
        excusedSet.add(String(r[1]).trim());
    });
  }

  // 오늘 출석/지각 기록이 있는 학생
  const attSheet   = SS.getSheetByName('Attendance');
  const checkedSet = new Set();
  getValidData(attSheet).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === today && String(r[1]).trim())
      checkedSet.add(String(r[1]).trim());
  });

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   count    = 0;

  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (checkedSet.has(name)) continue;
    if (excusedSet.has(name)) {
      // ★ '공결' 행을 남김 — 기록이 아예 없으면 스트릭이 끊기고 달력에 공백으로 보임
      //   (사후 등록 환불 경로가 '공결'로 남기는 것과 일관)
      attSheet.appendRow([today, name, '', '공결', '']);
      Logger.log('  ↳ 결석사유 면제(공결 기록): ' + name);
      continue;
    }

    const cur = Number(memData[i][3]) || 0;
    memSheet.getRange(i + 1, 4).setValue(cur + COIN.ABSENT);
    memData[i][3] = cur + COIN.ABSENT;
    // Attendance에 결석 기록
    attSheet.appendRow([today, name, '08:30:00', '결석', '']);
    // PointLog에 결석 기록
    SS.getSheetByName('PointLog').appendRow([today, name, COIN.ABSENT, '결석', '시스템', '완료']);
    _fsPushMemberRow(memData[i]);   // Firestore 즉시 반영
    count++;
    Logger.log('  ↳ 결석 처리: ' + name + ' (' + cur + ' → ' + (cur + COIN.ABSENT) + ')');
  }

  cacheRemove(['member_all', 'admin_data', 'att_all_' + today]);
  _markJobRan('absence', today);   // 끝까지 성공했을 때만 마킹 (중간 실패 시 재실행 가능)
  Logger.log('[dailyAbsencePenalty] 완료: ' + count + '명 처리');

  // ★ 결석 처리가 끝나 오늘 출결이 확정된 시점 — 교사에게 요약 알림 + 검토 화면 딥링크
  try {
    const finalAtt = getValidData(attSheet).slice(1)
      .filter(function(r) { return normalizeDateCell(r[0]) === today; });
    const sum = { 출석: 0, 지각: 0, 결석: 0, 공결: 0 };
    finalAtt.forEach(function(r) {
      const st = String(r[3] || '').trim();
      if (sum[st] !== undefined) sum[st]++;
    });
    const body = '출석 ' + sum.출석 + ' · 지각 ' + sum.지각 + ' · 결석 ' + sum.결석 + ' · 공결 ' + sum.공결 + ' — 탭하여 확인·정정';
    const link = 'https://tjddlf0224-sudo.github.io/Top_Class/?goto=att&date=' + today;
    _sendPushToStudent('교사', '📋 오늘 아침자습 출결 현황', body, link);
  } catch(e) { Logger.log('[dailyAbsencePenalty] 교사 알림 오류: ' + e.message); }
}


/**
 * [트리거 2] 매일 새벽 5시
 * 전날 플래너 미작성자 -1엽전
 * 결석 처리된 학생은 중복 차감 제외
 */
function dailyPlannerPenalty() { _withScriptLock(_dailyPlannerPenaltyCore); }
function _dailyPlannerPenaltyCore() {
  // ★ 대상일 = "전날" 고정. 새벽 5시 트리거가 6시를 넘겨 실행돼도
  //   (toDateStr의 6시 경계 때문에) 오늘 날짜로 잘못 잡아 전원 -1 되는 사고 방지.
  //   now-12h는 05~11시 실행 모두에서 정확히 전날을 가리킨다.
  const today = toDateStr(new Date(new Date().getTime() - 12 * 3600 * 1000));
  Logger.log('[dailyPlannerPenalty] 실행(대상일): ' + today);

  // ★ 주말(토=6, 일=0)이면 실행하지 않음
  const dayOfWeek = new Date(today + 'T00:00:00+09:00').getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log('[dailyPlannerPenalty] 주말 → 건너뜀');
    return;
  }

  // ★ 트리거 일시정지 기간이면 실행하지 않음
  if (isTriggerPaused(today)) {
    Logger.log('[dailyPlannerPenalty] 일시정지 기간 → 건너뜀');
    return;
  }

  // ★ 대상일이 공휴일이면 실행하지 않음
  if (isKoreanHoliday(today)) {
    Logger.log('[dailyPlannerPenalty] 공휴일 → 건너뜀');
    return;
  }

  // ★ 같은 대상일 중복 실행 방지
  if (_jobAlreadyRan('planner', today)) {
    Logger.log('[dailyPlannerPenalty] 이미 실행됨 → 건너뜀');
    return;
  }

  // 대상일에 플래너 제출한 학생
  const planSheet    = SS.getSheetByName('Planner');
  const submittedSet = new Set();
  getValidData(planSheet).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === today && String(r[1]).trim())
      submittedSet.add(String(r[1]).trim());
  });

  // 결석 처리된 학생(중복 차감 방지) + 이미 미작성 차감된 학생(행 단위 멱등성)
  // ★ '미출석' → '결석'으로 수정 (dailyAbsencePenalty가 '결석'으로 기록)
  const logSheet     = SS.getSheetByName('PointLog');
  const absentSet    = new Set();
  const penalizedSet = new Set();
  getValidData(logSheet).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) !== today) return;
    if (r[3] === '결석')         absentSet.add(String(r[1]).trim());
    if (r[3] === '플래너미작성') penalizedSet.add(String(r[1]).trim());
  });

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   count    = 0;

  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (submittedSet.has(name)) continue;
    if (absentSet.has(name)) {
      Logger.log('  ↳ 결석자 중복제외: ' + name);
      continue;
    }
    if (penalizedSet.has(name)) {
      Logger.log('  ↳ 이미 차감됨(멱등성): ' + name);
      continue;
    }

    const cur = Number(memData[i][3]) || 0;
    memSheet.getRange(i + 1, 4).setValue(cur + COIN.NO_PLANNER);
    memData[i][3] = cur + COIN.NO_PLANNER;
    logSheet.appendRow([today, name, COIN.NO_PLANNER, '플래너미작성', '시스템', '완료']);
    _fsPushMemberRow(memData[i]);   // Firestore 즉시 반영
    count++;
    Logger.log('  ↳ 플래너미작성: ' + name);
  }

  cacheRemove(['member_all', 'admin_data']);
  _markJobRan('planner', today);   // 끝까지 성공했을 때만 마킹
  Logger.log('[dailyPlannerPenalty] 완료: ' + count + '명 처리');
}


/**
 * [트리거 3] 매주 월요일 새벽 5시
 * 지난 주 월~금 출석 기록이 모두 '출석'인 학생 → +5엽전
 * 지각/결석 1회라도 있으면 해당 없음
 */
function weeklyAttendanceBonus() { _withScriptLock(_weeklyAttendanceBonusCore); }
function _weeklyAttendanceBonusCore() {
  const now   = new Date();
  const today = toDateStr(now);
  Logger.log('[weeklyAttendanceBonus] 실행: ' + today);

  // 지난 주 월~금 날짜 계산
  const weekDates = [];
  for (let d = 1; d <= 5; d++) {
    const target = new Date(now);
    target.setDate(now.getDate() - now.getDay() - (7 - d));
    weekDates.push(Utilities.formatDate(target, 'Asia/Seoul', 'yyyy-MM-dd'));
  }
  Logger.log('  주간 날짜: ' + weekDates.join(', '));

  // ★ 같은 주 중복 실행 방지 (이중 +5 지급 사고)
  if (_jobAlreadyRan('weeklyBonus', weekDates[0])) {
    Logger.log('[weeklyAttendanceBonus] 이미 실행됨 → 건너뜀');
    return;
  }

  const attSheet = SS.getSheetByName('Attendance');
  const attData  = getValidData(attSheet).slice(1);

  // ★ 수업일 판별: 그 날 출결 기록이 한 건이라도 있으면 수업일.
  //   공휴일·일시정지일은 아무 기록이 없으므로 개근 판정에서 제외
  //   → 공휴일 낀 주에도 나머지 수업일을 모두 출석하면 +5를 받는다.
  const schoolDays = weekDates.filter(function(d) {
    return attData.some(function(r) { return normalizeDateCell(r[0]) === d; });
  });
  Logger.log('  수업일: ' + (schoolDays.join(', ') || '없음'));
  if (!schoolDays.length) {
    Logger.log('[weeklyAttendanceBonus] 지난주 수업일 없음 → 건너뜀');
    return;
  }

  // 이미 이번 주기에 지급된 학생 (행 단위 멱등성)
  // ★ 날짜 '일치'가 아니라 '지난주 월요일 이후'로 판별 — 새벽 5시 실행(로그 날짜가
  //   toDateStr 경계로 일요일)과 6시 이후 재실행(월요일) 사이에도 가드가 유지되게.
  //   직전 주기의 지급 행은 그 전 일요일/월요일 날짜라 weekDates[0]보다 과거 → 안 걸림.
  const bonusGivenSet = new Set();
  getValidData(SS.getSheetByName('PointLog')).slice(1).forEach(function(r) {
    if (r[3] === '주간개근' && normalizeDateCell(r[0]) >= weekDates[0])
      bonusGivenSet.add(String(r[1]).trim());
  });

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  let   count    = 0;

  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (bonusGivenSet.has(name)) continue;

    // 지난 주 수업일 모두 '출석'인지 확인
    let allAttend = true;
    for (let di = 0; di < schoolDays.length; di++) {
      const rec = attData.find(function(r) {
        return normalizeDateCell(r[0]) === schoolDays[di] && String(r[1]).trim() === name;
      });
      if (!rec || rec[3] !== '출석') { allAttend = false; break; }
    }

    if (allAttend) {
      const cur = Number(memData[i][3]) || 0;
      memSheet.getRange(i + 1, 4).setValue(cur + COIN.WEEKLY_FULL);
      memData[i][3] = cur + COIN.WEEKLY_FULL;
      SS.getSheetByName('PointLog').appendRow([today, name, COIN.WEEKLY_FULL, '주간개근', '시스템', '완료']);
      _fsPushMemberRow(memData[i]);   // Firestore 즉시 반영
      try { _sendPushToStudent(name, '🎉 주간 개근!', '지난 주 개근으로 +5냥 받았어요!'); } catch(e) {}  // [E]
      count++;
      Logger.log('  ↳ 주간개근 +5엽전: ' + name);
    }
  }

  cacheRemove(['member_all', 'admin_data']);
  _markJobRan('weeklyBonus', weekDates[0]);   // 끝까지 성공했을 때만 마킹
  Logger.log('[weeklyAttendanceBonus] 완료: ' + count + '명 처리');
}


/* =========================================================
   17. 교과 선생님 QR 인증
   ========================================================= */
function verifyTeacher(p) {
  if (!p.name)  throw new Error('성함을 입력하시오.');
  if (!p.phone) throw new Error('전화번호 뒷 4자리를 입력하시오.');
  const sheet = SS.getSheetByName('Teachers');
  if (!sheet) throw new Error("'Teachers' 시트가 없소. 관리자에게 문의하시오.");
  const data = getValidData(sheet);
  for (let i = 1; i < data.length; i++) {
    const name  = String(data[i][0] || '').trim();
    const phone = String(data[i][1] || '').trim();
    if (name === p.name.trim() && phone === p.phone.trim())
      return { name, subject: String(data[i][2] || '').trim() };
  }
  throw new Error('일치하는 선생님 정보가 없소. 이름과 전화번호를 확인하시오.');
}


/* =========================================================
   트리거 일시정지 관리
   PropertiesService에 JSON 배열로 저장
   [{ start: 'yyyy-mm-dd', end: 'yyyy-mm-dd' }, ...]
   ========================================================= */

function getTriggerPauses() {
  const raw = PropertiesService.getScriptProperties().getProperty('trigger_pauses');
  if (!raw) return [];
  try { return JSON.parse(raw); } catch(e) { return []; }
}

function saveTriggerPause(p) {
  if (!p.start || !p.end) throw new Error('시작일과 종료일이 필요하오.');
  if (p.start > p.end)    throw new Error('종료일이 시작일보다 앞설 수 없소.');
  const list = getTriggerPauses();
  list.push({ start: p.start, end: p.end, label: p.label || '' });
  list.sort(function(a,b) { return a.start.localeCompare(b.start); });
  PropertiesService.getScriptProperties().setProperty('trigger_pauses', JSON.stringify(list));
  return { success: true, count: list.length };
}

function deleteTriggerPause(p) {
  const idx = Number(p.idx);
  const list = getTriggerPauses();
  if (idx < 0 || idx >= list.length) throw new Error('존재하지 않는 항목이오.');
  list.splice(idx, 1);
  PropertiesService.getScriptProperties().setProperty('trigger_pauses', JSON.stringify(list));
  return { success: true };
}

/** 오늘이 일시정지 기간인지 확인 */
function isTriggerPaused(today) {
  const list = getTriggerPauses();
  return list.some(function(p) { return p.start <= today && today <= p.end; });
}


/* =========================================================
   [유틸] 캐시 초기화
   ========================================================= */
function clearAllCache() {
  CacheService.getScriptCache().removeAll([
    'member_all', 'admin_data', 'stats_class', 'role_포도대장_v2',
    'market_inventory', 'market_log', 'counsel_all',
  ]);
  Logger.log('주요 캐시가 초기화되었소.');
}

/* =========================================================
   [1회용 유틸] 전체 학생 엽전 0으로 초기화
   ★ GAS 편집기에서 함수 선택 후 ▶실행 (doPost 미노출 — 실수 방지)
   - 모든 학생 엽전 = 0 (교사 제외)
   - 0이 아니던 학생은 PointLog에 '엽전리셋' 기록 남김(이력 일관성)
   - Firestore 즉시 반영 → 학생 화면도 실시간 0
   - 사면장(amnesty)은 건드리지 않음 (필요하면 아래 주석 해제)
   ========================================================= */
function resetAllCoins() {
  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  const logSheet = SS.getSheetByName('PointLog');
  const today    = toDateStr();
  let   count = 0;

  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;

    const prev = Number(memData[i][3]) || 0;
    if (prev !== 0) {
      logSheet.appendRow([today, name, -prev, '엽전리셋(0 초기화)', '교사', '완료']);
    }
    memSheet.getRange(i + 1, 4).setValue(0);   // D열 엽전 = 0
    memData[i][3] = 0;
    // memSheet.getRange(i + 1, 5).setValue(0); memData[i][4] = 0;  // ← 사면장도 0으로 하려면 주석 해제
    _fsPushMemberRow(memData[i]);              // Firestore 즉시 반영
    count++;
  }

  cacheRemove(['member_all', 'admin_data']);
  Logger.log('[resetAllCoins] ' + count + '명 엽전 0 초기화 완료');
  return count;
}

/* =========================================================
   ★ FIRESTORE 동기화 모듈
   =========================================================
   【스크립트 속성 설정 필수】
   GAS 편집기 → 프로젝트 설정 → 스크립트 속성 → 추가:
     FS_CLIENT_EMAIL = firebase-adminsdk-fbsvc@topclass-be740.iam.gserviceaccount.com
     FS_PRIVATE_KEY  = -----BEGIN PRIVATE KEY-----\n...(JSON의 private_key 값 전체)

   【트리거 등록】
     - syncAllToFirestore   : 매일 오전 3시 (시간 기반)
     - onMemberSheetEdit    : 스프레드시트 수정 시 (onEdit)
   ========================================================= */

const FS_PROJECT_ID = 'topclass-be740';

/** Firestore REST API용 OAuth2 액세스 토큰 발급 (50분 캐싱 — 엽전 실시간 push 시 재발급 최소화) */
function _getFsAccessToken() {
  const cached = CacheService.getScriptCache().get('fs_access_token');
  if (cached) return cached;

  const props       = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('FS_CLIENT_EMAIL');
  const privateKey  = props.getProperty('FS_PRIVATE_KEY').replace(/\\n/g, '\n');

  const now    = Math.floor(Date.now() / 1000);
  const header = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim  = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: clientEmail, sub: clientEmail,
    aud: 'https://oauth2.googleapis.com/token',
    scope: 'https://www.googleapis.com/auth/datastore',
    iat: now, exp: now + 3600
  }));
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(header + '.' + claim, privateKey)
  );
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + header + '.' + claim + '.' + sig
  });
  const token = JSON.parse(res.getContentText()).access_token;
  // 토큰 유효기간 60분 → 50분(3000초)만 캐싱하여 만료 여유 확보
  if (token) CacheService.getScriptCache().put('fs_access_token', token, 3000);
  return token;
}

/** 값 → Firestore Value 형식 변환 */
function _toFsValue(v) {
  if (v === null || v === undefined) return { nullValue: null };
  if (typeof v === 'boolean')        return { booleanValue: v };
  if (typeof v === 'number')         return { doubleValue: v };
  return { stringValue: String(v) };
}

/** Firestore 문서 upsert (PATCH) */
function _fsSet(collectionId, docId, fields) {
  const token    = _getFsAccessToken();
  const safeId   = String(docId).replace(/[\/\s#\[\]\*\?]/g, '_');
  const url      = 'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID
                 + '/databases/(default)/documents/' + collectionId + '/' + encodeURIComponent(safeId);
  const fsFields = {};
  Object.keys(fields).forEach(function(k) { fsFields[k] = _toFsValue(fields[k]); });
  UrlFetchApp.fetch(url, {
    method: 'patch', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ fields: fsFields }),
    muteHttpExceptions: true
  });
}

/**
 * Member 시트 한 행을 Firestore members/{학번}에 즉시 반영 (엽전 실시간 동기화)
 * row = [학번, 이름, 역할, 엽전, 사면장, 기숙사, 반, alert]
 * 엽전을 바꾸는 모든 함수에서 호출 → 아이들이 보는 잔액이 실시간 유지됨
 * Firestore 오류가 시트 처리(원본)를 막지 않도록 내부에서 try/catch.
 */
function _fsPushMemberRow(row) {
  try {
    if (!row) return;
    var id   = String(row[0] || '').trim();
    var name = String(row[1] || '').trim();
    if (!id || !name) return;
    _fsSet('members', id, {
      id:      id,
      name:    name,
      role:    String(row[2] || '관원').trim(),
      coins:   Number(row[3]) || 0,
      amnesty: Number(row[4]) || 0,
      dorm:    String(row[5] || '').trim().toUpperCase() === 'Y',
      track:   String(row[6] || '공무원').trim(),
      alert:   String(row[7] || '').trim()
    });
  } catch(e) {
    Logger.log('[FS] _fsPushMemberRow 오류(' + (row && row[1]) + '): ' + e.message);
  }
}

/** 이름으로 Member 행을 찾아 Firestore에 즉시 반영 (시트에서 최신값 재조회) */
function _fsPushMemberByName(name) {
  try {
    var memData = getValidData(SS.getSheetByName('Member'));
    for (var i = 1; i < memData.length; i++) {
      if (String(memData[i][1]).trim() === String(name).trim()) {
        _fsPushMemberRow(memData[i]);
        return;
      }
    }
  } catch(e) {
    Logger.log('[FS] _fsPushMemberByName 오류(' + name + '): ' + e.message);
  }
}

/** Firestore 컬렉션 전체 문서 삭제 */
function _fsDeleteCollection(collectionId) {
  const token = _getFsAccessToken();
  let   pageToken = '';
  do {
    const url  = 'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID
               + '/databases/(default)/documents/' + collectionId + '?pageSize=100'
               + (pageToken ? '&pageToken=' + pageToken : '');
    const res  = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    const body = JSON.parse(res.getContentText());
    if (!body.documents || body.documents.length === 0) break;
    body.documents.forEach(function(doc) {
      UrlFetchApp.fetch('https://firestore.googleapis.com/v1/' + doc.name, {
        method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
    });
    pageToken = body.nextPageToken || '';
  } while (pageToken);
}

/* ── 1. Member 시트 → Firestore members/ ── */
function syncMembersToFirestore() {
  const sheet = SS.getSheetByName('Member');
  if (!sheet) return Logger.log('[FS] Member 시트 없음');
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id  = String(row[0]).trim(), name = String(row[1]).trim();
    if (!id || !name) continue;
    _fsSet('members', id, {
      id: id, name: name,
      role:    String(row[2] || '관원').trim(),
      coins:   Number(row[3]) || 0,
      amnesty: Number(row[4]) || 0,
      dorm:    String(row[5] || '').trim().toUpperCase() === 'Y',
      track:   String(row[6] || '공무원').trim(),
      alert:   String(row[7] || '').trim()
    });
  }
  Logger.log('[FS] members 동기화 완료: ' + (data.length - 1) + '명');
}

/* ── 2. Teachers 시트 → Firestore teachers/ ── */
function syncTeachersToFirestore() {
  const sheet = SS.getSheetByName('Teachers');
  if (!sheet) return Logger.log('[FS] Teachers 시트 없음');
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    if (!name) continue;
    _fsSet('teachers', name, {
      name: name, phone: String(data[i][1] || '').trim(), subject: String(data[i][2] || '').trim()
    });
  }
  Logger.log('[FS] teachers 동기화 완료: ' + (data.length - 1) + '명');
}

/* ── 3. ExamScores 시트 → Firestore examScores/ ── */
function syncExamScoresToFirestore() {
  const sheet = SS.getSheetByName('ExamScores');
  if (!sheet) return Logger.log('[FS] ExamScores 시트 없음');
  const data = sheet.getDataRange().getValues();
  // ★ '전체 삭제 후 재작성' 금지 — 중간에 실패/타임아웃하면 성적 화면이 통째로 비어버림.
  //   upsert(덮어쓰기)로 먼저 다 반영하고, 시트에 없는 문서만 뒤에서 정리한다.
  const keep = new Set();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var examName = String(row[0] || '').trim();
    var examDate = normalizeDateCell(row[1]);
    var who      = String(row[2] || '').trim();
    if (!examName || !examDate || !who) continue;
    var isAvg = (who === '__avg__' || who === '__avgG__');
    var docId = (examName + '_' + examDate + '_' + who).replace(/[\/\s#\[\]\*\?]/g, '_');
    keep.add(docId);
    _fsSet('examScores', docId, {
      examName: examName, examDate: examDate,
      studentName: who, isAvg: isAvg,
      korean:  Number(row[3]) || 0, english: Number(row[4]) || 0,
      history: Number(row[5]) || 0, total:   Number(row[6]) || 0,
      rank:    isAvg ? 0 : (Number(row[7]) || 0)
    });
    count++;
  }
  // 시트에서 사라진 문서 정리 (upsert가 다 끝난 뒤라, 실패해도 최신 데이터는 이미 반영됨)
  try {
    const token = _getFsAccessToken();
    var pageToken = '';
    do {
      const url  = 'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID
                 + '/databases/(default)/documents/examScores?pageSize=300'
                 + (pageToken ? '&pageToken=' + pageToken : '');
      const res  = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
      const body = JSON.parse(res.getContentText());
      if (!body.documents || !body.documents.length) break;
      body.documents.forEach(function(doc) {
        const id = decodeURIComponent(String(doc.name).split('/').pop());
        if (!keep.has(id)) {
          UrlFetchApp.fetch('https://firestore.googleapis.com/v1/' + doc.name, {
            method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
          });
        }
      });
      pageToken = body.nextPageToken || '';
    } while (pageToken);
  } catch(e) { Logger.log('[FS] examScores 잔여 정리 오류(데이터는 안전): ' + e.message); }
  Logger.log('[FS] examScores 동기화 완료: ' + count + '행');
}

/* ── 4. ExamHistory 시트 → Firestore examHistory_공무원/공기업 ── */
function syncExamHistoryToFirestore(trackFilter) {
  var tracks = trackFilter ? [trackFilter] : ['공무원', '공기업'];
  tracks.forEach(function(track) {
    var sheetName = 'ExamHistory_' + track;
    var collName  = 'examHistory_' + track;
    var sheet     = SS.getSheetByName(sheetName);
    if (!sheet) return Logger.log('[FS] ' + sheetName + ' 없음');
    var data = sheet.getDataRange().getValues();
    _fsDeleteCollection(collName);
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var cohort  = String(row[0] || '').trim();
      var examName= String(row[1] || '').trim();
      var name    = String(row[2] || '').trim();
      if (!cohort || !examName || !name) continue;
      var korean  = Number(row[3]) || 0, english = Number(row[4]) || 0, history = Number(row[5]) || 0;
      var total   = Number(row[6]) || (korean + english + history);
      var passed  = String(row[7] || '').trim().toUpperCase() === 'Y';
      var docId   = (cohort + '_' + examName + '_' + name).replace(/[\/\s#\[\]\*\?]/g, '_');
      _fsSet(collName, docId, { cohort: cohort, examName: examName, name: name,
        korean: korean, english: english, history: history, total: total, passed: passed });
      count++;
    }
    Logger.log('[FS] ' + collName + ' 동기화 완료: ' + count + '행');
  });
}

/* ── 5. 성적 저장 시 해당 시험만 즉시 반영 ── */
function syncExamScoresAfterSave(examName, examDate, scores, classAvg, classAvgG) {
  try {
    scores.forEach(function(s) {
      var docId = (examName + '_' + examDate + '_' + s.name).replace(/[\/\s#\[\]\*\?]/g, '_');
      _fsSet('examScores', docId, {
        examName: examName, examDate: examDate, studentName: s.name, isAvg: false,
        korean: s.korean||0, english: s.english||0, history: s.history||0, total: s.total||0, rank: s.rank||0
      });
    });
    if (classAvg) {
      var avgId = (examName + '_' + examDate + '___avg__').replace(/[\/\s#\[\]\*\?]/g, '_');
      _fsSet('examScores', avgId, {
        examName: examName, examDate: examDate, studentName: '__avg__', isAvg: true,
        korean: classAvg.avgKorean||0, english: classAvg.avgEnglish||0,
        history: classAvg.avgHistory||0, total: classAvg.avgTotal||0, rank: 0
      });
    }
    if (classAvgG) {   // 공기업(NCS) 평균
      var avgGId = (examName + '_' + examDate + '___avgG__').replace(/[\/\s#\[\]\*\?]/g, '_');
      _fsSet('examScores', avgGId, {
        examName: examName, examDate: examDate, studentName: '__avgG__', isAvg: true,
        korean: classAvgG.avgKorean||0, english: classAvgG.avgEnglish||0,
        history: classAvgG.avgHistory||0, total: classAvgG.avgTotal||0, rank: 0
      });
    }
    Logger.log('[FS] examScores 부분 업데이트: ' + examName);
  } catch(e) {
    Logger.log('[FS] syncExamScoresAfterSave 오류: ' + e.message);
  }
}

/* ── 6. 전체 동기화 (매일 새벽 3시 트리거) ── */
function syncAllToFirestore() {
  Logger.log('=== Firestore 전체 동기화 시작 ===');
  try { syncMembersToFirestore();     } catch(e) { Logger.log('[FS] members 오류: ' + e.message); }
  try { syncTeachersToFirestore();    } catch(e) { Logger.log('[FS] teachers 오류: ' + e.message); }
  try { syncExamScoresToFirestore();  } catch(e) { Logger.log('[FS] examScores 오류: ' + e.message); }
  try { syncExamHistoryToFirestore(); } catch(e) { Logger.log('[FS] examHistory 오류: ' + e.message); }
  Logger.log('=== Firestore 전체 동기화 완료 ===');
}

/* ── 7. Member 시트 수정 시 해당 행만 즉시 반영 (onEdit 트리거) ── */
function onMemberSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'Member') return;
    var row = e.range.getRow();
    if (row < 2) return;
    var data = sheet.getRange(row, 1, 1, 8).getValues()[0];
    var id   = String(data[0]).trim(), name = String(data[1]).trim();
    if (!id || !name) return;
    _fsSet('members', id, {
      id: id, name: name,
      role:    String(data[2] || '관원').trim(),
      coins:   Number(data[3]) || 0,
      amnesty: Number(data[4]) || 0,
      dorm:    String(data[5] || '').trim().toUpperCase() === 'Y',
      track:   String(data[6] || '공무원').trim(),
      alert:   String(data[7] || '').trim()
    });
    Logger.log('[FS] member 즉시 반영: ' + name + '(' + id + ')');
  } catch(err) {
    Logger.log('[FS] onMemberSheetEdit 오류: ' + err.message);
  }
}

// PointLog 조회 (교사용)
// ★ 시트가 수천 행으로 커지면 전체 전송이 느려지므로 최근 500건만 반환
//   (row는 실제 시트 행 번호 그대로 → 수정/삭제 안전장치와 호환)
function getAdminPointLog(p) {
  const sheet = SS.getSheetByName('PointLog');
  if (!sheet) return [];
  const LIMIT = (p && Number(p.limit)) || 500;
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    result.push({
      row: i + 1,
      date: normalizeDateCell(row[0]),
      name: row[1], delta: row[2],
      reason: row[3], actor: row[4], status: row[5]
    });
  }
  return result.length > LIMIT ? result.slice(result.length - LIMIT) : result;
}

// PointLog 수정
function editPointLog(p) {
  const sheet = SS.getSheetByName('PointLog');
  const cur   = sheet.getRange(p.row, 1, 1, 6).getValues()[0];
  // ★ 안전장치: 프론트가 본 행과 시트 실제 행이 다르면(행 밀림) 오수정 방지
  //   (구버전 프론트는 expectName을 안 보내므로 그때는 기존 동작 유지)
  if (p.expectName !== undefined) {
    const okName  = String(cur[1]).trim() === String(p.expectName).trim();
    const okDelta = p.expectDelta === undefined || (Number(cur[2]) || 0) === Number(p.expectDelta);
    const okDate  = p.expectDate  === undefined || normalizeDateCell(cur[0]) === p.expectDate;
    if (!okName || !okDelta || !okDate)
      throw new Error('목록이 바뀌어 안전하게 수정할 수 없소. 새로고침 후 다시 시도하시오.');
  }
  // 잔액 반영은 시트의 실제 이름 기준 (프론트가 보낸 이름은 폴백)
  const name     = String(cur[1] || p.name || '').trim();
  const status   = String(cur[5] || '').trim();
  const oldDelta = Number(cur[2]) || 0;
  const diff     = Number(p.delta) - oldDelta;
  sheet.getRange(p.row, 3).setValue(p.delta);
  sheet.getRange(p.row, 4).setValue(p.reason);
  // 엽전 반영 — '대기'/'미승인' 행은 잔액에 적용된 적이 없으므로 생략
  //   (수정 후 승인하면 approveDeduct가 시트에서 delta를 다시 읽어 정상 반영)
  let editRow = -1;
  const memberSheet = SS.getSheetByName('Member');
  const members = memberSheet.getDataRange().getValues();
  if (diff !== 0 && status !== '대기' && status !== '미승인') {
    for (let i = 1; i < members.length; i++) {
      if (String(members[i][1]).trim() === name) {
        members[i][3] = (Number(members[i][3])||0) + diff;
        memberSheet.getRange(i+1, 4).setValue(members[i][3]);
        editRow = i;
        break;
      }
    }
  }
  cacheRemove(['member_all', 'admin_data', 'pointlog_' + name]);
  if (editRow >= 0) _fsPushMemberRow(members[editRow]);   // Firestore 즉시 반영
  return { success: true };
}

// PointLog 삭제 (엽전 복원)
function deletePointLog(p) {
  const sheet = SS.getSheetByName('PointLog');
  const cur   = sheet.getRange(p.row, 1, 1, 6).getValues()[0];
  // ★ 안전장치: 프론트가 본 행과 시트 실제 행이 다르면(트리거/동시편집으로 행 밀림) 오삭제 방지
  if (p.expectName !== undefined) {
    const okName  = String(cur[1]).trim() === String(p.expectName).trim();
    const okDelta = p.expectDelta === undefined || (Number(cur[2]) || 0) === Number(p.expectDelta);
    const okDate  = p.expectDate  === undefined || normalizeDateCell(cur[0]) === p.expectDate;
    if (!okName || !okDelta || !okDate)
      throw new Error('목록이 바뀌어 안전하게 삭제할 수 없소. 새로고침 후 다시 시도하시오.');
  }
  const delta  = Number(cur[2]) || 0;
  const name   = String(cur[1]).trim();
  const status = String(cur[5] || '').trim();
  sheet.deleteRow(p.row);
  // 엽전 복원 — '대기'/'미승인' 행은 잔액에 적용된 적이 없으므로 복원하지 않음
  //   (복원하면 적용된 적 없는 차감이 '되돌려져' 학생이 공짜 엽전을 얻음)
  const applied = status !== '대기' && status !== '미승인';
  let delRow = -1;
  const memberSheet = SS.getSheetByName('Member');
  const members = memberSheet.getDataRange().getValues();
  if (applied && delta !== 0) {
    for (let i = 1; i < members.length; i++) {
      if (String(members[i][1]).trim() === name) {
        members[i][3] = (Number(members[i][3])||0) - delta;
        memberSheet.getRange(i+1, 4).setValue(members[i][3]);
        delRow = i;
        break;
      }
    }
  }
  cacheRemove(['member_all', 'admin_data', 'pointlog_' + name]);
  if (delRow >= 0) _fsPushMemberRow(members[delRow]);   // Firestore 즉시 반영 (복원 후 잔액)
  return { success: true, restored: applied && delta !== 0 };
}

/* =========================================================
   PointLog 일괄 삭제 (한 번의 호출로 처리 → 빠름)
   p.rows = [시트행번호, ...]  (각 행의 엽전 변동량 복원)
   ========================================================= */
function deletePointLogBatch(p) {
  // p.items=[{row,name,delta}] 우선(안전장치용), 없으면 p.rows=[행번호] 호환
  let items = Array.isArray(p.items) ? p.items
            : (p.rows || []).map(function(r) { return { row: Number(r) }; });
  items = items.filter(function(it) { return Number(it.row) > 1; });
  if (!items.length) return { deleted: 0 };

  const sheet = SS.getSheetByName('PointLog');
  const all   = getValidData(sheet);   // 1회만 읽기

  // ★ 안전장치: 프론트가 본 행과 시트 실제 행이 다르면 전체 중단(부분 오삭제 방지)
  items.forEach(function(it) {
    const row = all[Number(it.row) - 1];
    if (!row) throw new Error('목록이 바뀌었소. 새로고침 후 다시 시도하시오.');
    if (it.name !== undefined) {
      const okName  = String(row[1]).trim() === String(it.name).trim();
      const okDelta = it.delta === undefined || (Number(row[2]) || 0) === Number(it.delta);
      if (!okName || !okDelta)
        throw new Error('목록이 바뀌어 안전하게 삭제할 수 없소. 새로고침 후 다시 시도하시오.');
    }
  });
  const rows = items.map(function(it) { return Number(it.row); });

  // 행별 (이름→복원할 엽전합) 집계
  // ★ '대기'/'미승인' 행은 잔액에 적용된 적이 없으므로 복원 대상에서 제외
  const restore = {};   // name -> 복원 delta 합 (삭제이므로 -원래delta)
  rows.forEach(function(r) {
    const row = all[r - 1];
    if (!row) return;
    const status = String(row[5] || '').trim();
    if (status === '대기' || status === '미승인') return;
    const name  = String(row[1] || '').trim();
    const delta = Number(row[2]) || 0;
    if (name) restore[name] = (restore[name] || 0) - delta;
  });

  // 큰 행부터 삭제 + 연속 구간은 deleteRows 한 번으로 (대량 삭제 시 잠금 점유 단축)
  rows.sort(function(a, b) { return b - a; });
  let ri = 0;
  while (ri < rows.length) {
    let rj = ri;
    while (rj + 1 < rows.length && rows[rj + 1] === rows[rj] - 1) rj++;
    sheet.deleteRows(rows[rj], rows[ri] - rows[rj] + 1);
    ri = rj + 1;
  }

  // 학생별 엽전 복원 + Firestore 즉시 반영 (1인당 1회)
  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  const cacheKeys = ['member_all', 'admin_data'];
  Object.keys(restore).forEach(function(name) {
    for (let i = 1; i < memData.length; i++) {
      if (String(memData[i][1]).trim() === name) {
        memData[i][3] = (Number(memData[i][3]) || 0) + restore[name];
        memSheet.getRange(i + 1, 4).setValue(memData[i][3]);
        _fsPushMemberRow(memData[i]);
        break;
      }
    }
    cacheKeys.push('pointlog_' + name);
  });
  cacheRemove(cacheKeys);
  return { deleted: rows.length };
}


/* =========================================================
   학생 엽전내역 수정요청 → 교사 승인
   EditRequest 시트: A요청일시 B학생명 C원날짜 D원변동량 E원사유
                     F요청유형(edit/delete) G희망변동량 H학생사유 I상태(대기/승인/반려)
   ========================================================= */
function _getEditRequestSheet() {
  let sheet = SS.getSheetByName('EditRequest');
  if (!sheet) {
    sheet = SS.insertSheet('EditRequest');
    sheet.appendRow(['요청일시', '학생명', '원날짜', '원변동량', '원사유', '요청유형', '희망변동량', '학생사유', '상태']);
  }
  return sheet;
}

/** 학생: 본인 엽전내역 1건에 대해 수정/삭제를 요청 */
function requestPointEdit(p) {
  if (!p.name)   throw new Error('이름이 없소.');
  if (!p.note)   throw new Error('요청 사유를 적어주시오.');
  const type = (p.type === 'delete') ? 'delete' : 'edit';
  if (type === 'edit') {
    if (p.newDelta === undefined || p.newDelta === null || p.newDelta === '' || isNaN(Number(p.newDelta)))
      throw new Error('희망하는 엽전 변동량을 입력하시오.');
  }
  const sheet = _getEditRequestSheet();
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([
    now, String(p.name).trim(),
    String(p.origDate || '').trim(),
    Number(p.origDelta) || 0,
    String(p.origReason || '').trim(),
    type,
    type === 'edit' ? Number(p.newDelta) : '',
    String(p.note).trim(),
    '대기'
  ]);
  try { _sendPushToStudent('교사', '📨 엽전 수정요청', String(p.name).trim() + ' 학생이 수정요청을 보냈어요.'); } catch(e) {}  // [C]
  return { success: true };
}

/** 교사: 대기중인 수정요청 목록 */
function getPointEditRequests() {
  const sheet = SS.getSheetByName('EditRequest');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8] || '').trim() !== '대기') continue;
    result.push({
      row:       i + 1,
      requestAt: String(data[i][0] || '').trim(),
      name:      String(data[i][1] || '').trim(),
      origDate:  normalizeDateCell(data[i][2]),
      origDelta: Number(data[i][3]) || 0,
      origReason:String(data[i][4] || '').trim(),
      type:      String(data[i][5] || 'edit').trim(),
      newDelta:  data[i][6] === '' ? null : (Number(data[i][6]) || 0),
      note:      String(data[i][7] || '').trim()
    });
  }
  result.sort(function(a, b) { return b.requestAt.localeCompare(a.requestAt); });
  return result;
}

/** 학생: 본인이 보낸 수정요청 목록 (대기/승인/반려 상태 추적용, 최근순) */
function getMyEditRequests(p) {
  if (!p.name) throw new Error('이름이 없소.');
  const sheet = SS.getSheetByName('EditRequest');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const name   = String(p.name).trim();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() !== name) continue;
    result.push({
      requestAt: String(data[i][0] || '').trim(),
      origDate:  normalizeDateCell(data[i][2]),
      origDelta: Number(data[i][3]) || 0,
      origReason:String(data[i][4] || '').trim(),
      type:      String(data[i][5] || 'edit').trim(),
      newDelta:  data[i][6] === '' ? null : (Number(data[i][6]) || 0),
      note:      String(data[i][7] || '').trim(),
      status:    String(data[i][8] || '대기').trim()
    });
  }
  result.sort(function(a, b) { return b.requestAt.localeCompare(a.requestAt); });
  return result;
}

/** 교사: 수정요청 승인(반영)/반려 */
function resolvePointEditRequest(p) {
  if (!p.row) throw new Error('요청 행이 없소.');
  const sheet = SS.getSheetByName('EditRequest');
  if (!sheet) throw new Error("'EditRequest' 시트가 없소.");
  const r = sheet.getRange(p.row, 1, 1, 9).getValues()[0];
  if (String(r[8] || '').trim() !== '대기') throw new Error('이미 처리된 요청이오.');

  const name      = String(r[1] || '').trim();
  const origDelta = Number(r[3]) || 0;
  const origReason= String(r[4] || '').trim();
  const type      = String(r[5] || 'edit').trim();
  const newDelta  = r[6] === '' ? null : (Number(r[6]) || 0);

  if (p.approve) {
    // 보정(compensating) 방식: 원래 기록은 두고, 차액만큼 새 기록을 더해 잔액을 맞춤
    let adjust = 0, reason = '';
    if (type === 'delete') {
      adjust = -origDelta;
      reason = '수정승인(취소): ' + origReason;
    } else {
      adjust = (newDelta || 0) - origDelta;
      reason = '수정승인: ' + origReason + ' → ' + newDelta + '냥';
    }
    if (adjust !== 0) {
      applyCoins(name, adjust, reason, '교사');   // 엽전·PointLog·캐시·Firestore 일괄 처리
    }
    sheet.getRange(p.row, 9).setValue('승인');
    try { _sendPushToStudent(name, '✅ 수정요청 반영', '엽전 수정요청이 승인됐어요.'); } catch(e) {}  // [B]
    return { success: true, status: '승인' };
  } else {
    sheet.getRange(p.row, 9).setValue('반려');
    try { _sendPushToStudent(name, '엽전 수정요청 결과', '요청이 반려됐어요.'); } catch(e) {}  // [B]
    return { success: true, status: '반려' };
  }
}


/* =========================================================
   방과후수업 일정표
   AfterSchool 시트:   A=id B=요일 C=시작 D=종료 E=과목 F=강사 G=장소 H=비고
   AfterSchoolEx 시트: A=id B=날짜 C=유형(휴강/특강) D=과목 E=강사 F=시작 G=종료 H=장소 I=비고
   - 주간 반복 시간표 + 날짜별 예외(휴강/특강)
   - 강사: Teachers 시트(이름+전화뒷4)로 로그인, 보기 전용 / 수정은 교사만
   - Firestore schedule/afterschool 문서로 실시간 동기화
   ========================================================= */
function _getAfterSchoolSheet() {
  let s = SS.getSheetByName('AfterSchool');
  if (!s) { s = SS.insertSheet('AfterSchool'); s.appendRow(['id','요일','시작','종료','과목','강사','장소','비고']); }
  return s;
}
function _getAfterSchoolExSheet() {
  let s = SS.getSheetByName('AfterSchoolEx');
  if (!s) { s = SS.insertSheet('AfterSchoolEx'); s.appendRow(['id','날짜','유형','과목','강사','시작','종료','장소','비고']); }
  return s;
}
function _asId() { return 'as' + Date.now() + Math.floor(Math.random() * 1000); }
function _hm(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'HH:mm');
  return String(v || '').trim();
}

/** 강사 로그인 (Teachers 시트 이름+전화뒷4 검증) */
function instructorLogin(p) {
  const t = verifyTeacher(p);   // {name, subject} 반환, 불일치 시 throw
  return { name: t.name, subject: t.subject, role: 'instructor' };
}

/** 전체 방과후 일정 조회 (주간 + 예외) */
function getAfterSchool() {
  const CACHE_KEY = 'afterschool';
  const cached = cacheGet(CACHE_KEY);
  if (cached) return cached;

  const weekly = getValidData(_getAfterSchoolSheet()).slice(1)
    .filter(function(r) { return String(r[0] || '').trim(); })
    .map(function(r) {
      return { id: String(r[0]).trim(), day: String(r[1] || '').trim(),
               start: _hm(r[2]), end: _hm(r[3]), subject: String(r[4] || '').trim(),
               teacher: String(r[5] || '').trim(), place: String(r[6] || '').trim(),
               note: String(r[7] || '').trim() };
    });
  const exceptions = getValidData(_getAfterSchoolExSheet()).slice(1)
    .filter(function(r) { return String(r[0] || '').trim(); })
    .map(function(r) {
      return { id: String(r[0]).trim(), date: normalizeDateCell(r[1]), type: String(r[2] || '').trim(),
               subject: String(r[3] || '').trim(), teacher: String(r[4] || '').trim(),
               start: _hm(r[5]), end: _hm(r[6]), place: String(r[7] || '').trim(),
               note: String(r[8] || '').trim() };
    });
  const result = { weekly: weekly, exceptions: exceptions };
  cacheSet(CACHE_KEY, result, 600);
  return result;
}

/** Firestore schedule/afterschool 문서에 통째로 반영 (실시간 구독용) */
function _pushAfterSchoolToFirestore() {
  try {
    const data = getAfterSchool();
    _fsSet('schedule', 'afterschool', {
      weekly:     JSON.stringify(data.weekly),
      exceptions: JSON.stringify(data.exceptions),
      updatedAt:  Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
    });
  } catch(e) { Logger.log('[FS] afterschool push 오류: ' + e.message); }
}

/** 주간 수업 추가/수정 (교사) */
function saveAfterSchoolClass(p) {
  if (!p.day)     throw new Error('요일이 필요하오.');
  if (!p.subject) throw new Error('과목명이 필요하오.');
  const sheet = _getAfterSchoolSheet();
  const data  = getValidData(sheet);
  const rowVals = [p.day, _hm(p.start), _hm(p.end), p.subject,
                   String(p.teacher || '').trim(), String(p.place || '').trim(), String(p.note || '').trim()];
  if (p.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(p.id).trim()) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([rowVals]);
        cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
        return { success: true, id: p.id };
      }
    }
    // 못 찾으면 = 클라이언트가 만든 새 id → 그 id로 신규 삽입 (낙관적 업데이트 지원)
    sheet.appendRow([String(p.id).trim()].concat(rowVals));
    cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
    return { success: true, id: String(p.id).trim() };
  }
  const id = _asId();
  sheet.appendRow([id].concat(rowVals));
  cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
  return { success: true, id: id };
}

/** 주간 수업 삭제 (교사) */
function deleteAfterSchoolClass(p) {
  if (!p.id) throw new Error('id가 필요하오.');
  const sheet = _getAfterSchoolSheet();
  const data  = getValidData(sheet);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.id).trim()) {
      sheet.deleteRow(i + 1);
      cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
      return { success: true };
    }
  }
  return { success: true, notFound: true };   // 이미 없으면 조용히 통과 (낙관적 삭제 대비)
}

/** 날짜 예외(휴강/특강) 추가/수정 (교사) */
function saveAfterSchoolException(p) {
  if (!p.date) throw new Error('날짜가 필요하오.');
  const type = (p.type === '특강') ? '특강' : '휴강';
  if (type === '휴강' && !p.subject) throw new Error('휴강할 과목을 지정하시오.');
  if (type === '특강' && !p.subject) throw new Error('특강 과목명이 필요하오.');
  const sheet = _getAfterSchoolExSheet();
  const data  = getValidData(sheet);
  const rowVals = [p.date, type, p.subject, String(p.teacher || '').trim(),
                   _hm(p.start), _hm(p.end), String(p.place || '').trim(), String(p.note || '').trim()];
  if (p.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(p.id).trim()) {
        sheet.getRange(i + 1, 2, 1, 8).setValues([rowVals]);
        cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
        return { success: true, id: p.id };
      }
    }
    // 못 찾으면 = 클라이언트가 만든 새 id → 그 id로 신규 삽입
    sheet.appendRow([String(p.id).trim()].concat(rowVals));
    cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
    return { success: true, id: String(p.id).trim() };
  }
  const id = _asId();
  sheet.appendRow([id].concat(rowVals));
  cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
  return { success: true, id: id };
}

/** 날짜 예외 삭제 (교사) */
function deleteAfterSchoolException(p) {
  if (!p.id) throw new Error('id가 필요하오.');
  const sheet = _getAfterSchoolExSheet();
  const data  = getValidData(sheet);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.id).trim()) {
      sheet.deleteRow(i + 1);
      cacheRemove(['afterschool']); _pushAfterSchoolToFirestore();
      return { success: true };
    }
  }
  return { success: true, notFound: true };   // 이미 없으면 조용히 통과
}


/* =========================================================
   푸시 알림 (FCM 웹 푸시)
   PushTokens 시트: A=token B=name C=updatedAt
   서비스계정(FS_CLIENT_EMAIL/FS_PRIVATE_KEY)으로 FCM v1 발송.
   ⚠️ GCP에서 'Firebase Cloud Messaging API'가 사용 설정돼 있어야 함.
   트리거 등록:
     - pushAfterSchoolNotice  : 매일 아침 (방과후 있는 날만, 반별 발송)
     - pushPlannerReminder    : 매일 저녁 21시 (플래너 미작성 독촉)
     - pushBeforeAfterSchool  : 10분마다 (수업 직전, 반별)
   ========================================================= */
function _getPushSheet() {
  let s = SS.getSheetByName('PushTokens');
  if (!s) { s = SS.insertSheet('PushTokens'); s.appendRow(['token', 'name', 'updatedAt']); }
  return s;
}

/** 클라이언트가 발급받은 FCM 토큰 저장(중복 제거) */
function savePushToken(p) {
  if (!p.token) throw new Error('토큰이 없소.');
  const sheet = _getPushSheet();
  const data  = getValidData(sheet);
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const tk    = String(p.token).trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tk) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[String(p.name || '').trim(), now]]);
      cacheRemove(['push_tokens_all']);
      return { success: true, updated: true };
    }
  }
  sheet.appendRow([tk, String(p.name || '').trim(), now]);
  cacheRemove(['push_tokens_all']);
  return { success: true };
}

/** FCM 발송용 OAuth 토큰 (firebase.messaging scope, 50분 캐싱) */
function _getFcmAccessToken() {
  const cached = CacheService.getScriptCache().get('fcm_access_token');
  if (cached) return cached;
  const props       = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('FS_CLIENT_EMAIL');
  const privateKey  = props.getProperty('FS_PRIVATE_KEY').replace(/\\n/g, '\n');
  const now    = Math.floor(Date.now() / 1000);
  const header = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim  = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: clientEmail, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600
  }));
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(header + '.' + claim, privateKey));
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + header + '.' + claim + '.' + sig
  });
  const token = JSON.parse(res.getContentText()).access_token;
  if (token) CacheService.getScriptCache().put('fcm_access_token', token, 3000);
  return token;
}

/** 토큰 목록[{row, token}]에 발송 (죽은 토큰은 자동 정리) */
function _sendPushTokens(tokenObjs, title, body, link) {
  if (!tokenObjs || !tokenObjs.length) return 0;
  const token   = _getFcmAccessToken();
  const url     = 'https://fcm.googleapis.com/v1/projects/' + FS_PROJECT_ID + '/messages:send';
  const baseUrl = 'https://tjddlf0224-sudo.github.io/Top_Class/';
  link = link || baseUrl;   // 지정하지 않으면 기존처럼 앱 첫 화면으로
  let sent = 0; const deadRows = [];
  tokenObjs.forEach(function(t) {
    if (!t.token) return;
    const payload = JSON.stringify({ message: {
      token: t.token,
      notification: { title: title, body: body },
      webpush: { notification: { icon: baseUrl + 'icon.png' }, fcmOptions: { link: link } }
    }});
    const resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token }, payload: payload, muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code === 200) sent++;
    else if (code === 404 && t.row) deadRows.push({ row: t.row, token: t.token });   // UNREGISTERED → 정리
    else if (code !== 404) Logger.log('[FCM] ' + code + ': ' + resp.getContentText().slice(0, 150));
  });
  if (deadRows.length) {
    const sheet = _getPushSheet();
    // ★ 캐시된 행 번호가 시트와 어긋났을 수 있으므로 토큰 일치 확인 후에만 삭제
    deadRows.sort(function(a, b) { return b.row - a.row; }).forEach(function(d) {
      if (String(sheet.getRange(d.row, 1).getValue()).trim() === d.token) sheet.deleteRow(d.row);
    });
    cacheRemove(['push_tokens_all']);
  }
  return sent;
}

/** 전체 발송 */
function _sendPushToAll(title, body, link) {
  const data = getValidData(_getPushSheet());
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const tk = String(data[i][0]).trim();
    if (tk) list.push({ row: i + 1, token: tk });
  }
  return _sendPushTokens(list, title, body, link);
}

/** 특정 이름(학생/교사)에게만 발송 — 본인 행동/대상 알림용 */
/** 푸시 토큰 시트 캐시 — 엽전 지급 등 잠금 잡은 액션 안에서 매번 전체 읽기 방지 */
function _getPushTokensCached() {
  const KEY = 'push_tokens_all';
  let data = cacheGet(KEY);
  if (!data) {
    data = getValidData(_getPushSheet());
    cacheSet(KEY, data, 600);
  }
  return data;
}

function _sendPushToStudent(name, title, body, link) {
  const data = _getPushTokensCached();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(name).trim()) {
      const tk = String(data[i][0]).trim();
      if (tk) list.push({ row: i + 1, token: tk });
    }
  }
  if (!list.length) return 0;
  return _sendPushTokens(list, title, body, link);
}

/** 수업의 반(트랙) 판별: 명시 track > 과목명 추론 > '공통' */
function _classTrack(c) {
  const tr = String((c && c.track) || '').trim();
  if (tr === '공무원' || tr === '공기업') return tr;
  const s = String((c && c.subject) || '');
  if (s.indexOf('공무원') >= 0) return '공무원';
  if (s.indexOf('공기업') >= 0) return '공기업';
  return '공통';
}

/** 특정 반(공무원/공기업) 학생 + 교사에게만 발송 */
function _sendPushToTrack(track, title, body, link) {
  const mem = getValidData(SS.getSheetByName('Member'));
  const names = new Set();
  for (let i = 1; i < mem.length; i++) {
    const name = String(mem[i][1] || '').trim();
    const tr   = String(mem[i][6] || '').trim() || '공무원';   // G열 = 반
    if (name && tr === track) names.add(name);
  }
  names.add('교사');   // 교사는 항상 수신
  const data = _getPushTokensCached();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (names.has(String(data[i][1] || '').trim())) {
      const tk = String(data[i][0]).trim();
      if (tk) list.push({ row: i + 1, token: tk });
    }
  }
  return _sendPushTokens(list, title, body, link);
}

/**
 * 방과후 일정 변경 시 관련 학생(track)+강사에게 즉시 푸시.
 * 프론트가 Firestore에 직접 저장(asPersist)한 직후 별도로 호출 —
 * 저장 자체는 이미 끝난 뒤라 이 호출이 실패해도 일정 데이터는 안전함(알림만 못 감).
 */
function notifyAfterSchoolChange(p) {
  const track = String(p.track || '').trim() || '공통';
  const title = String(p.title || '📅 방과후 일정 변경');
  const body  = String(p.body  || '').trim();
  if (!body) return { sent: 0 };

  let sent = 0;
  try { sent += (track === '공통') ? _sendPushToAll(title, body) : _sendPushToTrack(track, title, body); }
  catch(e) { Logger.log('[FCM] notifyAfterSchoolChange 학생발송 오류: ' + e.message); }

  const teacherName = String(p.teacherName || '').trim();
  // 강사명이 '교사'(담임)와 겹치면 _sendPushToTrack이 이미 보냈으므로 중복 발송 방지
  if (teacherName && teacherName !== '교사') {
    try { sent += _sendPushToStudent(teacherName, title, body); }
    catch(e) { Logger.log('[FCM] notifyAfterSchoolChange 강사발송 오류: ' + e.message); }
  }
  return { sent: sent };
}

/** Firestore schedule/afterschool 문서 읽기 (트리거에서 오늘 방과후 확인용) */
function _fsReadAfterSchool() {
  try {
    const token = _getFsAccessToken();
    const url = 'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID
              + '/databases/(default)/documents/schedule/afterschool';
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { weekly: [], exceptions: [] };
    const f = JSON.parse(res.getContentText()).fields || {};
    const w = (f.weekly && f.weekly.stringValue)     ? JSON.parse(f.weekly.stringValue)     : [];
    const e = (f.exceptions && f.exceptions.stringValue) ? JSON.parse(f.exceptions.stringValue) : [];
    return { weekly: w, exceptions: e };
  } catch(err) { Logger.log('[FCM] schedule read 오류: ' + err.message); return { weekly: [], exceptions: [] }; }
}

/** 오늘 날짜의 방과후 수업 목록 (주간 - 휴강 + 특강) */
function _getTodayAfterSchool(today) {
  const WD = ['일','월','화','수','목','금','토'];
  const wd = WD[new Date(today + 'T00:00:00+09:00').getDay()];
  const data = _fsReadAfterSchool();
  const cancels = (data.exceptions || []).filter(function(e) { return e.date === today && e.type === '휴강'; });
  // ★ 프론트 asResolveDay/_asItemsForDate와 동일하게 적용 기간(startDate~endDate) 반영 —
  //   기간 지난 반복수업으로 방과후 알림이 계속 발송되는 것 방지
  let items = (data.weekly || []).filter(function(c) { return c.day === wd; })
    .filter(function(c) { return (!c.startDate || today >= c.startDate) && (!c.endDate || today <= c.endDate); })
    .filter(function(c) { return !cancels.some(function(x) { return x.subject === c.subject; }); });
  (data.exceptions || []).filter(function(e) { return e.date === today && e.type === '특강'; })
    .forEach(function(e) { items.push(e); });
  return items;
}

/* (삭제됨) 아침 등교 안내(전체) — 주말·재량휴업일 오발송 방지를 위해 제거 */

/* =========================================================
   방과후 수업 출석 체크 (교사 전용, 기록용 — 엽전·상벌점에 영향 없음)
   AfterSchoolAttendance 시트: A=날짜 B=수업명 C=이름 D=상태(출석/결석) E=정정자 F=정정시각
   ========================================================= */
function _getAfterSchoolAttSheet() {
  let s = SS.getSheetByName('AfterSchoolAttendance');
  if (!s) { s = SS.insertSheet('AfterSchoolAttendance'); s.appendRow(['날짜','수업명','이름','상태','정정자','정정시각']); }
  return s;
}

/**
 * 특정 날짜의 방과후 수업별 출석 체크리스트 조회.
 * 대상 학생 = 그 수업의 트랙 기준(공통이면 전체 학생, 공무원/공기업이면 해당 반만).
 * 저장된 기록이 없으면 기본값 '출석'(대부분 참석하므로 결석자만 표시하는 쪽이 빠름).
 */
function getAfterSchoolAttendance(p) {
  if (!p.date) throw new Error('날짜가 없소.');
  const classes = _getTodayAfterSchool(p.date);
  if (!classes.length) return { date: p.date, classes: [] };

  const memRows = getValidData(SS.getSheetByName('Member')).slice(1)
    .filter(function(r) { const n = String(r[1] || '').trim(); return n && n !== '교사' && r[2] !== 'teacher'; });

  const saved = getValidData(_getAfterSchoolAttSheet()).slice(1)
    .filter(function(r) { return normalizeDateCell(r[0]) === p.date; });
  const savedMap = {};   // 수업명 -> { 이름: 상태 }
  saved.forEach(function(r) {
    const subj = String(r[1]).trim();
    if (!savedMap[subj]) savedMap[subj] = {};
    savedMap[subj][String(r[2]).trim()] = String(r[3]).trim();
  });

  const result = classes.map(function(c) {
    const track  = _classTrack(c);
    const roster = memRows
      .filter(function(r) { return track === '공통' || (String(r[6] || '').trim() || '공무원') === track; })
      .map(function(r) { return String(r[1]).trim(); })
      .sort(function(a, b) { return a.localeCompare(b, 'ko'); });
    const subjMap  = savedMap[c.subject] || {};
    const students = roster.map(function(name) {
      return { name: name, status: subjMap[name] || '출석' };
    });
    return { subject: c.subject, start: c.start, end: c.end, teacher: c.teacher, track: track, students: students };
  });
  return { date: p.date, classes: result };
}

/**
 * 방과후 수업 하나의 출석 상태 저장 — 그 날짜+수업명의 기존 기록을 통째로 교체.
 * p = { date, subject, records: [{name, status}], actor }
 */
function saveAfterSchoolAttendance(p) {
  if (!p.date || !p.subject) throw new Error('날짜와 수업명이 필요하오.');
  if (!Array.isArray(p.records)) throw new Error('출석 기록이 없소.');
  const sheet = _getAfterSchoolAttSheet();
  const data  = getValidData(sheet);
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizeDateCell(data[i][0]) === p.date && String(data[i][1]).trim() === p.subject)
      rowsToDelete.push(i + 1);
  }
  rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });

  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const actor = String(p.actor || '교사').trim();
  const rows  = p.records
    .filter(function(r) { return r && r.name; })
    .map(function(r) { return [p.date, p.subject, String(r.name).trim(), String(r.status || '출석').trim(), actor, now]; });
  if (rows.length) {
    _ensureRows(sheet, rows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
  return { success: true, count: rows.length };
}

/**
 * [트리거] 매일 오전 8시대 — 어제 방과후 수업이 있었다면 교사에게 출석 체크 알림.
 * 아침자습(오늘 기준)과 달리 '어제' 수업을 다루므로, 어제 날짜는 06시 경계와 무관하게
 * 현재 시각에서 24시간을 빼 계산한다(이 트리거 자체가 06시보다 한참 뒤인 8~9시 실행이라 안전).
 */
/* =========================================================
   ★ 방과후 학생 셀프 출석체크 (v2 앱 기능)
   수업 시작 18:30 / 마감 당일 24:00, 전부 한국시간.
   결석은 저장하지 않고 "체크인 기록이 없으면 결석"으로 계산한다
   → 자정이 지나면 자동으로 확정되므로 마감 트리거가 필요 없다.

   ⚠️ AS_SELFCHECK_ENABLED
     학생들이 아직 구 앱을 쓰는 동안에는 체크인 기록이 0건이라
     아침 알림이 "전원 결석"으로 잘못 나간다. 그래서 기본값을 false로 두고,
     v2를 학생들에게 여는 날 true로 바꾼다(이 한 줄만 바꾸면 됨).
   ========================================================= */
const AS_SELFCHECK_ENABLED = false;
const V2_URL = 'https://tjddlf0224-sudo.github.io/Top_Class/v2/';

/** 특정 날짜의 방과후 체크인 기록 조회 (Firestore asCheckIn) */
function _fsReadCheckIns(dateStr) {
  try {
    const token   = _getFsAccessToken();
    // 체크인 가능 구간은 그날 18:30 ~ 다음날 00:00 (KST)
    const from    = new Date(dateStr + 'T18:30:00+09:00');
    const nextDay = new Date(new Date(dateStr + 'T00:00:00+09:00').getTime() + 86400000);
    const to      = new Date(Utilities.formatDate(nextDay, 'Asia/Seoul', 'yyyy-MM-dd') + 'T00:00:00+09:00');

    const res = UrlFetchApp.fetch(
      'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID +
      '/databases/(default)/documents:runQuery',
      { method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ structuredQuery: {
          from: [{ collectionId: 'asCheckIn' }],
          where: { compositeFilter: { op: 'AND', filters: [
            { fieldFilter: { field: { fieldPath: 'at' }, op: 'GREATER_THAN_OR_EQUAL',
                             value: { timestampValue: from.toISOString() } } },
            { fieldFilter: { field: { fieldPath: 'at' }, op: 'LESS_THAN',
                             value: { timestampValue: to.toISOString() } } }
          ]}}
        }}),
        muteHttpExceptions: true });

    if (res.getResponseCode() !== 200) {
      Logger.log('[FS] asCheckIn 조회 실패: ' + res.getContentText().slice(0, 200));
      return [];
    }
    const out = [];
    (JSON.parse(res.getContentText()) || []).forEach(function(r) {
      if (!r.document || !r.document.fields) return;
      const f = r.document.fields;
      out.push({ name:    (f.name    && f.name.stringValue)    || '',
                 subject: (f.subject && f.subject.stringValue) || '' });
    });
    return out;
  } catch(e) {
    Logger.log('[FS] asCheckIn 조회 오류: ' + e.message);
    return [];
  }
}

/**
 * [트리거] 매일 18:30 — 방과후 수업 시작, 학생에게 출석체크 알림
 * ⚠️ v2를 학생들에게 열기 전까지는 이 트리거를 등록하지 말 것.
 */
function pushAfterSchoolCheckIn() {
  const today = toDateStr();
  const dow   = new Date(today + 'T00:00:00+09:00').getDay();
  if (dow === 0 || dow === 6) { Logger.log('[방과후체크인알림] 주말 → 건너뜀'); return; }
  if (isTriggerPaused(today)) { Logger.log('[방과후체크인알림] 일시정지 → 건너뜀'); return; }
  if (isKoreanHoliday(today)) { Logger.log('[방과후체크인알림] 공휴일 → 건너뜀'); return; }
  if (_jobAlreadyRan('asCheckInPush', today)) { Logger.log('[방과후체크인알림] 이미 실행됨 → 건너뜀'); return; }

  const classes = _getTodayAfterSchool(today);
  if (!classes.length) {
    Logger.log('[방과후체크인알림] 오늘 방과후 수업 없음');
    _markJobRan('asCheckInPush', today);
    return;
  }

  const byTrack = { '공무원': [], '공기업': [], '공통': [] };
  classes.forEach(function(c) { byTrack[_classTrack(c)].push(c.subject); });

  const link = V2_URL + '?goto=after';
  const msg  = function(subs) { return subs.join(', ') + ' — 앱에서 출석을 체크하세요 (오늘 24시 마감)'; };

  if (byTrack['공통'].length)   _sendPushToAll('✅ 방과후 출석체크', msg(byTrack['공통']), link);
  if (byTrack['공무원'].length) _sendPushToTrack('공무원', '✅ 방과후 출석체크', msg(byTrack['공무원']), link);
  if (byTrack['공기업'].length) _sendPushToTrack('공기업', '✅ 방과후 출석체크', msg(byTrack['공기업']), link);

  _markJobRan('asCheckInPush', today);
  Logger.log('[방과후체크인알림] 발송: ' + classes.map(function(c){ return c.subject; }).join(', '));
}

/** [트리거] 다음날 아침 — 교사에게 어제 방과후 출결 알림 */
function pushAfterSchoolAttendanceCheck() {
  const yesterday = toDateStr(new Date(new Date().getTime() - 24 * 3600 * 1000));

  if (isTriggerPaused(yesterday)) { Logger.log('[방과후출석체크] 일시정지 기간 → 건너뜀'); return; }
  if (_jobAlreadyRan('asAttCheck', yesterday)) { Logger.log('[방과후출석체크] 이미 실행됨 → 건너뜀'); return; }

  const classes = _getTodayAfterSchool(yesterday);
  if (!classes.length) {
    Logger.log('[방과후출석체크] 어제(' + yesterday + ') 방과후 수업 없음');
    _markJobRan('asAttCheck', yesterday);
    return;
  }

  let title, body, link;

  if (!AS_SELFCHECK_ENABLED) {
    // 구 앱 운영 중 — 종전대로 "체크해주세요"만. (체크인 기록이 없으므로 결석 집계를 하면 안 됨)
    title = '📝 방과후 출석 체크';
    body  = yesterday + ' 방과후 수업(' + classes.map(function(c){ return c.subject; }).join(', ') + ') 출석을 체크해주세요.';
    link  = 'https://tjddlf0224-sudo.github.io/Top_Class/?goto=asatt&date=' + yesterday;
  } else {
    // v2 운영 중 — 학생 셀프 체크인 결과를 집계해 결석자를 바로 알려준다
    const checked = {};   // 수업명 -> { 이름: true }
    _fsReadCheckIns(yesterday).forEach(function(r) {
      if (!checked[r.subject]) checked[r.subject] = {};
      checked[r.subject][r.name] = true;
    });

    const memRows = getValidData(SS.getSheetByName('Member')).slice(1)
      .filter(function(r) { const n = String(r[1] || '').trim(); return n && n !== '교사' && r[2] !== 'teacher'; });

    const lines = classes.map(function(c) {
      const track  = _classTrack(c);
      const roster = memRows
        .filter(function(r) { return track === '공통' || (String(r[6] || '').trim() || '공무원') === track; })
        .map(function(r) { return String(r[1]).trim(); });
      const map    = checked[c.subject] || {};
      const absent = roster.filter(function(n) { return !map[n]; });
      return absent.length
        ? c.subject + ' 결석 ' + absent.length + '명(' + absent.join(', ') + ')'
        : c.subject + ' 전원 출석';
    });

    title = '📝 어제 방과후 출결';
    body  = yesterday + ' · ' + lines.join(' / ');
    link  = V2_URL + '?goto=after&date=' + yesterday;
  }

  try { _sendPushToStudent('교사', title, body, link); }
  catch(e) { Logger.log('[방과후출석체크] 알림 오류: ' + e.message); }
  _markJobRan('asAttCheck', yesterday);
}

/* [트리거] 방과후 있는 날 아침 — 반별로 안내 (해당 반 학생 + 교사) */
function pushAfterSchoolNotice() {
  const today = toDateStr();
  const dow = new Date(today + 'T00:00:00+09:00').getDay();
  if (dow === 0 || dow === 6) return;
  const classes = _getTodayAfterSchool(today);
  if (!classes.length) return;

  const byTrack = { '공무원': [], '공기업': [], '공통': [] };
  classes.forEach(function(c) { byTrack[_classTrack(c)].push(c.subject); });

  if (byTrack['공통'].length)   _sendPushToAll('📅 오늘 방과후 수업', '오늘 방과후: ' + byTrack['공통'].join(', '));
  if (byTrack['공무원'].length) _sendPushToTrack('공무원', '📅 오늘 방과후 수업', '오늘 방과후(공무원): ' + byTrack['공무원'].join(', '));
  if (byTrack['공기업'].length) _sendPushToTrack('공기업', '📅 오늘 방과후 수업', '오늘 방과후(공기업): ' + byTrack['공기업'].join(', '));
}

/* 테스트용: 직접 실행하면 본인 기기들에 테스트 푸시 발송 */
function pushTest() {
  const n = _sendPushToAll('🔔 테스트 알림', '푸시가 정상 동작합니다!');
  Logger.log('테스트 발송: ' + n + '건');
}


/* [트리거-D] 매일 저녁 21시 — 플래너 미작성자 독촉 */
function pushPlannerReminder() {
  const today = toDateStr();
  const dow = new Date(today + 'T00:00:00+09:00').getDay();
  if (dow === 0 || dow === 6) return;
  if (isTriggerPaused(today)) return;
  const submitted = new Set();
  getValidData(SS.getSheetByName('Planner')).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === today && String(r[1]).trim()) submitted.add(String(r[1]).trim());
  });
  const mem = getValidData(SS.getSheetByName('Member'));
  let n = 0;
  for (let i = 1; i < mem.length; i++) {
    const name = String(mem[i][1] || '').trim();
    const role = String(mem[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (submitted.has(name)) continue;
    try { _sendPushToStudent(name, '📝 플래너 미작성', '오늘 플래너 작성했나요? 100% 달성하면 +2냥!'); n++; } catch(e) {}
  }
  Logger.log('[D] 플래너 독촉 ' + n + '명');
}

/* [트리거] 아침 등교 시간 — 아직 등교 인증 안 한 학생에게 푸시 알림.
   GAS 트리거로 등록: 시계 아이콘 → pushCheckInReminder → 시간 기반 → 일 단위 → '오전 8시~9시'.
   ★ 08:30 마감 이후엔 발송 안 함(독촉 무의미) + 당일 1회 가드(중복 방지).
   주말·공휴일·일시정지 제외, 이미 출석했거나 결석사유 등록된 학생 제외. */
function pushCheckInReminder() {
  const today = toDateStr();
  const dow = new Date(today + 'T00:00:00+09:00').getDay();
  if (dow === 0 || dow === 6) return;
  if (isTriggerPaused(today)) return;
  if (isKoreanHoliday(today)) return;

  // 08:30 마감 이후 실행이면 독촉 무의미 → 스킵
  const nowHM = Utilities.formatDate(new Date(), 'Asia/Seoul', 'HHmm');
  if (nowHM >= '0830') { Logger.log('[등교인증 독촉] 마감 후 → 스킵'); return; }

  // 당일 1회만 (트리거가 여러 번 돌아도 이중 발송 방지)
  if (_jobAlreadyRan('checkinReminder', today)) return;

  // 오늘 이미 출결 기록 있는 학생 (출석/지각/결석/공결)
  const checkedSet = new Set();
  getValidData(SS.getSheetByName('Attendance')).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === today && String(r[1]).trim())
      checkedSet.add(String(r[1]).trim());
  });
  // 결석사유 등록된 학생 (출석 불필요)
  const excusedSet = new Set();
  const absSheet = SS.getSheetByName('AbsenceReason');
  if (absSheet) getValidData(absSheet).slice(1).forEach(function(r) {
    if (normalizeDateCell(r[0]) === today && String(r[1]).trim())
      excusedSet.add(String(r[1]).trim());
  });

  const mem = getValidData(SS.getSheetByName('Member'));
  let n = 0;
  for (let i = 1; i < mem.length; i++) {
    const name = String(mem[i][1] || '').trim();
    const role = String(mem[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (checkedSet.has(name) || excusedSet.has(name)) continue;
    try { _sendPushToStudent(name, '🌅 등교 인증하세요!', '아직 등교 인증 전이에요. 08:30까지 인증하면 +1냥! (지각·결석 주의)'); n++; } catch(e) {}
  }
  _markJobRan('checkinReminder', today);
  Logger.log('[등교인증 독촉] ' + n + '명');
}

/* [트리거-F] 10분마다 — 방과후 수업 시작 ~10분 전 안내 (시간 입력된 수업만)
   같은 수업 중복발송 방지: 스크립트 속성에 당일 발송 기록 */
function pushBeforeAfterSchool() {
  const now   = new Date();
  const today = toDateStr(now);
  const dow = new Date(today + 'T00:00:00+09:00').getDay();
  if (dow === 0 || dow === 6) return;
  const classes = _getTodayAfterSchool(today);
  if (!classes.length) return;

  const hh = parseInt(Utilities.formatDate(now, 'Asia/Seoul', 'HH'), 10);
  const mm = parseInt(Utilities.formatDate(now, 'Asia/Seoul', 'mm'), 10);
  const nowMin = hh * 60 + mm;

  const props = PropertiesService.getScriptProperties();
  const key = 'asnotified_' + today;
  let sentMap = {};
  try { sentMap = JSON.parse(props.getProperty(key) || '{}'); } catch(e) {}
  let changed = false;

  classes.forEach(function(c) {
    const t = _hm(c.start);
    const m = t.match(/^(\d{1,2}):(\d{2})/);
    if (!m) return;   // 시작시간 없으면 건너뜀
    const startMin = parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
    const diff = startMin - nowMin;
    const k = (c.id || c.subject) + '_' + t;
    if (diff > 0 && diff <= 10 && !sentMap[k]) {
      try {
        const title = '⏰ 방과후 곧 시작', body = c.subject + ' 수업이 ' + t + '에 시작해요!';
        const tr = _classTrack(c);
        if (tr === '공통') _sendPushToAll(title, body);
        else _sendPushToTrack(tr, title, body);
      } catch(e) {}
      sentMap[k] = true; changed = true;
    }
  });
  if (changed) props.setProperty(key, JSON.stringify(sentMap));
}


/* =========================================================
   교사일지 (교감 말씀 등 기록)
   TeacherLog 시트: A=날짜 B=내용 C=작성시각
   ========================================================= */
function _getTeacherLogSheet() {
  let s = SS.getSheetByName('TeacherLog');
  if (!s) { s = SS.insertSheet('TeacherLog'); s.appendRow(['날짜', '내용', '작성시각']); }
  return s;
}
function getTeacherLog() {
  const sheet = _getTeacherLogSheet();
  const data  = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!String(data[i][1] || '').trim()) continue;
    result.push({ row: i + 1, date: normalizeDateCell(data[i][0]),
                  content: String(data[i][1]).trim(),
                  at: data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(data[i][2] || '') });
  }
  result.sort(function(a, b) { return (b.date + (b.at||'')).localeCompare(a.date + (a.at||'')); });
  return result;
}
function saveTeacherLog(p) {
  if (!p.content || !String(p.content).trim()) throw new Error('내용을 입력하시오.');
  const sheet = _getTeacherLogSheet();
  const date  = String(p.date || '').trim() || toDateStr();
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  if (p.row) {
    sheet.getRange(Number(p.row), 1, 1, 2).setValues([[date, String(p.content).trim()]]);
    return { success: true, updated: true };
  }
  sheet.appendRow([date, String(p.content).trim(), now]);
  return { success: true };
}
function deleteTeacherLog(p) {
  if (!p.row) throw new Error('행 번호가 필요하오.');
  _getTeacherLogSheet().deleteRow(Number(p.row));
  return { success: true };
}


/* =========================================================
   졸업생 성적 검색 (ExamHistory_공무원/공기업에서 이름으로)
   ExamHistory 컬럼: A기수 B시험명 C이름 D국어 E영어 F한국사 G합계 H합격
   ========================================================= */
function getGraduateScores(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('이름을 입력하시오.');
  const result = [];
  ['공무원', '공기업'].forEach(function(track) {
    const sheet = SS.getSheetByName('ExamHistory_' + track);
    if (!sheet) return;
    getValidData(sheet).slice(1).forEach(function(r) {
      if (String(r[2] || '').trim() !== name) return;
      const ko = Number(r[3]) || 0, en = Number(r[4]) || 0, hi = Number(r[5]) || 0;
      result.push({
        track: track, cohort: String(r[0] || '').trim(), examName: String(r[1] || '').trim(),
        korean: ko, english: en, history: hi,
        total: Number(r[6]) || (ko + en + hi),
        passed: String(r[7] || '').trim().toUpperCase() === 'Y'
      });
    });
  });
  return result;
}


/* =========================================================
   ★ 학년도 전환 (졸업/신입 교체) — 되돌리기 어려움, 신중히
   p.graduates   = [졸업할 학생 이름, ...]
   p.newStudents = [{ id, name, track('공무원'/'공기업'), dorm(bool) }, ...]
   p.cohort      = 기수 라벨 (예: '2026 졸업')
   동작: 졸업생 성적 → ExamHistory 이동 후 Member에서 제거 /
        잔류 학생 엽전·사면장 0 리셋 / 신입 추가 / 상담은 그대로 누적
   ========================================================= */
function _fsDeleteDoc(collectionId, docId) {
  try {
    const token = _getFsAccessToken();
    const safeId = String(docId).replace(/[\/\s#\[\]\*\?]/g, '_');
    const url = 'https://firestore.googleapis.com/v1/projects/' + FS_PROJECT_ID
              + '/databases/(default)/documents/' + collectionId + '/' + encodeURIComponent(safeId);
    UrlFetchApp.fetch(url, { method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  } catch(e) { Logger.log('[FS] _fsDeleteDoc 오류: ' + e.message); }
}

function promoteSchoolYear(p) {
  const cohort    = String(p.cohort || '').trim() || (new Date().getFullYear() + ' 졸업');
  const graduates = (p.graduates || []).map(function(s) { return String(s).trim(); }).filter(function(s) { return s; });
  const gradSet   = {};
  graduates.forEach(function(n) { gradSet[n] = true; });

  const memSheet = SS.getSheetByName('Member');
  const memData  = getValidData(memSheet);
  const today    = toDateStr();
  const plog     = SS.getSheetByName('PointLog');

  // 졸업생 이름→반(track), id 매핑 + 잔류 학생 엽전 리셋
  const gradInfo = {};   // name -> { track, id }
  const gradRows = [];   // Member 삭제할 행
  const gradIds  = [];
  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (!name || name === '교사' || role === 'teacher') continue;
    if (gradSet[name]) {
      gradInfo[name] = { track: String(memData[i][6] || '').trim() || '공무원', id: String(memData[i][0] || '').trim() };
      gradRows.push(i + 1);
      if (gradInfo[name].id) gradIds.push(gradInfo[name].id);
    } else {
      // 잔류 학생: 엽전·사면장 0 리셋 + 로그
      const prev = Number(memData[i][3]) || 0;
      if (prev !== 0) plog.appendRow([today, name, -prev, '학년도전환(엽전초기화)', '교사', '완료']);
      memSheet.getRange(i + 1, 4).setValue(0);
      memSheet.getRange(i + 1, 5).setValue(0);
    }
  }

  // 졸업생 성적 ExamScores → ExamHistory 이동
  const examSheet = SS.getSheetByName('ExamScores');
  let movedScores = 0;
  if (examSheet) {
    const ex = getValidData(examSheet);
    const exDeleteRows = [];
    for (let i = 1; i < ex.length; i++) {
      const who = String(ex[i][2] || '').trim();
      if (!gradSet[who]) continue;   // 졸업생 본인 점수만 (__avg__ 제외됨)
      const track = (gradInfo[who] && gradInfo[who].track) || '공무원';
      const hSheet = SS.getSheetByName('ExamHistory_' + track);
      if (hSheet) {
        const ko = Number(ex[i][3]) || 0, en = Number(ex[i][4]) || 0, hi = Number(ex[i][5]) || 0;
        hSheet.appendRow([cohort, String(ex[i][0] || '').trim(), who, ko, en, hi,
                          Number(ex[i][6]) || (ko + en + hi), '']);
        movedScores++;
      }
      exDeleteRows.push(i + 1);
    }
    exDeleteRows.sort(function(a, b) { return b - a; }).forEach(function(r) { examSheet.deleteRow(r); });
  }

  // 졸업생 Member 행 삭제 (큰 행부터)
  gradRows.sort(function(a, b) { return b - a; }).forEach(function(r) { memSheet.deleteRow(r); });

  // 신입 추가
  // ★ 학번/이름 중복 시 건너뜀 — 시간초과 후 교사가 재실행해도 신입생이 이중 등록되지 않게
  const curData = getValidData(memSheet);
  const existIds = {}, existNames = {};
  for (let i = 1; i < curData.length; i++) {
    existIds[String(curData[i][0] || '').trim()] = true;
    existNames[String(curData[i][1] || '').trim()] = true;
  }
  let added = 0;
  (p.newStudents || []).forEach(function(s) {
    const id = String(s.id || '').trim();
    const name = String(s.name || '').trim();
    if (!id || !name) return;
    if (existIds[id] || existNames[name]) return;   // 이미 등록됨 (재실행 멱등성)
    const track = (s.track === '공기업') ? '공기업' : '공무원';
    memSheet.appendRow([id, name, '관원', 0, 0, s.dorm ? 'Y' : '', track, '']);
    existIds[id] = true; existNames[name] = true;
    added++;
  });

  // 은행 저축 초기화 (지갑 0 리셋과 일관 — 새 학년도는 저축도 0에서 시작)
  const bankSheet = SS.getSheetByName('Bank');
  if (bankSheet && bankSheet.getLastRow() > 1) {
    bankSheet.deleteRows(2, bankSheet.getLastRow() - 1);
  }

  // Firestore: 졸업생 문서 삭제 + 나머지 재동기화
  gradIds.forEach(function(id) { _fsDeleteDoc('members', id); });
  try { syncMembersToFirestore(); } catch(e) { Logger.log('[FS] promote sync 오류: ' + e.message); }
  try { ['공무원', '공기업'].forEach(function(t) { CacheService.getScriptCache().remove('legacy_stats_' + t); }); } catch(e) {}
  cacheRemove(['member_all', 'admin_data', 'role_포도대장_v2', 'exam_list']);

  return { success: true, graduated: graduates.length, movedScores: movedScores, added: added, cohort: cohort };
}


/* =========================================================
   ★ 엽전 부여 사유 프리셋 (교사)
   Script Properties 'coin_presets' = [{label, amount}, ...]
   ========================================================= */
const DEFAULT_COIN_PRESETS = [
  { label: '수업 태도 우수', amount: 2 },
  { label: '자습 집중 우수', amount: 1 },
  { label: '청소/봉사 활동', amount: 2 },
  { label: '과제 미제출',   amount: -1 },
  { label: '수업 방해',     amount: -2 },
  { label: '지시 불이행',   amount: -1 },
];

function getCoinPresets() {
  const raw = PropertiesService.getScriptProperties().getProperty('coin_presets');
  if (!raw) return DEFAULT_COIN_PRESETS;
  try {
    const v = JSON.parse(raw);
    return Array.isArray(v) ? v : DEFAULT_COIN_PRESETS;
  } catch(e) { return DEFAULT_COIN_PRESETS; }
}

function saveCoinPresets(p) {
  if (!Array.isArray(p.presets)) throw new Error('프리셋 목록이 없소.');
  const clean = p.presets
    .map(function(x) {
      return { label: String(x.label || '').trim().slice(0, 30),
               amount: Math.max(-50, Math.min(50, Math.round(Number(x.amount) || 0))) };
    })
    .filter(function(x) { return x.label && x.amount !== 0; })
    .slice(0, 20);
  PropertiesService.getScriptProperties().setProperty('coin_presets', JSON.stringify(clean));
  return { success: true, count: clean.length };
}


/* =========================================================
   ★ 시험 D-day (교사 등록 → 홈 배너 표시)
   Script Properties 'exam_ddays' = [{id, name, date, track}, ...]
   track: '공통' | '공무원' | '공기업'
   ========================================================= */
function getExamDdays() {
  const raw = PropertiesService.getScriptProperties().getProperty('exam_ddays');
  if (!raw) return [];
  try {
    const v = JSON.parse(raw);
    return Array.isArray(v) ? v : [];
  } catch(e) { return []; }
}

function saveExamDday(p) {
  const name = String(p.name || '').trim();
  const date = String(p.date || '').trim();
  if (!name) throw new Error('시험 이름이 필요하오.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜 형식이 올바르지 않소. (yyyy-mm-dd)');
  const track = (p.track === '공무원' || p.track === '공기업') ? p.track : '공통';
  const list = getExamDdays();
  list.push({ id: 'dd' + Date.now(), name: name.slice(0, 40), date: date, track: track });
  list.sort(function(a, b) { return a.date.localeCompare(b.date); });
  if (list.length > 30) list.length = 30;
  PropertiesService.getScriptProperties().setProperty('exam_ddays', JSON.stringify(list));
  return { success: true, count: list.length };
}

function deleteExamDday(p) {
  if (!p.id) throw new Error('id가 필요하오.');
  const list = getExamDdays().filter(function(d) { return d.id !== p.id; });
  PropertiesService.getScriptProperties().setProperty('exam_ddays', JSON.stringify(list));
  return { success: true };
}


/* =========================================================
   ★ 엽전 은행 (저축 + 주간 이자)
   Bank 시트: A=이름 B=저축잔액 C=최근변동
   - 입금: 지갑(Member D열)에서 차감(PointLog '은행 저축') → 은행 잔액 증가
   - 출금: 은행 잔액 차감 → 지갑에 지급(PointLog '은행 출금')
   - 이자: weeklyBankInterest 트리거(매주 월요일)가 저축액의 BANK_RATE_PERCENT% 지급(내림)
   ========================================================= */
const BANK_RATE_PERCENT = 5;   // 주간 이자율(%) — 바꾸려면 이 숫자만 수정

function _getBankSheet() {
  let s = SS.getSheetByName('Bank');
  if (!s) { s = SS.insertSheet('Bank'); s.appendRow(['이름', '저축잔액', '최근변동']); }
  return s;
}

/** 은행 잔액 증감 (음수 잔액 방지). 반환: 변경 후 잔액 */
function _bankAdjust(name, delta) {
  const sheet = _getBankSheet();
  const data  = getValidData(sheet);
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name) {
      const cur  = Number(data[i][1]) || 0;
      const next = cur + delta;
      if (next < 0) throw new Error('저축 잔액이 부족하오. (저축 ' + cur + '냥)');
      sheet.getRange(i + 1, 2).setValue(next);
      sheet.getRange(i + 1, 3).setValue(now);
      return next;
    }
  }
  if (delta < 0) throw new Error('저축 잔액이 없소.');
  sheet.appendRow([name, delta, now]);
  return delta;
}

function getBank(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('이름이 없소.');
  const data = getValidData(_getBankSheet());
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name)
      return { balance: Number(data[i][1]) || 0, rate: BANK_RATE_PERCENT };
  }
  return { balance: 0, rate: BANK_RATE_PERCENT };
}

function getBankAll() {
  const data = getValidData(_getBankSheet());
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) result.push({ name: name, balance: Number(data[i][1]) || 0 });
  }
  return result;
}

function bankDeposit(p) {
  const name   = String(p.name || '').trim();
  const amount = Math.round(Number(p.amount));
  if (!name) throw new Error('이름이 없소.');
  if (!(amount > 0)) throw new Error('금액이 올바르지 않소.');
  // 지갑 잔액 확인
  const memData = getValidData(SS.getSheetByName('Member'));
  let coins = null;
  for (let i = 1; i < memData.length; i++) {
    if (String(memData[i][1]).trim() === name) { coins = Number(memData[i][3]) || 0; break; }
  }
  if (coins === null) throw new Error('관원을 찾지 못했소: ' + name);
  if (coins < amount) throw new Error('지갑 엽전이 부족하오. (보유 ' + coins + '냥)');
  // 지갑 차감(PointLog·캐시·Firestore 일괄) → 은행 증가
  const walletCoins = applyCoins(name, -amount, '은행 저축', '은행');
  const bankBalance = _bankAdjust(name, amount);
  return { walletCoins: walletCoins, bankBalance: bankBalance };
}

function bankWithdraw(p) {
  const name   = String(p.name || '').trim();
  const amount = Math.round(Number(p.amount));
  if (!name) throw new Error('이름이 없소.');
  if (!(amount > 0)) throw new Error('금액이 올바르지 않소.');
  // 은행 차감(잔액 부족 시 여기서 중단 → 지갑은 건드리지 않음) → 지갑 지급
  const bankBalance = _bankAdjust(name, -amount);
  const walletCoins = applyCoins(name, amount, '은행 출금', '은행');
  return { walletCoins: walletCoins, bankBalance: bankBalance };
}

/** [트리거] 매주 월요일 새벽 — 저축액의 BANK_RATE_PERCENT% 이자 지급 (내림) */
function weeklyBankInterest() { _withScriptLock(_weeklyBankInterestCore); }
function _weeklyBankInterestCore() {
  // ★ 같은 주 중복 실행 방지 (트리거 중복 등록 시 이자 이중 지급)
  //   키 = 이번 주 월요일 날짜: 실행 시각(새벽 5시/6시/다른 요일 재실행)과 무관하게 안정
  const _d = new Date();
  _d.setDate(_d.getDate() - ((_d.getDay() + 6) % 7));
  const weekKey = Utilities.formatDate(_d, 'Asia/Seoul', 'yyyy-MM-dd');
  if (_jobAlreadyRan('bankInterest', weekKey)) {
    Logger.log('[weeklyBankInterest] 이미 실행됨 → 건너뜀');
    return;
  }
  const sheet = _getBankSheet();
  const data  = getValidData(sheet);
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    const bal  = Number(data[i][1]) || 0;
    if (!name || bal <= 0) continue;
    const interest = Math.floor(bal * BANK_RATE_PERCENT / 100);
    if (interest <= 0) continue;
    sheet.getRange(i + 1, 2).setValue(bal + interest);
    sheet.getRange(i + 1, 3).setValue(now + ' (이자 +' + interest + ')');
    try { _sendPushToStudent(name, '🏦 저축 이자 도착', '+' + interest + '냥 (저축 ' + (bal + interest) + '냥)'); } catch(e) {}
    count++;
  }
  _markJobRan('bankInterest', weekKey);
  Logger.log('[weeklyBankInterest] ' + count + '명 이자 지급');
}


/* =========================================================
   ★ 주간 리포트 (지난주 월~금 요약)
   WeeklyReport 시트: A=주차시작 B=생성시각 C=내용
   ========================================================= */
function _getWeeklyReportSheet() {
  let s = SS.getSheetByName('WeeklyReport');
  if (!s) { s = SS.insertSheet('WeeklyReport'); s.appendRow(['주차시작', '생성시각', '내용']); }
  return s;
}

/** 지난주 월~금 데이터를 집계해 리포트 텍스트 생성
 *  '지난주' = 가장 최근에 끝난 월~금 (월~금 실행 시 지난주 / 토·일 실행 시 방금 끝난 이번 주) */
function buildWeeklyReport() {
  const now = new Date();
  const dow = now.getDay();   // 0=일 ~ 6=토
  const monday = new Date(now);
  monday.setDate(now.getDate() - ((dow + 6) % 7));          // 이번 주(월 시작) 월요일
  if (dow >= 1 && dow <= 5) monday.setDate(monday.getDate() - 7);   // 주중이면 지난주 월요일
  const weekDates = [];
  for (let d = 0; d < 5; d++) {
    const target = new Date(monday);
    target.setDate(monday.getDate() + d);                   // 월(d=0)~금(d=4)
    weekDates.push(Utilities.formatDate(target, 'Asia/Seoul', 'yyyy-MM-dd'));
  }
  const dateSet   = new Set(weekDates);
  const weekStart = weekDates[0], weekEnd = weekDates[4];

  // 학생 명단
  const memData  = getValidData(SS.getSheetByName('Member'));
  const students = [];
  for (let i = 1; i < memData.length; i++) {
    const name = String(memData[i][1] || '').trim();
    const role = String(memData[i][2] || '').trim();
    if (name && name !== '교사' && role !== 'teacher') students.push(name);
  }
  const stuSet = new Set(students);

  // 출결 집계
  // ★ 수업일 판별: weeklyAttendanceBonus와 동일하게 "출결 기록이 한 건이라도 있는 날"만 수업일로 삼음
  //   (공휴일·트리거 일시정지일은 기록 자체가 없으므로 개근 판정에서 자동 제외됨)
  const weekAttRows  = getValidData(SS.getSheetByName('Attendance')).slice(1)
    .filter(function(r) { return dateSet.has(normalizeDateCell(r[0])); });
  const schoolDayCount = new Set(weekAttRows.map(function(r) { return normalizeDateCell(r[0]); })).size;

  const att = {};
  students.forEach(function(n) { att[n] = { attend: 0, late: 0, absent: 0 }; });
  weekAttRows.forEach(function(r) {
    const name = String(r[1] || '').trim();
    if (!att[name]) return;
    const st = String(r[3] || '').trim();
    if (st === '출석') att[name].attend++;
    else if (st === '지각') att[name].late++;
    else if (st === '결석' || st === '미출석') att[name].absent++;
  });
  const perfect   = schoolDayCount > 0
    ? students.filter(function(n) { return att[n].attend === schoolDayCount; })
    : [];
  const lateTop   = students.filter(function(n) { return att[n].late   > 0; }).sort(function(a, b) { return att[b].late   - att[a].late;   }).slice(0, 3);
  const absentTop = students.filter(function(n) { return att[n].absent > 0; }).sort(function(a, b) { return att[b].absent - att[a].absent; }).slice(0, 3);

  // 플래너 집계
  const plan = {};
  students.forEach(function(n) { plan[n] = { count: 0, sum: 0 }; });
  getValidData(SS.getSheetByName('Planner')).slice(1).forEach(function(r) {
    const d = normalizeDateCell(r[0]);
    if (!dateSet.has(d)) return;
    const name = String(r[1] || '').trim();
    if (!plan[name]) return;
    plan[name].count++;
    plan[name].sum += Number(r[3]) || 0;
  });
  let planSum = 0, planCnt = 0;
  students.forEach(function(n) { planSum += plan[n].sum; planCnt += plan[n].count; });
  const classAvg  = planCnt ? Math.round(planSum / planCnt) : 0;
  const noPlanTop = students.filter(function(n) { return plan[n].count < 5; })
    .sort(function(a, b) { return plan[a].count - plan[b].count; }).slice(0, 3);

  // 엽전 주간 변동 (완료 건만)
  const coin = {};
  students.forEach(function(n) { coin[n] = 0; });
  getValidData(SS.getSheetByName('PointLog')).slice(1).forEach(function(r) {
    const d = normalizeDateCell(r[0]);
    if (!dateSet.has(d)) return;
    if (String(r[5] || '').trim() !== '완료') return;
    const name = String(r[1] || '').trim();
    if (coin[name] === undefined) return;
    coin[name] += Number(r[2]) || 0;
  });
  const gainTop = students.filter(function(n) { return coin[n] > 0; }).sort(function(a, b) { return coin[b] - coin[a]; }).slice(0, 3);
  const lossTop = students.filter(function(n) { return coin[n] < 0; }).sort(function(a, b) { return coin[a] - coin[b]; }).slice(0, 3);

  // 상담 완료 건수
  let counselDone = 0;
  const cSheet = SS.getSheetByName('Counseling');
  if (cSheet) {
    getValidData(cSheet).slice(1).forEach(function(r) {
      const d = normalizeDateCell(r[0]);
      if (!dateSet.has(d)) return;
      if (stuSet.has(String(r[1] || '').trim()) && String(r[2] || '').trim().toUpperCase() === 'Y') counselDone++;
    });
  }

  const fmtList = function(names, valueOf, unit) {
    return names.length
      ? names.map(function(n) { return n + '(' + valueOf(n) + unit + ')'; }).join(', ')
      : '없음';
  };
  const lines = [];
  lines.push('📊 주간 리포트 (' + weekStart + ' ~ ' + weekEnd + ') · 학생 ' + students.length + '명');
  lines.push('');
  lines.push('📅 출결');
  lines.push('· 개근(' + schoolDayCount + '일 출석): ' + (perfect.length ? perfect.join(', ') : '없음') + ' — ' + perfect.length + '명');
  lines.push('· 지각: ' + fmtList(lateTop, function(n) { return att[n].late; }, '회'));
  lines.push('· 결석: ' + fmtList(absentTop, function(n) { return att[n].absent; }, '회'));
  lines.push('');
  lines.push('📝 플래너');
  lines.push('· 학급 평균 달성률: ' + classAvg + '%');
  lines.push('· 작성 저조: ' + fmtList(noPlanTop, function(n) { return plan[n].count; }, '일 작성'));
  lines.push('');
  lines.push('🪙 엽전 (주간 변동)');
  lines.push('· 획득 상위: ' + fmtList(gainTop, function(n) { return '+' + coin[n]; }, '냥'));
  lines.push('· 차감 상위: ' + fmtList(lossTop, function(n) { return coin[n]; }, '냥'));
  lines.push('');
  lines.push('💬 상담 완료: ' + counselDone + '건');

  return { weekStart: weekStart, weekEnd: weekEnd, text: lines.join('\n') };
}

/** 리포트 생성·저장 (같은 주차가 있으면 덮어씀) */
function runWeeklyReport() {
  const r     = buildWeeklyReport();
  const sheet = _getWeeklyReportSheet();
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const data  = getValidData(sheet);
  for (let i = 1; i < data.length; i++) {
    if (normalizeDateCell(data[i][0]) === r.weekStart) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[now, r.text]]);
      return { weekStart: r.weekStart, text: r.text, updated: true };
    }
  }
  sheet.appendRow([r.weekStart, now, r.text]);
  return { weekStart: r.weekStart, text: r.text };
}

function getWeeklyReports() {
  const sheet = SS.getSheetByName('WeeklyReport');
  if (!sheet) return [];
  const data   = getValidData(sheet);
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const ws = normalizeDateCell(data[i][0]);
    if (!ws) continue;
    const at = data[i][1] instanceof Date
      ? Utilities.formatDate(data[i][1], 'Asia/Seoul', 'yyyy-MM-dd HH:mm')
      : String(data[i][1] || '');
    result.push({ weekStart: ws, at: at, text: String(data[i][2] || '') });
  }
  result.sort(function(a, b) { return b.weekStart.localeCompare(a.weekStart); });
  return result.slice(0, 12);
}

/* =========================================================
   ★ 연속 등교 스트릭 (홈 화면 🔥 배지)
   규칙: 평일 기준 출석/지각 = 유지+1, 공결 = 유지(가산 없음),
        결석/기록없음 = 중단. 오늘 아직 미인증이면 오늘만 관용.
   ========================================================= */
function getMyStreak(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('이름이 없소.');
  const data = getValidData(SS.getSheetByName('Attendance'));
  const map = {};
  const anyRecord = new Set();   // 전교생 기준 기록이 있는 날 = 수업일
  for (let i = 1; i < data.length; i++) {
    const ds = normalizeDateCell(data[i][0]);
    anyRecord.add(ds);
    if (String(data[i][1] || '').trim() === name) {
      map[ds] = String(data[i][3] || '').trim();
    }
  }
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  let streak = 0;
  let cur = new Date();
  for (let i = 0; i < 370; i++) {
    const ds  = Utilities.formatDate(cur, 'Asia/Seoul', 'yyyy-MM-dd');
    const dow = new Date(ds + 'T00:00:00+09:00').getDay();
    if (dow !== 0 && dow !== 6) {   // 주말은 건너뜀
      const st = map[ds] || '';
      if (st === '출석' || st === '지각') streak++;
      else if (st === '공결') { /* 스트릭 유지, 가산 없음 */ }
      else if (ds === todayStr) { /* 오늘 미인증은 아직 판정 보류 */ }
      else if (!anyRecord.has(ds)) { /* 아무도 기록 없는 날 = 공휴일·일시정지 → 유지 */ }
      else break;
    }
    cur = new Date(cur.getTime() - 86400000);
  }
  return { streak: streak };
}

/** [트리거] 매주 월요일 아침 — 주간 리포트 생성 + 교사 푸시 */
function weeklyTeacherReport() {
  const r = runWeeklyReport();
  try { _sendPushToStudent('교사', '📊 주간 리포트 도착', r.weekStart + ' 주 요약 완성 — 집무실 > 주간 리포트에서 확인하세요.'); } catch(e) {}
  Logger.log('[weeklyTeacherReport] 생성: ' + r.weekStart);
}
