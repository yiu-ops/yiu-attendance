// ==============================================
// 용인대학교 교직원 회의 출석 시스템 - Code.gs
// 버전: 2.0.0 (QR + GPS 이중 인증, 레이트 리미팅 추가)
// ==============================================

// 관리자 설정 (배포 전 수정 필요)
const SPREADSHEET_ID = '1J88woOKMJQNds9GH2fXMyAOkfZmdmeyIFzMHprsbRVw';

// 기본 설정값
const DEFAULT_SETTINGS = {
  meeting_title: '용인대학교 교직원 회의',
  meeting_description: '출석 정보를 정확히 입력해주세요',
  target_latitude: '37.2420',
  target_longitude: '127.1775',
  location_radius: '1200',
  admin_password: 'yongin2024',
  notification_email: '',
  current_event_id: '',
  max_attendance_per_event: '1000',
  session_timeout_hours: '24',
  qr_token_validity_minutes: '15',
  rate_limit_per_minute: '5',
  location_mode: 'outdoor',
  prevent_proxy_attendance: 'false'
};

// 시트 이름
const SHEET_NAMES = {
  SETTINGS: 'settings',
  ATTENDANCE: 'attendance',
  DEVICE_LOG: 'device_log',
  EMAIL_LOG: 'email_log',
  EVENT_CONTROL: 'event_control',
  ERROR_LOG: 'error_log',
  QR_TOKENS: 'qr_tokens',
  RATE_LIMIT: 'rate_limit'
};

// 캐시 제거됨 - 인스턴스별 캐시는 이벤트 ID 불일치 유발

/**
 * 웹 앱 진입점 - GET 요청
 */
function doGet(e) {
  try {
    var page = 'user';
    var method = null;
    var callback = null;
    var params = null;

    if (e && e.parameter) {
      page = e.parameter.page || 'user';
      method = e.parameter.method;
      callback = e.parameter.callback;
      params = e.parameter.params;
    }

    // JSONP 요청 처리
    if (method && callback) {
      var result;
      var parsedParams = null;

      if (params) {
        try {
          parsedParams = JSON.parse(decodeURIComponent(params));
        } catch (error) {
          console.error('Params 파싱 오류:', error);
          parsedParams = null;
        }
      }

      try {
        switch (method) {
          case 'testConnection':
            result = testConnection();
            break;
          case 'getSettings':
            result = getSettings();
            break;
          case 'saveAttendance':
            if (parsedParams) {
              result = saveAttendanceImproved(parsedParams);
            } else {
              result = { success: false, error: '출석 데이터가 필요합니다.' };
            }
            break;
          case 'startNewMeeting':
            var meetingTitle = parsedParams && parsedParams.meetingTitle ? parsedParams.meetingTitle : '';
            result = startNewMeetingFixed(meetingTitle);
            break;
          case 'getAttendanceStats':
            result = getAttendanceStats();
            break;
          case 'getSystemStatus':
            result = getSystemStatus();
            break;
          case 'authenticateAdmin':
            var password = parsedParams && parsedParams.password ? parsedParams.password : '';
            var isAuthenticated = authenticateAdmin(password);
            result = {
              success: isAuthenticated,
              authenticated: isAuthenticated,
              message: isAuthenticated ? '인증 성공' : '인증 실패'
            };
            break;
          case 'saveSettings':
            if (parsedParams && parsedParams.settings) {
              result = saveSettings(parsedParams.settings);
            } else {
              result = { success: false, error: '설정 데이터가 필요합니다.' };
            }
            break;
          case 'testEmailSystem':
            result = testEmailSystem();
            break;
          // QR 토큰 관련 API
          case 'generateQRToken':
            result = generateQRToken();
            break;
          case 'validateQRToken':
            var token = parsedParams && parsedParams.token ? parsedParams.token : '';
            result = validateQRToken(token);
            break;
          // 출석 조회 API
          case 'checkAttendance':
            if (parsedParams && parsedParams.name && parsedParams.department) {
              result = checkAttendance(parsedParams.name, parsedParams.department);
            } else {
              result = { success: false, error: '이름과 소속을 모두 입력해주세요.' };
            }
            break;
          // 다운로드 관련 API
          case 'getAttendanceForDownload':
            var startDate = parsedParams && parsedParams.startDate ? parsedParams.startDate : null;
            var endDate = parsedParams && parsedParams.endDate ? parsedParams.endDate : null;
            var eventOnly = parsedParams && parsedParams.eventOnly ? parsedParams.eventOnly : false;
            result = getAttendanceForDownload(startDate, endDate, eventOnly);
            break;
          default:
            result = { success: false, error: '지원하지 않는 메서드입니다: ' + method };
        }
      } catch (methodError) {
        console.error('메서드 실행 오류:', methodError);
        result = { success: false, error: '서버 처리 중 오류가 발생했습니다.' };
      }

      console.log('JSONP 응답:', method, result);

      var jsonpResponse = callback + '(' + JSON.stringify(result) + ');';
      return ContentService
        .createTextOutput(jsonpResponse)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    // 관리자 페이지 요청 처리
    if (page === 'admin') {
      return createAdminPage();
    } else {
      return HtmlService.createHtmlOutput(
        '<h1>용인대학교 교직원 회의 출석</h1>' +
        '<p>사용자 페이지는 별도 링크에서 접속하세요.</p>' +
        '<p><a href="?page=admin">관리자 페이지로 이동</a></p>'
      ).setTitle('용인대학교 교직원 회의 출석');
    }

  } catch (error) {
    console.error('doGet 오류:', error);

    if (e && e.parameter && e.parameter.callback) {
      var errorResponse = { success: false, error: '시스템 오류가 발생했습니다.' };
      var jsonpResponse = e.parameter.callback + '(' + JSON.stringify(errorResponse) + ');';
      return ContentService
        .createTextOutput(jsonpResponse)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return HtmlService.createHtmlOutput('<h1>시스템 오류</h1><p>페이지를 불러올 수 없습니다.</p>');
  }
}

/**
 * 관리자 페이지 생성
 */
function createAdminPage() {
  var adminHtml = HtmlService.createHtmlOutputFromFile('admin');
  return adminHtml
    .setTitle('용인대학교 교직원 회의 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==============================================
// 레이트 리미팅
// ==============================================

/**
 * 레이트 리미팅 확인
 * @param {string} deviceId - 기기 ID
 * @returns {object} - { allowed: boolean, remaining: number }
 */
function checkRateLimit(deviceId) {
  try {
    var settings = getSettings();
    var limitPerMinute = parseInt(settings.rate_limit_per_minute) || 5;

    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.RATE_LIMIT, [
      '기기ID', '시간', '요청횟수'
    ]);

    var now = Date.now();
    var oneMinuteAgo = now - 60000;

    var data = sheet.getDataRange().getValues();
    var requestCount = 0;

    // 최근 1분간의 요청 수 계산
    for (var i = 1; i < data.length; i++) {
      var rowDeviceId = data[i][0];
      var rowTime = new Date(data[i][1]).getTime();

      if (rowDeviceId === deviceId && rowTime > oneMinuteAgo) {
        requestCount++;
      }
    }

    // 새 요청 기록
    sheet.appendRow([deviceId, new Date(), 1]);

    // 오래된 레코드 정리 (1시간 이상 된 것)
    cleanupOldRateLimitRecords(sheet);

    if (requestCount >= limitPerMinute) {
      console.log('레이트 리밋 초과:', deviceId, requestCount);
      return {
        allowed: false,
        remaining: 0,
        message: '너무 많은 요청이 감지되었습니다. 1분 후 다시 시도해주세요.'
      };
    }

    return {
      allowed: true,
      remaining: limitPerMinute - requestCount - 1
    };

  } catch (error) {
    console.error('레이트 리밋 확인 오류:', error);
    // 오류 발생 시 허용 (사용자 경험 우선)
    return { allowed: true, remaining: 5 };
  }
}

/**
 * 오래된 레이트 리밋 레코드 정리
 */
function cleanupOldRateLimitRecords(sheet) {
  try {
    var oneHourAgo = Date.now() - 3600000;
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];

    for (var i = data.length - 1; i > 0; i--) {
      var rowTime = new Date(data[i][1]).getTime();
      if (rowTime < oneHourAgo) {
        rowsToDelete.push(i + 1);
      }
    }

    // 역순으로 삭제 (인덱스 문제 방지)
    for (var j = 0; j < rowsToDelete.length; j++) {
      sheet.deleteRow(rowsToDelete[j]);
    }
  } catch (error) {
    console.error('레이트 리밋 정리 오류:', error);
  }
}

// ==============================================
// QR 토큰 관리
// ==============================================

/**
 * QR 토큰 생성
 * @returns {object} - { success: boolean, token: string, expires: number }
 */
function generateQRToken() {
  try {
    var settings = getSettings();
    var validityMinutes = parseInt(settings.qr_token_validity_minutes) || 15;

    var token = 'YU_QR_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    var expires = Date.now() + (validityMinutes * 60 * 1000);
    var currentEventId = getCurrentEventId();

    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.QR_TOKENS, [
      '토큰', '생성시간', '만료시간', '이벤트ID', '상태', '사용횟수'
    ]);

    // 기존 활성 토큰 비활성화
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === 'active') {
        sheet.getRange(i + 1, 5).setValue('inactive');
      }
    }

    // 새 토큰 저장
    sheet.appendRow([
      token,
      new Date(),
      new Date(expires),
      currentEventId,
      'active',
      0
    ]);

    console.log('QR 토큰 생성:', token);

    return {
      success: true,
      token: token,
      expires: expires,
      validityMinutes: validityMinutes,
      eventId: currentEventId
    };

  } catch (error) {
    console.error('QR 토큰 생성 오류:', error);
    return {
      success: false,
      error: 'QR 토큰 생성에 실패했습니다.'
    };
  }
}

/**
 * QR 토큰 검증
 * @param {string} token - QR 토큰
 * @returns {object} - { valid: boolean, eventId: string }
 */
function validateQRToken(token) {
  try {
    if (!token || !token.startsWith('YU_QR_')) {
      return { valid: false, error: '유효하지 않은 토큰 형식입니다.' };
    }

    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAMES.QR_TOKENS);

    if (!sheet) {
      return { valid: false, error: 'QR 토큰 시스템이 초기화되지 않았습니다.' };
    }

    var data = sheet.getDataRange().getValues();
    var now = Date.now();

    for (var i = 1; i < data.length; i++) {
      var rowToken = data[i][0];
      var expiresAt = new Date(data[i][2]).getTime();
      var eventId = data[i][3];
      var status = data[i][4];
      var usageCount = parseInt(data[i][5]) || 0;

      if (rowToken === token) {
        if (status !== 'active') {
          return { valid: false, error: 'QR 코드가 비활성화되었습니다. 새 QR 코드를 요청하세요.' };
        }

        if (now > expiresAt) {
          // 만료된 토큰 비활성화
          sheet.getRange(i + 1, 5).setValue('expired');
          return { valid: false, error: 'QR 코드가 만료되었습니다. 새 QR 코드를 요청하세요.' };
        }

        // 사용 횟수 증가
        sheet.getRange(i + 1, 6).setValue(usageCount + 1);

        return {
          valid: true,
          eventId: eventId,
          remaining: Math.floor((expiresAt - now) / 60000)
        };
      }
    }

    return { valid: false, error: 'QR 코드를 찾을 수 없습니다.' };

  } catch (error) {
    console.error('QR 토큰 검증 오류:', error);
    return { valid: false, error: '토큰 검증 중 오류가 발생했습니다.' };
  }
}

// ==============================================
// 출석 데이터 저장
// ==============================================

/**
 * 출석 데이터 저장 (중복 체크 강화 + 레이트 리미팅)
 */
function saveAttendanceImproved(data) {
  try {
    console.log('출석 체크 시작:', data.name, data.department);

    // 입력값 검증
    var validationResult = validateAttendanceData(data);
    if (!validationResult.valid) {
      return {
        success: false,
        error: validationResult.error
      };
    }

    // 기기 분석
    var deviceAnalysis = analyzeDevice(data);
    var deviceId = generateImprovedDeviceId(data, deviceAnalysis);

    // 레이트 리미팅 확인
    var rateLimitResult = checkRateLimit(deviceId);
    if (!rateLimitResult.allowed) {
      return {
        success: false,
        error: rateLimitResult.message
      };
    }

    // QR 토큰 검증 (QR 출석인 경우)
    if (data.qrToken) {
      var qrResult = validateQRToken(data.qrToken);
      if (!qrResult.valid) {
        return {
          success: false,
          error: qrResult.error
        };
      }
    }

    // 위치 검증 (실내 모드에서는 건너뜀)
    var settings = getSettings();
    var isIndoorMode = settings.location_mode === 'indoor';
    var locationResult;

    if (isIndoorMode && data.locationMode === 'indoor') {
      // 실내 모드: GPS 검증 생략, QR 인증으로 대체
      locationResult = { valid: true, distance: 0 };
    } else {
      locationResult = validateLocationForDevice(
        data.latitude,
        data.longitude,
        data.accuracy,
        deviceAnalysis
      );

      if (!locationResult.valid) {
        return {
          success: false,
          error: locationResult.error || '위치 검증에 실패했습니다.'
        };
      }
    }

    // 스프레드시트 및 시트 준비
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.ATTENDANCE, [
      '성명', '소속', '이메일', '제출시간', '위도', '경도', '거리(m)', '정확도(m)', '기기정보', '세션ID', '이벤트ID', 'QR토큰'
    ]);

    // 현재 이벤트 ID
    var currentEventId = getCurrentEventId();

    console.log('중복 체크 시작:', deviceId, currentEventId);

    // 강화된 중복 체크
    var duplicateCheck = checkDeviceDuplicateImproved(sheet, deviceId, currentEventId, data.email);
    if (duplicateCheck.isDuplicate) {
      // 중복 시도 로그 기록
      logDeviceAttempt(data, deviceId, 'device_duplicate', currentEventId);

      return {
        success: false,
        error: '이 기기에서 현재 회의에 이미 출석 체크가 완료되었습니다.\n\n' +
               '기존 출석자: ' + duplicateCheck.existingName + '\n' +
               '출석 시간: ' + duplicateCheck.existingTime + '\n\n' +
               '한 기기당 한 번만 출석 체크가 가능합니다.'
      };
    }

    // 출석 데이터 저장
    var sessionId = generateSessionId();
    var emailValue = (data.email && data.email.trim()) ? data.email.trim() : '';

    var rowData = [
      data.name.trim(),
      data.department.trim(),
      emailValue,
      new Date(),
      data.latitude,
      data.longitude,
      Math.round(locationResult.distance),
      Math.round(data.accuracy || 0),
      deviceAnalysis.type + '|' + deviceAnalysis.browser + '|' + deviceId,
      sessionId,
      currentEventId,
      data.qrToken || ''
    ];

    sheet.appendRow(rowData);
    console.log('출석 데이터 저장 완료:', sessionId);

    // 이메일 발송 (이메일이 있는 경우에만)
    var emailResult = { success: false };
    if (emailValue) {
      emailResult = sendAttendanceEmail(data, sessionId, Math.round(locationResult.distance), currentEventId);
    }

    return {
      success: true,
      message: emailValue ? '출석이 완료되었습니다!\n확인 이메일이 발송되었습니다.' : '출석이 완료되었습니다!',
      distance: Math.round(locationResult.distance),
      sessionId: sessionId,
      eventId: currentEventId,
      emailSent: emailResult.success
    };

  } catch (error) {
    console.error('출석 저장 오류:', error);
    logError('saveAttendance', error, data);
    return {
      success: false,
      error: '출석 체크 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.'
    };
  }
}

/**
 * 입력값 검증 (이메일은 선택사항)
 */
function validateAttendanceData(data) {
  if (!data || typeof data !== 'object') {
    return { valid: false, error: '데이터 형식이 올바르지 않습니다.' };
  }

  if (!data.name || data.name.trim().length < 2) {
    return { valid: false, error: '이름을 2자 이상 입력해주세요.' };
  }

  if (!data.department || data.department.trim().length < 2) {
    return { valid: false, error: '소속을 2자 이상 입력해주세요.' };
  }

  // 이메일은 선택사항 - 입력된 경우에만 형식 검증
  if (data.email && data.email.trim().length > 0) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email.trim())) {
      return { valid: false, error: '올바른 이메일 형식이 아닙니다.' };
    }
  }

  return { valid: true };
}

/**
 * 기기 ID 기반 중복 체크 (명확한 로직)
 */
function checkDeviceDuplicateImproved(sheet, deviceId, eventId, email) {
  try {
    console.log('중복 체크 시작 - 기기ID:', deviceId, '이벤트ID:', eventId);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('출석 데이터 없음');
      return { isDuplicate: false };
    }

    // 전체 데이터 확인 (성능보다 정확성 우선)
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var existingEventId = (row[10] || '').toString();  // 이벤트ID 컬럼

      // 같은 이벤트인지 먼저 확인
      if (existingEventId === eventId) {
        var existingDeviceInfo = (row[8] || '').toString();  // 기기정보 컬럼

        console.log('기존 기기정보:', existingDeviceInfo);
        console.log('현재 기기ID:', deviceId);

        // 기기 ID 정확히 매칭 (문자열에 포함되어 있는지 확인)
        if (existingDeviceInfo.indexOf(deviceId) > -1) {
          console.log('중복 발견!');
          return {
            isDuplicate: true,
            existingName: (row[0] || '').toString(),
            existingTime: formatDateTime(new Date(row[3]))
          };
        }
      }
    }

    console.log('중복 없음');
    return { isDuplicate: false };

  } catch (error) {
    console.error('기기 중복 체크 오류:', error);
    return { isDuplicate: false };
  }
}

/**
 * 기기 중복 시도 로그 기록
 */
function logDeviceAttempt(data, deviceId, reason, eventId) {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.DEVICE_LOG, [
      '시간', '이름', '소속', '이메일', '기기ID', '사유', '이벤트ID'
    ]);

    var logEntry = [
      new Date(),
      data.name.trim(),
      data.department.trim(),
      (data.email && data.email.trim()) ? data.email.trim() : '',
      deviceId,
      reason === 'device_duplicate' ? '기기 중복' : reason,
      eventId
    ];

    logSheet.appendRow(logEntry);
    console.log('기기 시도 로그 기록:', reason);

  } catch (error) {
    console.error('기기 로그 기록 오류:', error);
  }
}

// ==============================================
// 출석 조회
// ==============================================

/**
 * 출석 조회 (이름 + 소속으로 검색)
 * @param {string} name - 성명
 * @param {string} department - 소속
 * @returns {object} - { success, found, record }
 */
function checkAttendance(name, department) {
  try {
    if (!name || !department) {
      return { success: false, error: '이름과 소속을 입력해주세요.' };
    }

    var trimmedName = name.trim();
    var trimmedDept = department.trim();

    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAMES.ATTENDANCE);

    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, found: false };
    }

    var settings = getSettings();
    var currentEventId = settings.current_event_id;
    var meetingTitle = settings.meeting_title || '교직원 회의';

    var data = sheet.getDataRange().getValues();

    // 현재 이벤트에서 이름+소속이 일치하는 최신 기록 검색
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var rowName = (row[0] || '').toString().trim();
      var rowDept = (row[1] || '').toString().trim();
      var rowEventId = (row[10] || '').toString();

      if (rowEventId === currentEventId && rowName === trimmedName && rowDept === trimmedDept) {
        return {
          success: true,
          found: true,
          record: {
            name: rowName,
            department: rowDept,
            time: formatDateTime(new Date(row[3])),
            distance: row[6] ? Math.round(parseFloat(row[6])) + 'm' : '-',
            meetingTitle: meetingTitle
          }
        };
      }
    }

    return { success: true, found: false };

  } catch (error) {
    console.error('출석 조회 오류:', error);
    return {
      success: false,
      error: '출석 조회 중 오류가 발생했습니다.'
    };
  }
}

// ==============================================
// 출석 데이터 다운로드
// ==============================================

/**
 * 출석 데이터 다운로드용 조회
 * @param {string} startDate - 시작 날짜 (선택)
 * @param {string} endDate - 종료 날짜 (선택)
 * @param {boolean} eventOnly - 현재 이벤트만 조회
 * @returns {object} - { success: boolean, data: array }
 */
function getAttendanceForDownload(startDate, endDate, eventOnly) {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAMES.ATTENDANCE);

    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        data: [],
        count: 0
      };
    }

    var settings = getSettings();
    var currentEventId = settings.current_event_id;

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var result = [];

    var startTimestamp = startDate ? new Date(startDate).getTime() : 0;
    var endTimestamp = endDate ? new Date(endDate).getTime() + 86400000 : Date.now() + 86400000;

    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      var rowTime = new Date(row[3]).getTime();
      var rowEventId = row[10] || '';

      // 이벤트 필터
      if (eventOnly && rowEventId !== currentEventId) {
        continue;
      }

      // 날짜 필터
      if (rowTime < startTimestamp || rowTime > endTimestamp) {
        continue;
      }

      result.push({
        name: row[0] || '',
        department: row[1] || '',
        email: row[2] || '',
        time: formatDateTime(new Date(row[3])),
        latitude: row[4] || '',
        longitude: row[5] || '',
        distance: row[6] || '',
        accuracy: row[7] || '',
        sessionId: row[9] || '',
        eventId: row[10] || ''
      });
    }

    return {
      success: true,
      data: result,
      count: result.length,
      currentEventId: currentEventId
    };

  } catch (error) {
    console.error('출석 데이터 다운로드 조회 오류:', error);
    return {
      success: false,
      error: '데이터 조회 중 오류가 발생했습니다.'
    };
  }
}

// ==============================================
// 통계 조회
// ==============================================

/**
 * 출석 통계 조회
 */
function getAttendanceStats() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var attendanceSheet = spreadsheet.getSheetByName(SHEET_NAMES.ATTENDANCE);
    var deviceSheet = spreadsheet.getSheetByName(SHEET_NAMES.DEVICE_LOG);
    var settings = getSettings();
    var currentEventId = settings.current_event_id;

    var stats = {
      totalAttendees: 0,
      deviceBlocked: 0,
      emailsSent: 0,
      currentEventAttendees: 0,
      recentAttendees: [],
      blockedDevices: [],
      hourlyStats: {}
    };

    // 출석 데이터 분석
    if (attendanceSheet && attendanceSheet.getLastRow() > 1) {
      var attendanceData = attendanceSheet.getDataRange().getValues();
      var currentEventCount = 0;
      var recentAttendees = [];
      var hourlyStats = {};

      for (var i = 1; i < attendanceData.length; i++) {
        var row = attendanceData[i];
        var eventId = row[10];
        var time = new Date(row[3]);

        if (eventId === currentEventId) {
          currentEventCount++;
          recentAttendees.push({
            name: row[0],
            department: row[1],
            email: row[2] || '(미입력)',
            time: formatDateTime(time),
            distance: row[6] ? Math.round(parseFloat(row[6])) : 'N/A'
          });

          // 시간대별 통계
          var hour = time.getHours();
          hourlyStats[hour] = (hourlyStats[hour] || 0) + 1;
        }
      }

      stats.totalAttendees = currentEventCount;
      stats.currentEventAttendees = currentEventCount;
      stats.recentAttendees = recentAttendees.reverse();
      stats.hourlyStats = hourlyStats;
    }

    // 차단된 기기 수 및 목록
    if (deviceSheet && deviceSheet.getLastRow() > 1) {
      var deviceData = deviceSheet.getDataRange().getValues();
      var blockedDevices = [];
      for (var i = 1; i < deviceData.length; i++) {
        var row = deviceData[i];
        var eventId = row[6];

        if (eventId === currentEventId) {
          blockedDevices.push({
            name: row[1],
            department: row[2],
            time: formatDateTime(new Date(row[0])),
            reason: row[5] || '알 수 없음'
          });
        }
      }

      stats.deviceBlocked = blockedDevices.length;
      stats.blockedDevices = blockedDevices.reverse();
    }

    // 이메일 발송 수 조회 (현재 이벤트 기준)
    var emailSheet = spreadsheet.getSheetByName(SHEET_NAMES.EMAIL_LOG);
    if (emailSheet && emailSheet.getLastRow() > 1) {
      var emailData = emailSheet.getDataRange().getValues();
      var emailCount = 0;
      for (var i = 1; i < emailData.length; i++) {
        if (emailData[i][4] === currentEventId) {
          emailCount++;
        }
      }
      stats.emailsSent = emailCount;
    }

    return {
      success: true,
      stats: stats
    };

  } catch (error) {
    console.error('출석 통계 조회 오류:', error);
    return {
      success: false,
      error: '통계 조회 중 오류가 발생했습니다: ' + error.message
    };
  }
}

// ==============================================
// 유틸리티 함수
// ==============================================

/**
 * 날짜 시간 형식 변환
 */
function formatDateTime(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return new Date(date).toLocaleString();
  }
}

/**
 * 오류 로깅
 */
function logError(functionName, error, data) {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var errorSheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.ERROR_LOG, [
      '시간', '함수명', '오류메시지', '사용자데이터', '스택트레이스'
    ]);

    var errorData = [
      new Date(),
      functionName,
      error.message || error.toString(),
      JSON.stringify(data).substring(0, 500),
      error.stack ? error.stack.substring(0, 500) : 'N/A'
    ];

    errorSheet.appendRow(errorData);
  } catch (logError) {
    console.error('오류 로깅 실패:', logError);
  }
}

/**
 * 연결 테스트
 */
function testConnection() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    return {
      success: true,
      message: '연결 성공! 스프레드시트: ' + spreadsheet.getName(),
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return {
      success: false,
      error: '스프레드시트 연결에 실패했습니다.'
    };
  }
}

/**
 * 관리자 인증
 */
function authenticateAdmin(password) {
  try {
    if (!password || typeof password !== 'string') {
      return false;
    }

    var trimmedPassword = password.toString().trim();
    if (trimmedPassword.length === 0) {
      return false;
    }

    try {
      var settings = getSettings();
      var adminPassword = settings.admin_password || 'yongin2024';
      return adminPassword === trimmedPassword;
    } catch (error) {
      return trimmedPassword === 'yongin2024';
    }

  } catch (error) {
    return false;
  }
}

/**
 * 설정 조회
 */
function getSettings() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.SETTINGS, ['키', '값']);

    var data = sheet.getDataRange().getValues();
    var settings = {};

    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];

      if (key && value !== undefined && value !== null) {
        settings[key] = value.toString();
      }
    }

    var finalSettings = {};
    for (var key in DEFAULT_SETTINGS) {
      finalSettings[key] = DEFAULT_SETTINGS[key];
    }
    for (var key in settings) {
      finalSettings[key] = settings[key];
    }

    return finalSettings;

  } catch (error) {
    return DEFAULT_SETTINGS;
  }
}

/**
 * 기기 분석
 */
function analyzeDevice(data) {
  var analysis = {
    type: 'unknown',
    browser: 'unknown',
    isMobile: false,
    isOldDevice: false
  };

  try {
    var userAgent = data.userAgent || '';

    if (/Android|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(userAgent)) {
      analysis.isMobile = true;
      analysis.type = 'mobile';
    } else {
      analysis.type = 'desktop';
    }

    if (userAgent.indexOf('Chrome') > -1) {
      analysis.browser = 'chrome';
    } else if (userAgent.indexOf('Safari') > -1) {
      analysis.browser = 'safari';
    } else if (userAgent.indexOf('Firefox') > -1) {
      analysis.browser = 'firefox';
    }

    // 구형 기기 감지
    if (/iPhone.*OS [1-9]_/i.test(userAgent) ||
        /Android [1-4]\./i.test(userAgent) ||
        /MSIE/i.test(userAgent) ||
        /Trident/i.test(userAgent)) {
      analysis.isOldDevice = true;
    }

  } catch (error) {
    console.error('기기 분석 오류:', error);
  }

  return analysis;
}

/**
 * 위치 검증 (기기별 최적화)
 */
function validateLocationForDevice(userLat, userLng, accuracy, deviceAnalysis) {
  try {
    var settings = getSettings();
    var targetLat = parseFloat(settings.target_latitude);
    var targetLng = parseFloat(settings.target_longitude);
    var baseRadius = parseFloat(settings.location_radius) || 1200;

    var distance = calculateDistance(userLat, userLng, targetLat, targetLng);
    var effectiveRadius = calculateDeviceSpecificRadius(baseRadius, accuracy, deviceAnalysis);

    if (distance <= effectiveRadius) {
      return {
        valid: true,
        distance: distance,
        effectiveRadius: effectiveRadius
      };
    } else {
      return {
        valid: false,
        distance: distance,
        effectiveRadius: effectiveRadius,
        error: '회의실에서 ' + Math.round(distance) + 'm 떨어져 있습니다. (허용: ' + Math.round(effectiveRadius) + 'm)'
      };
    }

  } catch (error) {
    return {
      valid: false,
      error: '위치 확인 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 기기별 허용 반경 계산
 */
function calculateDeviceSpecificRadius(baseRadius, accuracy, deviceAnalysis) {
  var multiplier = 1.0;

  if (deviceAnalysis.isOldDevice) {
    multiplier = 1.5;
  } else if (deviceAnalysis.isMobile) {
    multiplier = 1.2;
  } else {
    multiplier = 1.8;
  }

  if (accuracy > 500) {
    multiplier = multiplier * 1.3;
  } else if (accuracy > 200) {
    multiplier = multiplier * 1.1;
  }

  var effectiveRadius = baseRadius * multiplier;
  effectiveRadius = Math.max(300, effectiveRadius);
  effectiveRadius = Math.min(2500, effectiveRadius);

  return Math.round(effectiveRadius);
}

/**
 * 거리 계산 (Haversine formula)
 */
function calculateDistance(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;

  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);

  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c;
}

/**
 * 기기 ID 생성 (안정적)
 */
function generateImprovedDeviceId(data, deviceAnalysis) {
  try {
    var components = [
      data.userAgent ? data.userAgent.replace(/[\d\.\s]+/g, '').substring(0, 100) : 'unknown',
      data.screenInfo || 'unknown',
      deviceAnalysis.type || 'unknown',
      deviceAnalysis.browser || 'unknown',
      data.timeZone || 'UTC',
      data.language || 'ko',
      data.platform || 'unknown'
    ];

    var combined = components.join('|');

    var hash = 0;
    for (var i = 0; i < combined.length; i++) {
      var char = combined.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }

    var deviceId = 'YU_DEVICE_' + Math.abs(hash).toString(36);

    if (deviceId.length > 20) {
      deviceId = deviceId.substring(0, 20);
    }

    console.log('생성된 기기 ID:', deviceId);
    return deviceId;

  } catch (error) {
    console.error('기기 ID 생성 오류:', error);
    return 'YU_DEVICE_FALLBACK';
  }
}

/**
 * 현재 이벤트 ID 조회
 */
function getCurrentEventId() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.EVENT_CONTROL, [
      '이벤트ID', '생성시간', '상태', '회의제목', '생성자'
    ]);

    var settings = getSettings();
    var currentEventId = settings.current_event_id;

    if (!currentEventId || !isValidEventId(sheet, currentEventId)) {
      currentEventId = createNewEvent(sheet);
      var updatedSettings = {};
      for (var key in settings) {
        updatedSettings[key] = settings[key];
      }
      updatedSettings.current_event_id = currentEventId;
      saveSettings(updatedSettings);
    }

    return currentEventId;

  } catch (error) {
    return 'YU_EVENT_' + Date.now();
  }
}

/**
 * 새 이벤트 생성
 */
function createNewEvent(sheet) {
  try {
    var settings = getSettings();
    var eventId = 'YU_EVENT_' + Date.now();

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === 'active') {
        sheet.getRange(i + 1, 3).setValue('inactive');
      }
    }

    var newRow = [
      eventId,
      new Date(),
      'active',
      settings.meeting_title || '교직원 회의',
      'system'
    ];

    sheet.appendRow(newRow);
    return eventId;

  } catch (error) {
    return 'YU_EVENT_' + Date.now();
  }
}

/**
 * 이벤트 ID 유효성 검사
 */
function isValidEventId(sheet, eventId) {
  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][2] === 'active') {
        return true;
      }
    }
    return false;
  } catch (error) {
    return false;
  }
}

/**
 * 출석 확인 이메일 발송
 */
function sendAttendanceEmail(data, sessionId, distance, eventId) {
  try {
    if (!data.email || !data.email.trim()) {
      return { success: false, reason: 'no_email' };
    }

    var settings = getSettings();
    var subject = '[용인대학교] ' + settings.meeting_title + ' 출석 확인';

    var emailBody = '';
    emailBody += '용인대학교 교직원 회의 출석이 확인되었습니다.\n\n';
    emailBody += '성명: ' + data.name + '\n';
    emailBody += '소속: ' + data.department + '\n';
    emailBody += '출석 시간: ' + new Date().toLocaleString() + '\n';
    emailBody += '회의실 거리: ' + distance + 'm\n';
    emailBody += '세션 ID: ' + sessionId + '\n\n';
    emailBody += '© 2025 용인대학교 출석시스템';

    GmailApp.sendEmail(
      data.email.trim(),
      subject,
      emailBody,
      { name: '용인대학교 출석시스템' }
    );

    // 이메일 발송 로그 기록
    try {
      var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var emailLogSheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.EMAIL_LOG, [
        '시간', '수신자', '성명', '세션ID', '이벤트ID'
      ]);
      emailLogSheet.appendRow([new Date(), data.email.trim(), data.name, sessionId, eventId]);
    } catch (logError) {
      console.error('이메일 로그 기록 실패:', logError);
    }

    return { success: true };

  } catch (error) {
    console.error('이메일 발송 오류:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 설정 저장
 */
function saveSettings(settings) {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.SETTINGS, ['키', '값']);

    sheet.clear();
    sheet.getRange(1, 1, 1, 2).setValues([['키', '값']]);

    var data = [];
    for (var key in settings) {
      data.push([key, settings[key]]);
    }

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 2).setValues(data);
    }

    sheet.getRange(1, 1, 1, 2)
      .setBackground('#4A5568')
      .setFontColor('white')
      .setFontWeight('bold');

    return {
      success: true,
      message: '설정이 저장되었습니다'
    };

  } catch (error) {
    return {
      success: false,
      error: '설정 저장에 실패했습니다: ' + error.message
    };
  }
}

/**
 * 시트 생성 또는 조회
 */
function getOrCreateSheet(spreadsheet, sheetName, headers) {
  try {
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      if (headers && headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length)
          .setBackground('#4A5568')
          .setFontColor('white')
          .setFontWeight('bold');
      }
    }

    return sheet;
  } catch (error) {
    throw error;
  }
}

/**
 * 세션 ID 생성
 */
function generateSessionId() {
  return 'YU_SESSION_' + Date.now().toString(36) + '_' + Math.random().toString(36).substring(2, 8);
}

/**
 * 새 회의 시작
 */
function startNewMeetingFixed(meetingTitle) {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var eventSheet = getOrCreateSheet(spreadsheet, SHEET_NAMES.EVENT_CONTROL, [
      '이벤트ID', '생성시간', '상태', '회의제목', '생성자'
    ]);

    var data = eventSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === 'active') {
        eventSheet.getRange(i + 1, 3).setValue('inactive');
      }
    }

    var eventId = 'YU_EVENT_' + Date.now();
    var newRow = [
      eventId,
      new Date(),
      'active',
      meetingTitle || '교직원 회의',
      'admin'
    ];

    eventSheet.appendRow(newRow);

    var settings = getSettings();
    var updatedSettings = {};
    for (var key in settings) {
      updatedSettings[key] = settings[key];
    }
    updatedSettings.current_event_id = eventId;

    if (meetingTitle && meetingTitle.trim()) {
      updatedSettings.meeting_title = meetingTitle.trim();
    }

    saveSettings(updatedSettings);

    return {
      success: true,
      message: '새 회의가 시작되었습니다.',
      eventId: eventId
    };

  } catch (error) {
    return {
      success: false,
      error: '새 회의 시작 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 시스템 상태 조회
 */
function getSystemStatus() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = getSettings();

    return {
      success: true,
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      currentEventId: settings.current_event_id || 'N/A',
      timestamp: new Date().toISOString(),
      version: '2.0.0'
    };

  } catch (error) {
    return {
      success: false,
      error: '시스템 상태 확인 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 이메일 테스트
 */
function testEmailSystem() {
  try {
    var settings = getSettings();

    if (!settings.notification_email) {
      return {
        success: false,
        error: '알림 이메일을 먼저 설정해주세요.'
      };
    }

    var subject = '[테스트] 용인대학교 출석 시스템';
    var body = '테스트 이메일입니다.\n\n시간: ' + new Date().toLocaleString() + '\n상태: 정상 작동\n버전: 2.0.0\n\n--\n용인대학교 출석 시스템';

    GmailApp.sendEmail(settings.notification_email, subject, body);

    return {
      success: true,
      message: '테스트 이메일이 발송되었습니다!'
    };

  } catch (error) {
    return {
      success: false,
      error: '이메일 테스트 중 오류가 발생했습니다: ' + error.message
    };
  }
}
