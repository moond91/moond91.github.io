// ═══════════════════════════════════════════════
//  우토그록 트리거롱폼 v2.0.0
//  닉네임 검증 + 가입신청 + 승인관리 + 보너스연장 + 텔레그램 + 로그
//  Google Apps Script (웹앱 배포)
//
//  수강생명단 시트 구조:
//  A:강의명 | B:닉네임 | C:반 | D:상태 | E:신청일시 | F:승인일시 | G:만료일 | H:보너스(월)
//
//  ※ 로그인 검증: B열(닉네임) + CONFIG.PASSWORD
//  ※ 만료 체크: G열(만료일)
//  ※ C열(반): 일반반 / 프리미엄반
//  ※ D열(상태): 승인(초록) / 대기(주황) / 만료(빨강) / 관리자(파랑)
//  ※ 보너스 연장: H열에 숫자 입력 후 applyBonus() 실행 → G열 만료일 연장
// ═══════════════════════════════════════════════

const CONFIG = {
  VERSION: 'v2.0.0',
  COURSE_NAME: '트리거롱폼',
  SHEET_ID: '1dBO7b8ydKi1JpOpPVbaAqZjPVGrsxap10XR6fRtdNko',
  SHEET_NAME: '수강생명단',
  LOG_SHEET_NAME: '접속로그',
  APPLY_SHEET_NAME: '가입신청',
  VISIT_LOG_SHEET_NAME: '매뉴얼방문로그',
  PASSWORD: '0501',
  TELEGRAM_BOT_TOKEN: '8766741308:AAEMTS6uO7SL7iSQ6-pEEP0GGb2Wo34gfgw',
  TELEGRAM_CHAT_ID: '8501291449',

  // 수강생명단 열 매핑 (1-indexed)
  COL: {
    COURSE: 1,       // A 강의명
    NICKNAME: 2,     // B 닉네임
    CLASS: 3,        // C 반
    STATUS: 4,       // D 상태
    APPLY_DATE: 5,   // E 신청일시
    APPROVE_DATE: 6, // F 승인일시
    EXPIRE_DATE: 7,  // G 만료일
    BONUS: 8,        // H 보너스(월)
  }
};

// ═══════════════════════════════════════════════
//  웹앱 엔드포인트
// ═══════════════════════════════════════════════
//  시트 열 때 커스텀 메뉴 생성
// ═══════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔧 우토그록 관리')
    .addItem('✅ 일괄 승인 (대기 전체)', 'approveAll')
    .addItem('🎁 보너스 일괄 적용', 'applyBonus')
    .addSeparator()
    .addItem('📊 만료일 재계산 (전체)', 'recalcExpiry')
    .addItem('🧹 데이터 정리 (빈 행 제거)', 'cleanEmptyRows')
    .addSeparator()
    .addItem('🔧 시트 초기 설정', 'setupMemberSheet')
    .addItem('📡 텔레그램 테스트', 'testTelegram')
    .addToUi();
}

// ═══════════════════════════════════════════════
//  만료일 재계산 — G열 수식 다시 적용
// ═══════════════════════════════════════════════

function recalcExpiry() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var C = CONFIG.COL;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var formulas = [];
  for (var r = 2; r <= lastRow; r++) {
    formulas.push(['=IF(D' + r + '="관리자","무제한",IF(F' + r + '="","",TEXT(EDATE(F' + r + ',6+IF(H' + r + '="",0,H' + r + ')),"yyyy-mm-dd")))']);
  }
  sheet.getRange(2, C.EXPIRE_DATE, formulas.length, 1).setFormulas(formulas);
  SpreadsheetApp.flush();

  Logger.log('만료일 재계산 완료: ' + formulas.length + '행');
  SpreadsheetApp.getUi().alert('✅ 만료일 재계산 완료: ' + formulas.length + '명');
}

// ═══════════════════════════════════════════════
//  데이터 정리 — 빈 행 제거 (B열 닉네임 기준)
// ═══════════════════════════════════════════════

function cleanEmptyRows() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var count = 0;
  for (var r = lastRow; r >= 2; r--) {
    var nick = sheet.getRange(r, CONFIG.COL.NICKNAME).getValue();
    if (!nick || String(nick).trim() === '') {
      sheet.deleteRow(r);
      count++;
    }
  }

  Logger.log('빈 행 ' + count + '개 제거');
  SpreadsheetApp.getUi().alert('🧹 빈 행 ' + count + '개 제거 완료');
}

// ═══════════════════════════════════════════════
//  텔레그램 테스트
// ═══════════════════════════════════════════════

function testTelegram() {
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  sendTelegram('🔔 [' + CONFIG.COURSE_NAME + ' ' + CONFIG.VERSION + '] 텔레그램 테스트\n🕐 ' + now + '\n✅ 정상 작동');
  SpreadsheetApp.getUi().alert('✅ 텔레그램 테스트 메시지 전송 완료!');
}

// ═══════════════════════════════════════════════
//  웹앱 엔드포인트
// ═══════════════════════════════════════════════

function doGet(e) {
  const action = (e.parameter.action || 'login').trim();
  const nickname = (e.parameter.nickname || '').trim();
  const password = (e.parameter.password || '').trim();
  const course = (e.parameter.course || '').trim();
  const source = (e.parameter.source || 'web').trim();

  if (action === 'apply') return handleApply(nickname, course, source);
  if (action === 'visit') return handleVisit(e.parameter);
  return handleLogin(nickname, password, course, source);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const source = data.source || 'web';
    if (data.action === 'apply') return handleApply(data.nickname, data.course, source);
    if (data.action === 'visit') return handleVisit(data);
    return handleLogin(data.nickname, data.password, data.course, source);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ═══════════════════════════════════════════════
//  로그인 — B열(닉네임) 검색, G열(만료일) 체크
// ═══════════════════════════════════════════════

function handleLogin(nickname, password, course, source) {
  if (!nickname || !password) return jsonResponse({ success: false, error: 'missing_params' });

  const sourceLabel = source === 'extension' ? '🧩 확장프로그램' : '🌐 매뉴얼';
  const C = CONFIG.COL;

  if (password !== CONFIG.PASSWORD) {
    writeLog(nickname, course, '❌ 비번 오류', sourceLabel);
    return jsonResponse({ success: false, error: 'wrong_password' });
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    writeLog(nickname, course, '❌ 미등록', sourceLabel);
    return jsonResponse({ success: false, error: 'not_found' });
  }

  // B열(닉네임) ~ G열(만료일) 읽기 — D열(상태)도 포함
  const data = sheet.getRange(2, C.NICKNAME, lastRow - 1, C.EXPIRE_DATE - C.NICKNAME + 1).getValues();
  const expIdx = C.EXPIRE_DATE - C.NICKNAME; // G열의 배열 인덱스
  const statusIdx = C.STATUS - C.NICKNAME;   // D열의 배열 인덱스

  let found = false;
  let expired = false;
  let isAdmin = false;

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === nickname) {
      found = true;
      const status = String(data[i][statusIdx]).trim();

      // 관리자는 만료 체크 안 함
      if (status === '관리자') {
        isAdmin = true;
        break;
      }

      const expDate = data[i][expIdx];
      // "무제한"이면 만료 안 됨
      if (String(expDate).trim() === '무제한') {
        break;
      }
      if (expDate && expDate instanceof Date && new Date() > expDate) {
        expired = true;
      }
      break;
    }
  }

  if (!found) {
    writeLog(nickname, course, '❌ 미등록', sourceLabel);
    return jsonResponse({ success: false, error: 'not_found' });
  }

  if (expired) {
    writeLog(nickname, course, '⏰ 만료', sourceLabel);
    return jsonResponse({ success: false, error: 'expired' });
  }

  writeLog(nickname, course, '✅ 접속', sourceLabel);

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  sendTelegram(`📋 [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] ${sourceLabel} 접속\n👤 닉네임: ${nickname}\n📚 강의: ${course}\n🕐 ${now}`);

  return jsonResponse({ success: true });
}

// ═══════════════════════════════════════════════
//  가입 신청
// ═══════════════════════════════════════════════

function handleApply(nickname, course, source) {
  if (!nickname) return jsonResponse({ success: false, error: 'missing_nickname' });

  const sourceLabel = source === 'extension' ? '🧩 확장프로그램' : '🌐 매뉴얼';
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let applySheet = ss.getSheetByName(CONFIG.APPLY_SHEET_NAME);

  if (!applySheet) {
    applySheet = ss.insertSheet(CONFIG.APPLY_SHEET_NAME);
    applySheet.appendRow(['신청일시', '닉네임', '강의명', '경로', '상태']);
    applySheet.getRange('1:1').setFontWeight('bold');
    applySheet.getRange(1, 1, 1, 5).setBackground('#4285F4').setFontColor('#FFFFFF');
    applySheet.setFrozenRows(1);
  }

  const lastRow = applySheet.getLastRow();
  if (lastRow >= 2) {
    const existing = applySheet.getRange('B2:B' + lastRow).getValues().flat().map(v => String(v).trim());
    if (existing.includes(nickname.trim())) {
      return jsonResponse({ success: false, error: 'already_applied' });
    }
  }

  const memberSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const memberLastRow = memberSheet.getLastRow();
  if (memberLastRow >= 2) {
    const nicks = memberSheet.getRange(2, CONFIG.COL.NICKNAME, memberLastRow - 1, 1).getValues().flat().map(v => String(v).trim());
    if (nicks.includes(nickname.trim())) {
      return jsonResponse({ success: false, error: 'already_registered' });
    }
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  applySheet.appendRow([now, nickname, course || '', sourceLabel, '대기']);

  writeLog(nickname, course, '📝 가입신청', sourceLabel);

  sendTelegram(`📝 [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] 가입 신청\n👤 닉네임: ${nickname}\n📚 강의: ${course}\n📍 경로: ${sourceLabel}\n🕐 ${now}\n\n👉 수강생명단 시트에 추가해주세요!`);

  return jsonResponse({ success: true });
}

// ═══════════════════════════════════════════════
//  매뉴얼 방문 로그
// ═══════════════════════════════════════════════

function handleVisit(params) {
  const referrer = (params.referrer || '').trim();
  const userAgent = (params.userAgent || '').trim();
  const pageUrl = (params.pageUrl || '').trim();

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let visitSheet = ss.getSheetByName(CONFIG.VISIT_LOG_SHEET_NAME);

  if (!visitSheet) {
    visitSheet = ss.insertSheet(CONFIG.VISIT_LOG_SHEET_NAME);
    visitSheet.appendRow(['일시', '강의명', '페이지URL', '리퍼러', 'UserAgent']);
    visitSheet.getRange('1:1').setFontWeight('bold');
    visitSheet.getRange(1, 1, 1, 5).setBackground('#4285F4').setFontColor('#FFFFFF');
    visitSheet.setFrozenRows(1);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  visitSheet.appendRow([now, CONFIG.COURSE_NAME, pageUrl, referrer, userAgent]);

  const isMobile = /Mobile|Android|iPhone/i.test(userAgent);
  const deviceLabel = isMobile ? '📱 모바일' : '💻 PC';

  sendTelegram(`📖 [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] 매뉴얼 방문\n${deviceLabel}\n🔗 ${pageUrl || '(직접접속)'}\n↩️ ${referrer || '(없음)'}\n🕐 ${now}`);

  return jsonResponse({ success: true });
}

// ═══════════════════════════════════════════════
//  접속로그
// ═══════════════════════════════════════════════

function writeLog(nickname, course, status, source) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    logSheet.appendRow(['일시', '닉네임', '강의명', '상태', '경로']);
    logSheet.getRange('1:1').setFontWeight('bold');
    logSheet.getRange(1, 1, 1, 5).setBackground('#4285F4').setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([now, nickname, course || '', status, source || '']);
}

// ═══════════════════════════════════════════════
//  승인 처리 (수동 실행)
// ═══════════════════════════════════════════════

function approveAll() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const applySheet = ss.getSheetByName(CONFIG.APPLY_SHEET_NAME);
  const memberSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const C = CONFIG.COL;

  if (!applySheet) { Logger.log('가입신청 탭 없음'); return; }

  const lastRow = applySheet.getLastRow();
  if (lastRow < 2) { Logger.log('신청 없음'); return; }

  const data = applySheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();
  const expDate = new Date(now);
  expDate.setMonth(expDate.getMonth() + 6);

  const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
  const expStr = Utilities.formatDate(expDate, 'Asia/Seoul', 'yyyy-MM-dd');

  let count = 0;

  for (let i = 0; i < data.length; i++) {
    if (data[i][4] === '대기') {
      const nick = data[i][1];
      const course = data[i][2] || CONFIG.COURSE_NAME;
      const applyDate = data[i][0];

      // A:강의명 B:닉네임 C:반(빈칸) D:상태 E:신청일시 F:승인일시 G:만료일 H:보너스(빈칸)
      memberSheet.appendRow([course, nick, '', '승인', applyDate, nowStr, expStr, '']);

      applySheet.getRange(i + 2, 5).setValue('승인완료');
      count++;
    }
  }

  Logger.log(count + '명 승인 완료');

  if (count > 0) {
    sendTelegram(`✅ [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] ${count}명 일괄 승인 완료\n📅 승인일: ${nowStr}\n📅 만료일: ${expStr}`);
  }
}

// ═══════════════════════════════════════════════
//  보너스 연장 (수동 실행)
//  H열(보너스 월)에 숫자 입력 후 이 함수 실행
//  → G열(만료일) 기준으로 해당 월수만큼 연장 → H열 비움
// ═══════════════════════════════════════════════

function applyBonus() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const C = CONFIG.COL;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('데이터 없음'); return; }

  const nicknames = sheet.getRange(2, C.NICKNAME, lastRow - 1, 1).getValues();
  const expDates = sheet.getRange(2, C.EXPIRE_DATE, lastRow - 1, 1).getValues();
  const bonuses = sheet.getRange(2, C.BONUS, lastRow - 1, 1).getValues();

  let count = 0;
  const results = [];

  for (let i = 0; i < bonuses.length; i++) {
    const bonusVal = bonuses[i][0];

    if (bonusVal && !isNaN(Number(bonusVal)) && Number(bonusVal) > 0) {
      const months = Math.floor(Number(bonusVal));
      const nick = String(nicknames[i][0]).trim();
      const row = i + 2;

      let baseDate = expDates[i][0];

      if (baseDate && baseDate instanceof Date) {
        const newExpDate = new Date(baseDate);
        newExpDate.setMonth(newExpDate.getMonth() + months);
        const newExpStr = Utilities.formatDate(newExpDate, 'Asia/Seoul', 'yyyy-MM-dd');
        const oldExpStr = Utilities.formatDate(baseDate, 'Asia/Seoul', 'yyyy-MM-dd');

        sheet.getRange(row, C.EXPIRE_DATE).setValue(newExpStr);
        sheet.getRange(row, C.BONUS).setValue('');
        results.push(`${nick}: ${oldExpStr} → ${newExpStr} (+${months}개월)`);
      } else {
        const today = new Date();
        const newExpDate = new Date(today);
        newExpDate.setMonth(newExpDate.getMonth() + months);
        const newExpStr = Utilities.formatDate(newExpDate, 'Asia/Seoul', 'yyyy-MM-dd');

        sheet.getRange(row, C.EXPIRE_DATE).setValue(newExpStr);
        sheet.getRange(row, C.BONUS).setValue('');
        results.push(`${nick}: (오늘기준) → ${newExpStr} (+${months}개월)`);
      }
      count++;
    }
  }

  Logger.log(count + '명 보너스 연장 완료');

  if (count > 0) {
    sendTelegram(`🎁 [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] 보너스 연장 ${count}명\n\n${results.join('\n')}`);
  } else {
    Logger.log('H열(보너스)에 숫자가 입력된 행이 없습니다.');
  }
}

// ═══════════════════════════════════════════════
//  수강생명단 셋업 + 닉네임 복구 (최초 1회)
//
//  ⚠️ 시트를 완전히 지우고 새로 씁니다!
// ═══════════════════════════════════════════════

function setupMemberSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  const now = new Date();
  const expDate = new Date(now);
  expDate.setMonth(expDate.getMonth() + 6);
  const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
  const expStr = Utilities.formatDate(expDate, 'Asia/Seoul', 'yyyy-MM-dd');

  // 트리거롱폼 수강생 전원
  const nicknames = [
    '우미1','김히엉','Goldy','Hoyt','Karline','Miraflores','Secret Gorilla','Yum',
    'crystal0213','dochipapa','fermata','kingplay','luck맘','mtools','nameisbk','sall',
    '감동지나','감자나무떡','강남콩','강철궁딩이','거나쏠','겨울','경기도민인','계획된우연',
    '고라파덕','고야','고튼이','골드맘','곰고미','구백프로','구일일','그라티아777','기로로',
    '길작가','김꿀벌','김독끼','김연나','김캐리','김한다','깐따삐야','꿈꾸는 문어','뀰기',
    '나는나야','나도해','나라올라','나령','나에게도','나옹이','날다진','내추럴스윽팍','내퍼',
    '냐냐2','냐하이','뉴저지쌤','늘보랑','다누리','다린','다비드강','다빈치','다엘',
    '다이렉트콜','다크8250','다하온','다해냄','단순하게','단호한연어','달려라사슴','달빛',
    '달소리','달콩빠빠','당산로송혜교','대전93','데니','데헷이','도롱이','도마드나','도이석',
    '도지파파','돈지','돌돌','돌멩','돗토리','동감','동규니','둥둥고래','드림로더',
    '디고고','디노하','디얼쓰','딩딩푸야수박이네','딸하나','떨리치','또콩이아빠','라라라123',
    '라이라이언','라이키','러블리러브2','럭키한정군','레메','레이지캣','레제','로니','로블리맘',
    '로하로','루네스','루밍','루튜브','리채','린','마수니','만타코랄','망냉이','망냐뇽',
    '망뎅이','매니머니','머니들','머니의힘','머니커','머니토리','먼지','모나코왕','모닝라떼',
    '모락모락','모모아맘','모아모a','모짜','몽몽그리','무사222','미러클요이','미소하나','미쯔',
    '바다소리','바이올렛2727','반전','베레베레','베스킨','별천지','보경운','복덩이','복승아',
    '봄내','부루스타','부산오이와당근','브랜드데이','블랙고','블랙엔젤','블루라이트','블리블리',
    '블링블링','비채맘','비케이','빈그림','빛나는이야기','빨강맛','빵미녀','빼삐','뽀기성',
    '뽀름','쁘니쏜','쁘미0','삐용','산료','선짱','섬유유연제','성공한덕후','성공한아빠',
    '세눈까마귀','소금항아리','소심한토끼','소쏘','소피','소히','솔라리움','쇼츠코인','수더기',
    '수림','수혜짱','순대버핏','슈블리','슈크림카레','시마다','싹싹쏙쏙 프로강','씬넘버',
    '아랑동','아미','아미슈','아스트라','아울투유','아이캔두잇','아지8648','안프','앙찌모찌',
    '애니타임','엄마의도전','엄지손','에나','에나아삽','에센','에코엔코','엘리12','여름가을하늘',
    '열두시간','영영1','오뉴러브','오늘행복','오밥','오부자로','오오오영','오키바리도키','오타루',
    '온도도','온포인트','와사장','왕밤1','요미링','요플렛','용의기사','용코','우락이','우맘이',
    '우키키키2','월삼천가자','월천부인','월천줌마','유님','유자차차','유카엘','유툴룰루',
    '유혈낭자','윤슬하','윰주','으쓱우쓱','은율잉','은전갈','이고요','인유티','일나제','임찬',
    '잇타임','잼살','저스트두win','정미인','정워리','정이사','제2','제로콜라','조각모음',
    '조네오','좋은기운','주니파','주베르','주차금지','준밀리언5','중껑마','지니남편감','쨔요민',
    '쭈니유니맘','찌니0','찌니찌닝','찌뭉','찐심','차근이','차우','차진교','찬찬감동',
    '천억버는아라맘','청이','초롱초롱초롱','최고','춘식이에오','카이젠','캔두우','커피목수',
    '컨트리맨','컴트루','케리아','케이즈','콤마','쿠니','쿠키','쿠키쿠키','쿨파워냉방',
    '퀄퀄퀄','킹계발','킹그릇','탁구','테스스스','토리콩','튜플','티보라','파람','펠리즈',
    '포포씨','포포제니','폴리','푸른잎','풍요','프로꼼지락','하다','하둥파','하루야챙',
    '한네기','해링','해품써','해피경자','해피데이12','행복바람','행복찾기1','행복해지기',
    '헐그덕뜨','헤더킴','헤롱헤롱','헤이지니쩡','헬리네이션','혀니닝','혜잔','혜지니',
    '호이끼약','호주아재','혼돈','화이팅','황금빛','효군','희망찬이','희파람','흰사자',
    '히히히히흐히','힐링여행','힐스'
  ];

  // ── 1단계: 시트 전체 지우기 (validation 포함) ──
  sheet.clear();
  sheet.clearConditionalFormatRules();
  // 기존 드롭다운 validation 완전 제거
  var maxR = sheet.getMaxRows();
  var maxC = sheet.getMaxColumns();
  if (maxR > 0 && maxC > 0) {
    sheet.getRange(1, 1, maxR, maxC).clearDataValidations();
  }
  SpreadsheetApp.flush();

  // ── 2단계: 헤더 ──
  const headers = ['강의명', '닉네임', '반', '상태', '신청일시', '승인일시', '만료일', '보너스(월)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF');

  // ── 3단계: 수강생 데이터 입력 ──
  const admins = ['우미1', '김히엉'];
  const rows = nicknames.map(function(nick) {
    var status = admins.indexOf(nick) >= 0 ? '관리자' : '승인';
    // A:강의명 B:닉네임 C:반(빈칸) D:상태 E:신청일시 F:승인일시 G:만료일(수식) H:보너스(빈칸)
    return [CONFIG.COURSE_NAME, nick, '', status, nowStr, nowStr, '', ''];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);

    // G열(만료일) = 승인일(F열) + 6개월 + 보너스(H열) 자동 계산
    // D열이 "관리자"면 "무제한"
    // H열 보너스 숫자만큼 추가 연장
    var formulas = [];
    for (var r = 0; r < rows.length; r++) {
      var rowNum = r + 2;
      formulas.push(['=IF(D' + rowNum + '="관리자","무제한",IF(F' + rowNum + '="","",TEXT(EDATE(F' + rowNum + ',6+IF(H' + rowNum + '="",0,H' + rowNum + ')),"yyyy-mm-dd")))']);
    }
    sheet.getRange(2, 7, rows.length, 1).setFormulas(formulas);
  }

  const totalRows = Math.max(rows.length, 500);

  // 빈 행에도 G열 수식 미리 넣기 (새 수강생 추가 대비)
  var emptyFormulas = [];
  for (var r2 = rows.length; r2 < totalRows; r2++) {
    var rowNum2 = r2 + 2;
    emptyFormulas.push(['=IF(D' + rowNum2 + '="관리자","무제한",IF(F' + rowNum2 + '="","",TEXT(EDATE(F' + rowNum2 + ',6+IF(H' + rowNum2 + '="",0,H' + rowNum2 + ')),"yyyy-mm-dd")))']);
  }
  if (emptyFormulas.length > 0) {
    sheet.getRange(rows.length + 2, 7, emptyFormulas.length, 1).setFormulas(emptyFormulas);
  }

  // ── 4단계: 열 너비 ──
  sheet.setColumnWidth(1, 120);  // A 강의명
  sheet.setColumnWidth(2, 160);  // B 닉네임
  sheet.setColumnWidth(3, 100);  // C 반
  sheet.setColumnWidth(4, 80);   // D 상태
  sheet.setColumnWidth(5, 120);  // E 신청일시
  sheet.setColumnWidth(6, 120);  // F 승인일시
  sheet.setColumnWidth(7, 120);  // G 만료일
  sheet.setColumnWidth(8, 90);   // H 보너스(월)

  // ── 5단계: 드롭다운 ──

  // A열(강의명)
  var courseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['트리거롱폼', '실버인포롱폼'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 1, totalRows, 1).setDataValidation(courseRule);

  // C열(반)
  var classRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['일반반', '프리미엄반'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 3, totalRows, 1).setDataValidation(classRule);

  // D열(상태)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['승인', '대기', '만료', '관리자'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 4, totalRows, 1).setDataValidation(statusRule);

  // H열(보너스)
  var bonusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '12'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 8, totalRows, 1).setDataValidation(bonusRule);

  // ── 6단계: D열(상태) 조건부 서식 — 색상 ──
  var statusRange = sheet.getRange('D2:D' + (totalRows + 1));

  // 승인 = 초록
  var ruleApproved = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('승인')
    .setBackground('#D9EAD3')
    .setFontColor('#38761D')
    .setRanges([statusRange])
    .build();

  // 대기 = 주황
  var ruleWaiting = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('대기')
    .setBackground('#FCE5CD')
    .setFontColor('#B45F06')
    .setRanges([statusRange])
    .build();

  // 만료 = 빨강
  var ruleExpired = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('만료')
    .setBackground('#F4CCCC')
    .setFontColor('#CC0000')
    .setRanges([statusRange])
    .build();

  // 관리자 = 파랑
  var ruleAdmin = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('관리자')
    .setBackground('#CFE2F3')
    .setFontColor('#1155CC')
    .setRanges([statusRange])
    .build();

  // ── 7단계: C열(반) 조건부 서식 — 색상 ──
  var classRange = sheet.getRange('C2:C' + (totalRows + 1));

  // 일반반 = 회색
  var ruleNormal = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('일반반')
    .setBackground('#EFEFEF')
    .setFontColor('#666666')
    .setRanges([classRange])
    .build();

  // 프리미엄반 = 골드
  var rulePremium = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('프리미엄반')
    .setBackground('#FFF2CC')
    .setFontColor('#BF8F00')
    .setRanges([classRange])
    .build();

  // 조건부 서식 일괄 적용
  sheet.setConditionalFormatRules([
    ruleApproved, ruleWaiting, ruleExpired, ruleAdmin,
    ruleNormal, rulePremium
  ]);

  // ── 8단계: 닉네임 가나다순 정렬 ──
  sheet.getRange(2, 1, rows.length, 8).sort({column: 2, ascending: true});

  // ── 9단계: 1행 고정 ──
  sheet.setFrozenRows(1);

  Logger.log('✅ 셋업 완료: ' + nicknames.length + '명 등록');

  sendTelegram(`🔧 [${CONFIG.COURSE_NAME} ${CONFIG.VERSION}] 시트 셋업 완료\n📊 ${nicknames.length}명 등록\n📋 A강의명 B닉네임 C반 D상태 E신청일 F승인일 G만료일 H보너스`);
}

// ═══════════════════════════════════════════════
//  onEdit — 자동 트리거 (시트 편집 시 자동 실행)
//
//  1) H열(보너스) 숫자 입력 → G열(만료일) 자동 연장 → H열 비움
//  2) D열(상태) "관리자" 선택 → G열(만료일) "무제한" 자동 입력
// ═══════════════════════════════════════════════

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row < 2) return;

  var C = CONFIG.COL;

  // ── H열(보너스) 입력 시: 텔레그램 알림만 (만료일은 수식이 자동 계산) ──
  if (col === C.BONUS) {
    var bonusVal = e.range.getValue();
    if (bonusVal && !isNaN(Number(bonusVal)) && Number(bonusVal) > 0) {
      var months = Math.floor(Number(bonusVal));
      var nick = sheet.getRange(row, C.NICKNAME).getValue();

      // 상태가 만료였으면 승인으로 변경
      var statusCell = sheet.getRange(row, C.STATUS);
      if (statusCell.getValue() === '만료') {
        statusCell.setValue('승인');
      }

      // 수식 재계산 대기 후 새 만료일 읽기
      SpreadsheetApp.flush();
      var newExp = sheet.getRange(row, C.EXPIRE_DATE).getDisplayValue();

      sendTelegram('🎁 [' + CONFIG.COURSE_NAME + ' ' + CONFIG.VERSION + '] 보너스 연장\n👤 ' + nick + '\n📅 만료일: ' + newExp + ' (+' + months + '개월 보너스)');
    }
  }
}

// ═══════════════════════════════════════════════
//  텔레그램 / 유틸
// ═══════════════════════════════════════════════

function sendTelegram(text) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: CONFIG.TELEGRAM_CHAT_ID, text: text }),
    muteHttpExceptions: true
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
