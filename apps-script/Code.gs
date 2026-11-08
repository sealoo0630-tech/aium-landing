/**
 * AiUM 랜딩페이지 → 구글 시트 저장
 * 시트 탭: 방문 / 클릭 / 신청 / 사전예약 / 요약
 * 신청·사전예약이 들어오면 시트 소유자에게 알림 메일을 보낸다.
 */

const TABS = {
  // 새 칸은 항상 맨 뒤에 붙인다. 요약 수식이 열 글자(M, S, T)를 쓰기 때문이다.
  visit:      { name: '방문',     cols: ['시각','세션','utm_source','utm_medium','utm_campaign','utm_content','referrer','기기','persona'],
                keys: ['ts','sid','utm_source','utm_medium','utm_campaign','utm_content','referrer','device','persona'] },
  cta:        { name: '클릭',     cols: ['시각','세션','버튼','utm_source','utm_campaign'],
                keys: ['ts','sid','cta','utm_source','utm_campaign'] },
  plan_click: { name: '클릭',     cols: null, keys: ['ts','sid','plan','utm_source','utm_campaign'] },
  lead:       { name: '신청',     cols: ['시각','세션','브랜드','담당자','이메일','연락처','카테고리','셀러규모','라이브','현재방법','설명','마케팅동의','utm_source','utm_medium','utm_campaign','utm_content','기기','처리상태','하는 일','섭외 빈도','인터뷰 동의','회사 규모','persona'],
                keys: ['ts','sid','brand','name','email','phone','category','sellerSize','liveRequired','currentMethod','message','marketingConsent','utm_source','utm_medium','utm_campaign','utm_content','device','_status','role','frequency','interviewOk','companySize','persona'] },
  preorder:   { name: '사전예약', cols: ['시각','세션','요금제','브랜드','이메일','가격 반응','utm_source','utm_campaign','기기','persona'],
                keys: ['ts','sid','plan','brand','email','priceFit','utm_source','utm_campaign','device','persona'] }
};

/**
 * 11/7 추가 칸 반영: 이미 있는 시트에서 한 번 실행한다. 기존 행과 요약 수식은 건드리지 않는다.
 * - 방문·신청·사전예약 탭: 비어 있는 제목 칸에만 새 열 제목을 쓴다.
 * - 요약 탭: G15(persona별), J15(회사 규모별) 집계. 그 자리에 이미 무언가 있으면 건너뛴다.
 */
function addColumns1107() {
  const ss = ss_();
  ['visit', 'lead', 'preorder'].forEach(k => {
    const t = TABS[k];
    const sh = ss.getSheetByName(t.name);
    if (!sh) return;
    t.cols.forEach((c, i) => {
      const cell = sh.getRange(1, i + 1);
      if (cell.getValue() === '') cell.setValue(c).setFontWeight('bold');
    });
  });
  const s = ss.getSheetByName('요약');
  if (!s) { Logger.log('요약 탭 없음: setup()을 먼저 실행'); return; }
  const put = (a1, title, col) => {
    const head = s.getRange(a1), body = head.offset(1, 0, 6, 2);
    if (head.getValue() !== '' || body.getValues().some(r => r.some(v => v !== ''))) { Logger.log(a1 + ' 자리에 값이 있어 건너뜀'); return; }
    head.setValue(title).setFontWeight('bold');
    head.offset(1, 0).setFormula(`=IFERROR(QUERY(신청!${col}2:${col}, "select ${col}, count(${col}) where ${col} is not null group by ${col} label count(${col}) '신청'", 0), "아직 없음")`);
  };
  put('G15', '광고 페르소나(persona)별 신청 수', 'W');
  put('J15', '회사 규모별 신청 수', 'V');
  Logger.log('11/7 칸 추가 완료');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const t = TABS[d.type];
    if (!t) return out('unknown type');
    d.ts = new Date();
    d._status = '샘플 발송 전';
    const sh = sheet_(t.name, (t.cols || TABS.cta.cols));
    sh.appendRow(t.keys.map(k => clean_(d[k])));
    if (d.type === 'lead' || d.type === 'preorder') notify_(d);
    return out('ok');
  } catch (err) {
    return out('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out('AiUM landing endpoint ok'); }

/** 시트에 붙은 스크립트면 그 시트, 따로 만든 스크립트면 「AiUM 랜딩 DB」를 새로 만들어 기억 */
function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    id = SpreadsheetApp.create('AiUM 랜딩 DB').getId();
    props.setProperty('SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}

/** 처음 한 번 실행: 탭과 요약(퍼널) 만들기 */
function setup() {
  Object.values(TABS).forEach(t => t.cols && sheet_(t.name, t.cols));
  const ss = ss_();
  const s = ss.getSheetByName('요약') || ss.insertSheet('요약', 0);
  s.clear();
  s.getRange('A1:D1').setValues([['단계','수','전환율','목표(사업계획서)']]).setFontWeight('bold');
  s.getRange('A2:D6').setValues([
    ['방문(세션)', '=COUNTA(방문!B2:B)', '', ''],
    ['신청 버튼 클릭', '=COUNTA(클릭!C2:C)-B5', '=IFERROR(B3/B2,"")', ''],
    ['샘플 리포트 신청', '=COUNTA(신청!B2:B)', '=IFERROR(B4/B2,"")', '방문 대비 5% 이상, 30건'],
    ['요금제 버튼 클릭', '=COUNTIF(클릭!C2:C,"Starter")+COUNTIF(클릭!C2:C,"Growth")+COUNTIF(클릭!C2:C,"Enterprise")', '', ''],
    ['사전예약', '=COUNTA(사전예약!B2:B)', '=IFERROR(B6/B4,"")', '신청 대비 10% 이상, 3건']
  ]);
  s.getRange('C2:C6').setNumberFormat('0.0%');
  s.getRange('A8').setValue('광고 채널별 신청 수').setFontWeight('bold');
  s.getRange('A9').setFormula('=IFERROR(QUERY(신청!M2:M, "select M, count(M) where M is not null group by M label count(M) \'신청\'", 0), "아직 없음")');
  s.getRange('A15').setValue('하는 일(페르소나)별 신청 수').setFontWeight('bold');
  s.getRange('A16').setFormula('=IFERROR(QUERY(신청!S2:S, "select S, count(S) where S is not null group by S label count(S) \'신청\'", 0), "아직 없음")');
  s.getRange('D15').setValue('섭외 빈도별 신청 수').setFontWeight('bold');
  s.getRange('D16').setFormula('=IFERROR(QUERY(신청!T2:T, "select T, count(T) where T is not null group by T label count(T) \'신청\'", 0), "아직 없음")');
  // 이미 있는 「신청」 탭에 새 열 제목 채우기
  const lead = ss.getSheetByName(TABS.lead.name);
  if (lead) lead.getRange(1, 1, 1, TABS.lead.cols.length).setValues([TABS.lead.cols]).setFontWeight('bold');
  addColumns1107();
  s.autoResizeColumns(1, 4);
  const extra = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (extra && ss.getSheets().length > 1) ss.deleteSheet(extra);
  Logger.log('완료. 시트 주소: ' + ss.getUrl());
}

function sheet_(name, cols) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(cols);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 수식 주입 방지 + 길이 제한
function clean_(v) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) return v;
  let s = String(v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function notify_(d) {
  try {
    const me = Session.getEffectiveUser().getEmail();
    const title = d.type === 'lead' ? `[AiUM] 샘플 리포트 신청: ${d.brand}` : `[AiUM] 사전예약(${d.plan}): ${d.brand}`;
    const body = d.type === 'lead'
      ? `하는 일: ${d.role || '-'} (섭외 빈도: ${d.frequency || '-'}, 인터뷰: ${d.interviewOk || 'N'})\n브랜드: ${d.brand}\n담당자: ${d.name}\n이메일: ${d.email}\n카테고리: ${d.category}\n셀러 규모: ${d.sellerSize || '-'}\n라이브: ${d.liveRequired || '-'}\n현재 방법: ${d.currentMethod || '-'}\n회사 규모: ${d.companySize || '-'}\n설명: ${d.message || '-'}\n유입: ${d.utm_source || '직접'} / ${d.utm_campaign || '-'} / persona ${d.persona || '-'}\n\n48시간 안에 샘플 리포트를 보내세요.`
      : `요금제: ${d.plan}\n브랜드: ${d.brand}\n이메일: ${d.email}\n요금 방식(11/7 전 신청은 가격 반응): ${d.priceFit}\n유입: ${d.utm_source || '직접'}`;
    MailApp.sendEmail(me, title, body);
  } catch (e) {}
}

function out(msg) {
  return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT);
}
