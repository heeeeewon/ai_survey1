/* AI 기반 보건조사방법론 2차 과제 및 중간발표 피드백
 * 기존 스프레드시트에 연결된 Apps Script의 Code.gs를 교체합니다.
 * setup()은 기존 Feedback 데이터와 Settings 조별 코드를 보존합니다.
 * 2차 제출물은 Feedback_Sprint2에 저장합니다.
 */
const COURSE_ID = 'daejin-ai-survey-2026-sprint1';
const COURSE_ID_2 = 'daejin-ai-survey-2026-sprint2';
const GROUPS = ['1','2','3','4'];
const SETTINGS_SHEET = 'Settings';
const FEEDBACK_SHEET = 'Feedback';
const FEEDBACK_SHEET_2 = 'Feedback_Sprint2';
const VISIBLE_ON_SUBMIT = true;

const CRITERIA_V1 = [
  ['topic','연구주제의 적절성과 범위'], ['question','연구문제의 구체성'],
  ['concept','핵심 개념의 명확성'], ['hypothesis','연구가설의 적절성과 일관성'],
  ['evidence','선행연구와 가설의 연결'], ['ai','AI 활용과 인간의 검증 및 판단']
];
const CRITERIA_V2 = [
  ['question','연구문제와 가설 및 수정 내용'], ['design','조사설계와 선택 근거'],
  ['population','목표 모집단과 표집틀'], ['sampling','표본추출 단위와 방법 및 크기'],
  ['collection','자료수집방법과 선택 근거'], ['ai','AI 제안에 대한 판단과 기록']
];

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, FEEDBACK_SHEET, headers_(false));
  ensureSheet_(ss, FEEDBACK_SHEET_2, headers_(true));
  let settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.appendRow(['team','code']);
    GROUPS.forEach(team => settings.appendRow([team, makeCode_()]));
    settings.setFrozenRows(1);
  }
}

function headers_(sprint2) {
  const h = ['submittedAt','recordId','fromTeam','toTeam','projectTitle'];
  (sprint2 ? CRITERIA_V2 : CRITERIA_V1).forEach(([id]) => {
    if (!sprint2) h.push(id+'Status');
    h.push(id+'Comment');
  });
  return h.concat(['strength','clarification','priority','suggestion','rewrite','visible','payload']);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function course_(id) {
  if (id === COURSE_ID_2) return {id:id, title:'2차 과제 및 중간발표 피드백', sheet:FEEDBACK_SHEET_2, criteria:CRITERIA_V2, sprint2:true};
  if (id === COURSE_ID) return {id:id, title:'1차 스프린트 피드백', sheet:FEEDBACK_SHEET, criteria:CRITERIA_V1, sprint2:false};
  throw new Error('이 수업의 제출 자료가 아닙니다.');
}

function doGet() {
  return page_('조별 피드백', '<h1>조별 피드백</h1><p>피드백 작성 화면에서 조별 코드를 입력하십시오.</p>');
}

function doPost(e) {
  try {
    const p = e && e.parameter || {};
    if (p.action === 'submit') return submit_(p);
    if (p.action === 'inbox') return inbox_(p);
    return page_('요청 오류', '<h1>요청 오류</h1><p>올바르지 않은 요청입니다.</p>');
  } catch (err) {
    return page_('처리 오류', '<h1>처리 오류</h1><p>'+esc_(err && err.message ? err.message : String(err))+'</p>');
  }
}

function submit_(p) {
  const code = String(p.access_code || '').trim();
  if (!code || !p.payload) throw new Error('조별 코드와 제출 내용이 필요합니다.');
  const team = teamForCode_(code);
  if (!team) throw new Error('조별 코드가 올바르지 않습니다.');
  const payloadText = String(p.payload);
  if (payloadText.length > 32000) throw new Error('제출 내용이 너무 깁니다. 입력란의 글자 수를 확인하십시오.');
  let raw;
  try { raw = JSON.parse(payloadText); } catch (_) { throw new Error('제출 형식을 확인하십시오.'); }
  if (!raw || typeof raw !== 'object' || !raw.feedback) throw new Error('제출 형식을 확인하십시오.');
  const course = course_(String(raw.courseId || ''));
  if (p.course_id && String(p.course_id) !== course.id) throw new Error('피드백 차수가 일치하지 않습니다.');
  if (course.sprint2 && raw.schema !== 'daejin.sprint2.feedback.v1') throw new Error('2차 피드백 작성 화면을 확인하십시오.');
  const d = raw.feedback;
  if (String(d.fromTeam) !== team) throw new Error('선택한 보내는 조와 조별 코드가 일치하지 않습니다.');
  if (!GROUPS.includes(String(d.toTeam)) || String(d.toTeam) === team) throw new Error('받는 조를 확인하십시오. 자기 조에는 제출할 수 없습니다.');
  const id = limited_(raw.recordId,90,'접수번호');
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw new Error('접수번호 형식을 확인하십시오.');

  const feedback = {role:'peer',unit:'team',fromTeam:team,toTeam:String(d.toTeam),criteria:{}};
  feedback.projectTitle = limited_(d.projectTitle,180,'연구주제');
  for (const field of ['strength','clarification','suggestion','rewrite']) feedback[field] = limited_(d[field],1500,field);
  feedback.priority = limited_(d.priority,80,'우선 확인 영역');
  const rawCriteria = d.criteria && typeof d.criteria === 'object' ? d.criteria : {};
  course.criteria.forEach(([key,title]) => {
    const item = rawCriteria[key] && typeof rawCriteria[key] === 'object' ? rawCriteria[key] : {};
    feedback.criteria[key] = {comment:limited_(item.comment,1500,title)};
    // 기존 1차 행 구조는 유지하되 수신 화면에서는 이 필드를 출력하지 않습니다.
    if (!course.sprint2) feedback.criteria[key].status = clean_(item.status,30);
  });
  if (course.sprint2) {
    if (!feedback.projectTitle || !feedback.strength || !feedback.clarification || !feedback.suggestion) throw new Error('필수 피드백 내용을 모두 작성하십시오.');
    if (!course.criteria.some(([key])=>key===feedback.priority)) throw new Error('우선 확인·보완할 영역을 선택하십시오.');
    course.criteria.forEach(([key,title]) => {if (!feedback.criteria[key].comment) throw new Error(title+'의 근거 또는 질문을 작성하십시오.');});
  }
  const rec = {
    schema: course.sprint2 ? 'daejin.sprint2.feedback.v1' : 'daejin.sprint1.feedback.v1',
    courseId:course.id, recordId:id, preparedAt:clean_(raw.preparedAt,50), feedback:feedback
  };
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(course.sheet);
  if (!sh) throw new Error('교수자가 먼저 setup()을 실행해야 합니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const values = sh.getDataRange().getValues();
    const h = values[0].map(String), ix = Object.fromEntries(h.map((name,i)=>[name,i]));
    if (ix.recordId === undefined || ix.payload === undefined || ix.visible === undefined || ix.toTeam === undefined) throw new Error('피드백 시트의 열 이름을 확인하십시오.');
    for (let i=1;i<values.length;i++) {
      if (String(values[i][ix.recordId]) !== id) continue;
      let previous;
      try { previous = JSON.parse(String(values[i][ix.payload])); } catch (_) { throw new Error('동일 접수번호의 기존 자료를 확인하십시오.'); }
      if (previous.courseId !== rec.courseId || JSON.stringify(previous.feedback) !== JSON.stringify(rec.feedback)) throw new Error('이미 사용한 접수번호입니다. 작성 화면을 다시 열고 제출하십시오.');
      return receipt_(id,true);
    }
    const now = new Date();
    rec.receivedAt = now.toISOString();
    const map = {
      submittedAt:now,recordId:id,fromTeam:team,toTeam:feedback.toTeam,projectTitle:feedback.projectTitle,
      strength:feedback.strength,clarification:feedback.clarification,priority:feedback.priority,
      suggestion:feedback.suggestion,rewrite:feedback.rewrite,visible:VISIBLE_ON_SUBMIT,payload:JSON.stringify(rec)
    };
    course.criteria.forEach(([key]) => {
      map[key+'Comment'] = feedback.criteria[key].comment;
      if (!course.sprint2) map[key+'Status'] = feedback.criteria[key].status;
    });
    sh.appendRow(h.map(name => sheetValue_(Object.prototype.hasOwnProperty.call(map,name) ? map[name] : '')));
    SpreadsheetApp.flush();
    return receipt_(id,false);
  } finally { lock.releaseLock(); }
}

function inbox_(p) {
  const team = teamForCode_(String(p.access_code || '').trim());
  if (!team) throw new Error('조회 코드가 올바르지 않습니다.');
  // 기존 1차 HTML은 course_id를 보내지 않았으므로 기본값을 그대로 둡니다.
  const course = course_(String(p.course_id || COURSE_ID));
  const sh = SpreadsheetApp.getActive().getSheetByName(course.sheet);
  if (!sh || sh.getLastRow() < 2) return inboxPage_(team,[],course);
  const values = sh.getDataRange().getValues(), h = values[0].map(String);
  const ix = Object.fromEntries(h.map((name,i)=>[name,i]));
  const records = [];
  for (let i=1;i<values.length;i++) {
    const row=values[i], visible=row[ix.visible]===true || String(row[ix.visible]).toUpperCase()==='TRUE';
    if (String(row[ix.toTeam])!==team || !visible) continue;
    try {
      const rec=JSON.parse(String(row[ix.payload] || '{}'));
      if (rec && rec.courseId===course.id && rec.feedback && String(rec.feedback.toTeam)===team) records.push(rec);
    } catch (_) {}
  }
  records.reverse();
  return inboxPage_(team,records,course);
}

function inboxPage_(team,records,course) {
  let body='<h1>'+esc_(team)+'조 받은 피드백</h1><p class="muted">'+esc_(course.title)+'</p>';
  if (!records.length) body+='<div class="empty">현재 확인할 수 있는 피드백이 없습니다.</div>';
  records.forEach(rec=>{
    const d=rec.feedback || {}, c=d.criteria || {};
    body+='<article><h2>'+esc_(d.projectTitle || '연구주제 미입력')+'</h2><h3>항목별 검토</h3>';
    course.criteria.forEach(([id,title])=>{
      body+='<h4>'+esc_(title)+'</h4><p>'+nl_((c[id] || {}).comment)+'</p>';
    });
    body+='<h3>잘한 점</h3><p>'+nl_(d.strength)+'</p><h3>확인할 점</h3><p>'+nl_(d.clarification)+'</p>';
    const selected=course.criteria.find(([id])=>id===d.priority);
    if (selected) body+='<h3>우선 확인·보완할 영역</h3><p>'+esc_(selected[1])+'</p>';
    body+='<h3>수정 제안</h3><p>'+nl_(d.suggestion)+'</p>';
    if (d.rewrite) body+='<h3>수정 문장 예시</h3><p>'+nl_(d.rewrite)+'</p>';
    body+='</article>';
  });
  // 보내는 조, 코드, 접수번호, 원본 JSON 및 평가 상태는 수신 화면에 포함하지 않습니다.
  return page_(team+'조 받은 피드백',body);
}

function receipt_(id,alreadyReceived) {
  return page_('접수 완료','<h1>접수 완료</h1><p>'+(alreadyReceived ? '이미 접수된 피드백을 확인했습니다. 중복 저장하지 않았습니다.' : '피드백이 접수되었습니다.')+'</p><p>접수번호: <strong>'+esc_(id)+'</strong></p><p class="muted">이 창을 닫고 작성 화면으로 돌아가십시오.</p>');
}

function teamForCode_(code) {
  const target=String(code || '').trim();
  if (!target) return '';
  const sh=SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow()<2) throw new Error('교수자가 먼저 setup()을 실행해야 합니다.');
  const values=sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues();
  const matches=values.filter(([team,saved])=>GROUPS.includes(String(team).trim()) && String(saved).trim()===target);
  return matches.length===1 ? String(matches[0][0]).trim() : '';
}

function makeCode_() { return Utilities.getUuid().replace(/-/g,'').slice(0,10).toUpperCase(); }
function clean_(v,max) { return String(v==null ? '' : v).trim().slice(0,max); }
function limited_(v,max,label) {
  const s=String(v==null ? '' : v).trim();
  if (s.length>max) throw new Error(label+'의 글자 수를 확인하십시오.');
  return s;
}
function sheetValue_(v) {
  // 학생 입력이 스프레드시트 수식으로 실행되지 않도록 텍스트로 저장합니다.
  return typeof v==='string' && /^[=+@-]/.test(v) ? "'"+v : v;
}
function esc_(v) { return String(v==null ? '' : v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function nl_(v) { return esc_(v || '미작성').replace(/\n/g,'<br>'); }
function page_(title,body) {
  const html='<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="referrer" content="no-referrer"><title>'+esc_(title)+'</title><style>'+
    'body{margin:0;background:#f4f7f9;color:#153440;font:15px/1.75 "Malgun Gothic","Noto Sans KR",system-ui,sans-serif;word-break:keep-all;overflow-wrap:anywhere}main{max-width:860px;margin:36px auto;padding:28px;background:#fff;border:1px solid #d8e3e8;border-radius:12px}h1{font-size:28px;margin:0 0 12px}h2{font-size:20px;margin:0 0 18px}h3{font-size:16px;margin:24px 0 7px}h4{font-size:14px;margin:17px 0 4px}p{margin:5px 0}.muted{font-size:12px;color:#516873}.empty{padding:18px;margin-top:22px;background:#eef8f9;border-left:3px solid #087c8c}article{border-top:1px solid #d8e3e8;padding:28px 0}article:first-of-type{border-top:0;padding-top:24px}@media(max-width:600px){main{margin:0;border:0;border-radius:0;padding:22px 16px}h1{font-size:24px}}'+
    '</style></head><body><main>'+body+'</main></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
