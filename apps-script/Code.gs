/** LK 테크 Google Sheets 저장 API. 이 파일에는 관리자 비밀키가 없습니다. */
const SHEET_ID = '14nus_yGwAIZu5J29FULfIFVBRYFBXkXpgAvXuQAGcqI';
const WEB_ORIGIN = 'https://johnpark236-tech.github.io';
function book_(){ return SpreadsheetApp.openById(SHEET_ID); }
function tab_(name){const s=book_().getSheetByName(name);if(!s)throw Error('필요한 시트 없음: '+name);return s;}
function rows_(name){const s=tab_(name),n=s.getLastRow();if(n>20000)throw Error('관리자에게 시트 정리를 요청해 주세요.');return n<2?[]:s.getRange(2,1,n-1,s.getLastColumn()).getValues().map((v,i)=>({v,row:i+2}));}
/** 수식 주입 방지 + 0으로 시작하는 숫자(전화번호)가 숫자로 바뀌어 앞자리 0이 사라지지 않도록 텍스트로 저장 */
function cellText_(x){return typeof x==='string'&&(/^[=+@-]/.test(x)||/^0[0-9]+$/.test(x))?"'"+x:x;}
/** 이미 숫자로 저장된 전화번호(예: 1071423011)는 앞에 0을 복원 */
function phone_(v){if(v===''||v==null)return '';const t=String(v);return typeof v==='number'&&/^[1-9][0-9]{7,9}$/.test(t)?'0'+t:t;}
function add_(name,v){const s=tab_(name);s.getRange(s.getLastRow()+1,1,1,v.length).setValues([v.map(cellText_)]);}
function update_(name,row,v){tab_(name).getRange(row,1,1,v.length).setValues([v.map(cellText_)]);}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function stamp_(v){return v instanceof Date?v.toISOString():String(v||'');}
function day_(d){return Utilities.formatDate(d,'Asia/Seoul','yyyy-MM-dd');}
function storedDay_(v){return v instanceof Date?day_(v):String(v||'').slice(0,10);}
function time_(){return new Date().toISOString();}
function uuid_(){return Utilities.getUuid();}
function status_(s){return ({'대기':'pending','승인대기':'pending','인정':'approved','반려':'rejected','해제':'disabled','활동중':'active'})[s]||s;}
function fail_(message){throw Error(message);}
function ownerEditor_(){const a=Session.getActiveUser().getEmail(),b=Session.getEffectiveUser().getEmail();if(!a||a!==b)fail_('소유자만 실행할 수 있습니다.');}
function install(){ownerEditor_();ensureSheets_();book_().setSpreadsheetTimeZone('Asia/Seoul');tab_('설정').getRange('B2').setValue('웹 앱 배포 후 연결 URL 등록 필요');onOpen();return '설치 완료. 운영자 계정: '+owners_().join(', ')+(clientId_()?'':' / GOOGLE_CLIENT_ID 설정 필요');}
function onOpen(){SpreadsheetApp.getUi().createMenu('LK 테크').addItem('로그인 설정 확인','showLoginSetup').addToUi();}
function showLoginSetup(){ownerEditor_();SpreadsheetApp.getUi().alert('LK 테크 로그인 설정','운영자 구글 계정: '+owners_().join(', ')+'\nGoogle Client ID: '+(clientId_()||'미설정'),SpreadsheetApp.getUi().ButtonSet.OK);}
/** 홈페이지 Google 로그인(OAuth 클라이언트 ID). 공개 값이며 비밀키가 아님. 스크립트 속성 GOOGLE_CLIENT_ID로도 설정 가능. */
const GOOGLE_CLIENT_ID = '524254414923-7725ihp35uanp4pfa0i1e6pfhrbos59l.apps.googleusercontent.com';
function clientId_(){return PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID')||GOOGLE_CLIENT_ID;}
/** 운영자 = 웹 앱을 배포한 구글 계정 + 스크립트 속성 OWNER_EMAILS(쉼표 구분)에 적은 계정 */
function owners_(){const list=String(PropertiesService.getScriptProperties().getProperty('OWNER_EMAILS')||'').split(',');list.push(Session.getEffectiveUser().getEmail()||'');return [...new Set(list.map(x=>x.trim().toLowerCase()).filter(Boolean))];}
/** Google ID 토큰을 구글 서버에서 검증하고 이메일을 돌려줌. 검증 결과는 만료 전까지 캐시. */
function googleEmail_(token){
  if(typeof token!=='string'||token.length>4096||!/^[\w-]+\.[\w-]+\.[\w-]+$/.test(token))fail_('로그인 정보가 올바르지 않습니다. 다시 로그인하세요.');
  const cid=clientId_();if(!cid)fail_('구글 로그인 설정(Client ID)이 아직 없습니다.');
  const cache=CacheService.getScriptCache(),key='idt:'+hash_(token).slice(0,40),hit=cache.get(key);if(hit)return hit;
  const res=UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token='+encodeURIComponent(token),{muteHttpExceptions:true});
  if(res.getResponseCode()!==200)fail_('구글 로그인이 만료되었습니다. 다시 로그인하세요.');
  const t=JSON.parse(res.getContentText()),exp=Number(t.exp)*1000;
  if(t.aud!==cid||!['accounts.google.com','https://accounts.google.com'].includes(t.iss)||String(t.email_verified)!=='true'||!(exp>Date.now()))fail_('구글 로그인이 확인되지 않습니다. 다시 로그인하세요.');
  const email=String(t.email).toLowerCase();cache.put(key,email,Math.max(1,Math.min(3600,Math.floor((exp-Date.now())/1000)-30)));return email;
}
function actor_(p){
  if(typeof p.token!=='string'||!/^([a-f0-9]{64})$/.test(p.token))fail_('기기 저장 기능을 허용하고 새로고침하세요.');
  const email=p.idToken?googleEmail_(p.idToken):'',owner=!!email&&owners_().includes(email);
  const id=owner?'owner':email?'google:'+email:'guest:'+hash_(p.token),entry=email&&!owner?rows_('관리자신청').find(r=>r.v[0]===id):null;
  return {id,email,owner,admin:owner||status_(entry?.v[4])==='active',name:owner?'운영자':entry?.v[2]||'',staffStatus:entry?status_(entry.v[4]):null};
}
function profile_(b){const name=typeof b.name==='string'?b.name.trim():'',phone=typeof b.phone==='string'?b.phone.replace(/[\s-]/g,''):'';if(!name||name.length>40)fail_('이름을 1~40자로 입력하세요.');if(!/^0[0-9]{8,10}$/.test(phone))fail_('올바른 국내 전화번호를 입력하세요.');if(b.consent!==true)fail_('정보 저장에 동의해 주세요.');return {name,phone};}
function limited_(id){const cache=CacheService.getScriptCache(),key='limit:'+id+':'+day_(new Date()),n=Number(cache.get(key)||0);if(n>=20)fail_('신청 횟수가 많습니다. 잠시 후 다시 시도하세요.');cache.put(key,String(n+1),21600);}
function notice_(id,type,name,phone,now){add_('알림대기',[id,now,type,name,phone,'대기','']);}
function chat_(id){const msgs=rows_('상담기록').filter(r=>r.v[0]===id).map(r=>({time:stamp_(r.v[1]),role:r.v[2],text:String(r.v[3]),responder:r.v[4]||''}));const app=rows_('작업자신청').filter(r=>r.v[1]===id).at(-1);return {id,messages:msgs,application:app?{name:app.v[3],phone:phone_(app.v[4]),submitted:stamp_(app.v[2])}:null,pending:msgs.length>0&&msgs.at(-1).role!=='admin',updated:msgs.at(-1)?.time||stamp_(app?.v[2])||'',created:msgs[0]?.time||stamp_(app?.v[2])||''};}
function appendMessage_(id,role,text,name){add_('상담기록',[id,time_(),role,text,name||'']);}
function distance_(a,b){const r=x=>x*Math.PI/180,h=Math.sin(r(b.lat-a.lat)/2)**2+Math.cos(r(a.lat))*Math.cos(r(b.lat))*Math.sin(r(b.lng-a.lng)/2)**2;return 6371000*2*Math.atan2(Math.sqrt(h),Math.sqrt(1-h));}
function site_(v){return {id:v[0],name:v[1],address:v[2],lat:Number(v[3]),lng:Number(v[4]),radius:Number(v[5])};}
function event_(v){const e={eventId:v[0],day:storedDay_(v[3]),kind:v[4],name:v[2],time:stamp_(v[5]),site:{id:v[6],name:v[7],address:v[8],lat:Number(v[17]),lng:Number(v[18]),radius:Number(v[19])},location:{lat:Number(v[9]),lng:Number(v[10]),accuracy:Number(v[11]),distance:Number(v[12])},status:status_(v[13])};if(v[14])e.review={time:stamp_(v[14]),name:v[15],reason:v[16]||''};return e;}
function dispatch_(p){
 const a=actor_(p),path=String(p.path||'').split('?')[0],method=p.method||'GET',b=p.body||{},now=time_(),parts=String(p.path||'').split('?')[1]||'';const monthParam=parts.match(/(?:^|&)month=([^&]+)/),month=monthParam?decodeURIComponent(monthParam[1]):day_(new Date()).slice(0,7);
 if(path==='me')return {admin:a.admin,owner:a.owner,guest:!a.email,email:a.email,staffStatus:a.staffStatus,setup:false};
 if(path==='apply'&&method==='POST'){const v=profile_(b);limited_(a.id);const apps=rows_('작업자신청').filter(r=>r.v[1]===a.id);if(apps.length&&Date.now()-Date.parse(stamp_(apps.at(-1).v[2]))<30000)fail_('이미 접수했습니다. 잠시 후 다시 시도하세요.');const id=uuid_();add_('작업자신청',[id,a.id,now,v.name,v.phone,'접수','']);notice_(id,'작업자',v.name,v.phone,now);appendMessage_(a.id,'user','기술자 지원\n이름: '+v.name+'\n전화번호: '+v.phone);appendMessage_(a.id,'bot','관리자가 검토 후 전화로 연락하여 근무 여부를 안내해 드립니다.');return {ok:true,message:'관리자가 검토 후 바로 연락하여 근무 여부를 안내해 드립니다.'};}
 if(path==='staff/request'&&method==='POST'){if(!a.email)fail_('관리자 신청은 구글 로그인 후 할 수 있습니다.');if(a.owner||a.staffStatus)fail_('이미 신청되었거나 등록된 관리자입니다.');const v=profile_(b);limited_(a.id);add_('관리자신청',[a.id,now,v.name,v.phone,'pending','','']);notice_(a.id,'관리자',v.name,v.phone,now);return {ok:true};}
 if(path==='staff/add'&&method==='POST'){if(!a.owner)fail_('관리자 추가는 운영자만 할 수 있습니다.');const email=String(b.email||'').trim().toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||email.length>120)fail_('구글 계정 이메일을 확인하세요.');if(owners_().includes(email))fail_('운영자 계정은 따로 추가할 필요가 없습니다.');const name=String(b.name||'').trim().slice(0,40)||email.split('@')[0],phone=String(b.phone||'').replace(/[\s-]/g,'');if(phone&&!/^0[0-9]{8,10}$/.test(phone))fail_('올바른 국내 전화번호를 입력하세요.');const id='google:'+email,row=rows_('관리자신청').find(r=>r.v[0]===id);if(row){row.v[4]='active';row.v[5]=now;row.v[6]=a.name;if(b.name)row.v[2]=name;if(phone)row.v[3]=phone;update_('관리자신청',row.row,row.v);}else add_('관리자신청',[id,now,name,phone,'active',now,a.name]);return {ok:true,id};}
 if(path==='staff'){if(!a.owner)fail_('관리자 승인 권한이 없습니다.');if(method==='GET')return rows_('관리자신청').map(r=>({id:r.v[0],name:r.v[2],phone:phone_(r.v[3]),status:status_(r.v[4]),requested:stamp_(r.v[1])}));if(method==='POST'){if(!['active','disabled'].includes(b.status)||b.id==='owner')fail_('승인·해제만 가능합니다.');const row=rows_('관리자신청').find(r=>r.v[0]===b.id);if(!row)fail_('신청을 찾을 수 없습니다.');row.v[4]=b.status;row.v[5]=now;row.v[6]=a.name;update_('관리자신청',row.row,row.v);return {ok:true};}}
 if(path==='chat'&&method==='GET')return chat_(a.id);
 if(path==='chat'&&method==='DELETE'){const apps=rows_('작업자신청').filter(r=>r.v[1]===a.id),ids=apps.map(r=>r.v[0]);for(const [sheet,filter] of [['상담기록',r=>r.v[0]===a.id],['작업자신청',r=>r.v[1]===a.id],['알림대기',r=>ids.includes(r.v[0])]])for(const row of rows_(sheet).filter(filter).reverse())tab_(sheet).deleteRow(row.row);return {ok:true};}
 if((path==='chat'||path==='reply')&&method==='POST'){if(path==='reply'&&!a.admin)fail_('관리자만 답변할 수 있습니다.');if(path==='chat'&&b.consent!==true)fail_('정보 저장에 동의하세요.');const text=typeof b.text==='string'?b.text.trim():'';if(!text||text.length>1500)fail_('1~1500자로 입력하세요.');const id=path==='reply'?String(b.id||''):a.id;if(!/^(guest:[a-f0-9]{64}|owner|google:[^\s@]+@[^\s@]+)$/.test(id))fail_('상담 대상을 확인하세요.');const chat=chat_(id);if(chat.messages.length>=300)fail_('상담 기록 한도입니다.');if(path==='chat'&&chat.updated&&Date.now()-Date.parse(chat.updated)<1500)fail_('잠시 후 다시 보내세요.');appendMessage_(id,path==='reply'?'admin':'user',text,a.name);if(path==='chat')appendMessage_(id,'bot',answer_(text)||'담당자 확인이 필요한 문의로 접수했습니다. 확인 후 안내드립니다.');return chat_(id);}
 if(path==='admin'&&method==='GET'){if(!a.admin)fail_('승인된 관리자만 볼 수 있습니다.');const ids=[...new Set(rows_('상담기록').map(r=>r.v[0]).concat(rows_('작업자신청').map(r=>r.v[1])))];return ids.map(chat_).sort((x,y)=>y.updated.localeCompare(x.updated));}
 if(path==='workplaces'){if(method==='GET')return rows_('근무지').map(r=>site_(r.v));if(method==='POST'){if(!a.owner)fail_('근무지 수정 권한이 없습니다.');if(!b.name||!b.address||String(b.name).length>60||String(b.address).length>200||!Number.isFinite(b.lat)||!Number.isFinite(b.lng)||Math.abs(b.lat)>90||Math.abs(b.lng)>180||!Number.isFinite(b.radius)||b.radius<30||b.radius>1000)fail_('근무지명·실제 주소·좌표·반경을 확인하세요.');const existing=b.id?rows_('근무지').find(r=>r.v[0]===b.id):null;if(b.id&&!existing)fail_('근무지가 없습니다.');const site={id:b.id||uuid_(),name:b.name,address:b.address,lat:b.lat,lng:b.lng,radius:Math.round(b.radius)},v=[site.id,site.name,site.address,site.lat,site.lng,site.radius];if(existing)update_('근무지',existing.row,v);else add_('근무지',v);return site;}}
 if(path==='attendance' || path==='attendance/admin'){if(method!=='GET'||!/^20[0-9]{2}-(0[1-9]|1[0-2])$/.test(month))fail_('조회 월을 확인하세요.');if(path==='attendance/admin'&&!a.admin)fail_('관리자만 볼 수 있습니다.');const filtered=rows_('근태기록').filter(r=>storedDay_(r.v[3]).slice(0,7)===month&&(path==='attendance/admin'||r.v[1]===a.id));if(path==='attendance')return {id:a.id,month,events:filtered.map(r=>event_(r.v))};const ids=[...new Set(filtered.map(r=>r.v[1]))];return ids.map(id=>({id,month,events:filtered.filter(r=>r.v[1]===id).map(r=>event_(r.v))}));}
 if(path==='attendance/record'&&method==='POST'){if(b.consent!==true)fail_('위치·시간 정보 저장에 동의하세요.');if(!['in','out'].includes(b.kind))fail_('출퇴근 구분을 확인하세요.');if(!Number.isFinite(b.lat)||!Number.isFinite(b.lng)||Math.abs(b.lat)>90||Math.abs(b.lng)>180||!Number.isFinite(b.accuracy)||b.accuracy<=0||b.accuracy>100)fail_('위치 오차가 100m 이하일 때 현장에서 다시 시도하세요.');const sr=rows_('근무지').find(r=>r.v[0]===b.siteId);if(!sr)fail_('등록된 근무지를 선택하세요.');const s=site_(sr.v),distance=distance_(s,{lat:b.lat,lng:b.lng});if(distance>s.radius)fail_('근무지 허용 반경 밖입니다.');const day=day_(new Date()),today=rows_('근태기록').filter(r=>r.v[1]===a.id&&storedDay_(r.v[3])===day);if(today.some(r=>r.v[4]===b.kind&&status_(r.v[13])!=='rejected'))fail_('오늘 기록이 이미 있습니다.');const arrival=today.find(r=>r.v[4]==='in'&&status_(r.v[13])!=='rejected');if(b.kind==='out'&&(!arrival||arrival.v[6]!==s.id))fail_('오늘 출근한 근무지에서 퇴근을 기록하세요.');if(today.length>=6)fail_('일일 기록 한도입니다.');const app=rows_('작업자신청').filter(r=>r.v[1]===a.id).at(-1),name=app?.v[3]||a.name||String(b.name||'').trim();if(!name||name.length>40)fail_('이름을 입력하세요.');const v=[uuid_(),a.id,name,day,b.kind,now,s.id,s.name,s.address,b.lat,b.lng,b.accuracy,Math.round(distance),'pending','','','',s.lat,s.lng,s.radius];add_('근태기록',v);return event_(v);}
 if(path==='attendance/review'&&method==='POST'){if(!a.admin)fail_('관리자만 확인할 수 있습니다.');if(b.id===a.id)fail_('본인의 기록은 다른 관리자가 확인해야 합니다.');if(!['approved','rejected'].includes(b.status))fail_('승인·반려를 선택하세요.');const r=rows_('근태기록').find(r=>r.v[0]===b.eventId&&r.v[1]===b.id&&storedDay_(r.v[3]).slice(0,7)===b.month);if(!r||status_(r.v[13])!=='pending')fail_('기록이 없거나 이미 확인했습니다.');if(b.status==='rejected'&&!String(b.reason||'').trim())fail_('반려 사유를 입력하세요.');r.v[13]=b.status;r.v[14]=now;r.v[15]=a.name;r.v[16]=String(b.reason||'').slice(0,200);update_('근태기록',r.row,r.v);return event_(r.v);}
 if(path==='workers')return workers_(a,method,b);
 if(path==='payroll'){if(method!=='GET'||!a.admin)fail_('승인된 관리자만 볼 수 있습니다.');if(!/^20[0-9]{2}-(0[1-9]|1[0-2])$/.test(month))fail_('조회 월을 확인하세요.');return payroll_(month);}
 fail_('지원하지 않는 요청입니다.');
}
function doGet(){return ContentService.createTextOutput(JSON.stringify({ok:true,storage:'Google Sheets',version:1})).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){let p,output;const lock=LockService.getScriptLock();try{const raw=e.parameter?.payload||e.postData?.contents||'';if(raw.length>12000)fail_('요청이 너무 깁니다.');p=JSON.parse(raw);if(p.origin!==WEB_ORIGIN)fail_('요청 출처 오류');if(!lock.tryLock(10000))fail_('다른 요청 처리 중입니다. 다시 시도하세요.');output={ok:true,data:dispatch_(p)};}catch(error){output={ok:false,error:String(error.message||error)};}finally{if(lock.hasLock())lock.releaseLock();}if(p?.bridgeNonce&&/^[a-f0-9]{32}$/.test(p.bridgeNonce)){const safe=JSON.stringify({bridgeNonce:p.bridgeNonce,result:output}).replace(/</g,'\\u003c');return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>window.top.postMessage('+safe+','+JSON.stringify(WEB_ORIGIN)+');</script>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);}

function answer_(t){if(/회사|업체|모집주체|LK|엘케이/.test(t))return 'LK 테크 모집 안내입니다. 업무는 반도체 자동화설비 조립·제작이며, 근무 형태는 일용직(근로계약 체결, 고용·산재보험 가입)입니다. 기술자 지원에서 이름과 연락처를 남겨주시면 담당자가 확인 후 연락드립니다.';if(/나이|연령|성별|남성|여성|지원 자격/.test(t))return '나이·성별 제한 없이 설비 조립 작업이 가능한 분을 모집합니다. 초보도 지원할 수 있으며 성실한 출근과 근태를 중요하게 봅니다. 채용은 담당자 확인 후 확정됩니다.';if(/급여|일급|얼마|교통비|식대/.test(t)&&! /지급|언제|잔업|세금|실수령/.test(t))return '일급은 14만원부터 시작하며(세전, 식대·교통비 포함), 자재 입고부터 조립·설치·셋업까지 전 공정을 순환하며 배웁니다. 일급은 그날 맡은 공정이 아니라 본인의 숙련도·경력을 기준으로 정하며, 숙련도가 오르면 일급도 오릅니다. 유경험자는 면접 시 협의하며, 성별과 관계없이 같은 기준을 적용합니다. 일용근로소득은 하루 15만원까지 비과세라 세금 부담이 거의 없습니다.';if(/프리랜서|일용|근무 형태|고용 형태|보험|연락처|전화번호|카카오/.test(t))return '근무 형태는 일용직입니다. 근로계약서를 작성하고 고용보험·산재보험에 가입합니다. 기술자 지원에서 이름과 연락처를 남겨주시면 담당자가 확인 후 연락드립니다.';if(/숙소|기숙사/.test(t))return '숙소는 제공하지 않습니다.';if(/준비|공구|렌치|가위/.test(t))return '기본공구, 렌치 세트, 작업용 가위를 준비해 주세요.';if(/근무시간|근무일|요일|평일|주말|토요일|일요일|몇 시|몇시/.test(t))return '기본 평일(월~금), 근무시간은 08:30~17:30입니다. 휴게시간은 1시간이며, 잔업·주말 근무 여부는 담당자 확인이 필요합니다.';if(/근무지|지역|어디/.test(t))return '모집 지역은 둔포·직산·평택입니다. 실제 배치 주소는 담당자가 확인해 안내합니다.';if(/하는 일|업무|조립/.test(t))return '반도체 자동화설비 조립·제작 업무입니다. LK 테크가 도급받은 현장에서 LK 테크 현장책임자의 지도 아래 근무하며, 삼성 직접 채용 공고가 아닙니다.';if(/초보|경력/.test(t))return '초보 지원이 가능합니다. 성실한 출근과 근태를 중요하게 봅니다. 채용과 출근 확정은 담당자가 안내합니다.';return null;}

/* ===== 인력 관리: 작업자 명부 · 숙련 단계별 일급 · 월간 정산 ===== */
const SCHEMA = {
  '작업자신청':['신청ID','방문자ID','접수시각','이름','전화번호','상태','관리메모'],
  '관리자신청':['방문자ID','신청시각','이름','전화번호','권한상태','승인시각','승인자'],
  '근태기록':['기록ID','방문자ID','이름','근무일','구분','기록시각','근무지ID','근무지명','주소','위도','경도','위치오차(m)','거리(m)','승인상태','확인시각','확인자','반려사유','근무지위도','근무지경도','허용반경(m)'],
  '근무지':['근무지ID','근무지명','주소','위도','경도','허용반경(m)'],
  '상담기록':['방문자ID','시각','발신구분','내용','담당자'],
  '알림대기':['알림ID','접수시각','구분','이름','전화번호','발송상태','발송시각'],
  '작업자':['작업자ID','방문자ID','이름','전화번호','등급','일급','상태','등록일','메모'],
  '설정':['항목','값']
};
/** 없는 시트만 만들고 기존 시트는 건드리지 않습니다. */
function ensureSheets_(){const b=book_();for(const [name,head] of Object.entries(SCHEMA)){if(!b.getSheetByName(name)){const s=b.insertSheet(name);s.getRange(1,1,1,head.length).setValues([head]);s.setFrozenRows(1);}}}

/** 숙련 단계별 기본 일급(세전, 식대·교통비 포함). 그날 맡은 공정과 무관하게 본인 단계 기준. 작업자별 수정 가능. */
const GRADES = {
  '입문':{label:'입문 · 처음 시작, 선임 보조',wage:140000},
  '숙련':{label:'숙련 · 공구 사용·가조립 단독 가능',wage:150000},
  '전문':{label:'전문 · 핵심 조립·셋업 단독 가능',wage:170000}
};
const RULES = {dailyExempt:150000,rate:0.06,taxCredit:0.55,localRate:0.1,employmentRate:0.009,smallTaxLimit:1000,enrollDays:8,enrollHours:60,warnDays:6,breakHours:1};

function worker_(v){return {id:String(v[0]),visitorId:String(v[1]||''),name:String(v[2]),phone:phone_(v[3]),grade:String(v[4]),wage:Number(v[5]),status:status_(v[6]),registered:stamp_(v[7]),memo:String(v[8]||'')};}

function workers_(a,method,b){
  if(!a.admin)fail_('승인된 관리자만 볼 수 있습니다.');
  if(method==='GET'){
    const list=rows_('작업자').map(r=>worker_(r.v)),known=new Set(list.map(w=>w.visitorId).filter(Boolean)),seen=new Set();
    // 관리자(승인 대기·활동 중)와 운영자는 작업자 채용 목록에서 제외: 계정 또는 전화번호가 같으면 관리자로 봄
    const digits=v=>phone_(v).replace(/[^0-9]/g,''),staff=rows_('관리자신청').filter(r=>status_(r.v[4])!=='disabled');
    const staffIds=new Set(staff.map(r=>String(r.v[0]))),staffPhones=new Set(staff.map(r=>digits(r.v[3])).filter(Boolean));
    let hiddenStaff=0;
    const applicants=rows_('작업자신청').slice().reverse().filter(r=>{const id=String(r.v[1]);if(known.has(id)||seen.has(id))return false;seen.add(id);if(id==='owner'||staffIds.has(id)||staffPhones.has(digits(r.v[4]))){hiddenStaff++;return false;}return true;}).map(r=>({visitorId:String(r.v[1]),name:String(r.v[3]),phone:phone_(r.v[4]),submitted:stamp_(r.v[2])}));
    return {grades:GRADES,workers:list,applicants,hiddenStaff};
  }
  if(method!=='POST')fail_('지원하지 않는 요청입니다.');
  if(!a.owner)fail_('작업자 등록·수정은 운영자만 할 수 있습니다.');
  const name=typeof b.name==='string'?b.name.trim():'',phone=typeof b.phone==='string'?b.phone.replace(/[\s-]/g,''):'';
  if(!name||name.length>40)fail_('이름을 1~40자로 입력하세요.');
  if(phone&&!/^0[0-9]{8,10}$/.test(phone))fail_('올바른 국내 전화번호를 입력하세요.');
  if(!GRADES[b.grade])fail_('등급을 선택하세요.');
  const wage=Number(b.wage);if(!Number.isInteger(wage)||wage<10000||wage>1000000)fail_('일급을 1만~100만원 사이 정수로 입력하세요.');
  const status=b.status==='ended'?'ended':'active',memo=String(b.memo||'').slice(0,200);
  const visitorId=String(b.visitorId||'');if(visitorId&&!/^guest:[a-f0-9]{64}$/.test(visitorId))fail_('방문자 식별자를 확인하세요.');
  const all=rows_('작업자');
  if(visitorId&&all.some(r=>String(r.v[1])===visitorId&&String(r.v[0])!==String(b.id||'')))fail_('이미 등록된 작업자입니다.');
  if(b.id){const row=all.find(r=>String(r.v[0])===String(b.id));if(!row)fail_('작업자를 찾을 수 없습니다.');const v=[row.v[0],visitorId||row.v[1],name,phone,b.grade,wage,status,row.v[7],memo];update_('작업자',row.row,v);return worker_(v);}
  const v=[uuid_(),visitorId,name,phone,b.grade,wage,status,time_(),memo];add_('작업자',v);
  if(visitorId){const app=rows_('작업자신청').filter(r=>String(r.v[1])===visitorId).at(-1);if(app){app.v[5]='채용';update_('작업자신청',app.row,app.v);}}
  return worker_(v);
}

/** 출근·퇴근 시각(ISO)으로 하루 근로시간 계산. 6시간 이상이면 휴게 1시간 차감, 0.5시간 단위. */
function workHours_(inTime,outTime){const span=(Date.parse(outTime)-Date.parse(inTime))/3600000;if(!(span>0))return 0;const h=span>=6?span-RULES.breakHours:span;return Math.max(0,Math.round(h*2)/2);}
/** 일용근로소득 하루 소득세(원 미만 버림): (일급-15만)×6%×(1-55%) */
function dayIncomeTax_(wage){return Math.floor(Math.max(0,wage-RULES.dailyExempt)*RULES.rate*(1-RULES.taxCredit));}
/** 월 1회 지급 기준 정산. 10원 미만 절사, 소득세 1,000원 미만은 소액부징수. */
function settle_(wage,dayCount){
  const gross=wage*dayCount,raw=dayIncomeTax_(wage)*dayCount;
  let incomeTax=Math.floor(raw/10)*10;if(incomeTax<RULES.smallTaxLimit)incomeTax=0;
  const localTax=Math.floor(incomeTax*RULES.localRate/10)*10,employment=Math.floor(gross*RULES.employmentRate/10)*10;
  return {gross,incomeTax,localTax,employment,net:gross-incomeTax-localTax-employment};
}

function payroll_(month){
  const workers=rows_('작업자').map(r=>worker_(r.v)),byVisitor=new Map(workers.filter(w=>w.visitorId).map(w=>[w.visitorId,w]));
  const people=new Map();
  for(const r of rows_('근태기록')){
    if(storedDay_(r.v[3]).slice(0,7)!==month||status_(r.v[13])!=='approved')continue;
    const id=String(r.v[1]),e=event_(r.v);if(!people.has(id))people.set(id,{name:e.name,days:new Map()});
    const p=people.get(id),d=p.days.get(e.day)||{};d[e.kind]=e.time;p.days.set(e.day,d);
  }
  const out=[];
  for(const [visitorId,p] of people){
    const days=[...p.days].filter(([,d])=>d.in&&d.out).map(([day,d])=>({day,hours:workHours_(d.in,d.out)})).sort((x,y)=>x.day.localeCompare(y.day));
    if(!days.length)continue;
    const w=byVisitor.get(visitorId),dayCount=days.length,hours=days.reduce((s,d)=>s+d.hours,0);
    const alert=dayCount>=RULES.enrollDays||hours>=RULES.enrollHours?'enroll':dayCount>=RULES.warnDays?'warn':null;
    const base={visitorId,name:w?w.name:p.name,registered:!!w,workerId:w?w.id:'',phone:w?w.phone:'',grade:w?w.grade:'',wage:w?w.wage:null,days,dayCount,hours,alert};
    out.push(w?{...base,...settle_(w.wage,dayCount)}:{...base,gross:null,incomeTax:null,localTax:null,employment:null,net:null});
  }
  out.sort((x,y)=>x.name.localeCompare(y.name,'ko'));
  const sum=k=>out.reduce((s,p)=>s+(p[k]||0),0);
  return {month,rules:RULES,people:out,totals:{people:out.length,days:sum('dayCount'),gross:sum('gross'),incomeTax:sum('incomeTax'),localTax:sum('localTax'),employment:sum('employment'),net:sum('net')}};
}
