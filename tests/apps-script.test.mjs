import fs from 'node:fs';import vm from 'node:vm';import crypto from 'node:crypto';import assert from 'node:assert/strict';
const schema=JSON.parse(fs.readFileSync(new URL('../apps-script/schema.json',import.meta.url),'utf8'));const data=new Map(Object.entries(schema).map(([n,h])=>[n,[h]]));
const sheets=new Map([...data].map(([name,rows])=>[name,{getLastRow:()=>rows.length,getLastColumn:()=>rows[0].length,getRange:(r,c,h=1,w=1)=>({getValues:()=>rows.slice(r-1,r-1+h).map(row=>row.slice(c-1,c-1+w)),setValues:v=>{v.forEach((row,i)=>{rows[r-1+i]||=[];row.forEach((x,j)=>rows[r-1+i][c-1+j]=typeof x==='string'&&x.startsWith("'")?x.slice(1):x)})}}),deleteRow:r=>rows.splice(r-1,1)}]));
const CID='test-client.apps.googleusercontent.com',props=new Map([['GOOGLE_CLIENT_ID',CID]]);const cache=new Map();let fetches=0;
const OWN='h.owner.s',STF='h.staff.s',BAD='h.wrongaud.s',OLD='h.expired.s',later=Math.floor(Date.now()/1000)+3000;
const tokens={[OWN]:{aud:CID,iss:'https://accounts.google.com',email:'Owner@Example.com',email_verified:'true',exp:String(later)},[STF]:{aud:CID,iss:'accounts.google.com',email:'staff@example.com',email_verified:'true',exp:String(later)},[BAD]:{aud:'other',iss:'accounts.google.com',email:'x@example.com',email_verified:'true',exp:String(later)},[OLD]:{aud:CID,iss:'accounts.google.com',email:'staff@example.com',email_verified:'true',exp:'1000'}};
const ctx={console,Date,Math,JSON,Number,String,Error,Set,encodeURIComponent,Session:{getEffectiveUser:()=>({getEmail:()=>'owner@example.com'})},DriveApp:(()=>{const files=new Map();const folder={createFile:blob=>{const id='f'+files.size;files.set(id,blob);return {getId:()=>id}}};return {files,getFileById:id=>id===undefined?null:(files.has(id)?{getBlob:()=>({getBytes:()=>files.get(id).bytes})}:{getParents:()=>({next:()=>({getFoldersByName:()=>({hasNext:()=>true,next:()=>folder})})})})}})(),UrlFetchApp:{fetch:url=>{fetches++;const t=tokens[decodeURIComponent(url.split('id_token=')[1])];return {getResponseCode:()=>t?200:400,getContentText:()=>JSON.stringify(t||{error:'invalid_token'})}}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v)})},SpreadsheetApp:{openById:()=>({getSheetByName:n=>sheets.get(n)})},CacheService:{getScriptCache:()=>({get:k=>cache.get(k),put:(k,v)=>cache.set(k,v)})},Utilities:{DigestAlgorithm:{SHA_256:1},Charset:{UTF_8:1},getUuid:()=>crypto.randomUUID(),computeDigest:(_,text)=>[...crypto.createHash('sha256').update(Array.isArray(text)?Buffer.from(text.map(x=>(x+256)%256)):text).digest()].map(x=>x>127?x-256:x),base64Decode:b=>[...Buffer.from(b,'base64')].map(x=>x>127?x-256:x),base64Encode:a=>Buffer.from(a.map(x=>(x+256)%256)).toString('base64'),newBlob:(bytes,type,name)=>({bytes,type,name}),formatDate:d=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(d)}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync(new URL('../apps-script/Code.gs',import.meta.url),'utf8'),ctx);
const tokenA='a'.repeat(64),tokenB='b'.repeat(64),call=(path,body={},method='POST',token=tokenA,idToken='')=>ctx.dispatch_({path,method,body,token,idToken});
assert.equal(call('me',{},'GET').admin,false);assert.equal(call('me',{},'GET',tokenA,OWN).email,'owner@example.com');
assert.throws(()=>call('me',{},'GET',tokenA,BAD),/확인되지/);assert.throws(()=>call('me',{},'GET',tokenA,OLD),/확인되지/);assert.throws(()=>call('me',{},'GET',tokenA,'not-a-jwt'),/올바르지/);
const f0=fetches;call('me',{},'GET',tokenA,OWN);assert.equal(fetches,f0,'verified token is cached');
assert.throws(()=>call('staff/request',{name:'담당자',phone:'01098765432',consent:true},'POST',tokenB),/구글 로그인/);assert.equal(call('me',{},'GET',tokenA,OWN).owner,true);
assert.throws(()=>call('apply',{name:'홍길동',consent:true}));assert.throws(()=>call('apply',{name:'홍길동',phone:'01012345678',consent:false}));
assert.equal(call('apply',{name:'홍길동',phone:'010-1234-5678',consent:true}).ok,true);assert.equal(data.get('작업자신청')[1][4],'01012345678');assert.equal(data.get('알림대기')[1][2],'작업자');assert.equal(call('chat',{},'GET',tokenB,STF).messages.length,0);
assert.equal(call('staff/request',{name:'담당자',phone:'01098765432',consent:true},'POST',tokenB,STF).ok,true);assert.throws(()=>call('admin',{},'GET',tokenB,STF));
const staffId=call('staff',{},'GET',tokenA,OWN)[0].id;assert.equal(call('staff',{id:staffId,status:'active'},'POST',tokenA,OWN).ok,true);assert.equal(call('me',{},'GET',tokenB,STF).admin,true);assert.throws(()=>call('staff',{},'GET',tokenB,STF));
const s=call('workplaces',{name:'시험 현장',address:'테스트용 주소',lat:36.88,lng:127.15,radius:100},'POST',tokenA,OWN);const point={siteId:s.id,name:'홍길동',lat:s.lat,lng:s.lng,accuracy:10,consent:true};
assert.throws(()=>call('attendance/record',{...point,kind:'out'}));assert.throws(()=>call('attendance/record',{...point,kind:'in',lat:38}));assert.throws(()=>call('attendance/record',{...point,kind:'in',accuracy:200}));
const entry=call('attendance/record',{...point,kind:'in'});assert.equal(entry.status,'pending');assert.throws(()=>call('attendance/record',{...point,kind:'in'}));const out=call('attendance/record',{...point,kind:'out'});const month=entry.day.slice(0,7);assert.equal(call('attendance?month='+month,{},'GET',tokenB,STF).events.length,0);
const personId=data.get('근태기록')[1][1];assert.equal(call('attendance/review',{id:personId,eventId:entry.eventId,month,status:'approved'},'POST',tokenB,STF).status,'approved');assert.equal(call('attendance/review',{id:personId,eventId:out.eventId,month,status:'approved'},'POST',tokenB,STF).status,'approved');
assert.equal(call('attendance?month='+month,{},'GET').events.filter(e=>e.status==='approved').length,2);assert.throws(()=>call('attendance/admin?month='+month,{},'GET'));
// 인력 관리: 작업자 등록(운영자만), 공정별 일급, 월간 정산, 월 8일 알림
const roster=call('workers',{},'GET',tokenB,STF);assert.equal(roster.applicants.length,1);assert.equal(roster.applicants[0].visitorId,personId);
assert.throws(()=>call('workers',{visitorId:personId,name:'홍길동',grade:'숙련',wage:160000},'POST',tokenB,STF),/운영자/);
assert.throws(()=>call('workers',{visitorId:personId,name:'홍길동',grade:'반장',wage:160000},'POST',tokenA,OWN),/등급/);
const w=call('workers',{visitorId:personId,name:'홍길동',phone:'010-1234-5678',grade:'숙련',wage:160000},'POST',tokenA,OWN);assert.equal(w.wage,160000);
assert.equal(data.get('작업자신청')[1][5],'채용');assert.equal(call('workers',{},'GET',tokenB,STF).applicants.length,0);
assert.throws(()=>call('workers',{visitorId:personId,name:'중복',grade:'입문',wage:140000},'POST',tokenA,OWN),/이미 등록/);
const today=entry.day;let added=0;for(let d=1;d<=28&&added<7;d++){const day=month+'-'+String(d).padStart(2,'0');if(day===today)continue;added++;
 for(const [kind,hm] of [['in','08:30'],['out','17:30']])data.get('근태기록').push(['x'+d+kind,personId,'홍길동',day,kind,new Date(day+'T'+hm+':00+09:00').toISOString(),s.id,s.name,'주소',s.lat,s.lng,10,0,'approved','','','',s.lat,s.lng,100]);}
const pay=call('payroll?month='+month,{},'GET',tokenB,STF);const p=pay.people[0];
assert.equal(p.dayCount,8);assert.equal(p.alert,'enroll');assert.equal(p.gross,1280000);assert.equal(p.incomeTax,2160);assert.equal(p.localTax,210);assert.equal(p.employment,11520);assert.equal(p.net,1266110);assert.equal(p.days.find(d=>d.day!==today).hours,8);
assert.equal(ctx.settle_(150000,20).incomeTax,0);assert.equal(ctx.settle_(160000,3).incomeTax,0);assert.equal(ctx.settle_(170000,5).incomeTax,2700);
assert.throws(()=>call('payroll?month='+month,{},'GET',tokenA));
console.log('PASS: Sheets writes, required contact and consent, private chat, admin approval, location radius, duplicate clocks and monthly approval, worker roster, payroll tax and 8-day alert');
// 운영자가 이메일로 관리자 바로 추가
assert.throws(()=>call('staff/add',{email:'new@example.com'},'POST',tokenB,STF),/운영자만/);
assert.throws(()=>call('staff/add',{email:'not-an-email'},'POST',tokenA,OWN),/이메일/);
assert.equal(call('staff/add',{email:'New@Example.com',name:'신규 관리자'},'POST',tokenA,OWN).id,'google:new@example.com');
assert.equal(call('staff',{},'GET',tokenA,OWN).find(r=>r.id==='google:new@example.com').status,'active');
call('staff',{id:'google:new@example.com',status:'disabled'},'POST',tokenA,OWN);call('staff/add',{email:'new@example.com'},'POST',tokenA,OWN);
assert.equal(call('staff',{},'GET',tokenA,OWN).filter(r=>r.id==='google:new@example.com').length,1);
console.log('PASS: owner adds admin by email');
// 전화번호 앞자리 0 보존
assert.equal(ctx.cellText_('01071423011'),"'01071423011");assert.equal(ctx.phone_(1071423011),'01071423011');assert.equal(ctx.phone_('01012345678'),'01012345678');assert.equal(ctx.phone_(''),'');
data.get('관리자신청').push(['google:old@example.com','2026-10-31','기존',1071423011,'pending','','']);assert.equal(call('staff',{},'GET',tokenA,OWN).find(r=>r.id==='google:old@example.com').phone,'01071423011');
console.log('PASS: phone leading zero');
// 관리자로 등록된 사람은 작업자 채용 목록에서 제외 (전화번호 또는 계정 일치)
data.get('작업자신청').push(['s1','guest:'+'c'.repeat(64),'2026-10-31T01:00:00Z','김석호',1033334444,'접수',''],['s2','guest:'+'d'.repeat(64),'2026-10-31T02:00:00Z','일반지원자','01055556666','접수','']);
data.get('관리자신청').push(['google:ksh@example.com','2026-10-31','김석호','01033334444','pending','','']);
const r2=call('workers',{},'GET',tokenA,OWN);assert.deepEqual(r2.applicants.map(a=>a.name),['일반지원자']);assert.ok(r2.hiddenStaff>=1);
call('staff',{id:'google:ksh@example.com',status:'disabled'},'POST',tokenA,OWN);assert.ok(call('workers',{},'GET',tokenA,OWN).applicants.some(a=>a.name==='김석호'),'disabled staff can be hired again');
console.log('PASS: staff hidden from applicant list');
// 전자 근로계약서: 관리자만 저장, 무결성(SHA-256) 확인, PDF만 허용, 다운로드 코드로 근로자 수령
{const pdf=Buffer.from('%PDF-1.4 test contract');const b64=pdf.toString('base64'),sha=crypto.createHash('sha256').update(pdf).digest('hex');const cid=crypto.randomUUID();
const body={id:cid,workerId:'w1',name:'홍길동',phone:'010-1234-5678',grade:'숙련',wage:160000,start:'2026-11-02',end:'2026-11-30',employerSigner:'박상열',pdf:b64,sha256:sha};
assert.throws(()=>call('contracts',body,'POST',tokenA),/관리자/);
assert.throws(()=>call('contracts',{...body,sha256:'0'.repeat(64)},'POST',tokenA,OWN),/무결성/);
const nopdf=Buffer.from('hello');assert.throws(()=>call('contracts',{...body,pdf:nopdf.toString('base64'),sha256:crypto.createHash('sha256').update(nopdf).digest('hex')},'POST',tokenA,OWN),/PDF/);
assert.throws(()=>call('contracts',{...body,end:'2026-10-01'},'POST',tokenA,OWN),/기간/);
const saved=call('contracts',body,'POST',tokenA,OWN);assert.match(saved.code,/^[A-Za-z0-9]{12}$/);assert.throws(()=>call('contracts',body,'POST',tokenA,OWN),/이미/);
const row=data.get('근로계약서').at(-1);assert.equal(row[3],'01012345678');assert.notEqual(row[14],saved.code,'code stored only as hash');
assert.equal(call('contracts',{},'GET',tokenB,STF)[0].sha256,sha);
assert.equal(call('contracts/file?id='+cid,{},'GET',tokenA,OWN).base64,b64);
const got=call('contract/public',{code:saved.code},'POST',tokenB);assert.equal(got.base64,b64);assert.equal(got.name,'홍길동');
assert.throws(()=>call('contract/public',{code:'AAAAAAAAAAAA'},'POST',tokenB),/찾을 수 없습니다/);
console.log('PASS: e-contract save, integrity, worker download code');}
// 근로계약서 양식(사업주 정보): 관리자별 저장·수정·삭제, 다른 관리자 양식은 보이지 않음
assert.throws(()=>call('templates',{},'GET',tokenA),/관리자/);
assert.throws(()=>call('templates',{name:'LK 테크'},'POST',tokenA,OWN),/대표자/);
assert.throws(()=>call('templates',{name:'LK 테크',ceo:'박상열',bizNo:'123'},'POST',tokenA,OWN),/10자리/);
const t1=call('templates',{name:'LK 테크',ceo:'박상열',bizNo:'1234567890',address:'충남 천안시',phone:'010-1111-2222',site:'직산 현장'},'POST',tokenA,OWN);
assert.equal(t1.bizNo,'123-45-67890');assert.equal(t1.label,'LK 테크 · 직산 현장');assert.equal(t1.phone,'01011112222');
assert.equal(call('templates',{},'GET',tokenB,STF).length,0,'staff cannot see owner template');
assert.throws(()=>call('templates',{id:t1.id,name:'X',ceo:'Y'},'POST',tokenB,STF),/찾을 수 없습니다/);
assert.equal(call('templates',{id:t1.id,name:'LK 테크',ceo:'박상열',site:'평택 현장'},'POST',tokenA,OWN).site,'평택 현장');
assert.equal(call('templates',{},'GET',tokenA,OWN).length,1);call('templates',{id:t1.id,delete:true},'POST',tokenA,OWN);assert.equal(call('templates',{},'GET',tokenA,OWN).length,0);
console.log('PASS: per-admin contract templates');
