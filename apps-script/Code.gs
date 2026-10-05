/** Neighbors API — Google Apps Script (V8). مرتبط بالشيت (Extensions > Apps Script). */
const T={cfg:'إعدادات',mem:'أعضاء',pay:'دفعات',prj:'مشاريع',exp:'مصروفات',usr:'مستخدمون',log:'سجل'};
const H={cfg:['المفتاح','القيمة'],mem:['id','الاسم','الهاتف','نشط','رقم_البيت'],pay:['التاريخ','id_عضو','id_مشروع','المبلغ','الشهر'],
  prj:['id','الاسم','التكلفة','مساهمة_فردية','رصيد_افتتاحي'],exp:['التاريخ','البند','المبلغ'],usr:['username','salt','hash'],log:['الوقت','المستوى','المستخدم','الحدث']};
const TZ='Africa/Algiers';
let LOG=[];
const log=(l,u,e)=>LOG.push([new Date(),l,u||'',e]);
const flush=()=>{if(!LOG.length)return;const s=sh('log');s.getRange(s.getLastRow()+1,1,LOG.length,4).setValues(LOG);LOG=[]};
const sh=k=>SpreadsheetApp.getActive().getSheetByName(T[k]);
const rows=k=>sh(k).getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
const ym=v=>v instanceof Date?Utilities.formatDate(v,TZ,'yyyy-MM'):String(v);
const ymd=v=>v instanceof Date?Utilities.formatDate(v,TZ,'yyyy-MM-dd'):String(v);
const json=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const num=v=>{const n=Number(v);if(!isFinite(n)||n<=0)throw new Error('bad amount');return n};
const txt=s=>{s=String(s||'').trim().slice(0,80);return /^[=+\-@]/.test(s)?"'"+s:s};
const hash=(salt,pw)=>{let h=salt+pw;for(let i=0;i<1000;i++)h=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,h));return h};

/* ---------- نقاط الدخول ---------- */
function doGet(e){
  try{ if(e.parameter.action==='public')return json(summary()); return json({error:'bad action'}) }
  catch(err){log('ERROR','','public: '+err);return json({error:'server'})} finally{flush()}
}
function doPost(e){
  let u='';
  try{
    const b=JSON.parse(e.postData.contents);
    if(b.action==='login')return json({ok:true,data:login(String(b.username),String(b.password))});
    u=CacheService.getScriptCache().get('t:'+b.token);
    if(!u)return json({error:'401'});
    return json({ok:true,data:route(b,u)});
  }catch(err){log('ERROR',u,String(err));return json({ok:false,error:String(err.message||err)})}
  finally{flush()}
}

/* ---------- المصادقة ---------- */
function login(user,pw){
  const c=CacheService.getScriptCache(),k='f:'+user,n=+c.get(k)||0;
  if(n>=5)throw new Error('locked');
  const r=rows('usr').find(x=>x[0]===user);
  if(!r||hash(r[1],pw)!==r[2]){c.put(k,n+1,900);log('WARN',user,'login_fail');throw new Error('invalid')}
  c.remove(k);
  const t=Utilities.getUuid()+Utilities.getUuid();
  c.put('t:'+t,user,21600);log('INFO',user,'login');
  return{token:t};
}
/** شغّلها من المحرر بعد ضبط INIT_USER و INIT_PASS في Project Settings > Script Properties (تُحذف كلمة السر بعد الاستعمال). */
function createAdmin(){
  const p=PropertiesService.getScriptProperties(),u=p.getProperty('INIT_USER'),pw=p.getProperty('INIT_PASS');
  if(!u||!pw)throw new Error('Set INIT_USER and INIT_PASS in Script Properties');
  const salt=Utilities.getUuid(),s=sh('usr'),i=rows('usr').findIndex(r=>r[0]===u);
  const row=[u,salt,hash(salt,pw)];
  if(i<0)s.appendRow(row);else s.getRange(i+2,1,1,3).setValues([row]);
  p.deleteProperty('INIT_PASS');log('INFO','system','admin_set:'+u);flush();
}

/* ---------- المنطق ---------- */
function summary(){
  const cfg=Object.fromEntries(rows('cfg')),month=Utilities.formatDate(new Date(),TZ,'yyyy-MM');
  const pay=rows('pay'),act=rows('mem').filter(r=>r[3]!==false&&r[3]!=='لا');
  const sum=f=>pay.filter(f).reduce((a,r)=>a+Number(r[3]),0);
  const mon=pay.filter(r=>!r[2]&&ym(r[4])===month);
  const got=mon.reduce((a,r)=>a+Number(r[3]),0),target=Number(cfg.monthly_target)||0;
  const income=Number(cfg.general_opening||0)+sum(r=>!r[2]);
  const spent=rows('exp').reduce((a,r)=>a+Number(r[2]),0);
  return{schema_version:1,generated_at:new Date().toISOString(),month,notice:String(cfg.notice||''),
    subscriptions:{target,collected:got,rate:target?Math.min(got/target,1):0,paid_members:new Set(mon.map(r=>r[1])).size,active_members:act.length},
    general:{income,expenses:spent,balance:income-spent},
    projects:rows('prj').map(r=>{const c=Number(r[4]||0)+sum(p=>p[2]===r[0]),cost=Number(r[2]);
      return{name:r[1],cost,collected:c,deficit:Math.max(cost-c,0),rate:cost?Math.min(c/cost,1):0}})};
}
function append(u,k,row,ev,rid){
  const L=LockService.getScriptLock(),c=CacheService.getScriptCache();L.waitLock(10000);
  try{if(rid&&c.get('r:'+rid)){log('WARN',u,'duplicate_blocked:'+ev);return{saved:true,dup:true}}
    sh(k).appendRow(row);if(rid)c.put('r:'+rid,'1',1800);log('INFO',u,ev)}finally{L.releaseLock()}
  return{saved:true};
}
function route(b,u){
  switch(b.action){
    case 'lists':return{summary:summary(),members:rows('mem').map(r=>[r[0],r[1],r[4]||'']),projects:rows('prj').map(r=>[r[0],r[1]]),
      recent:rows('pay').slice(-15).reverse().map(r=>[ymd(r[0]),r[1],r[2],r[3],ym(r[4])])};
    case 'addMember':{
      const name=txt(b.name);if(!name)throw new Error('bad name');
      const id='M'+(Math.max(0,...rows('mem').map(r=>+String(r[0]).slice(1)||0))+1);
      const house=txt(b.house);if(!house)throw new Error('bad house');
      return append(u,'mem',[id,name,txt(b.phone),true,house],'addMember:'+id,b.rid)}
    case 'addPayment':{
      const m=String(b.member),j=String(b.project||''),mo=String(b.month);
      if(!rows('mem').some(r=>r[0]===m))throw new Error('bad member');
      if(j&&!rows('prj').some(r=>r[0]===j))throw new Error('bad project');
      if(!/^\d{4}-\d{2}$/.test(mo))throw new Error('bad month');
      return append(u,'pay',[new Date(),m,j,num(b.amount),mo],'addPayment:'+m+':'+(j||'monthly'),b.rid)}
    case 'addExpense':{
      const item=txt(b.item);if(!item)throw new Error('bad item');
      const d=/^\d{4}-\d{2}-\d{2}$/.test(b.date)?new Date(b.date):new Date();
      return append(u,'exp',[d,item,num(b.amount)],'addExpense',b.rid)}
    case 'setNotice':{
      let t=String(b.text||'').trim().slice(0,300);if(/^[=+\-@]/.test(t))t="'"+t;
      const s=sh('cfg'),i=rows('cfg').findIndex(r=>r[0]==='notice');
      if(i<0)s.appendRow(['notice',t]);else s.getRange(i+2,2).setValue(t);
      log('INFO',u,'setNotice');return{saved:true}}
    default:throw new Error('bad action');
  }
}

/* ---------- التهيئة (مرة واحدة) ---------- */
function setup(){
  const ss=SpreadsheetApp.getActive();
  for(const k in T){if(!ss.getSheetByName(T[k])){const s=ss.insertSheet(T[k]);s.getRange(1,1,1,H[k].length).setValues([H[k]]).setFontWeight('bold');s.setFrozenRows(1)}}
  if(sh('cfg').getLastRow()<2)sh('cfg').getRange(2,1,2,2).setValues([['monthly_target',58770],['general_opening',0]]);
  if(sh('prj').getLastRow()<2)sh('prj').getRange(2,1,2,5).setValues([['P1','مشروع الكمرات',212000,4000,92000],['P2','مشروع البوابة',90000,4000,46770]]);
  const m=sh('mem');if(!m.getRange('E1').getValue())m.getRange('E1').setValue('رقم_البيت').setFontWeight('bold');
  if(!rows('cfg').some(r=>r[0]==='notice'))sh('cfg').appendRow(['notice','']);
  sh('pay').getRange('E:E').setNumberFormat('@');sh('mem').getRange('C:C').setNumberFormat('@');
  ['usr','log'].forEach(k=>sh(k).hideSheet());
}
