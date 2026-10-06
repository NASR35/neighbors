/** Neighbors API v2 — Google Apps Script (V8). مرتبط بالشيت (Extensions > Apps Script). */
const T={cfg:'إعدادات',hs:'بيوت',ty:'أنواع',py:'مدفوعات',ex:'مصروفات',usr:'مستخدمون',log:'سجل'};
const H={cfg:['المفتاح','القيمة'],hs:['البيت','الاسم','الهاتف','شهر_البداية'],
  ty:['id','النوع','الاسم','المبلغ','التكلفة','البداية','نشط','رصيد_افتتاحي'],
  py:['وصل','التاريخ','البيت','id_النوع','الشهر','المبلغ','ملغاة'],
  ex:['التاريخ','البند','المبلغ'],usr:['username','salt','hash'],log:['الوقت','المستوى','المستخدم','الحدث']};
const TZ='Africa/Algiers',PRJ='مشروع',MON='شهري';
let LOG=[];
const log=(l,u,e)=>LOG.push([new Date(),l,u||'',e]);
const flush=()=>{if(!LOG.length)return;const s=sh('log');s.getRange(s.getLastRow()+1,1,LOG.length,4).setValues(LOG);LOG=[]};
const sh=k=>SpreadsheetApp.getActive().getSheetByName(T[k]);
const rows=k=>sh(k).getDataRange().getValues().slice(1).filter(r=>r[0]!=='');
const fmt=(v,p)=>v instanceof Date?Utilities.formatDate(v,TZ,p):String(v);
const ym=v=>fmt(v,'yyyy-MM'),ymd=v=>fmt(v,'yyyy-MM-dd');
const NOW=()=>ym(new Date()),TODAY=()=>ymd(new Date());
const mi=s=>{const p=String(s).split('-');return +p[0]*12+(+p[1]-1)};
const ms=x=>Math.floor(x/12)+'-'+String(x%12+1).padStart(2,'0');
const json=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const num=v=>{const n=Number(v);if(!isFinite(n)||n<=0)throw new Error('bad amount');return n};
const txt=s=>{s=String(s||'').trim().slice(0,80);return /^[=+\-@]/.test(s)?"'"+s:s};
const hash=(salt,pw)=>{let h=salt+pw;for(let i=0;i<1000;i++)h=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,h));return h};

/* ---------- نقاط الدخول ---------- */
function doGet(e){
  try{ if(e.parameter.action==='public')return json(pub(load())); return json({error:'bad action'}) }
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
/** شغّلها بعد ضبط INIT_USER و INIT_PASS في Script Properties (تُحذف كلمة السر بعد الاستعمال). */
function createAdmin(){
  const p=PropertiesService.getScriptProperties(),u=p.getProperty('INIT_USER'),pw=p.getProperty('INIT_PASS');
  if(!u||!pw)throw new Error('Set INIT_USER and INIT_PASS in Script Properties');
  const salt=Utilities.getUuid(),s=sh('usr'),i=rows('usr').findIndex(r=>r[0]===u),row=[u,salt,hash(salt,pw)];
  if(i<0)s.appendRow(row);else s.getRange(i+2,1,1,3).setValues([row]);
  p.deleteProperty('INIT_PASS');log('INFO','system','admin_set:'+u);flush();
}

/* ---------- القراءة والحساب ---------- */
function load(){
  const hs=rows('hs').map(r=>({id:String(r[0]),name:String(r[1]||''),phone:String(r[2]||''),start:r[3]?ym(r[3]):''}));
  const ty=rows('ty').map(r=>({id:String(r[0]),kind:r[1],name:String(r[2]),amount:+r[3]||0,cost:+r[4]||0,start:r[5]?ymd(r[5]):'',active:r[6]!==false&&r[6]!=='لا',open:+r[7]||0}));
  const all=rows('py').map(r=>({ref:r[0],date:ymd(r[1]),house:String(r[2]),type:String(r[3]),month:r[4]?ym(r[4]):'',amount:+r[5]||0,off:!!r[6]}));
  return{hs,ty,all,py:all.filter(p=>!p.off)};
}
/* الدين يظهر من اليوم الأول من الشهر الموالي؛ البيت بلا شهر بداية لا يُحتسب عليه شيء */
function due(h,t,now){
  if(!h.start)return 0;
  if(t.kind===PRJ)return t.start&&TODAY()<t.start?0:t.amount;
  const s=t.start&&t.start.slice(0,7)>h.start?t.start.slice(0,7):h.start;
  return Math.max(0,mi(now)-mi(s))*t.amount;
}
function houseRows(D,now){
  const P={};D.py.forEach(p=>{const k=p.house+'|'+p.type;P[k]=(P[k]||0)+p.amount});
  return D.hs.map(h=>{
    const lines=D.ty.filter(t=>t.active).map(t=>{const d=due(h,t,now),p=P[h.id+'|'+t.id]||0;return{t:t.name,id:t.id,due:d,paid:p,debt:Math.max(0,d-p)}});
    return{...h,lines,debt:lines.reduce((a,l)=>a+l.debt,0)};
  });
}
function pub(D){
  const now=NOW(),cfg=Object.fromEntries(rows('cfg')),HR=houseRows(D,now);
  const monIds=new Set(D.ty.filter(t=>t.kind===MON).map(t=>t.id)),per=D.ty.filter(t=>t.kind===MON&&t.active).reduce((a,t)=>a+t.amount,0);
  const exp=m=>per*D.hs.filter(h=>h.start&&h.start<=m).length;
  const monPay=D.py.filter(p=>monIds.has(p.type)),col=m=>monPay.filter(p=>p.month===m).reduce((a,p)=>a+p.amount,0);
  const e=exp(now),c=col(now),income=Number(cfg.general_opening||0)+monPay.reduce((a,p)=>a+p.amount,0);
  const spent=rows('ex').reduce((a,r)=>a+Number(r[2]),0),part=D.hs.filter(h=>h.start).length;
  return{schema_version:2,generated_at:new Date().toISOString(),month:now,notice:String(cfg.notice||''),
    houses:{total:D.hs.length,participating:part},
    monthly:{expected:e,collected:c,rate:e?Math.min(c/e,1):0,paid_houses:new Set(monPay.filter(p=>p.month===now).map(p=>p.house)).size},
    debts:{total:HR.reduce((a,h)=>a+h.debt,0),houses:HR.filter(h=>h.debt>0).length},
    general:{income,expenses:spent,balance:income-spent},
    projects:D.ty.filter(t=>t.kind===PRJ).map(t=>{const ps=D.py.filter(p=>p.type===t.id),got=t.open+ps.reduce((a,p)=>a+p.amount,0);
      return{name:t.name,active:t.active,cost:t.cost,collected:got,deficit:Math.max(t.cost-got,0),rate:t.cost?Math.min(got/t.cost,1):0,per:t.amount,paid_houses:new Set(ps.map(p=>p.house)).size,participants:part}}),
    chart:[5,4,3,2,1,0].map(i=>{const m=ms(mi(now)-i);return{m,expected:exp(m),collected:col(m)}})};
}

/* ---------- الكتابة (قفل + منع التكرار) ---------- */
function W(u,rid,ev,fn){
  const L=LockService.getScriptLock(),c=CacheService.getScriptCache();L.waitLock(10000);
  try{
    if(rid&&c.get('r:'+rid)){log('WARN',u,'duplicate_blocked:'+ev);return{saved:true,dup:true}}
    const out=fn()||{};SpreadsheetApp.flush();if(rid)c.put('r:'+rid,'1',1800);log('INFO',u,ev);
    return{saved:true,...out};
  }finally{L.releaseLock()}
}
const rowOf=(k,id)=>{const s=sh(k),v=s.getRange(1,1,s.getLastRow(),1).getValues().flat(),i=v.findIndex((x,j)=>j>0&&String(x)===String(id));if(i<0)throw new Error('not found');return i+1};
const month=v=>{v=String(v||'');if(v&&!/^\d{4}-\d{2}$/.test(v))throw new Error('bad month');return v};

function route(b,u){
  const D=load(),now=NOW(),tn=id=>(D.ty.find(t=>t.id===id)||{}).name||id,hn=id=>(D.hs.find(h=>h.id===id)||{}).name||'';
  switch(b.action){
    case 'dash':{const HR=houseRows(D,now);
      return{pub:pub(D),top:HR.filter(h=>h.debt>0).sort((a,c)=>c.debt-a.debt).slice(0,5).map(h=>({id:h.id,name:h.name,phone:h.phone,debt:h.debt}))}}
    case 'houses':return houseRows(D,now).map(h=>({id:h.id,name:h.name,phone:h.phone,start:h.start,debt:h.debt}));
    case 'house':{const h=houseRows(D,now).find(x=>x.id===b.id);if(!h)throw new Error('not found');
      return{h,pays:D.all.filter(p=>p.house===b.id).slice(-40).reverse().map(p=>({...p,type:tn(p.type)}))}}
    case 'types':return D.ty.map(t=>{const ps=D.py.filter(p=>p.type===t.id);
      return{...t,collected:t.open+ps.reduce((a,p)=>a+p.amount,0),paid_houses:new Set(ps.map(p=>p.house)).size}});
    case 'payments':return D.all.slice(-100).reverse().map(p=>({...p,name:hn(p.house),type:tn(p.type)}));

    case 'addHouse':{
      const id=String(b.id||'').trim().toUpperCase();if(!/^[A-D]\d{2,3}$/.test(id))throw new Error('bad house');
      if(D.hs.some(h=>h.id===id))throw new Error('house exists');
      const st=month(b.start);
      return W(u,b.rid,'addHouse:'+id,()=>{sh('hs').appendRow([id,txt(b.name),txt(b.phone),st])})}
    case 'editHouse':{const st=month(b.start);
      return W(u,b.rid,'editHouse:'+b.id,()=>{sh('hs').getRange(rowOf('hs',b.id),2,1,3).setValues([[txt(b.name),txt(b.phone),st]])})}
    case 'addType':{
      const kind=b.kind===PRJ?PRJ:MON,name=txt(b.name);if(!name)throw new Error('bad name');
      const amount=num(b.amount),isP=kind===PRJ,cost=isP?num(b.cost):'';
      const st=isP?String(b.start||''):'';if(isP&&!/^\d{4}-\d{2}-\d{2}$/.test(st))throw new Error('bad date');
      const open=isP?Math.max(0,Number(b.open)||0):0,id='T'+(Math.max(0,...D.ty.map(t=>+t.id.slice(1)||0))+1);
      return W(u,b.rid,'addType:'+id,()=>{sh('ty').appendRow([id,kind,name,amount,cost,st,true,open])})}
    case 'toggleType':
      return W(u,b.rid,'toggleType:'+b.id,()=>{const r=rowOf('ty',b.id),c=sh('ty').getRange(r,7);c.setValue(!(c.getValue()!==false&&c.getValue()!=='لا'))});
    case 'pay':{
      const h=D.hs.find(x=>x.id===b.house),t=D.ty.find(x=>x.id===b.type&&x.active);if(!h||!t)throw new Error('bad input');
      return W(u,b.rid,'pay:'+h.id+':'+t.id,()=>{
        const s=sh('py'),last=s.getLastRow(),ref0=last>1?Math.max(0,...s.getRange(2,1,last-1,1).getValues().flat().map(Number).filter(isFinite)):0,out=[],now2=new Date();
        if(t.kind===PRJ)out.push([ref0+1,now2,h.id,t.id,now,num(b.amount),'']);
        else{
          const mo=month(b.month);if(!mo)throw new Error('bad month');
          const n=Math.max(1,Math.min(24,+b.months||1)),amt=n===1&&b.amount?num(b.amount):t.amount;
          const have=new Set(D.py.filter(p=>p.house===h.id&&p.type===t.id).map(p=>p.month));
          for(let i=0;i<n;i++){const m=ms(mi(mo)+i);if(!have.has(m))out.push([ref0+out.length+1,now2,h.id,t.id,m,amt,''])}
          if(!out.length)throw new Error('already paid');
        }
        s.getRange(last+1,1,out.length,7).setValues(out);return{refs:out.map(r=>r[0])};
      })}
    case 'cancelPay':
      return W(u,b.rid,'cancelPay:'+b.ref,()=>{const s=sh('py'),v=s.getRange(1,1,s.getLastRow(),1).getValues().flat(),i=v.findIndex((x,j)=>j>0&&Number(x)===Number(b.ref));
        if(i<0)throw new Error('not found');s.getRange(i+1,7).setValue('ملغاة')});
    case 'addExpense':{
      const item=txt(b.item);if(!item)throw new Error('bad item');
      const d=/^\d{4}-\d{2}-\d{2}$/.test(b.date)?new Date(b.date):new Date(),a=num(b.amount);
      return W(u,b.rid,'addExpense',()=>{sh('ex').appendRow([d,item,a])})}
    case 'setNotice':{
      let t=String(b.text||'').trim().slice(0,300);if(/^[=+\-@]/.test(t))t="'"+t;
      return W(u,b.rid,'setNotice',()=>{const i=rows('cfg').findIndex(r=>r[0]==='notice');
        if(i<0)sh('cfg').appendRow(['notice',t]);else sh('cfg').getRange(i+2,2).setValue(t)})}
    default:throw new Error('bad action');
  }
}

/* ---------- التهيئة (مرة واحدة، آمنة للتكرار) ---------- */
function setup(){
  const ss=SpreadsheetApp.getActive();
  for(const k in T){if(!ss.getSheetByName(T[k])){const s=ss.insertSheet(T[k]);s.getRange(1,1,1,H[k].length).setValues([H[k]]).setFontWeight('bold');s.setFrozenRows(1)}}
  if(!rows('cfg').some(r=>r[0]==='general_opening'))sh('cfg').appendRow(['general_opening',0]);
  if(!rows('cfg').some(r=>r[0]==='notice'))sh('cfg').appendRow(['notice','']);
  if(sh('ty').getLastRow()<2)sh('ty').appendRow(['T1',MON,'الاشتراك الشهري',1000,'','',true,0]);
  sh('hs').getRange('C:D').setNumberFormat('@');sh('ty').getRange('F:F').setNumberFormat('@');sh('py').getRange('E:E').setNumberFormat('@');
  ['usr','log'].forEach(k=>sh(k).hideSheet());
}
