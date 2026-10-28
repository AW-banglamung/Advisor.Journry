/** Advisor Journey – ระบบหลังบ้าน (Google Apps Script) */
const H = ['id','token','unit','name','nick','age','dob','phone','edu','job','created','done'];
const NAMES = ['ช่วงเริ่มต้นสนใจ','เตรียมสอบ+อบรม ก่อนออกรหัส','อบรมและสอบใบอนุญาต','เตรียมความรู้ก่อนออกรหัส (Starter #1)','สมัครสมาชิก','หลังได้รับรหัสตัวแทน','เตรียมเริ่มขาย (Starter #2)','ฝึกภาคสนาม JOIN Fieldwork','FUFT เดือนที่ 1','FUFT เดือนที่ 2','FUFT เดือนที่ 3','สิทธิ์การขาย Unitlink'];
// [กำหนดเสร็จ(วันนับจากวันเพิ่มชื่อ), จำนวนรายการ] ต้องตรงกับ index.html
const DUE = [[30,5],[44,7],[58,5],[72,6],[79,3],[86,7],[93,4],[107,3],[135,9],[165,6],[197,4],[227,13]];
const DAY = 864e5;

function P(k){ return PropertiesService.getScriptProperties().getProperty(k); }
function sheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('Agents');
  if(!s){ s = ss.insertSheet('Agents'); s.setFrozenRows(1); }
  s.getRange(1,1,1,H.length).setValues([H]);
  return s;
}
function rows(){
  const v = sheet().getDataRange().getValues(); v.shift();
  return v.filter(r=>r[0]).map(r=>{ const o={}; H.forEach((h,i)=>o[h]=r[i]);
    try{o.done = JSON.parse(o.done||'{}')}catch(e){o.done={}}
    o.created = o.created instanceof Date ? o.created.toISOString() : String(o.created);
    return o; });
}
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function isOwner(pin){ return P('PIN') && String(pin) === P('PIN'); }

function doGet(e){
  const p = e.parameter||{};
  if(p.action==='ranks') return out({ranks: JSON.parse(P('RANKS')||'null')});
  if(p.action==='list'){
    if(!isOwner(p.pin)) return out({error:'pin'});
    return out({agents: rows()});
  }
  if(p.action==='get'){
    const a = rows().find(r=>p.token && r.token===p.token);
    return out(a ? {agent:a} : {error:'notfound'});
  }
  return out({ok:true});
}

function doPost(e){
  const b = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const s = sheet();
    if(b.action==='add'){
      if(!isOwner(b.pin)) return out({error:'pin'});
      const a = b.agent, id='a'+Date.now().toString(36), token=Utilities.getUuid().replace(/-/g,'');
      s.appendRow([id,token,a.unit,a.name,a.nick,a.age,a.dob,"'"+a.phone,a.edu,a.job,new Date().toISOString(),'{}']);
      return out({ok:true,id});
    }
    if(b.action==='ranks'){
      if(!isOwner(b.pin)) return out({error:'pin'});
      PropertiesService.getScriptProperties().setProperty('RANKS', JSON.stringify(b.ranks)); return out({ok:true});
    }
    const ids = s.getRange(1,1,s.getLastRow(),2).getValues();
    let r = -1;
    for(let i=1;i<ids.length;i++){
      if(b.token ? ids[i][1]===b.token : ids[i][0]===b.id){ r=i+1; break; }
    }
    if(r<0) return out({error:'notfound'});
    if(!b.token && !isOwner(b.pin)) return out({error:'pin'});
    if(b.action==='tick'){
      const c = s.getRange(r,12); const d = JSON.parse(c.getValue()||'{}');
      if(b.on) d[b.key]=new Date().toISOString(); else delete d[b.key];
      c.setValue(JSON.stringify(d)); return out({ok:true});
    }
    if(b.action==='del'){ if(!isOwner(b.pin)) return out({error:'pin'}); s.deleteRow(r); return out({ok:true}); }
  } finally { lock.releaseLock(); }
  return out({error:'bad'});
}

/** ขั้นที่ยังไม่เสร็จขั้นแรกของตัวแทน */
function cur(a){
  const start = new Date(a.created).getTime();
  for(let si=0; si<DUE.length; si++){
    let c=0; for(let i=0;i<DUE[si][1];i++) if(a.done[si+'-'+i]) c++;
    if(c<DUE[si][1]){ const due=start+DUE[si][0]*DAY;
      return {si,c,t:DUE[si][1],due,d:Math.ceil((due-Date.now())/DAY),name:NAMES[si]}; }
  }
  return null;
}
const dstr = ms => Utilities.formatDate(new Date(ms),'Asia/Bangkok','d/M/yyyy');
/** ข้อความเตือนน่ารักๆ พร้อมคัดลอกไปวางใน LINE */
function msg(a,f){
  const n = a.nick||a.name;
  const when = f.d<0 ? '⏰ เลยกำหนดมา '+(-f.d)+' วันแล้วค่ะ' : f.d===0 ? '🎯 วันนี้ครบกำหนดแล้วนะคะ' : '🗓 ครบกำหนดอีก '+f.d+' วัน ('+dstr(f.due)+')';
  const url = P('PAGE_URL') ? '\n👉 ดูรายการที่เหลือ: '+P('PAGE_URL')+'?t='+a.token : '';
  return '🌸 สวัสดีค่ะคุณ'+n+'\n📚 ขั้นที่ '+(f.si+1)+': '+f.name+'\n'+when+'\n✅ ทำแล้ว '+f.c+'/'+f.t+' รายการ'+url+
    (f.d<0 ? '\n🌷 ไม่เป็นไรนะคะ ค่อยๆ ทำไปด้วยกัน ติดตรงไหนทักพี่ได้เลย 🤗' : '\nสู้ๆ นะคะ 💪🥰');
}

/** รันทุกเช้า: ส่งอีเมลสรุปให้เจ้าของ พร้อมข้อความเตือนของแต่ละคน */
function dailyDigest(){
  const late=[], soon=[];
  rows().forEach(a=>{ const f=cur(a); if(!f) return;
    const head = '■ '+a.name+' ('+a.unit+') – ขั้นที่ '+(f.si+1)+' ทำแล้ว '+f.c+'/'+f.t;
    const item = head+(f.d<0?' เลยกำหนด '+(-f.d)+' วัน':' ครบกำหนดอีก '+f.d+' วัน')+'\n\n'+msg(a,f);
    if(f.d<0) late.push(item); else if(f.d<=3) soon.push(item); });
  if(!late.length && !soon.length) return;
  const sep = '\n\n----------------------------------------\n\n';
  MailApp.sendEmail(P('OWNER_EMAIL'),'Advisor Journey: ต้องตาม '+late.length+' คน',
    'เลยกำหนด ('+late.length+' คน)\n\n'+(late.join(sep)||'-')+sep+'ใกล้ครบกำหนด 3 วัน ('+soon.length+' คน)\n\n'+(soon.join(sep)||'-'));
}

/** รันครั้งเดียว: ตั้งรหัส PIN, อีเมล และเวลาส่งสรุป (แก้ 2 ค่าด้านล่างก่อนรัน) */
function setup(){
  PropertiesService.getScriptProperties().setProperties({PIN:'1234', OWNER_EMAIL:'your@email.com'});
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(8).create();
}
/** (ไม่บังคับ) ใส่ลิงก์เว็บของคุณ เพื่อให้ข้อความเตือนมีลิงก์ส่วนตัวของตัวแทนแนบไปด้วย */
function setPageUrl(){
  PropertiesService.getScriptProperties().setProperty('PAGE_URL','https://ชื่อคุณ.github.io/advisor/');
}
