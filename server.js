const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'reservations.json');
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';

fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify({updatedAt:null,reservations:[]}, null, 2));
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => cb(null, /\.xlsx?$/i.test(file.originalname || ''))
});

app.disable('x-powered-by');
app.use(express.json());
app.use((req,res,next)=>{
  res.setHeader('X-Content-Type-Options','nosniff');
  res.setHeader('Referrer-Policy','same-origin');
  next();
});
app.use(express.static(path.join(__dirname,'public')));

function readData(){
  try { return JSON.parse(fs.readFileSync(DATA_FILE,'utf8')); }
  catch { return {updatedAt:null,reservations:[]}; }
}
function writeData(data){
  const tmp = `${DATA_FILE}.${crypto.randomUUID()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data,null,2),'utf8');
  fs.renameSync(tmp, DATA_FILE);
}
function norm(v){ return v == null ? '' : String(v).trim(); }
function num(v){ const m=norm(v).replace(/,/g,'').match(/\d+/); return m?Number(m[0]):0; }
function col(row,names){ for(const n of names) if(Object.prototype.hasOwnProperty.call(row,n)) return row[n]; return ''; }
function parseSlot(raw){
  const s=norm(raw);
  const m=s.match(/(\d{1,2})월\s*(\d{1,2})일\s*(?:\([^)]*\))?\s*(\d{1,2}):(\d{2})/);
  if(!m) return null;
  const month=Number(m[1]), day=Number(m[2]), hour=Number(m[3]), minute=Number(m[4]);
  return {month,day,hour,minute,key:`${month}-${day}-${hour}:${String(minute).padStart(2,'0')}`,label:`${hour}:${String(minute).padStart(2,'0')}`};
}
function parseExcel(buffer){
  const wb=XLSX.read(buffer,{type:'buffer',cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
  return rows.map((row,i)=>{
    const name=norm(col(row,['예매자 성함','예매자 성함(*)']));
    const count=num(col(row,['예매 인원','예매 인원(*)']));
    const rawDates=norm(col(row,['예매 일자','예매 일자(*)']));
    const rawFriends=norm(col(row,['지인','지인 선택']));
    const slots=rawDates.split('|').map(parseSlot).filter(Boolean);
    const friends=rawFriends.split('|').map(norm).filter(Boolean);
    return {id:i+2,name,count,friends,slots};
  }).filter(r=>r.name && r.count>0 && r.slots.length);
}
function adminOk(req){
  if(!ADMIN_PASSWORD) return false;
  const a=Buffer.from(String(req.headers['x-admin-password']||''));
  const b=Buffer.from(ADMIN_PASSWORD);
  return a.length===b.length && crypto.timingSafeEqual(a,b);
}

app.get('/health',(req,res)=>res.json({ok:true}));
app.get('/api/status',(req,res)=>{
  const d=readData();
  res.json({updatedAt:d.updatedAt,reservationCount:d.reservations.length,totalPeople:d.reservations.reduce((s,r)=>s+r.count,0)});
});
app.get('/api/search',(req,res)=>{
  const q=norm(req.query.name);
  if(!q || q.length>50) return res.status(400).json({error:'배우 이름을 입력해주세요.'});
  const d=readData();
  const matches=d.reservations.filter(r=>r.friends.includes(q));
  const reservations=matches.map(r=>({id:r.id,name:r.name,count:r.count,slots:r.slots}));
  reservations.sort((a,b)=>{
    const A=a.slots[0]||{},B=b.slots[0]||{};
    return (A.month||99)-(B.month||99)||(A.day||99)-(B.day||99)||(A.hour||99)-(B.hour||99)||(A.minute||99)-(B.minute||99)||a.name.localeCompare(b.name,'ko');
  });
  res.json({name:q,totalBookings:matches.length,totalPeople:matches.reduce((s,r)=>s+r.count,0),reservations,updatedAt:d.updatedAt});
});
app.post('/api/upload',upload.single('file'),(req,res)=>{
  if(!adminOk(req)) return res.status(401).json({error:'관리자 비밀번호가 올바르지 않습니다.'});
  if(!req.file) return res.status(400).json({error:'엑셀 파일을 선택해주세요.'});
  try {
    const reservations=parseExcel(req.file.buffer);
    if(!reservations.length) throw new Error('유효한 예매 데이터가 없습니다.');
    const updatedAt=new Date().toISOString();
    writeData({updatedAt,reservations});
    res.json({ok:true,updatedAt,reservationCount:reservations.length,totalPeople:reservations.reduce((s,r)=>s+r.count,0)});
  } catch(e) { res.status(400).json({error:e.message||'엑셀 처리에 실패했습니다.'}); }
});
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,'0.0.0.0',()=>console.log(`Listening on ${PORT}`));
