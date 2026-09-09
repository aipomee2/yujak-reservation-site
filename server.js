const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const EXCEL_FILE = path.join(__dirname, 'data', 'reservations.xlsx');
const EXCLUDED_NAMES = new Set(['김맹훈', '정해진']);

function text(v) { return v == null ? '' : String(v).trim(); }
function num(v) {
  const m = text(v).replace(/,/g, '').match(/\d+(?:\.\d+)?/);
  return m ? Number(m[0]) : 0;
}
function col(row, names) {
  for (const n of names) if (Object.prototype.hasOwnProperty.call(row, n)) return row[n];
  return '';
}
function parseSlot(raw) {
  const s = text(raw);
  const m = s.match(/(\d{1,2})월\s*(\d{1,2})일\s*(?:\([^)]*\))?\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const month = Number(m[1]), day = Number(m[2]), hour = Number(m[3]), minute = Number(m[4]);
  return { month, day, hour, minute, key: `${month}-${day}-${hour}:${String(minute).padStart(2,'0')}`, label: `${String(hour).padStart(2,'0')}:${String(minute).padStart(2,'0')}` };
}
function readExcel() {
  if (!fs.existsSync(EXCEL_FILE)) throw new Error('data/reservations.xlsx 파일을 찾을 수 없습니다.');
  const wb = XLSX.readFile(EXCEL_FILE, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  const reservations = rows.map((row, index) => {
    const name = text(col(row, ['예매자 성함(*)', '예매자 성함']));
    const count = num(col(row, ['예매 인원(*)', '예매 인원']));
    const dateRaw = text(col(row, ['예매 일자(*)', '예매 일자']));
    const friendRaw = text(col(row, ['지인 선택', '지인']));
    const slots = dateRaw.split('|').map(parseSlot).filter(Boolean);
    const friends = friendRaw.split('|').map(text).filter(Boolean);
    return { id: index + 2, name, count, friends, slots };
  }).filter(r => r.name && r.count > 0 && r.slots.length && !EXCLUDED_NAMES.has(r.name));

  const responseTimes = rows.map(r => text(col(r, ['응답일시']))).filter(Boolean).sort();
  const latestResponse = responseTimes.length ? responseTimes[responseTimes.length - 1] : null;
  const fileMtime = fs.statSync(EXCEL_FILE).mtime.toISOString();
  return { reservations, latestResponse, fileMtime, sourceRows: rows.length };
}

app.disable('x-powered-by');
app.use((req,res,next)=>{ res.setHeader('X-Content-Type-Options','nosniff'); res.setHeader('Referrer-Policy','same-origin'); next(); });
app.use(express.static(path.join(__dirname, 'public'), { extensions:['html'] }));

app.get('/health', (req,res)=>res.json({ok:true}));

app.get('/api/status', (req,res)=>{
  try {
    const d = readExcel();
    res.json({ reservationCount:d.reservations.length, totalPeople:d.reservations.reduce((s,r)=>s+r.count,0), latestResponse:d.latestResponse, fileMtime:d.fileMtime });
  } catch(e) { res.status(500).json({error:e.message}); }
});

app.get('/api/search', (req,res)=>{
  try {
    const q = text(req.query.name);
    if (!q || q.length > 50) return res.status(400).json({error:'배우 이름을 입력해주세요.'});
    const d = readExcel();
    const matches = d.reservations.filter(r => r.friends.includes(q));
    matches.sort((a,b)=>{
      const A=a.slots[0]||{}, B=b.slots[0]||{};
      return (A.month-B.month)||(A.day-B.day)||(A.hour-B.hour)||(A.minute-B.minute)||a.name.localeCompare(b.name,'ko');
    });
    res.json({name:q,totalBookings:matches.length,totalPeople:matches.reduce((s,r)=>s+r.count,0),reservations:matches.map(r=>({id:r.id,name:r.name,count:r.count,slots:r.slots})),latestResponse:d.latestResponse,fileMtime:d.fileMtime});
  } catch(e) { res.status(500).json({error:e.message}); }
});

app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,'0.0.0.0',()=>console.log(`Yujak site listening on ${PORT}`));
