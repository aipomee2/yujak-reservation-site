const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'reservations.json');
const SEED_FILE = path.join(__dirname, 'seed-data', 'reservations.json');

if (process.env.NODE_ENV === 'production' && !ADMIN_PASSWORD) {
  console.error('ADMIN_PASSWORD is required in production.');
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) {
  if (fs.existsSync(SEED_FILE)) fs.copyFileSync(SEED_FILE, DATA_FILE);
  else fs.writeFileSync(DATA_FILE, JSON.stringify({ updatedAt: null, reservations: [] }, null, 2));
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /\.xlsx?$/i.test(file.originalname || '');
    cb(ok ? null : new Error('엑셀 파일(.xlsx 또는 .xls)만 업로드할 수 있습니다.'), ok);
  }
});

app.disable('x-powered-by');
app.use(express.json({ limit: '100kb' }));
app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Referrer-Policy', 'same-origin');
  res.setHeader('Cache-Control', 'no-store');
  next();
});
app.use(express.static(path.join(__dirname, 'public'), { extensions: ['html'] }));

function readData() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return { updatedAt: null, reservations: [] }; }
}
function writeData(data) {
  const tmp = `${DATA_FILE}.${crypto.randomUUID()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tmp, DATA_FILE);
}
function normalize(v) { return v == null ? '' : String(v).trim(); }
function numberFrom(v) {
  const m = normalize(v).replace(/,/g, '').match(/\d+/);
  return m ? Number(m[0]) : 0;
}

function parseSlot(raw) {
  const s = normalize(raw);
  // Example: 10월 11일 (일) 18:30 [전형진 ...]
  const m = s.match(/(\d{1,2})월\s*(\d{1,2})일\s*(?:\([^)]*\))?\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return {
    raw: s,
    month: Number(m[1]),
    day: Number(m[2]),
    hour: Number(m[3]),
    minute: Number(m[4]),
    key: `${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}:${m[4]}`,
    label: `${Number(m[3])}:${m[4]}`
  };
}

function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  if (!rows.length) throw new Error('엑셀에 데이터가 없습니다.');

  const required = ['예매자 성함(*)', '예매 인원(*)', '예매 일자(*)', '지인 선택'];
  const missing = required.filter(k => !(k in rows[0]));
  if (missing.length) throw new Error(`필수 열이 없습니다: ${missing.join(', ')}`);

  return rows.map((row, idx) => {
    const rawDates = normalize(row['예매 일자(*)']);
    const slots = rawDates.split('|').map(parseSlot).filter(Boolean);
    const friends = normalize(row['지인 선택']).split('|').map(normalize).filter(Boolean);
    return {
      id: idx + 2,
      name: normalize(row['예매자 성함(*)']),
      count: numberFrom(row['예매 인원(*)']),
      friends,
      slots
    };
  }).filter(r => r.name && r.slots.length && r.count > 0);
}

function adminOk(req) {
  return !!ADMIN_PASSWORD && crypto.timingSafeEqual(
    Buffer.from(String(req.headers['x-admin-password'] || '')),
    Buffer.from(ADMIN_PASSWORD)
  );
}

app.get('/api/status', (req, res) => {
  const data = readData();
  res.json({ updatedAt: data.updatedAt, reservationCount: data.reservations.length });
});

app.get('/api/search', (req, res) => {
  const q = normalize(req.query.name);
  if (!q || q.length > 50) return res.status(400).json({ error: '이름을 입력해주세요.' });
  const data = readData();
  const matches = data.reservations.filter(r => r.friends.includes(q));
  const reservations = matches.map(r => ({ id: r.id, name: r.name, count: r.count, slots: r.slots }));

  reservations.sort((a, b) => {
    const A = a.slots[0] || {}; const B = b.slots[0] || {};
    return (A.month ?? 99) - (B.month ?? 99) || (A.day ?? 99) - (B.day ?? 99) || (A.hour ?? 99) - (B.hour ?? 99) || (A.minute ?? 99) - (B.minute ?? 99) || a.name.localeCompare(b.name, 'ko');
  });

  res.json({
    name: q,
    totalBookings: matches.length,
    totalPeople: matches.reduce((sum, r) => sum + r.count, 0),
    reservations,
    updatedAt: data.updatedAt
  });
});

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!adminOk(req)) return res.status(401).json({ error: '관리자 비밀번호가 올바르지 않습니다.' });
  if (!req.file) return res.status(400).json({ error: '엑셀 파일을 선택해주세요.' });
  try {
    const reservations = parseExcel(req.file.buffer);
    if (!reservations.length) throw new Error('유효한 예매 데이터가 없습니다.');
    const updatedAt = new Date().toISOString();
    writeData({ updatedAt, reservations });
    res.json({ ok: true, updatedAt, reservationCount: reservations.length });
  } catch (e) {
    res.status(400).json({ error: e.message || '엑셀 처리에 실패했습니다.' });
  }
});

app.get('/health', (req, res) => res.json({ ok: true }));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, '0.0.0.0', () => console.log(`Yujak reservation site listening on ${PORT}`));
