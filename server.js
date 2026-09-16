const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT || 3000);

const EXCEL_FILE = path.join(__dirname, 'data', 'reservations.xlsx');

// 검색 결과에서 제외할 이름
const EXCLUDED_NAMES = new Set([
  '김맹훈',
  '정해진'
]);

function text(v) {
  return v == null ? '' : String(v).trim();
}

function num(v) {
  const m = text(v)
    .replace(/,/g, '')
    .match(/\d+(?:\.\d+)?/);

  return m ? Number(m[0]) : 0;
}

// 엑셀 응답일시를 timestamp로 변환
function responseTimestamp(v) {
  // Date 객체
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return v.getTime();
  }

  // Excel serial date
  if (typeof v === 'number' && Number.isFinite(v)) {
    const ms = Math.round(
      (v - 25569) * 86400 * 1000
    );

    return Number.isFinite(ms) ? ms : NaN;
  }

  // 문자열
  const s = text(v);

  if (!s) {
    return NaN;
  }

  const t = Date.parse(s);

  return Number.isNaN(t) ? NaN : t;
}

function responseDate(v) {
  const ts = responseTimestamp(v);

  if (Number.isNaN(ts)) {
    return null;
  }

  return new Date(ts).toISOString();
}

// 엑셀 컬럼 찾기
function col(row, names) {
  for (const name of names) {
    if (
      Object.prototype.hasOwnProperty.call(row, name)
    ) {
      return row[name];
    }
  }

  return '';
}

// 예매 날짜/시간 파싱
function parseSlot(raw) {
  const s = text(raw);

  const m = s.match(
    /(\d{1,2})월\s*(\d{1,2})일\s*(?:\([^)]*\))?\s*(\d{1,2}):(\d{2})/
  );

  if (!m) {
    return null;
  }

  const month = Number(m[1]);
  const day = Number(m[2]);
  const hour = Number(m[3]);
  const minute = Number(m[4]);

  return {
    month,
    day,
    hour,
    minute,

    key:
      `${month}-${day}-${hour}:${String(minute).padStart(2, '0')}`,

    label:
      `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`
  };
}

// 엑셀 읽기
function readExcel() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(
      'data/reservations.xlsx 파일을 찾을 수 없습니다.'
    );
  }

  const workbook = XLSX.readFile(
    EXCEL_FILE,
    {
      cellDates: true
    }
  );

  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(
    worksheet,
    {
      defval: ''
    }
  );

  // =========================
  // 예매 데이터
  // =========================

  const reservations = rows
    .map((row, index) => {

      const name = text(
        col(row, [
          '예매자 성함(*)',
          '예매자 성함'
        ])
      );

      const count = num(
        col(row, [
          '예매 인원(*)',
          '예매 인원'
        ])
      );

      const dateRaw = text(
        col(row, [
          '예매 일자(*)',
          '예매 일자'
        ])
      );

      const friendRaw = text(
        col(row, [
          '지인 선택',
          '지인'
        ])
      );

      const slots = dateRaw
        .split('|')
        .map(parseSlot)
        .filter(Boolean);

      const friends = friendRaw
        .split('|')
        .map(text)
        .filter(Boolean);

      return {
        id: index + 2,
        name,
        count,
        friends,
        slots
      };
    })
    .filter(r =>
      r.name &&
      r.count > 0 &&
      r.slots.length &&
      !EXCLUDED_NAMES.has(r.name)
    );

  // =========================
  // 가장 최근 응답일시
  // =========================

  const parsedResponses = rows
    .map(row => {

      const raw = col(row, [
        '응답일시'
      ]);

      const timestamp =
        responseTimestamp(raw);

      if (Number.isNaN(timestamp)) {
        return null;
      }

      return {
        timestamp,
        value: responseDate(raw)
      };
    })
    .filter(Boolean)
    .sort(
      (a, b) =>
        a.timestamp - b.timestamp
    );

  const latestResponse =
    parsedResponses.length
      ? parsedResponses[
          parsedResponses.length - 1
        ].value
      : null;

  // =========================
  // 회차별 예매 현황
  // =========================

  const episodeMap = new Map();

  for (const reservation of reservations) {

    for (const slot of reservation.slots) {

      if (!episodeMap.has(slot.key)) {

        episodeMap.set(
          slot.key,
          {
            key: slot.key,

            month: slot.month,
            day: slot.day,

            hour: slot.hour,
            minute: slot.minute,

            label:
              `${slot.month}월 ${slot.day}일 ` +
              `${slot.label}`,

            people: 0,
            bookings: 0
          }
        );
      }

      const episode =
        episodeMap.get(slot.key);

      episode.people +=
        reservation.count;

      episode.bookings += 1;
    }
  }

  // 날짜/시간 순서대로 정렬
  const episodes =
    [...episodeMap.values()]
      .sort((a, b) => {

        return (
          (a.month - b.month) ||
          (a.day - b.day) ||
          (a.hour - b.hour) ||
          (a.minute - b.minute)
        );
      })
      .map((episode, index) => ({
        episode: index + 1,
        ...episode
      }));

  // 엑셀 파일 자체의 수정 시간
  const fileMtime =
    fs.statSync(EXCEL_FILE)
      .mtime
      .toISOString();

  return {
    reservations,
    latestResponse,
    fileMtime,
    sourceRows: rows.length,
    episodes
  };
}

// =========================
// 기본 설정
// =========================

app.disable('x-powered-by');

app.use((req, res, next) => {

  res.setHeader(
    'X-Content-Type-Options',
    'nosniff'
  );

  res.setHeader(
    'Referrer-Policy',
    'same-origin'
  );

  next();
});

// public 폴더
app.use(
  express.static(
    path.join(__dirname, 'public'),
    {
      extensions: ['html']
    }
  )
);

// =========================
// 서버 상태 확인
// =========================

app.get(
  '/health',
  (req, res) => {

    res.json({
      ok: true
    });
  }
);

// =========================
// 전체 예매 현황
// =========================

app.get(
  '/api/status',
  (req, res) => {

    try {

      const data = readExcel();

      const totalPeople =
        data.reservations.reduce(
          (sum, reservation) =>
            sum + reservation.count,
          0
        );

      res.json({

        reservationCount:
          data.reservations.length,

        totalPeople,

        latestResponse:
          data.latestResponse,

        fileMtime:
          data.fileMtime,

        episodes:
          data.episodes
      });

    } catch (error) {

      console.error(error);

      res.status(500).json({
        error: error.message
      });
    }
  }
);

// =========================
// 배우 이름 검색
// =========================

app.get(
  '/api/search',
  (req, res) => {

    try {

      const q =
        text(req.query.name);

      if (!q || q.length > 50) {

        return res.status(400).json({
          error:
            '배우 이름을 입력해주세요.'
        });
      }

      const data =
        readExcel();

      const matches =
        data.reservations
          .filter(reservation =>
            reservation.friends.includes(q)
          );

      // 날짜/시간 순 정렬
      matches.sort((a, b) => {

        const A =
          a.slots[0] || {};

        const B =
          b.slots[0] || {};

        return (
          (A.month - B.month) ||
          (A.day - B.day) ||
          (A.hour - B.hour) ||
          (A.minute - B.minute) ||
          a.name.localeCompare(
            b.name,
            'ko'
          )
        );
      });

      const totalPeople =
        matches.reduce(
          (sum, reservation) =>
            sum + reservation.count,
          0
        );

      res.json({

        name: q,

        totalBookings:
          matches.length,

        totalPeople,

        reservations:
          matches.map(reservation => ({
            id: reservation.id,
            name: reservation.name,
            count: reservation.count,
            slots: reservation.slots
          })),

        latestResponse:
          data.latestResponse,

        fileMtime:
          data.fileMtime
      });

    } catch (error) {

      console.error(error);

      res.status(500).json({
        error: error.message
      });
    }
  }
);

// =========================
// 마지막 fallback
// =========================
// Express 5에서도 문제가 생기지 않도록
// app.get('*') 대신 app.use 사용

app.use(
  (req, res, next) => {

    if (req.method !== 'GET') {
      return next();
    }

    res.sendFile(
      path.join(
        __dirname,
        'public',
        'index.html'
      )
    );
  }
);

// =========================
// 서버 시작
// =========================

app.listen(
  PORT,
  '0.0.0.0',
  () => {

    console.log(
      `Yujak site listening on ${PORT}`
    );
  }
);
