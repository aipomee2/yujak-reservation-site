const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT || 3000);

const EXCEL_FILE = path.join(
  __dirname,
  'data',
  'reservations.xlsx'
);

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

function responseTimestamp(v) {

  if (
    v instanceof Date &&
    !Number.isNaN(v.getTime())
  ) {
    return v.getTime();
  }

  if (
    typeof v === 'number' &&
    Number.isFinite(v)
  ) {
    const ms =
      Math.round(
        (v - 25569) *
        86400 *
        1000
      );

    return ms;
  }

  const s = text(v);

  if (!s) {
    return NaN;
  }

  const t = Date.parse(s);

  return Number.isNaN(t)
    ? NaN
    : t;
}

function responseDate(v) {

  const ts =
    responseTimestamp(v);

  if (Number.isNaN(ts)) {
    return null;
  }

  return new Date(ts).toISOString();
}

function col(row, names) {

  for (const name of names) {

    if (
      Object.prototype.hasOwnProperty.call(
        row,
        name
      )
    ) {
      return row[name];
    }
  }

  return '';
}

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
      `${month}-${day}-${hour}:${String(
        minute
      ).padStart(2, '0')}`,

    label:
      `${String(hour).padStart(2, '0')}:${String(
        minute
      ).padStart(2, '0')}`
  };
}

function readExcel() {

  if (!fs.existsSync(EXCEL_FILE)) {

    throw new Error(
      'data/reservations.xlsx 파일을 찾을 수 없습니다.'
    );
  }

  const workbook =
    XLSX.readFile(
      EXCEL_FILE,
      {
        cellDates: true
      }
    );

  const sheetName =
    workbook.SheetNames[0];

  const worksheet =
    workbook.Sheets[sheetName];

  const rows =
    XLSX.utils.sheet_to_json(
      worksheet,
      {
        defval: ''
      }
    );

  const reservations =
    rows
      .map((row, index) => {

        const name =
          text(
            col(row, [
              '예매자 성함(*)',
              '예매자 성함'
            ])
          );

        const count =
          num(
            col(row, [
              '예매 인원(*)',
              '예매 인원'
            ])
          );

        const dateRaw =
          text(
            col(row, [
              '예매 일자(*)',
              '예매 일자'
            ])
          );

        const friendRaw =
          text(
            col(row, [
              '지인 선택',
              '지인'
            ])
          );

        const slots =
          dateRaw
            .split('|')
            .map(parseSlot)
            .filter(Boolean);

        const friends =
          friendRaw
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

  const parsedResponses =
    rows
      .map(row => {

        const raw =
          col(row, [
            '응답일시'
          ]);

        const timestamp =
          responseTimestamp(raw);

        if (Number.isNaN(timestamp)) {
          return null;
        }

        return {
          timestamp,
          value:
            responseDate(raw)
        };
      })
      .filter(Boolean)
      .sort(
        (a, b) =>
          a.timestamp -
          b.timestamp
      );

  const latestResponse =
    parsedResponses.length
      ? parsedResponses[
          parsedResponses.length - 1
        ].value
      : null;

  const episodeMap =
    new Map();

  for (
    const reservation
    of reservations
  ) {

    for (
      const slot
      of reservation.slots
    ) {

      if (
        !episodeMap.has(
          slot.key
        )
      ) {

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
        episodeMap.get(
          slot.key
        );

      episode.people +=
        reservation.count;

      episode.bookings += 1;
    }
  }

  const episodes =
    [...episodeMap.values()]
      .sort((a, b) =>
        (a.month - b.month) ||
        (a.day - b.day) ||
        (a.hour - b.hour) ||
        (a.minute - b.minute)
      )
      .map((episode, index) => ({
        episode: index + 1,
        ...episode
      }));

  const fileMtime =
    fs.statSync(EXCEL_FILE)
      .mtime
      .toISOString();

  return {

    reservations,

    latestResponse,

    fileMtime,

    sourceRows:
      rows.length,

    episodes
  };
}

/* =========================
   기본 설정
========================= */

app.disable(
  'x-powered-by'
);

app.use(
  (req, res, next) => {

    res.setHeader(
      'X-Content-Type-Options',
      'nosniff'
    );

    res.setHeader(
      'Referrer-Policy',
      'same-origin'
    );

    next();
  }
);

app.use(
  express.static(
    path.join(
      __dirname,
      'public'
    )
  )
);

/* =========================
   Health
========================= */

app.get(
  '/health',
  (req, res) => {

    res.json({
      ok: true
    });
  }
);

/* =========================
   전체 현황
========================= */

app.get(
  '/api/status',
  (req, res) => {

    try {

      const data =
        readExcel();

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

/* =========================
   검색

   type=friend
   → 지인 성함 검색

   type=booker
   → 예매자 성함 검색
========================= */

app.get(
  '/api/search',
  (req, res) => {

    try {

      const q =
        text(req.query.name);

      const type =
        text(req.query.type) ||
        'friend';

      if (!q) {

        return res.status(400).json({
          error:
            '검색어를 입력해주세요.'
        });
      }

      if (q.length > 50) {

        return res.status(400).json({
          error:
            '검색어가 너무 깁니다.'
        });
      }

      const data =
        readExcel();

      let matches;

      /* =========================
         지인 이름 검색
      ========================= */

      if (type === 'friend') {

        matches =
          data.reservations
            .filter(
              reservation =>
                reservation.friends.includes(q)
            );

      }

      /* =========================
         예매자 이름 검색
      ========================= */

      else if (type === 'booker') {

        matches =
          data.reservations
            .filter(
              reservation =>
                reservation.name === q
            );

      }

      else {

        return res.status(400).json({
          error:
            '잘못된 검색 방식입니다.'
        });
      }

      matches.sort(
        (a, b) => {

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
        }
      );

      const totalPeople =
        matches.reduce(
          (sum, reservation) =>
            sum + reservation.count,
          0
        );

      res.json({

        name: q,

        type,

        totalBookings:
          matches.length,

        totalPeople,

        reservations:
          matches.map(
            reservation => ({

              id:
                reservation.id,

              name:
                reservation.name,

              count:
                reservation.count,

              friends:
                reservation.friends,

              slots:
                reservation.slots

            })
          ),

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

/* =========================
   점심 / 저녁 비교
========================= */

app.get(
  '/api/compare-performances',
  (req, res) => {

    try {

      const data =
        readExcel();

      const targetNames =
        new Set([

          '이희성',
          '이유민',
          '서태원',
          '유정수',
          '권오성',
          '김현지',
          '이종현',
          '정우영',
          '이예림',
          '윤강보',
          '김명훈',
          '정해령',
          '이슬아',
          '김여림',
          '서채림',
          '김준희',
          '강태희',
          '박경희',
          '김재현',
          '백경환'

        ]);

      const result = {

        lunch: {
          bookings: 0,
          people: 0,
          names: []
        },

        dinner: {
          bookings: 0,
          people: 0,
          names: []
        },

        both: []
      };

      for (
        const reservation
        of data.reservations
      ) {

        if (
          !targetNames.has(
            reservation.name
          )
        ) {
          continue;
        }

        const lunchSlots =
          reservation.slots.filter(
            slot =>
              slot.hour >= 11 &&
              slot.hour < 16
          );

        const dinnerSlots =
          reservation.slots.filter(
            slot =>
              slot.hour >= 17 &&
              slot.hour < 22
          );

        if (
          lunchSlots.length
        ) {

          result.lunch.bookings++;

          result.lunch.people +=
            reservation.count;

          result.lunch.names.push({

            name:
              reservation.name,

            count:
              reservation.count,

            slots:
              lunchSlots

          });
        }

        if (
          dinnerSlots.length
        ) {

          result.dinner.bookings++;

          result.dinner.people +=
            reservation.count;

          result.dinner.names.push({

            name:
              reservation.name,

            count:
              reservation.count,

            slots:
              dinnerSlots

          });
        }

        if (
          lunchSlots.length &&
          dinnerSlots.length
        ) {

          result.both.push(
            reservation.name
          );
        }
      }

      res.json(result);

    } catch (error) {

      console.error(error);

      res.status(500).json({
        error:
          error.message
      });
    }
  }
);

/* =========================
   SPA fallback
========================= */

app.use(
  (req, res, next) => {

    if (
      req.method !== 'GET'
    ) {
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

/* =========================
   서버 실행
========================= */

app.listen(
  PORT,
  '0.0.0.0',
  () => {

    console.log(
      `Yujak site listening on ${PORT}`
    );

  }
);
