const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT || 3000);

const EXCEL_FILE = path.join(__dirname, 'data', 'reservations.xlsx');

// 필요하면 기존 제외 명단 유지
const EXCLUDED_NAMES = new Set([
  '김맹훈',
  '정해진'
]);

function text(value) {
  return value == null ? '' : String(value).trim();
}

function num(value) {
  const match = text(value)
    .replace(/,/g, '')
    .match(/\d+(?:\.\d+)?/);

  return match ? Number(match[0]) : 0;
}

function getColumn(row, names) {
  for (const name of names) {
    if (Object.prototype.hasOwnProperty.call(row, name)) {
      return row[name];
    }
  }

  return '';
}

function responseTimestamp(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getTime();
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.round((value - 25569) * 86400 * 1000);
  }

  const valueText = text(value);

  if (!valueText) {
    return NaN;
  }

  const timestamp = Date.parse(valueText);

  return Number.isNaN(timestamp)
    ? NaN
    : timestamp;
}

function responseDate(value) {
  const timestamp = responseTimestamp(value);

  if (Number.isNaN(timestamp)) {
    return null;
  }

  return new Date(timestamp).toISOString();
}

function parseSlot(raw) {
  const value = text(raw);

  /*
    예:
    9월 26일 (토) 13:00
    9월 26일 13:00
  */

  const match = value.match(
    /(\d{1,2})월\s*(\d{1,2})일\s*(?:\([^)]*\))?\s*(\d{1,2}):(\d{2})/
  );

  if (!match) {
    return null;
  }

  const month = Number(match[1]);
  const day = Number(match[2]);
  const hour = Number(match[3]);
  const minute = Number(match[4]);

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

  const reservations = rows
    .map((row, index) => {

      const name = text(
        getColumn(row, [
          '예매자 성함(*)',
          '예매자 성함'
        ])
      );

      const count = num(
        getColumn(row, [
          '예매 인원(*)',
          '예매 인원'
        ])
      );

      const dateRaw = text(
        getColumn(row, [
          '예매 일자(*)',
          '예매 일자'
        ])
      );

      const friendRaw = text(
        getColumn(row, [
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

        slots,

        hasFriendSelection:
          friends.length > 0
      };
    })
    .filter(reservation => {

      return (
        reservation.name &&
        reservation.count > 0 &&
        reservation.slots.length > 0 &&
        !EXCLUDED_NAMES.has(reservation.name)
      );
    });

  /*
    마지막 응답일시
  */

  const responseList = rows
    .map(row => {

      const raw = getColumn(row, [
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
    .sort((a, b) =>
      a.timestamp - b.timestamp
    );

  const latestResponse =
    responseList.length > 0
      ? responseList[responseList.length - 1].value
      : null;

  /*
    회차별 집계
  */

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
              `${slot.month}월 ${slot.day}일 ${slot.label}`,

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
      .map((episode, index) => {

        return {
          episode: index + 1,
          ...episode
        };
      });

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


/* --------------------------------------------------
   기본 설정
-------------------------------------------------- */

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


/* --------------------------------------------------
   정적 파일
-------------------------------------------------- */

app.use(
  express.static(
    path.join(__dirname, 'public'),
    {
      extensions: ['html']
    }
  )
);


/* --------------------------------------------------
   Health Check
-------------------------------------------------- */

app.get('/health', (req, res) => {

  res.json({
    ok: true
  });
});


/* --------------------------------------------------
   전체 현황
-------------------------------------------------- */

app.get('/api/status', (req, res) => {

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
});


/* --------------------------------------------------
   검색

   type=friend
   → 배우 이름 검색

   type=booker
   → 예매자 성함 검색
-------------------------------------------------- */

app.get('/api/search', (req, res) => {

  try {

    const query =
      text(req.query.name);

    const type =
      text(req.query.type) || 'friend';

    if (!query) {

      return res.status(400).json({
        error: '검색어를 입력해주세요.'
      });
    }

    if (query.length > 50) {

      return res.status(400).json({
        error: '검색어가 너무 깁니다.'
      });
    }

    const normalized =
      query.replace(/\s+/g, '');

    const data =
      readExcel();

    const matches =
      data.reservations.filter(
        reservation => {

          /*
            예매자 성함 검색
          */

          if (type === 'booker') {

            return reservation.name
              .replace(/\s+/g, '')
              .includes(normalized);
          }

          /*
            배우 이름 검색

            기존 지인 선택 데이터를 기준으로 검색
          */

          return reservation.friends.some(
            friend =>
              friend
                .replace(/\s+/g, '')
                .includes(normalized)
          );
        }
      );

    /*
      날짜순 정렬
    */

    matches.sort((a, b) => {

      const A = a.slots[0] || {};
      const B = b.slots[0] || {};

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

      name: query,

      type,

      totalBookings:
        matches.length,

      totalPeople,

      reservations:
        matches.map(reservation => ({
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
});


/* --------------------------------------------------
   지인 미선택 예매자

   지인 선택 값이 비어있는 사람만 표시
-------------------------------------------------- */

app.get('/api/other', (req, res) => {

  try {

    const data =
      readExcel();

    const reservations =
      data.reservations
        .filter(reservation => {

          return (
            !reservation.friends ||
            reservation.friends.length === 0
          );
        })
        .sort((a, b) => {

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
        })
        .map(reservation => ({

          id:
            reservation.id,

          name:
            reservation.name,

          count:
            reservation.count,

          slots:
            reservation.slots

        }));

    const totalPeople =
      reservations.reduce(
        (sum, reservation) =>
          sum + reservation.count,
        0
      );

    res.json({

      total:
        reservations.length,

      totalPeople,

      reservations,

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
});


/* --------------------------------------------------
   바닐라 비교

   대상 20명
-------------------------------------------------- */

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

        /*
          점심
          11:00 ~ 15:59
        */

        const lunchSlots =
          reservation.slots.filter(
            slot =>
              slot.hour >= 11 &&
              slot.hour < 16
          );


        /*
          저녁
          17:00 ~ 21:59
        */

        const dinnerSlots =
          reservation.slots.filter(
            slot =>
              slot.hour >= 17 &&
              slot.hour < 22
          );


        if (lunchSlots.length > 0) {

          result.lunch.bookings += 1;

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


        if (dinnerSlots.length > 0) {

          result.dinner.bookings += 1;

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
          lunchSlots.length > 0 &&
          dinnerSlots.length > 0
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
        error: error.message
      });
    }
  }
);


/* --------------------------------------------------
   SPA fallback
-------------------------------------------------- */

app.use((req, res, next) => {

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
});


/* --------------------------------------------------
   서버 실행
-------------------------------------------------- */

app.listen(
  PORT,
  '0.0.0.0',
  () => {

    console.log(
      `Yujak site listening on ${PORT}`
    );

  }
);
