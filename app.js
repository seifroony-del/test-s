const DAY_SECONDS = 86400;
const TICKET_SECONDS = 90;
const TICKET_DAY_FACTOR = TICKET_SECONDS / DAY_SECONDS;

function systemDayFractionFromTickets(ticketCount) {
  return Number(ticketCount || 0) * TICKET_DAY_FACTOR;
}

function systemSecondsFromTickets(ticketCount) {
  return Math.round(Number(ticketCount || 0) * TICKET_SECONDS);
}

function normaliseStatus(value) {
  return String(value ?? '').trim().toLowerCase();
}

function formatTime(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) {
    return '0:00:00';
  }

  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remaining = Math.floor(seconds % 60);

  return `${hours}:${String(minutes).padStart(2, '0')}:${String(remaining).padStart(2, '0')}`;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = event => {
      try {
        resolve(
          XLSX.read(
            new Uint8Array(event.target.result),
            {
              type: 'array',
              cellDates: true,
              raw: false
            }
          )
        );
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function readBundledStructure() {
  let response;

  try {
    response = await fetch(
      encodeURI('STR Loss.xlsx'),
      { cache: 'no-store' }
    );
  } catch (error) {
    throw new Error(
      'تعذر تحميل STR Loss.xlsx تلقائياً. إذا كنت تفتح الصفحة مباشرة من الملفات المحلية فشغّلها عبر localhost أو ارفع Structure يدويًا.'
    );
  }

  if (!response.ok) {
    throw new Error('لم يتم العثور على STR Loss.xlsx داخل جذر المستودع');
  }

  const buffer = await response.arrayBuffer();
  return XLSX.read(
    new Uint8Array(buffer),
    {
      type: 'array',
      cellDates: true,
      raw: false
    }
  );
}

function getOrdersFrom...