if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const fs = require('fs');
const tmp = require('tmp');

const app = express();
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN;
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID;
const WEBEX_BOT_NAME = 'bot_small';

// ✅ แปลง \\n → \n ใน private_key เพื่อให้ใช้กับ Render ได้
const rawCreds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
rawCreds.private_key = rawCreds.private_key.replace(/\\n/g, '\n');

const auth = new google.auth.GoogleAuth({
  credentials: rawCreds,
  scopes: ['https://www.googleapis.com/auth/drive.readonly']
});
const drive = google.drive({ version: 'v3', auth });

let BOT_PERSON_ID = '';

async function getBotPersonId() {
  const res = await axios.get('https://webexapis.com/v1/people/me', {
    headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
  });
  BOT_PERSON_ID = res.data.id;
  console.log('🤖 BOT_PERSON_ID:', BOT_PERSON_ID);
}

async function sendLongMessage({ roomId, toPersonId, text }) {
  const chunks = text.match(/([\s\S]{1,7000})(?:\n|$)/g);
  for (const chunk of chunks) {
    try {
      const payload = roomId
        ? { roomId, text: chunk }
        : { toPersonId, text: chunk };

      await axios.post('https://webexapis.com/v1/messages', payload, {
        headers: {
          Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
        }
      });
    } catch (err) {
      console.error('❌ ERROR ส่งข้อความ:', err.response?.data || err.message);
    }
  }
}

function formatDateTH(date) {
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

function generateDateVariants(dateStr) {
  const parts = dateStr.split('/');
  if (parts.length !== 3) return [];
  let [d, m, y] = parts;
  const year = parseInt(y);
  if (isNaN(year)) return [];

  d = d.padStart(2, '0');
  m = m.padStart(2, '0');

  const variants = [`${d}/${m}/${y}`];
  if (year > 2100) {
    variants.push(`${d}/${m}/${year - 543}`);
  } else if (year < 2100 && year < 2500) {
    variants.push(`${d}/${m}/${year + 543}`);
  }
  return variants;
}

async function searchInGoogleSheet(keyword, sheetName, options = { onlyDate: false, column: undefined }) {
  const res = await drive.files.export({
    fileId: GOOGLE_SHEET_FILE_ID,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }, { responseType: 'stream' });

  const tmpFile = tmp.fileSync({ postfix: '.xlsx' });
  const dest = fs.createWriteStream(tmpFile.name);

  await new Promise((resolve, reject) => {
    res.data.pipe(dest).on('finish', resolve).on('error', reject);
  });

  const workbook = XLSX.readFile(tmpFile.name);
  const sheetNamesToSearch = sheetName ? [sheetName] : workbook.SheetNames;
  let allResults = [];

  for (const name of sheetNamesToSearch) {
    const sheet = workbook.Sheets[name];
    if (!sheet || !sheet['!ref']) continue;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headers = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const top = sheet[XLSX.utils.encode_cell({ r: 0, c: C })]?.v || '';
      const bottom = sheet[XLSX.utils.encode_cell({ r: 1, c: C })]?.v || '';
      headers.push(`${top} ${bottom}`.trim());
    }

    const json = XLSX.utils.sheet_to_json(sheet, {
      defval: '-',
      cellDates: true,
      header: headers,
      range: 2
    });

    const filtered = keyword === '*' ? json : json.filter(row =>
      Object.entries(row).some(([key, val]) => {
        if (options.column && !key.includes(options.column)) return false;

        const variants = generateDateVariants(keyword);

        if (options.onlyDate) {
          if (val instanceof Date) return variants.includes(formatDateTH(val));
          if (typeof val === 'string') {
            const [d, m, y] = val.split('/');
            if (d && m && y) {
              const parsed = new Date(`${y}-${m}-${d}`);
              if (!isNaN(parsed)) return variants.includes(formatDateTH(parsed));
            }
          }
          if (typeof val === 'number') {
            const excelDate = new Date(Math.round((val - 25569) * 86400 * 1000));
            if (!isNaN(excelDate)) return variants.includes(formatDateTH(excelDate));
          }
          return false;
        }

        if (val instanceof Date) return variants.includes(formatDateTH(val));
        return String(val).toLowerCase().includes(keyword.toLowerCase());
      })
    );

    if (filtered.length > 0) {
      const resultText = [`📄 แผ่นงาน: ${name}\n`];
      for (const row of filtered) {
        const out = {
          งาน: '',
          WBS: '',
          อนุมัติ: '',
          ชำระ: '',
          รับแฟ้ม: '',
          ระยะทาง: { HT: '', LT: '' },
          เสา: [],
          ผู้ควบคุม: '',
          หมายเหตุ: ''
        };

        for (const [key, val] of Object.entries(row)) {
          let displayVal = val;

          if (val instanceof Date) displayVal = formatDateTH(val);
          else if (typeof val === 'number' && key.includes('ชำระ')) {
            displayVal = `฿${val.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;
          } else if (typeof val === 'number' && key.match(/(ลว|แฟ้ม|อนุมัติ|วันที่)/)) {
            const excelDate = new Date(Math.round((val - 25569) * 86400 * 1000));
            if (!isNaN(excelDate)) displayVal = formatDateTH(excelDate);
          }

          if (key.includes('ชื่องาน')) out.งาน = displayVal;
          else if (key.includes('WBS')) out.WBS = displayVal;
          else if (key.includes('อนุมัติ')) out.อนุมัติ = displayVal;
          else if (key.includes('ชำระ')) out.ชำระ = displayVal;
          else if (key.includes('รับแฟ้ม')) out.รับแฟ้ม = displayVal;
          else if (key.includes('HT')) out.ระยะทาง.HT = displayVal;
          else if (key.includes('LT')) out.ระยะทาง.LT = displayVal;
          else if (key.trim().startsWith('เสา') && displayVal !== '-') out.เสา.push(`[${key.trim()}:${displayVal}]`);
          else if (key.includes('ควบคุม')) out.ผู้ควบคุม = displayVal;
          else if (key.includes('หมายเหตุ')) out.หมายเหตุ = displayVal;
        }

        if (out.งาน.length > 100) out.งาน = out.งาน.slice(0, 100) + '...';

        resultText.push(
          `🔹 ชื่องาน: ${out.งาน}\n` +
          `🧾 WBS: ${out.WBS}\n` +
          `📅 อนุมัติ: ${out.อนุมัติ} | ชำระ: ${out.ชำระ} | รับแฟ้ม: ${out.รับแฟ้ม}\n` +
          `📏 ระยะ HT: ${out.ระยะทาง.HT} | LT: ${out.ระยะทาง.LT}` +
          (out.เสา.length ? ` | เสา: ${out.เสา.join(' ')}` : '') +
          (out.ผู้ควบคุม ? `\n👤 พชง.ควบคุม: ${out.ผู้ควบคุม}` : '') +
          (out.หมายเหตุ ? `\n📝 หมายเหตุ: ${out.หมายเหตุ}` : '') +
          `\n---`
        );
      }
      allResults.push(resultText.join('\n'));
    }
  }

  tmpFile.removeCallback();
  return allResults.length ? allResults.join('\n\n') : '❌ ไม่พบข้อมูลที่ค้นหา';
}

app.post('/webhook', async (req, res) => {
  console.log('✅ Webhook Triggered');
  const message = req.body.data;
  const roomId = message.roomId;
  const personId = message.personId;

  if (!BOT_PERSON_ID || personId === BOT_PERSON_ID) return res.sendStatus(200);

  try {
    const msgRes = await axios.get(`https://webexapis.com/v1/messages/${message.id}`, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });

    const mentionedPeople = msgRes.data.mentionedPeople || [];
    if (!mentionedPeople.includes(BOT_PERSON_ID)) {
      console.log('📭 ข้ามข้อความที่ไม่ mention bot');
      return res.sendStatus(200);
    }

    const text = msgRes.data.text.trim();
    const cleanedText = text.replace(WEBEX_BOT_NAME, '').trim();
    const parts = cleanedText.split(/\s+/);
    const command = parts[0]?.toLowerCase();

    if (command === 'ค้นหา') {
      const keyword = parts[1];
      const arg = parts.slice(2).join(' ').trim();
      const isDatePattern = /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(keyword);
      const options = isDatePattern ? { onlyDate: true } : {};

      if (keyword === '-') {
        const sheetName = arg;
        const result = await searchInGoogleSheet('*', sheetName);
        await sendLongMessage({ roomId, toPersonId: personId, text: result });
      } else {
        if (arg) options.column = arg;
        const result = await searchInGoogleSheet(keyword, undefined, options);
        await sendLongMessage({ roomId, toPersonId: personId, text: result });
      }
    } else if (command === 'help') {
      const helpText = '📝 คำสั่ง:\n' +
        'ค้นหา <คำค้นหา> [ชื่อแผ่นงาน หรือชื่อคอลัมน์]\n' +
        'ค้นหา - <ชื่อแผ่นงาน>\n' +
        'พิมพ์ help เพื่อดูคำสั่ง';
      await sendLongMessage({ roomId, toPersonId: personId, text: helpText });
    }

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ ERROR ใน webhook:', err.response?.data || err.message);
    res.sendStatus(500);
  }
});

app.listen(PORT, async () => {
  await getBotPersonId();
  console.log(`✅ Bot server running on port ${PORT}`);
});