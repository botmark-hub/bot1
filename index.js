// โหลด .env เฉพาะเวลา dev
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

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive.readonly']
});
const drive = google.drive({ version: 'v3', auth });

async function sendLongMessage(roomId, text) {
  const chunks = text.match(/([\s\S]{1,7000})(?:\n|$)/g);
  for (const chunk of chunks) {
    await axios.post('https://webexapis.com/v1/messages', {
      roomId,
      text: chunk
    }, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });
  }
}

async function downloadFile(fileId) {
  if (!fileId) throw new Error('❌ Missing required parameter: fileId');

  const tmpFile = tmp.fileSync({ postfix: '.xlsx' });
  const dest = fs.createWriteStream(tmpFile.name);

  const meta = await drive.files.get({ fileId, fields: 'mimeType' });
  const mimeType = meta.data.mimeType;

  let res;
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    res = await drive.files.export({
      fileId,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }, { responseType: 'stream' });
  } else {
    res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'stream' });
  }

  await new Promise((resolve, reject) => {
    res.data.pipe(dest).on('finish', resolve).on('error', reject);
  });

  return tmpFile.name;
}

async function searchAndReadFileByName(_, keyword, sheetName) {
  const filePath = await downloadFile(GOOGLE_SHEET_FILE_ID);
  const workbook = XLSX.readFile(filePath);

  const sheetNamesToSearch = sheetName ? [sheetName] : workbook.SheetNames;
  let allResults = [];

  for (const name of sheetNamesToSearch) {
    const sheet = workbook.Sheets[name];
    if (!sheet) continue;

    const sheetRange = XLSX.utils.decode_range(sheet['!ref']);
    const headers = [];
    let lastNonEmpty = '';
    for (let C = sheetRange.s.c; C <= sheetRange.e.c; ++C) {
      const cell1 = sheet[XLSX.utils.encode_cell({ r: 0, c: C })];
      const cell2 = sheet[XLSX.utils.encode_cell({ r: 1, c: C })];
      const val1 = (cell1?.v || '').toString().trim();
      const val2 = (cell2?.v || '').toString().trim();
      if (val1) lastNonEmpty = val1;
      const header = `${lastNonEmpty} ${val2}`.trim();
      headers.push(header);
    }

    const data = XLSX.utils.sheet_to_json(sheet, {
      header: headers,
      range: 2,
      defval: ''
    });

    const filtered = keyword
      ? data.filter(row => Object.values(row).some(val => val.toString().toLowerCase().includes(keyword.toLowerCase())))
      : data;

    if (!filtered.length) continue;

    const usedHeaders = headers.filter(h => filtered.some(row => row[h] !== ''));
    const tableHeader = usedHeaders.join(' | ');
    const tableRows = filtered.map((row, i) =>
      `${i + 1} | ` + usedHeaders.map(h => (row[h] || '').toString().replace(/\|/g, '｜').replace(/\n/g, ' ')).join(' | ')
    );

    const result = `📑 แผ่นงาน: ${name}\n\n${tableHeader}\n${'-'.repeat(tableHeader.length)}\n${tableRows.join('\n—\n')}`;
    allResults.push(result);
  }

  if (!allResults.length) {
    return `❌ ไม่พบข้อมูล${keyword ? `ที่มีคำว่า "${keyword}" ` : ''}ในไฟล์`;
  }

  return allResults.join('\n\n');
}