// index.js (เวอร์ชันรวมคำสั่งทั้งหมดที่คุณต้องการ)

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const fs = require('fs');
const tmp = require('tmp');
const FormData = require('form-data');

const app = express();
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN;
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID;
const BOT_ID = process.env.BOT_ID;

const rawCreds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
rawCreds.private_key = rawCreds.private_key.replace(/\\n/g, '\n');

const auth = new google.auth.GoogleAuth({
  credentials: rawCreds,
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
  ]
});

const sheets = google.sheets({ version: 'v4', auth });

async function getAllSheetNames(spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets.map(sheet => sheet.properties.title);
}

async function getSheetDataByName(sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_FILE_ID,
    range: `${sheetName}`
  });
  const [header, ...rows] = res.data.values;
  return rows.map(row => {
    const obj = {};
    header.forEach((h, i) => {
      obj[h] = row[i] || '';
    });
    return obj;
  });
}

function formatRow(row, sheetName, index) {
  return `📄 แถว ${index + 2} | ชื่องาน: ${row['ชื่องาน'] || '-'} | WBS: ${row['WBS'] || '-'}\n` +
         `ชำระเงิน/ลว: ${row['ชำระเงิน/ลว'] || '-'} | อนุมัติ/ลว.: ${row['อนุมัติ/ลว.'] || '-'} | รับแฟ้ม: ${row['รับแฟ้ม'] || '-'}`;
}

async function sendMessage(roomId, message) {
  await axios.post('https://webexapis.com/v1/messages', {
    roomId,
    text: message
  }, {
    headers: {
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
    }
  });
}

async function sendFile(roomId, filePath, caption) {
  const form = new FormData();
  form.append('roomId', roomId);
  form.append('text', caption);
  form.append('files', fs.createReadStream(filePath));

  await axios.post('https://webexapis.com/v1/messages', form, {
    headers: {
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
      ...form.getHeaders()
    }
  });
}

async function updateCell(sheetName, columnName, rowIndex, newValue) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_FILE_ID,
    range: `${sheetName}`
  });

  const [headers] = res.data.values;
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) return false;

  const cell = String.fromCharCode(65 + colIndex) + (parseInt(rowIndex) + 1);

  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_FILE_ID,
    range: `${sheetName}!${cell}`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [[newValue]]
    }
  });
  return true;
}

app.post('/webex', async (req, res) => {
  const { data } = req.body;
  if (data.personId === BOT_ID) return res.sendStatus(200);

  const msg = await axios.get(`https://webexapis.com/v1/messages/${data.id}`, {
    headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
  });

  const text = msg.data.text.trim();
  const roomId = msg.data.roomId;

  if (text === 'help') {
    const help = `📘 คำสั่งที่ใช้ได้:
1. 🔍 ค้นหา <คำ> → ค้นหาคำในทุกชีต
2. 📄 ค้นหา <ชื่อชีต> → แสดงข้อมูลทั้งหมด ถ้าเกิน 7000 ตัวอักษรจะส่งเป็น .txt
3. ✏️ แก้ไข <ชื่อชีต> <คอลัมน์> <แถวที่> <ข้อความ> → แก้ข้อมูลใน cell
4. ℹ️ help → แสดงวิธีใช้`;
    await sendMessage(roomId, help);
    return res.sendStatus(200);
  }

  if (text.startsWith('ค้นหา ')) {
    const args = text.split(' ');
    const sheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);

    // คำสั่ง: ค้นหา <ชื่อชีต>
    if (sheetNames.includes(args[1]) && args.length === 2) {
      const rows = await getSheetDataByName(args[1]);
      const output = rows.map((r, i) => formatRow(r, args[1], i)).join('\n\n');

      if (output.length > 7000) {
        const tmpFile = tmp.fileSync({ postfix: '.txt' });
        fs.writeFileSync(tmpFile.name, output, 'utf8');
        await sendFile(roomId, tmpFile.name, '📎 ข้อมูลยาวเกิน Webex รองรับ จึงแนบเป็นไฟล์');
        tmpFile.removeCallback();
      } else {
        await sendMessage(roomId, output);
      }
      return res.sendStatus(200);
    }

    // คำสั่ง: ค้นหา <คำ>
    const keyword = text.replace('ค้นหา ', '').trim();
    for (const sheetName of sheetNames) {
      const rows = await getSheetDataByName(sheetName);
      const matched = rows.filter(row => Object.values(row).some(v => v.includes(keyword)));
      if (matched.length > 0) {
        const result = matched.map((r, i) => formatRow(r, sheetName, i)).join('\n\n');
        await sendMessage(roomId, result.substring(0, 7000));
        return res.sendStatus(200);
      }
    }
    await sendMessage(roomId, `❌ ไม่พบข้อมูลที่ตรงกับ "${keyword}"`);
    return res.sendStatus(200);
  }

  if (text.startsWith('แก้ไข ')) {
    const parts = text.split(' ');
    if (parts.length < 5) {
      await sendMessage(roomId, '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <คอลัมน์> <แถวที่> <ข้อความ>');
      return res.sendStatus(200);
    }
    const [_, sheetName, colName, rowIndex, ...rest] = parts;
    const newValue = rest.join(' ');
    const success = await updateCell(sheetName, colName, rowIndex, newValue);
    if (success) {
      await sendMessage(roomId, `✅ แก้ไขข้อมูลแถว ${rowIndex} ในคอลัมน์ "${colName}" ของชีต "${sheetName}" แล้ว`);
    } else {
      await sendMessage(roomId, `❌ ไม่พบคอลัมน์ "${colName}" ในชีต "${sheetName}"`);
    }
    return res.sendStatus(200);
  }

  res.sendStatus(200);
});

app.listen(PORT, () => console.log(`🚀 Bot running on port ${PORT}`));