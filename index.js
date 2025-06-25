if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const fs = require('fs');
const tmp = require('tmp');

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
    'https://www.googleapis.com/auth/spreadsheets.readonly',
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
  // 🔧 ปรับข้อความให้เหมาะกับงานของคุณ
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 2})
ชื่องาน: ${row['ชื่องาน'] || '-'} | WBS: ${row['WBS'] || '-'}
ชำระเงิน/ลว: ${row['ชำระเงิน/ลว'] || '-'} | อนุมัติ/ลว.: ${row['อนุมัติ/ลว.'] || '-'} | รับแฟ้ม: ${row['รับแฟ้ม'] || '-'}
หม้อแปลง: ${row['หม้อแปลง'] || '-'} | ระยะทาง HT: ${row['ระยะทาง HT'] || '-'} | ระยะทาง LT: ${row['ระยะทาง LT'] || '-'}`
    .replace(/\n{2,}/g, '\n'); // ป้องกันเว้นหลายบรรทัด
}

async function sendMessage(roomId, message) {
  await axios.post(
    'https://webexapis.com/v1/messages',
    {
      roomId,
      text: message
    },
    {
      headers: {
        Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
      }
    }
  );
}

async function sendFile(roomId, filePath, message) {
  const form = new FormData();
  form.append('roomId', roomId);
  form.append('text', message);
  form.append('files', fs.createReadStream(filePath));

  await axios.post('https://webexapis.com/v1/messages', form, {
    headers: {
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
      ...form.getHeaders()
    }
  });
}

app.post('/webex', async (req, res) => {
  const { data } = req.body;

  if (data.personId === BOT_ID) return res.sendStatus(200);

  const messageRes = await axios.get(`https://webexapis.com/v1/messages/${data.id}`, {
    headers: {
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
    }
  });

  const text = messageRes.data.text.trim();
  const roomId = messageRes.data.roomId;

  // ✅ ฟีเจอร์: ค้นหา <ชื่อชีต>
  if (text.startsWith("ค้นหา ")) {
    const sheetName = text.replace("ค้นหา ", "").trim();

    const sheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
    if (!sheetNames.includes(sheetName)) {
      await sendMessage(roomId, `❌ ไม่พบชีตชื่อ "${sheetName}"`);
      return res.sendStatus(200);
    }

    const rows = await getSheetDataByName(sheetName);
    if (rows.length === 0) {
      await sendMessage(roomId, `📭 ไม่พบข้อมูลในชีต "${sheetName}"`);
      return res.sendStatus(200);
    }

    const messages = rows.map((row, i) => formatRow(row, sheetName, i)).join('\n\n');

    if (messages.length > 7000) {
      const tmpFile = tmp.fileSync({ postfix: '.txt' });
      fs.writeFileSync(tmpFile.name, messages, 'utf8');

      await sendFile(roomId, tmpFile.name, "📎 ข้อมูลยาวเกิน Webex รองรับ จึงแนบมาเป็นไฟล์แทน");

      tmpFile.removeCallback();
    } else {
      await sendMessage(roomId, messages);
    }

    return res.sendStatus(200);
  }

  // 👉 ฟีเจอร์อื่น เช่น help หรือค้นหาคำทั่ว ๆ ไป สามารถใส่เพิ่มตรงนี้

  res.sendStatus(200);
});

app.listen(PORT, () => {
  console.log(`Bot running on port ${PORT}`);
});