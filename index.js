// index.js
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN;
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID;
const WEBEX_BOT_NAME = 'bot';
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
const drive = google.drive({ version: 'v3', auth });

async function getAllSheetNames(spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets.map(sheet => sheet.properties.title);
}

async function getSheetData(spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}`
  });
  const rows = res.data.values;
  if (!rows || rows.length === 0) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = (row[i] || '').replace(/\n/g, ' ').trim();
    });
    return rowData;
  });
}

function formatRow(row, sheetName, index) {
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 2})\n` +
    `📝 ชื่องาน: ${row['ชื่องาน'] || ''} | 🧾 WBS: ${row['WBS'] || ''}\n` +
    `💰 ชำระเงิน/ลว.: ${row['ชำระเงิน/ลว.'] || ''} | ✅ อนุมัติ/ลว.: ${row['อนุมัติ/ลว.'] || ''} | 📂 รับแฟ้ม: ${row['รับแฟ้ม'] || ''}\n` +
    `🔌 หม้อแปลง: ${row['หม้อแปลง'] || ''} | ⚡ ระยะทาง HT: ${row['ระยะทาง HT'] || ''} | ⚡ ระยะทาง LT: ${row['ระยะทาง LT'] || ''}\n` +
    `🪵 เสา 8 : ${row['เสา 8'] || '-'} | 🪵 เสา 9 : ${row['เสา 9'] || '-'} | 🪵 เสา 12 : ${row['เสา 12'] || '-'} | 🪵 เสา 12.20 : ${row['เสา 12.20'] || '-'}\n` +
    `👷‍♂️ พชง.ควบคุม: ${row['พชง.ควบคุม'] || ''}\n` +
    `📌 สถานะงาน: ${row['สถานะงาน'] || ''} | 📊 เปอร์เซ็นงาน: ${row['เปอร์เซ็นงาน'] || ''}\n` +
    `🗒️ หมายเหตุ: ${row['หมายเหตุ'] || ''}\n`;
}

async function sendMessageInChunks(roomId, text) {
  const MAX_LENGTH = 7000;
  let index = 0;
  while (index < text.length) {
    const chunk = text.substring(index, index + MAX_LENGTH);
    try {
      console.log(`✅ ส่งข้อความ chunk ขนาด: ${chunk.length} ตัวอักษร`);
      await axios.post('https://webexapis.com/v1/messages', {
        roomId,
        text: chunk
      }, {
        headers: {
          Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
        }
      });
    } catch (error) {
      console.error('❗ Error sending chunk:', error.response?.data || error.message);
    }
    index += MAX_LENGTH;
  }
}

app.post('/webex', async (req, res) => {
  const messageId = req.body.data.id;
  try {
    const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, {
      headers: {
        Authorization: `Bearer ${WEBEX_BOT_TOKEN}`
      }
    });

    const message = messageRes.data.text.trim();
    const roomId = messageRes.data.roomId;
    const personId = messageRes.data.personId;

    if (personId === BOT_ID) return res.sendStatus(200);

    const command = message.replace(/@?\b${WEBEX_BOT_NAME}\b/i, '').trim();

    if (command.toLowerCase() === 'help') {
      const helpText = `🧠 คำสั่งที่ใช้ได้:\n` +
        `• ค้นหา <คำ> → ค้นหาข้อมูลในทุกชีต\n` +
        `• ค้นหา <ชื่อชีต> → แสดงข้อมูลทั้งหมดในชีต\n` +
        `• ค้นหา <ชื่อชีต> <ชื่อคอลัมน์> → แสดงเฉพาะคอลัมน์ที่ต้องการ\n` +
        `• help → แสดงคำสั่งทั้งหมด`;
      await sendMessageInChunks(roomId, helpText);
    } else if (command.startsWith('ค้นหา')) {
      const args = command.split(' ');
      const keywords = args.slice(1);

      if (keywords.length === 0) {
        await sendMessageInChunks(roomId, '⚠️ กรุณาระบุคำค้นหาด้วย');
        return res.sendStatus(200);
      }

      const allSheets = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
      let foundText = '';

      for (const sheetName of allSheets) {
        const data = await getSheetData(GOOGLE_SHEET_FILE_ID, sheetName);

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const rowText = Object.values(row).join(' ');

          if (keywords.every(kw => rowText.includes(kw))) {
            foundText += formatRow(row, sheetName, i) + '\n';
          }
        }
      }

      if (!foundText) {
        foundText = '🔍 ไม่พบข้อมูลที่ค้นหา';
      }

      await sendMessageInChunks(roomId, foundText);
    } else {
      await sendMessageInChunks(roomId, '⚠️ ไม่รู้จักคำสั่ง พิมพ์ `help` เพื่อดูตัวอย่าง');
    }

    res.sendStatus(200);
  } catch (err) {
    console.error('❗ ERROR:', err);
    res.sendStatus(500);
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
});
