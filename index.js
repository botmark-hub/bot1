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
const WEBEX_BOT_NAME = 'bot_small';
const BOT_ID = (process.env.BOT_ID || '').trim();

const rawCreds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
rawCreds.private_key = rawCreds.private_key.replace(/\n/g, '\n');

const auth = new google.auth.GoogleAuth({
  credentials: rawCreds,
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
  ]
});

const sheets = google.sheets({ version: 'v4', auth });

function flattenText(text) {
  return (text || '').toString().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

function formatRow(row, sheetName, index) {
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 2})\n` +
    `📝 ชื่องาน: ${flattenText(row['ชื่องาน'])} | 🨾 WBS: ${flattenText(row['WBS'])}\n` +
    `💰 ชำระเงิน/\u0e25ว.: ${flattenText(row['ชำระเงิน/\u0e25ว.'])} | ✅ อนุมัติ/\u0e25ว.: ${flattenText(row['อนุมัติ/\u0e25ว.'])} | 📂 รับแฟ้ม: ${flattenText(row['รับแฟ้ม'])}\n` +
    `🔌 หม้อแปลง: ${flattenText(row['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${flattenText(row['ระยะทาง HT'])} | ⚡ ระยะทาง LT: ${flattenText(row['ระยะทาง LT'])}\n` +
    `🩵 เสา 8 : ${flattenText(row['เสา 8']) || '-'} | 🩵 เสา 9 : ${flattenText(row['เสา 9']) || '-'} | 🩵 เสา 12 : ${flattenText(row['เสา 12']) || '-'} | 🩵 เสา 12.20 : ${flattenText(row['เสา 12.20']) || '-'}\n` +
    `👷‍♂️ พชง.ควบคุม: ${flattenText(row['พชง.ควบคุม'])}\n` +
    `📌 สถานะงาน: ${flattenText(row['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(row['เปอร์เซ็นงาน'])}\n` +
    `🗒️ หมายเทส: ${flattenText(row['หมายเทส'])}`;
}

async function getAllSheetNames(spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets.map(sheet => sheet.properties.title);
}

async function getSheetWithHeaders(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A1:Z`
  });

  const rows = res.data.values;
  if (!rows || rows.length < 2) return [];

  const headers = rows[0].map(h => h.trim());
  const dataRows = rows.slice(1);

  return dataRows.map(row => {
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = row[i] || '';
    });
    return rowData;
  });
}

async function sendMessageInChunks(roomId, fullMessage) {
  const CHUNK_LIMIT = 7000;
  const lines = fullMessage.split('\n\n');
  let buffer = '';
  for (const line of lines) {
    if ((buffer + '\n\n' + line).length > CHUNK_LIMIT) {
      await axios.post('https://webexapis.com/v1/messages', {
        roomId,
        text: buffer
      }, {
        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
      });
      buffer = line;
    } else {
      buffer += (buffer ? '\n\n' : '') + line;
    }
  }
  if (buffer) {
    await axios.post('https://webexapis.com/v1/messages', {
      roomId,
      text: buffer
    }, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });
  }
}

app.post('/webex', async (req, res) => {
  try {
    const data = req.body.data;
    const personId = (data.personId || '').trim();
    if (personId === BOT_ID) return res.status(200).send('Ignore self-message');

    const messageId = data.id;
    const roomId = data.roomId;

    const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });
    let messageText = messageRes.data.text;
    if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME)) {
      messageText = messageText.substring(WEBEX_BOT_NAME.length).trim();
    }

    let responseText = '';
    const [command, ...args] = messageText.split(' ');
    const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);

    if (command === 'help') {
      responseText = `\ud83d\udccc \u0e04\u0e33\u0e2a\u0e31\u0e48\u0e07\u0e17\u0e35\u0e48\u0e43\u0e0a\u0e49\u0e44\u0e14\u0e49:\n` +
        `1. @bot_small \u0e04\u0e49\u0e19\u0e2b\u0e32 <\u0e04\u0e33> \u2192 \u0e04\u0e49\u0e19\u0e2b\u0e32\u0e17\u0e38\u0e01\u0e0a\u0e35\u0e15\u0e17\u0e35\u0e48\u0e21\u0e35\n` +
        `2. @bot_small \u0e04\u0e49\u0e19\u0e2b\u0e32 <\u0e0a\u0e37\u0e48\u0e2d\u0e0a\u0e35\u0e15> \u2192 \u0e41\u0e2a\u0e14\u0e07\u0e17\u0e38\u0e01\u0e23\u0e32\u0e22\u0e01\u0e32\u0e23\n` +
        `3. @bot_small \u0e04\u0e49\u0e19\u0e2b\u0e32 <\u0e0a\u0e37\u0e48\u0e2d\u0e0a\u0e35\u0e15> <\u0e04\u0e2d\u0e25\u0e31\u0e21\u0e19\u0e4c> \u2192 \u0e41\u0e2a\u0e14\u0e07\u0e40\u0e09\u0e1e\u0e32\u0e30\u0e04\u0e2d\u0e25\u0e31\u0e21\u0e19\u0e4c\n` +
        `4. @bot_small \u0e41\u0e01\u0e49\u0e44\u0e02 <\u0e0a\u0e37\u0e48\u0e2d\u0e0a\u0e35\u0e15> <\u0e04\u0e2d\u0e25\u0e31\u0e21\u0e19\u0e4c> <\u0e41\u0e16\u0e27> <\u0e02\u0e49\u0e2d\u0e04\u0e27\u0e32\u0e21> \u2192 \u0e41\u0e01\u0e49\u0e44\u0e02\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\n`;
    } else if (command === 'ค้นหา') {
      const keyword = args.join(' ').replace(/\s+/g, ' ').trim();
      const sheetNameFromArgs = keyword;
      if (args.length === 2 && allSheetNames.includes(args[0])) {
        const data = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, args[0]);
        responseText = data.map((row, idx) => `${args[1]}: ${flattenText(row[args[1]])}`).join('\n');
      } else if (allSheetNames.includes(sheetNameFromArgs)) {
        const data = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetNameFromArgs);
        responseText = data.length > 0
          ? data.map((row, idx) => formatRow(row, sheetNameFromArgs, idx)).join('\n\n')
          : `⚠️ ไม่พบข้อมูลในชีต "${sheetNameFromArgs}"`;
      } else {
        let results = [];
        for (const sheetName of allSheetNames) {
          const data = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
          data.forEach((row, idx) => {
            const match = Object.values(row).some(v => flattenText(v).includes(keyword));
            if (match) results.push(formatRow(row, sheetName, idx));
          });
        }
        responseText = results.length ? results.join('\n\n') : '❌ ไม่พบข้อมูลที่ต้องการ';
      }
    } else {
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "bot help"';
    }

    await sendMessageInChunks(roomId, responseText);
    res.status(200).send('OK');
  } catch (error) {
    console.error('❗ ERROR:', error?.stack || error?.message || error);
    res.status(500).send('Error');
  }
});

app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));