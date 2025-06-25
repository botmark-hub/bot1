if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const { google } = require('googleapis');
const fs = require('fs');
const FormData = require('form-data');

const app = express();
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN;
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID;
const WEBEX_BOT_NAME = 'bot_small';
const BOT_ID = (process.env.BOT_ID || '').trim();

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

function flattenText(text) {
  return (text || '').toString().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

function formatRow(row, sheetName, index) {
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 2})\n` +
    `📝 ชื่องาน: ${flattenText(row['ชื่องาน'])} | 🧾 WBS: ${flattenText(row['WBS'])}\n` +
    `💰 ชำระเงิน/ลว.: ${flattenText(row['ชำระเงิน/ลว'])} | ✅ อนุมัติ/ลว.: ${flattenText(row['อนุมัติ/ลว.'])} | 📂 รับแฟ้ม: ${flattenText(row['รับแฟ้ม'])}\n` +
    `🔌 หม้อแปลง: ${flattenText(row['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${flattenText(row['ระยะทาง HT'])} | ⚡ ระยะทาง LT: ${flattenText(row['ระยะทาง LT'])}\n` +
    `🩵 เสา 8 : ${flattenText(row['เสา 8']) || '-'} | 🩵 เสา 9 : ${flattenText(row['เสา 9']) || '-'} | 🩵 เสา 12 : ${flattenText(row['เสา 12']) || '-'} | 🩵 เสา 12.20 : ${flattenText(row['เสา 12.20']) || '-'}\n` +
    `👷‍♂️ พชง.ควบคุม: ${flattenText(row['พชง.ควบคุม'])}\n` +
    `📌 สถานะงาน: ${flattenText(row['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(row['เปอร์เซ็นงาน'])}\n` +
    `📜 หมายเหตุ: ${flattenText(row['หมายเหตุ'])}`;
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

async function sendMessageInChunks(roomId, message) {
  const CHUNK_LIMIT = 6000;
  for (let i = 0; i < message.length; i += CHUNK_LIMIT) {
    const chunk = message.substring(i, i + CHUNK_LIMIT);
    try {
      await axios.post('https://webexapis.com/v1/messages', {
        roomId,
        markdown: chunk
      }, {
        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
      });
    } catch (err) {
      console.error('❌ Webex error:', err.response?.data || err.message);
    }
  }
}

async function sendFileAttachment(roomId, filename, content) {
  const tempFilePath = `/tmp/${filename}`;
  fs.writeFileSync(tempFilePath, content, 'utf8');

  const form = new FormData();
  form.append('roomId', roomId);
  form.append('files', fs.createReadStream(tempFilePath));

  try {
    await axios.post('https://webexapis.com/v1/messages', form, {
      headers: {
        Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
        ...form.getHeaders()
      }
    });
    console.log('📎 ส่งไฟล์แนบเรียบร้อย');
  } catch (err) {
    console.error('❌ ส่งไฟล์แนบล้มเหลว:', err.response?.data || err.message);
  } finally {
    fs.unlinkSync(tempFilePath);
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
      responseText = '📌 คำสั่งที่ใช้ได้:\n' +
        '1. ค้นหา <คำ> → ค้นหาทั้งหมด\n' +
        '2. ค้นหา <ชื่อชีต> → แสดงทั้งหมดในชีต\n' +
        '3. ค้นหา <ชื่อชีต> <ชื่อคอลัมน์> → แสดงเฉพาะคอลัมน์\n' +
        '4. แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n' +
        '5. help → วิธีใช้ทั้งหมด';
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
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "help"';
    }

    if (responseText.length > 20000) {
      await axios.post('https://webexapis.com/v1/messages', {
        roomId,
        markdown: '📎 ข้อมูลยาวเกินกว่าที่ Webex รองรับ จึงแนบมาเป็นไฟล์แทน'
      }, {
        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
      });
      await sendFileAttachment(roomId, 'ข้อมูลจากบอท.txt', responseText);
    } else {
      await sendMessageInChunks(roomId, responseText);
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('❗ ERROR:', error?.stack || error?.message || error);
    res.status(500).send('Error');
  }
});

app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));