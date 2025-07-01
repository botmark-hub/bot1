// === เริ่มต้นการตั้งค่า ===
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const fs = require('fs');
const FormData = require('form-data');
const { google } = require('googleapis');

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

// === ฟังก์ชันจัดการข้อความและข้อมูล ===
function flattenText(text) {
  return (text || '').toString().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

function getCellByHeader2(rowArray, headerRow2, keyword) {
  const idx = headerRow2.findIndex(h =>
    h.trim().toLowerCase().includes(keyword.toLowerCase())
  );
  return idx !== -1 ? flattenText(rowArray[idx]) : '-';
}

function formatRow(rowObj, headerRow2, index, sheetName) {
  const rowArray = Object.values(rowObj);
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 3})\n` +
    `📝 ชื่องาน: ${flattenText(rowObj['ชื่องาน'])} | 🧾 WBS: ${flattenText(rowObj['WBS'])}\n` +
    `💰 ชำระเงิน/ลว.: ${flattenText(rowObj['ชำระเงิน/ลว.'])} | ✅ อนุมัติ/ลว.: ${flattenText(rowObj['อนุมัติ/ลว.'])} | 📂 รับแฟ้ม: ${flattenText(rowObj['รับแฟ้ม'])}\n` +
    `🔌 หม้อแปลง: ${flattenText(rowObj['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${getCellByHeader2(rowArray, headerRow2, 'HT')} | ⚡ ระยะทาง LT: ${getCellByHeader2(rowArray, headerRow2, 'LT')}\n` +
    `🪵 เสา 8 : ${getCellByHeader2(rowArray, headerRow2, '8')} | 🪵 เสา 9 : ${getCellByHeader2(rowArray, headerRow2, '9')} | 🪵 เสา 12 : ${getCellByHeader2(rowArray, headerRow2, '12')} | 🪵 เสา 12.20 : ${getCellByHeader2(rowArray, headerRow2, '12.20')}\n` +
    `👷‍♂️ พชง.ควบคุม: ${flattenText(rowObj['พชง.ควบคุม'])}\n` +
    `📌 สถานะงาน: ${flattenText(rowObj['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(rowObj['เปอร์เซ็นงาน'])}\n` +
    `🗒️ หมายเหตุ: ${flattenText(rowObj['หมายเหตุ'])}`;
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
  if (!rows || rows.length < 3) return { data: [], rawHeaders2: [] };

  const headerRow1 = rows[0];
  const headerRow2 = rows[1];

  const headers = headerRow1.map((h1, i) => {
    const h2 = headerRow2[i] || '';
    return h2 ? `${h1} ${h2}`.trim() : h1.trim();
  });

  const dataRows = rows.slice(2);

  return {
    data: dataRows.map(row => {
      const rowData = {};
      headers.forEach((header, i) => {
        rowData[header] = row[i] || '';
      });
      return rowData;
    }),
    rawHeaders2: headerRow2
  };
}

async function sendMessageInChunks(roomId, message) {
  const CHUNK_LIMIT = 7000;
  for (let i = 0; i < message.length; i += CHUNK_LIMIT) {
    const chunk = message.substring(i, i + CHUNK_LIMIT);
    await axios.post('https://webexapis.com/v1/messages', {
      roomId,
      text: chunk
    }, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });
  }
}

async function sendFileAttachment(roomId, filename, content) {
  const filePath = `/tmp/${filename}`;
  fs.writeFileSync(filePath, content, 'utf8');

  const form = new FormData();
  form.append('roomId', roomId);
  form.append('files', fs.createReadStream(filePath));

  await axios.post('https://webexapis.com/v1/messages', form, {
    headers: {
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
      ...form.getHeaders()
    }
  });

  fs.unlinkSync(filePath);
}

// === Route หลักของ Webex Webhook ===
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

    const [command, ...args] = messageText.split(' ');
    const keyword = args.join(' ').trim();
    const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
    let responseText = '';

    if (command === 'help') {
      responseText = `📌 คำสั่งที่ใช้ได้:\n` +
        `1. @bot_small ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
        `2. @bot_small ค้นหา <ชื่อชีต> → ดึงข้อมูลทั้งหมดในชีต (แนบไฟล์ถ้ายาว)\n` +
        `3. @bot_small แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n` +
        `4. @bot_small help → แสดงวิธีใช้ทั้งหมด`;
    } else if (command === 'ค้นหา') {
      if (allSheetNames.includes(keyword)) {
        const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);
        const resultText = data.map((row, idx) => formatRow(row, rawHeaders2, idx, keyword)).join('\n\n');
        if (resultText.length > 7000) {
          await axios.post('https://webexapis.com/v1/messages', {
            roomId,
            markdown: '📎 ข้อมูลยาวเกิน แนบเป็นไฟล์แทน'
          }, {
            headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
          });
          await sendFileAttachment(roomId, 'ข้อมูล.txt', resultText);
          return res.status(200).send('sent file');
        } else {
          responseText = resultText;
        }
      } else {
        let results = [];
        for (const sheetName of allSheetNames) {
          const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
          data.forEach((row, idx) => {
            const match = Object.values(row).some(v => flattenText(v).includes(keyword));
            if (match) results.push(formatRow(row, rawHeaders2, idx, sheetName));
          });
        }
        responseText = results.length ? results.join('\n\n') : '❌ ไม่พบข้อมูลที่ต้องการ';
      }
    } else if (command === 'แก้ไข') {
      if (args.length < 4) {
        responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
      } else {
        // รองรับชื่อชีต 2 คำ หรือ 1 คำ
        let sheetName = '';
        let columnName = '';
        let rowIndex = 0;
        let newValue = '';

        let trySheetName = args[0] + ' ' + args[1];
        if (allSheetNames.includes(trySheetName)) {
          sheetName = trySheetName;
          columnName = args[2];
          rowIndex = parseInt(args[3]);
          newValue = args.slice(4).join(' ');
        } else if (allSheetNames.includes(args[0])) {
          sheetName = args[0];
          columnName = args[1];
          rowIndex = parseInt(args[2]);
          newValue = args.slice(3).join(' ');
        } else {
          responseText = `❌ ไม่พบชีตชื่อ "${args[0]}" หรือ "${trySheetName}"`;
        }

        if (sheetName) {
          const res = await sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEET_FILE_ID,
            range: `${sheetName}!A1:Z2`
          });
          const header1 = res.data.values[0];
          const header2 = res.data.values[1];
          const headers = header1.map((h1, i) => `${h1} ${header2[i] || ''}`.trim());
          const colIndex = headers.findIndex(h => h.includes(columnName));
          if (colIndex === -1) {
            responseText = `❌ ไม่พบคอลัมน์ "${columnName}"`;
          } else {
            const colLetter = String.fromCharCode(65 + colIndex);
            const range = `${sheetName}!${colLetter}${rowIndex}`;
            await sheets.spreadsheets.values.update({
              spreadsheetId: GOOGLE_SHEET_FILE_ID,
              range,
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [[newValue]] }
            });
            responseText = `✅ แก้ไขแล้ว: ${range} → ${newValue}`;
          }
        }
      }
    } else {
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "help"';
    }

    await sendMessageInChunks(roomId, responseText);
    res.status(200).send('OK');
  } catch (err) {
    console.error('❌ ERROR:', err.stack || err.message);
    res.status(500).send('Error');
  }
});

// === เริ่มทำงาน ===
app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));