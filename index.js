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

function getCell(row, keyword) {
  const normalized = (text) => text.replace(/\s+/g, '').toLowerCase();
  const match = Object.keys(row).find(k => {
    const key = normalized(k);
    const numKeyword = normalized(keyword).replace('.', '');
    const numKey = key.replace('.', '');
    return numKey.endsWith(numKeyword);
  });
  return flattenText(row[match]) || '-';
}

function formatRow(row, sheetName, index) {
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 3})\n` +
    `📝 ชื่องาน: ${flattenText(row['ชื่องาน'])} | 🧾 WBS: ${flattenText(row['WBS'])}\n` +
    `💰 ชำระเงิน/ลว.: ${flattenText(row['ชำระเงิน/ลว.'])} | ✅ อนุมัติ/ลว.: ${flattenText(row['อนุมัติ/ลว.'])} | 📂 รับแฟ้ม: ${flattenText(row['รับแฟ้ม'])}\n` +
    `🔌 หม้อแปลง: ${flattenText(row['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${getCell(row, 'HT')} | ⚡ ระยะทาง LT: ${getCell(row, 'LT')}\n` +
    `🪵 เสา 8 : ${getCell(row, '8')} | 🪵 เสา 9 : ${getCell(row, '9')} | 🪵 เสา 12 : ${getCell(row, '12')} | 🪵 เสา 12.20 : ${getCell(row, '12.20')}\n` +
    `👷‍♂️ พชง.ควบคุม: ${flattenText(row['พชง.ควบคุม'])}\n` +
    `📌 สถานะงาน: ${flattenText(row['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(row['เปอร์เซ็นงาน'])}\n` +
    `🗒️ หมายเหตุ: ${flattenText(row['หมายเหตุ'])}`;
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
  if (!rows || rows.length < 3) return [];

  const headerRow1 = rows[0];
  const headerRow2 = rows[1];

  const headers = headerRow1.map((h1, i) => {
    const h2 = headerRow2[i] || '';
    return h2 ? `${h1} ${h2}`.trim() : h1.trim();
  });

  const dataRows = rows.slice(2);

  return dataRows.map(row => {
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = row[i] || '';
    });
    return rowData;
  });
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
  } catch (err) {
    console.error('❌ ส่งไฟล์แนบล้มเหลว:', err.response?.data || err.message);
  } finally {
    fs.unlinkSync(tempFilePath);
  }
}

app.post('/webex', async (req, res) => {
  console.log('📥 Webhook trigger');
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
    const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
    let responseText = '';

    if (command === 'help') {
      responseText = '📌 คำสั่งที่ใช้ได้:\n' +
        '1. ค้นหา <คำ> → ค้นหาทั้งหมด\n' +
        '2. ค้นหา <ชื่อชีต> → ส่งข้อมูลทั้งชีตเป็นไฟล์ .txt\n' +
        '3. แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ> → แก้ไขค่า\n' +
        '4. help → แสดงวิธีใช้';
    } else if (command === 'ค้นหา') {
      const keyword = args.join(' ').trim();

      if (allSheetNames.includes(keyword)) {
        const data = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);
        const content = data.map((row, idx) => formatRow(row, keyword, idx)).join('\n\n');
        if (content.length > 0) {
          await axios.post('https://webexapis.com/v1/messages', {
            roomId,
            markdown: '📎 แนบไฟล์ข้อมูลจากชีต: ' + keyword
          }, {
            headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
          });
          await sendFileAttachment(roomId, `${keyword}.txt`, content);
          return res.status(200).send('OK');
        } else {
          responseText = `⚠️ ไม่พบข้อมูลในชีต "${keyword}"`;
        }
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
    } else if (command === 'แก้ไข') {
      if (args.length < 4) {
        responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
      } else {
        const [sheetName, columnName, rowNumberStr, ...valueParts] = args;
        const newValue = valueParts.join(' ');
        const rowNumber = parseInt(rowNumberStr);

        if (!allSheetNames.includes(sheetName)) {
          responseText = `❌ ไม่พบชีตชื่อ "${sheetName}"`;
        } else {
          const headerRes = await sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEET_FILE_ID,
            range: `${sheetName}!A1:Z2`
          });
          const headers = headerRes.data.values?.[1] || [];
          const headerList = headers.map(h => h.trim());

          const columnIndex = headerList.findIndex(h =>
            h.toLowerCase() === columnName.toLowerCase() ||
            h.toLowerCase().includes(columnName.toLowerCase())
          );

          if (columnIndex === -1) {
            responseText = `❌ ไม่พบคอลัมน์ "${columnName}"`;
          } else {
            const columnLetter = String.fromCharCode(65 + columnIndex);
            const targetCell = `${columnLetter}${rowNumber}`;
            await sheets.spreadsheets.values.update({
              spreadsheetId: GOOGLE_SHEET_FILE_ID,
              range: `${sheetName}!${targetCell}`,
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [[newValue]] }
            });
            responseText = `✅ แก้ไข ${sheetName}!${targetCell} เป็น "${newValue}" แล้ว`;
          }
        }
      }
    } else {
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "help"';
    }

    await sendMessageInChunks(roomId, responseText);
    res.status(200).send('OK');
  } catch (err) {
    console.error('❗ ERROR:', err?.stack || err?.message || err);
    res.status(500).send('Internal error');
  }
});

app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));