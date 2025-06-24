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

async function getSheetWithCombinedHeaders(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A1:Z2`
  });

  const rows = res.data.values;
  if (!rows || rows.length < 2) return [];

  const combinedHeaders = rows[0].map((h, i) =>
    `${(h || '').trim().replace(/\s+/g, ' ')} ${(rows[1][i] || '').trim().replace(/\s+/g, ' ')}`.trim()
  );

  const resData = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A3:Z`
  });
  const dataRows = resData.data.values || [];

  return dataRows.map(row => {
    const rowData = {};
    combinedHeaders.forEach((header, i) => {
      rowData[header] = row[i] || '';
    });
    return rowData;
  });
}

function formatRow(row, sheetName, index) {
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 3})\n` +
    ` ชื่องาน: ${row['ชื่องาน']} | WBS: ${row['WBS']}\n` +
    ` ชำระเงิน/ลว: ${row['ชำระเงิน/ลว']} | อนุมัติ/ลว.: ${row['อนุมัติ/ลว.']} | รับแฟ้ม: ${row['รับแฟ้ม']}\n` +
    ` หม้อแปลง: ${row['หม้อแปลง']} | ระยะทาง HT: ${row['ระยะทาง HT']} | ระยะทาง LT: ${row['ระยะทาง LT']}\n` +
    ` เสา 8 : ${row['เสา 8'] || '-'} | เสา 9 : ${row['เสา 9'] || '-'} | เสา 12 : ${row['เสา 12'] || '-'} | เสา 12.20 : ${row['เสา 12.20'] || '-'}\n` +
    ` พชง.ควบคุม: ${row['พชง.ควบคุม']}\n` +
    ` สถานะงาน: ${row['สถานะงาน']} | เปอร์เซ็นงาน: ${row['เปอร์เซ็นงาน']}\n` +
    ` หมายเหตุ: ${row['หมายเหตุ']}`;
}

app.post('/webex', async (req, res) => {
  try {
    const data = req.body.data;

    if (data.personId === BOT_ID) {
      return res.status(200).send('Ignore self-message');
    }

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
      responseText = `📌 คำสั่งที่ใช้ได้:\n` +
        `1. คำสั่ง ค้นหา <คำ> → ค้นหาคำในทุกแถว\n` +
        `2. คำสั่ง ค้นหา <ชื่อชีต> → แสดงข้อมูลทั้งหมด\n` +
        `3. คำสั่ง ค้นหา <ชื่อชีต> <ชื่อคอลัมน์> → แสดงเฉพาะคอลัมน์นั้น\n` +
        `4. คำสั่ง แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ> → แก้ไขข้อมูลในเซลล์\n` +
        `5. คำสั่ง help → แสดงวิธีใช้ทั้งหมด`;
    } else if (command === 'ค้นหา') {
      const keyword = args.join(' ');

      if (args.length === 2 && allSheetNames.includes(args[0])) {
        const data = await getSheetWithCombinedHeaders(sheets, GOOGLE_SHEET_FILE_ID, args[0]);
        responseText = data.map((row, idx) => `${args[1]}: ${row[args[1]]}`).join('\n');
      } else if (args.length === 1 && allSheetNames.includes(args[0])) {
        const data = await getSheetWithCombinedHeaders(sheets, GOOGLE_SHEET_FILE_ID, args[0]);
        responseText = data.map((row, idx) => formatRow(row, args[0], idx)).join('\n\n');
      } else {
        let results = [];
        for (const sheetName of allSheetNames) {
          const data = await getSheetWithCombinedHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
          data.forEach((row, idx) => {
            const match = Object.values(row).some(v => v.includes(keyword));
            if (match) results.push(formatRow(row, sheetName, idx));
          });
        }
        responseText = results.length ? results.join('\n\n') : '❌ ไม่พบข้อมูลที่ต้องการ';
      }
    } else if (command === 'แก้ไข') {
      if (args.length < 5) {
        responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง ควรเป็น:\nแก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
      } else {
        const sheetName = `${args[0]} ${args[1]}`;
        const columnName = args[2];
        const rowNumberStr = args[3];
        const valueParts = args.slice(4);
        const newValue = valueParts.join(' ');
        const rowNumber = parseInt(rowNumberStr);

        if (!allSheetNames.includes(sheetName)) {
          responseText = `❌ ไม่พบชีตชื่อ "${sheetName}"`;
        } else {
          const res = await sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEET_FILE_ID,
            range: `${sheetName}!A1:Z2`
          });
          const headers = res.data.values;
          if (!headers || headers.length < 2) {
            responseText = '❌ ไม่สามารถโหลด header ได้';
          } else {
            const combinedHeaders = headers[0].map((h, i) =>
              `${(h || '').trim()} ${(headers[1][i] || '').trim()}`.trim()
            );

            const columnIndex = combinedHeaders.findIndex(h =>
              h.toLowerCase() === columnName.toLowerCase() ||
              h.toLowerCase().endsWith(' ' + columnName.toLowerCase()) ||
              h.toLowerCase().includes(columnName.toLowerCase())
            );

            if (columnIndex === -1) {
              responseText = `❌ ไม่พบคอลัมน์ "${columnName}" ในชีต "${sheetName}"`;
            } else {
              const matchedHeader = combinedHeaders[columnIndex];
              const columnLetter = String.fromCharCode(65 + columnIndex);
              const targetCell = `${columnLetter}${rowNumber}`;
              await sheets.spreadsheets.values.update({
                spreadsheetId: GOOGLE_SHEET_FILE_ID,
                range: `${sheetName}!${targetCell}`,
                valueInputOption: 'USER_ENTERED',
                requestBody: { values: [[newValue]] }
              });
              responseText = `✅ แก้ไข ${sheetName}!${targetCell} (${matchedHeader}) เป็น "${newValue}" แล้ว`;
            }
          }
        }
      }
    } else {
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "bot help"';
    }

    await axios.post('https://webexapis.com/v1/messages', {
      roomId,
      text: responseText
    }, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });

    res.status(200).send('OK');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error');
  }
});

app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));