// === เริ่มต้นการตั้งค่า ===
// โหลดค่าตัวแปรจากไฟล์ .env ถ้าไม่ใช่โหมด production
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

// เรียกใช้ library ที่จำเป็น
const express = require('express'); // ใช้สร้างเว็บเซิร์ฟเวอร์
const bodyParser = require('body-parser'); // ช่วยแปลงข้อมูลที่ส่งเข้ามาให้เป็น JSON
const axios = require('axios'); // ใช้ส่ง HTTP request ไปยัง Webex API
const fs = require('fs'); // ใช้จัดการไฟล์ เช่น สร้าง/ลบไฟล์
const FormData = require('form-data'); // ใช้ส่งไฟล์แนบแบบฟอร์ม
const { google } = require('googleapis'); // ใช้เชื่อมต่อกับ Google Sheets / Google Drive

// สร้างเว็บเซิร์ฟเวอร์
const app = express();
app.use(bodyParser.json()); // ให้ express อ่านข้อมูลแบบ JSON ได้

// ตั้งค่าตัวแปรสำคัญ
const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN; // token สำหรับเข้าถึง Webex
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID; // ID ของ Google Sheet
const WEBEX_BOT_NAME = 'bot_small'; // ชื่อเรียกบอทใน Webex
const BOT_ID = (process.env.BOT_ID || '').trim(); // รหัสของบอท (เพื่อเช็คว่าใครส่งข้อความ)

// ตั้งค่าการเชื่อมต่อ Google API ด้วยบัญชีบริการ (Service Account)
const rawCreds = JSON.parse(process.env.GOOGLE_CREDENTIALS); // โหลดข้อมูล JSON ที่เป็น Key ของ Google
rawCreds.private_key = rawCreds.private_key.replace(/\\n/g, '\n'); // แก้ format ของ private key

const auth = new google.auth.GoogleAuth({
  credentials: rawCreds,
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets', // สิทธิ์อ่าน/เขียน Google Sheets
    'https://www.googleapis.com/auth/drive.readonly' // สิทธิ์อ่าน Google Drive
  ]
});

// สร้างตัวแปรสำหรับใช้งาน Google Sheets API
const sheets = google.sheets({ version: 'v4', auth });

// === ฟังก์ชันจัดการข้อความและข้อมูล ===

// ลบช่องว่างซ้ำๆ และขึ้นบรรทัดใหม่
function flattenText(text) {
  return (text || '').toString().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

// ค้นหาค่าตามหัวตาราง row ที่ 2 โดยดูว่า keyword อยู่ใน header ตรงไหน แล้วดึงค่าจากแถวนั้นมา
function getCellByHeader2(rowArray, headerRow2, keyword) {
  const idx = headerRow2.findIndex(h =>
    h.trim().toLowerCase().includes(keyword.toLowerCase())
  );
  return idx !== -1 ? flattenText(rowArray[idx]) : '-';
}

// สร้างข้อความแสดงผล 1 แถวข้อมูลอย่างสวยงาม
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

// ดึงรายชื่อชีตทั้งหมดใน Google Sheets
async function getAllSheetNames(spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets.map(sheet => sheet.properties.title);
}

// ดึงข้อมูลทั้งชีต พร้อมชื่อหัวตาราง 2 แถวแรก
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

  const dataRows = rows.slice(2); // ข้ามหัวตาราง

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

// ส่งข้อความยาว แบ่งเป็นตอนๆ ไม่เกิน 7000 ตัวอักษร
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

// ส่งไฟล์แนบให้ผู้ใช้ทาง Webex
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

  fs.unlinkSync(filePath); // ลบไฟล์ออกหลังส่งเสร็จ
}

// === ส่วนหลัก รับข้อความจาก Webex แล้วตอบกลับ ===
app.post('/webex', async (req, res) => {
  try {
    const data = req.body.data;
    const personId = (data.personId || '').trim();
    if (personId === BOT_ID) return res.status(200).send('Ignore self-message'); // ถ้าเป็นข้อความจากตัวเอง ไม่ตอบกลับ

    const messageId = data.id;
    const roomId = data.roomId;

    // ดึงข้อความที่ผู้ใช้ส่งเข้ามา
    const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });

    let messageText = messageRes.data.text;
    if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME)) {
      messageText = messageText.substring(WEBEX_BOT_NAME.length).trim();
    }

    // แยกคำสั่งจากข้อความ
    const [command, ...args] = messageText.split(' ');
    const keyword = args.join(' ').trim();
    const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
    let responseText = '';

    // === ตรวจสอบคำสั่งที่ผู้ใช้ส่ง ===
    if (command === 'help') {
      responseText = `📌 คำสั่งที่ใช้ได้:\n` +
        `1. @bot_small ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
        `2. @bot_small ค้นหา <ชื่อชีต> → ดึงข้อมูลทั้งหมดในชีต (แนบไฟล์ถ้ายาว)\n` +
        `3. @bot_small แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n` +
        `4. @bot_small help → แสดงวิธีใช้ทั้งหมด`;
    } else if (command === 'ค้นหา') {
      // ถ้าพิมพ์ชื่อชีต จะดึงข้อมูลทั้งชีต
      if (allSheetNames.includes(keyword)) {
        const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);
        const resultText = data.map((row, idx) => formatRow(row, rawHeaders2, idx, keyword)).join('\n\n');
        if (resultText.length > 7000) {
          // ถ้ายาวเกินแนบเป็นไฟล์แทน
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
        // ถ้าพิมพ์คำค้นหาแบบทั่วไป จะค้นหาทุกชีต
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
      // การแก้ไขข้อมูล
      if (args.length < 4) {
        responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
      } else {
        // รองรับชื่อชีต 1 หรือ 2 คำ
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
          // ตรวจสอบว่า column มีจริงหรือไม่
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
      // ถ้าพิมพ์ไม่ถูกสั่ง
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "help"';
    }

    // ส่งข้อความตอบกลับ
    await sendMessageInChunks(roomId, responseText);
    res.status(200).send('OK');
  } catch (err) {
    console.error('❌ ERROR:', err.stack || err.message);
    res.status(500).send('Error');
  }
});

// === เริ่มให้เซิร์ฟเวอร์ทำงานที่พอร์ตที่กำหนด ===
app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));