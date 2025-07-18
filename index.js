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
  // กำหนดข้อมูลรับรอง (credentials) ที่จะใช้ในการยืนยันตัวตนกับ Google API
  credentials: rawCreds,

  // ระบุขอบเขต (scopes) ของสิทธิ์ที่ต้องการให้แอปนี้เข้าถึง
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    // 👉 ให้สิทธิ์อ่านและเขียนข้อมูลใน Google Sheets เช่น อ่านข้อมูล, แก้ไข, เพิ่มแถว ฯลฯ

    'https://www.googleapis.com/auth/drive.readonly'
    // 👉 ให้สิทธิ์ "อ่านอย่างเดียว" บน Google Drive เช่น ค้นหาไฟล์, อ่าน metadata ของไฟล์, อ่านเนื้อหาไฟล์ (ถ้ารูปแบบรองรับ)
  ]
});

/// เรียกใช้งาน Google Sheets API เวอร์ชัน 4 พร้อมกำหนดการยืนยันตัวตนด้วยตัวแปร 'auth'
const sheets = google.sheets({ version: 'v4', auth });

function flattenText(text) {
  // แปลงค่า text ให้เป็น string และกำหนดค่าเริ่มต้นเป็น '' หาก text เป็น null หรือ undefined
  return (text || '')
    .toString()                         // แปลงค่าให้เป็น string เผื่อว่า text เป็นตัวเลขหรือ object อื่นๆ
    .replace(/\n/g, ' ')                // แทนที่ทุกบรรทัดใหม่ (\n) ด้วยช่องว่าง (space)
    .replace(/\s+/g, ' ')               // แทนที่ช่องว่างทุกแบบ (space, tab, newline) ที่ต่อกันหลายตัว ด้วยช่องว่างเพียง 1 ตัว
    .trim();                            // ลบช่องว่างที่อยู่หน้าสุดและท้ายสุดของ string ออก
}

// ฟังก์ชัน getCellByHeader2 ใช้เพื่อดึงค่าจาก rowArray โดยค้นหาตำแหน่งของ header ที่มีคำว่า keyword อยู่ใน headerRow2
function getCellByHeader2(rowArray, headerRow2, keyword) {
  // ค้นหาดัชนีของ header ที่มีข้อความตรงกับ keyword (ไม่สนใจตัวพิมพ์เล็ก-ใหญ่ และตัดช่องว่างหน้าหลัง)
  const idx = headerRow2.findIndex(h =>
    h.trim().toLowerCase().includes(keyword.toLowerCase())
  );

  // ถ้าพบ index ที่ตรงกับ keyword (ไม่เป็น -1) ให้คืนค่าข้อมูลใน rowArray ที่ตำแหน่งนั้น (ผ่านการ flattenText)
  // ถ้าไม่พบ ให้คืนเครื่องหมาย '-'
  return idx !== -1 ? flattenText(rowArray[idx]) : '-';
}

// สร้างข้อความแสดงผล 1 แถวข้อมูลอย่างสวยงาม
function formatRow(rowObj, headerRow2, index, sheetName) {
  // แปลงค่า rowObj ให้เป็น array ของ value เพื่อใช้เข้าถึง cell ตามลำดับ index
  const rowArray = Object.values(rowObj);

  // เริ่มสร้างข้อความสรุปข้อมูลในรูปแบบ string โดยรวมข้อมูลหลักจาก rowObj และ headerRow2
  return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 3})\n` + // ระบุชื่อชีตและแถวที่พบข้อมูล (index + 3 เพราะเผื่อ header 2 แถวแรก)

    // แสดงชื่อของงาน และ WBS
    `📝 ชื่องาน: ${flattenText(rowObj['ชื่องาน'])} | 🧾 WBS: ${flattenText(rowObj['WBS'])}\n` +

    // แสดงข้อมูลเกี่ยวกับการชำระเงิน, วันที่อนุมัติ, การรับแฟ้ม
    `💰 ชำระเงิน/ลว.: ${flattenText(rowObj['ชำระเงิน/ลว.'])} | ✅ อนุมัติ/ลว.: ${flattenText(rowObj['อนุมัติ/ลว.'])} | 📂 รับแฟ้ม: ${flattenText(rowObj['รับแฟ้ม'])}\n` +

    // แสดงข้อมูลหม้อแปลง และระยะทางสายไฟ HT / LT โดยใช้ฟังก์ชัน getCellByHeader2 สำหรับ column ที่อาจอยู่ตำแหน่งไม่แน่นอน
    `🔌 หม้อแปลง: ${flattenText(rowObj['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${getCellByHeader2(rowArray, headerRow2, 'HT')} | ⚡ ระยะทาง LT: ${getCellByHeader2(rowArray, headerRow2, 'LT')}\n` +

    // แสดงจำนวนเสาแต่ละประเภท (8, 9, 12, 12.20 เมตร)
    `🪵 เสา 8 : ${getCellByHeader2(rowArray, headerRow2, '8')} | 🪵 เสา 9 : ${getCellByHeader2(rowArray, headerRow2, '9')} | 🪵 เสา 12 : ${getCellByHeader2(rowArray, headerRow2, '12')} | 🪵 เสา 12.20 : ${getCellByHeader2(rowArray, headerRow2, '12.20')}\n` +

    // แสดงชื่อผู้ควบคุมงาน
    `👷‍♂️ พชง.ควบคุม: ${flattenText(rowObj['พชง.ควบคุม'])}\n` +

    // แสดงสถานะของงาน และเปอร์เซ็นต์ความคืบหน้า
    `📌 สถานะงาน: ${flattenText(rowObj['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(rowObj['เปอร์เซ็นงาน'])}\n` +

    // แสดงหมายเหตุประกอบงาน (ถ้ามี)
    `🗒️ หมายเหตุ: ${flattenText(rowObj['หมายเหตุ'])}`;
}

// ฟังก์ชันแบบ async เพื่อดึงชื่อของทุกแผ่นงาน (sheet) ใน Google Spreadsheet
async function getAllSheetNames(spreadsheetId) {
  // เรียก Google Sheets API เพื่อดึงข้อมูลของ spreadsheet ตาม id ที่ส่งมา
  const res = await sheets.spreadsheets.get({ spreadsheetId });

  // จากข้อมูลที่ได้มา (res.data.sheets) ทำการ map เพื่อดึงชื่อของแต่ละแผ่นงาน (sheet title)
  return res.data.sheets.map(sheet => sheet.properties.title);
}


// ดึงข้อมูลทั้งชีต พร้อมชื่อหัวตาราง 2 แถวแรก
// ฟังก์ชัน async สำหรับดึงข้อมูลจาก Google Sheet โดยใช้ header 2 แถว (2 แถวแรก)
async function getSheetWithHeaders(sheets, spreadsheetId, sheetName) {
  // เรียกข้อมูลจาก Google Sheets API ในช่วง A1 ถึง Z (สามารถปรับขอบเขตตามจริงได้)
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,                    // รหัส spreadsheet ที่ต้องการดึงข้อมูล
    range: `${sheetName}!A1:Z`        // ขอบเขตข้อมูลจากชีตที่กำหนด เริ่มตั้งแต่ A1 ถึง Z (รวม header ทั้ง 2 แถว)
  });

  const rows = res.data.values;        // ข้อมูลแถวทั้งหมดในช่วงที่ดึงมา
  if (!rows || rows.length < 3)        // ตรวจสอบว่ามีข้อมูลอย่างน้อย 3 แถว (2 แถวเป็น header + 1 แถวข้อมูล)
    return { data: [], rawHeaders2: [] }; // ถ้าน้อยกว่า 3 แถว ให้คืนค่าเปล่า

  const headerRow1 = rows[0];          // แถวแรกเป็น header ชั้นที่ 1
  const headerRow2 = rows[1];          // แถวที่สองเป็น header ชั้นที่ 2

  // รวม header ทั้งสองแถวเข้าด้วยกัน เช่น "หมวด รายรับ"
  const headers = headerRow1.map((h1, i) => {
    const h2 = headerRow2[i] || '';    // ถ้า header ชั้นที่ 2 ไม่มีข้อมูล ให้ใช้ค่าว่าง
    return h2 ? `${h1} ${h2}`.trim() : h1.trim(); // รวม header ชั้น 1 และ 2 ถ้ามี
  });

  const dataRows = rows.slice(2);      // ตัดเอาเฉพาะแถวข้อมูล (ข้าม 2 แถวแรกที่เป็น header)

  return {
    // แปลงข้อมูลแต่ละแถวให้เป็น object โดยใช้ชื่อ header ที่รวมแล้วเป็น key
    data: dataRows.map(row => {
      const rowData = {};
      headers.forEach((header, i) => {
        rowData[header] = row[i] || ''; // ถ้าค่านั้นไม่มี (undefined) ให้ใส่เป็น string ว่าง
      });
      return rowData;
    }),
    rawHeaders2: headerRow2            // ส่งกลับ header ชั้นที่ 2 เผื่อใช้งานต่อ
  };
}

// ส่งข้อความยาว แบ่งเป็นตอนๆ ไม่เกิน 7000 ตัวอักษร
async function sendMessageInChunks(roomId, message) {
  // กำหนดจำนวนตัวอักษรสูงสุดในแต่ละข้อความที่ส่งได้ (limit ที่ 7000 ตัวอักษร)
  const CHUNK_LIMIT = 7000;

  // วนลูปแบ่งข้อความใหญ่เป็นส่วนๆ ทีละ 7000 ตัวอักษร
  for (let i = 0; i < message.length; i += CHUNK_LIMIT) {
    // ตัดข้อความจากตำแหน่ง i ถึง i+7000 (หรือจนถึงจบข้อความ)
    const chunk = message.substring(i, i + CHUNK_LIMIT);

    // ส่งข้อความ chunk ที่ตัดออกไปผ่าน Webex API แบบ asynchronous รอจนส่งเสร็จก่อนส่ง chunk ถัดไป
    await axios.post('https://webexapis.com/v1/messages', {
      roomId,   // ส่งข้อความเข้าในห้อง chat ที่ระบุ
      text: chunk  // เนื้อหาข้อความ chunk ที่จะส่ง
    }, {
      // ใส่ header สำหรับยืนยันตัวตนด้วยโทเคนของ Webex Bot
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });
  }
}

// ส่งไฟล์แนบให้ผู้ใช้ทาง Webex
async function sendFileAttachment(roomId, filename, content) {
  // สร้าง path ของไฟล์ชั่วคราวที่จะเก็บเนื้อหาไฟล์ใน /tmp ตามชื่อไฟล์ที่ระบุ
  const filePath = `/tmp/${filename}`;

  // เขียนข้อมูล content ลงในไฟล์ที่สร้างขึ้นในระบบไฟล์ ด้วย encoding แบบ utf8
  fs.writeFileSync(filePath, content, 'utf8');

  // สร้าง FormData เพื่อใช้ส่งข้อมูล multipart/form-data สำหรับการอัปโหลดไฟล์
  const form = new FormData();

  // แนบข้อมูลรหัสห้องแชท (roomId) ลงใน FormData
  form.append('roomId', roomId);

  // แนบไฟล์ที่สร้างขึ้นในระบบไฟล์ ไปเป็น field 'files' ใน FormData
  form.append('files', fs.createReadStream(filePath));

  // ส่ง HTTP POST request ไปยัง Webex API เพื่อส่งข้อความพร้อมไฟล์แนบ
  await axios.post('https://webexapis.com/v1/messages', form, {
    headers: {
      // กำหนด Header Authorization ใช้ Bearer token สำหรับบอทที่อนุญาต
      Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
      // แนบ header ที่จำเป็นสำหรับ multipart/form-data จาก form
      ...form.getHeaders()
    }
  });

  // ลบไฟล์ชั่วคราวออกจากระบบไฟล์ หลังจากส่งไฟล์เสร็จเรียบร้อย
  fs.unlinkSync(filePath);
}

// === ส่วนหลัก รับข้อความจาก Webex แล้วตอบกลับ ===
app.post('/webex', async (req, res) => {
  try {
    // ดึงข้อมูลส่วน data จาก request body ที่ส่งมา
    const data = req.body.data;

    // ดึง personId (ID ของคนส่งข้อความ) แล้วตัดช่องว่างหัวท้ายออก
    const personId = (data.personId || '').trim();

    // เช็คว่าถ้า personId ตรงกับ BOT_ID (หมายถึงข้อความมาจากบอทเอง) 
    // ให้ไม่ตอบกลับโดยส่งสถานะ 200 พร้อมข้อความ 'Ignore self-message'
    if (personId === BOT_ID) return res.status(200).send('Ignore self-message');

    // ดึง ID ของข้อความ (messageId) จากข้อมูล data
    const messageId = data.id;

    // ดึง ID ของห้องแชท (roomId) จากข้อมูล data
    const roomId = data.roomId;

    // เรียกข้อมูลข้อความ (message) จาก Webex API โดยใช้ messageId ที่ระบุ และส่ง token เพื่อยืนยันตัวตน
    const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, {
      headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
    });

    // ดึงข้อความ (text) จากข้อมูลที่ได้จาก API
    let messageText = messageRes.data.text;

    // ตรวจสอบว่า ข้อความนั้นขึ้นต้นด้วยชื่อบอท (WEBEX_BOT_NAME) หรือไม่ (ไม่สนใจตัวพิมพ์ใหญ่เล็ก)
    // เช่น ถ้าผู้ใช้พิมพ์ "@BotName สั่งงานอะไรสักอย่าง" ก็จะลบ "@BotName" ออก
    if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME)) {
      // ตัดชื่อบอทออกจากข้อความ แล้วลบช่องว่างส่วนเกินที่เหลือข้างหน้า-ข้างหลัง
      messageText = messageText.substring(WEBEX_BOT_NAME.length).trim();
    }

    // แยกคำสั่งจากข้อความ
    // แยกคำแรกออกเป็นคำสั่ง (command) และคำที่เหลือเป็นอาร์กิวเมนต์ (args)
    const [command, ...args] = messageText.split(' ');

    // รวมอาร์กิวเมนต์ทั้งหมดเป็นข้อความเดียว (keyword) และตัดช่องว่างซ้ายขวาออก
    const keyword = args.join(' ').trim();

    // ดึงชื่อแผ่นงานทั้งหมดจาก Google Sheets โดยใช้ฟังก์ชัน getAllSheetNames และส่ง Google Sheet ไอดีไป
    const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);

    // เตรียมตัวแปรเก็บข้อความตอบกลับ (ยังเป็นค่าว่าง)
    let responseText = '';


    // === ตรวจสอบคำสั่งที่ผู้ใช้ส่ง ===
    if (command === 'help') {
      // ตรวจสอบว่าคำสั่งที่รับเข้ามาเป็น 'help' หรือไม่
      responseText = `📌 คำสั่งที่ใช้ได้:\n` +
        // กำหนดข้อความตอบกลับโดยเริ่มด้วยหัวข้อคำสั่งที่ใช้ได้
        `1. @bot_small ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
        // คำสั่งค้นหาคำในทุกชีตข้อมูล
        `2. @bot_small ค้นหา <ชื่อชีต> → ดึงข้อมูลทั้งหมดในชีต (แนบไฟล์ถ้ายาว)\n` +
        // คำสั่งค้นหาข้อมูลทั้งหมดในชีตที่ระบุ พร้อมส่งเป็นไฟล์ถ้าข้อมูลยาวเกินไป
        `3. @bot_small แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n` +
        // คำสั่งแก้ไขข้อมูลในชีตตามตำแหน่งที่ระบุ (ชีต, คอลัมน์, แถว, ข้อความใหม่)
        `4. @bot_small help → แสดงวิธีใช้ทั้งหมด`;
      // คำสั่งแสดงข้อความช่วยเหลือ (ตัวนี้เอง)
    }
    else if (command === 'ค้นหา') {
      // เช็คว่าคำสั่งเป็น 'ค้นหา'

      // ถ้าคำที่พิมพ์มา (keyword) เป็นชื่อชีตที่มีอยู่ใน Google Sheets
      if (allSheetNames.includes(keyword)) {
        // เรียกฟังก์ชัน getSheetWithHeaders เพื่อดึงข้อมูลพร้อมหัวตาราง (headers) ของชีตนั้นๆ
        const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);

        // นำข้อมูลแต่ละแถวมาแปลงให้อยู่ในรูปแบบข้อความที่อ่านง่าย โดยใช้ฟังก์ชัน formatRow
        // แล้วนำแต่ละแถวมาต่อกันด้วยขึ้นบรรทัดใหม่ 2 ครั้ง (\n\n)
        const resultText = data.map((row, idx) => formatRow(row, rawHeaders2, idx, keyword)).join('\n\n');

        // ถ้าความยาวของข้อความเกิน 7000 ตัวอักษร (เกินขีดจำกัดข้อความใน Webex)
        if (resultText.length > 7000) {
          // ส่งข้อความแจ้งว่า ข้อมูลยาวเกิน จะแนบเป็นไฟล์แทน
          await axios.post('https://webexapis.com/v1/messages', {
            roomId,
            markdown: '📎 ข้อมูลยาวเกิน แนบเป็นไฟล์แทน'
          }, {
            headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
          });

          // เรียกฟังก์ชันส่งไฟล์แนบชื่อ 'ข้อมูล.txt' โดยมีเนื้อหาเป็นข้อความข้อมูลที่ดึงมา
          await sendFileAttachment(roomId, 'ข้อมูล.txt', resultText);

          // ส่งสถานะตอบกลับว่าไฟล์ถูกส่งเรียบร้อยแล้ว
          return res.status(200).send('sent file');
        } else {
          // ถ้าข้อความไม่ยาวเกิน ส่งข้อความข้อมูลทั้งหมดกลับในรูปแบบข้อความธรรมดา
          responseText = resultText;
        }
      } else {
        // กรณีที่ผู้ใช้พิมพ์คำค้นหาแบบทั่วไป (ไม่ระบุชื่อชีตเฉพาะ)
        // จะทำการค้นหาข้อมูลในทุกชีตที่มีอยู่ในสเปรดชีต

        let results = [];  // สร้างตัวแปรเก็บผลลัพธ์ที่เจอทั้งหมด

        // วนลูปทุกชื่อชีตในไฟล์ Google Sheet
        for (const sheetName of allSheetNames) {
          // เรียกฟังก์ชันดึงข้อมูลจากชีตนั้นๆ พร้อมหัวข้อคอลัมน์
          const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);

          // วนลูปข้อมูลแต่ละแถวในชีต
          data.forEach((row, idx) => {
            // ตรวจสอบว่าข้อความคำค้นหา (keyword) อยู่ในค่าของ cell ใดๆ ในแถวนั้นหรือไม่
            // flattenText คือฟังก์ชันช่วยแปลงข้อมูลให้อยู่ในรูปแบบข้อความที่ง่ายต่อการค้นหา
            const match = Object.values(row).some(v => flattenText(v).includes(keyword));

            // ถ้าพบคำค้นหาในแถวนั้น
            if (match)
              // นำแถวนั้นไปฟอร์แมตรูปแบบข้อความให้สวยงาม พร้อมใส่ข้อมูลเช่น ชื่อชีตและเลขแถว
              results.push(formatRow(row, rawHeaders2, idx, sheetName));
          });
        }

        // กำหนดข้อความตอบกลับ
        // ถ้ามีข้อมูลเจอให้รวมผลลัพธ์ทั้งหมดมาแสดงคั่นด้วยบรรทัดว่าง 2 บรรทัด
        // ถ้าไม่เจอข้อมูลเลย ให้ตอบข้อความแจ้งว่าไม่พบข้อมูล
        responseText = results.length ? results.join('\n\n') : '❌ ไม่พบข้อมูลที่ต้องการ';
      }
    } else if (command === 'แก้ไข') {
      // ถ้าคำสั่งคือ 'แก้ไข' ให้เข้ามาทำส่วนนี้เพื่อแก้ไขข้อมูลในชีต

      if (args.length < 4) {
        // ถ้าจำนวน argument น้อยกว่า 4 แสดงว่ารูปแบบคำสั่งไม่ครบถ้วน
        responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
        // แจ้งเตือนผู้ใช้ว่าคำสั่งไม่ถูกต้อง เพราะขาดข้อมูลบางอย่าง
      } else {
        // กรณีคำสั่งมี argument ครบถ้วน

        // ประกาศตัวแปรเพื่อเก็บข้อมูลจากคำสั่ง
        let sheetName = '';  // เก็บชื่อชีตที่จะเข้าไปแก้ไข
        let columnName = ''; // เก็บชื่อคอลัมน์ที่จะแก้ไข
        let rowIndex = 0;    // เก็บเลขแถวที่จะแก้ไข
        let newValue = '';   // เก็บข้อความใหม่ที่จะนำไปแทนที่ในเซลล์นั้น

        // สร้างชื่อชีตแบบรวม 2 คำ เช่น "เดือน ปี" จาก args[0] และ args[1]
        let trySheetName = args[0] + ' ' + args[1];

        // เช็คว่า trySheetName นี้อยู่ใน allSheetNames (รายชื่อชีตทั้งหมด) หรือไม่
        if (allSheetNames.includes(trySheetName)) {
          // ถ้าเจอ ชื่อชีตจะใช้ trySheetName ที่รวม 2 คำนี้
          sheetName = trySheetName;
          // กำหนดชื่อคอลัมน์เป็น args[2]
          columnName = args[2];
          // แปลง args[3] เป็นเลขแถว (integer)
          rowIndex = parseInt(args[3]);
          // เอาค่าใหม่ที่ต้องการแก้ไข ตั้งแต่ args[4] เป็นต้นไป รวมเป็นข้อความเดียว
          newValue = args.slice(4).join(' ');
        }
        // ถ้าไม่เจอแบบรวม 2 คำ ให้เช็คว่า args[0] เป็นชื่อชีตใน allSheetNames หรือไม่
        else if (allSheetNames.includes(args[0])) {
          // ถ้าเจอ ให้ใช้ args[0] เป็นชื่อชีต
          sheetName = args[0];
          // กำหนดชื่อคอลัมน์เป็น args[1]
          columnName = args[1];
          // แปลง args[2] เป็นเลขแถว
          rowIndex = parseInt(args[2]);
          // เอาค่าใหม่ที่ต้องการแก้ไข ตั้งแต่ args[3] เป็นต้นไป รวมเป็นข้อความเดียว
          newValue = args.slice(3).join(' ');
        }
        // ถ้าไม่เจอทั้งสองแบบ
        else {
          // แจ้งว่าไม่พบชีตชื่อ args[0] หรือ trySheetName
          responseText = `❌ ไม่พบชีตชื่อ "${args[0]}" หรือ "${trySheetName}"`;
        }


        if (sheetName) {
          // ถ้ามีชื่อ sheet ส่งเข้ามา (เช็คว่า sheetName ไม่ใช่ค่าว่างหรือ undefined)

          // เรียก Google Sheets API เพื่อดึงค่าช่วงข้อมูลจาก sheet นั้น (ช่วง A1 ถึง Z2)
          const res = await sheets.spreadsheets.values.get({
            spreadsheetId: GOOGLE_SHEET_FILE_ID,  // ไอดีของไฟล์ Google Sheets ที่จะดึงข้อมูล
            range: `${sheetName}!A1:Z2`           // กำหนดช่วงข้อมูลที่ต้องการดึง (แถว 1 ถึง 2 คอลัมน์ A ถึง Z ของ sheet นี้)
          });
          // เอาข้อมูลแถวแรกมาเก็บในตัวแปร header1 (มักจะเป็นชื่อคอลัมน์หลัก)
          const header1 = res.data.values[0];

          // เอาข้อมูลแถวที่สองมาเก็บในตัวแปร header2 (อาจจะเป็นรายละเอียดเพิ่มเติมของคอลัมน์)
          const header2 = res.data.values[1];

          // รวมชื่อคอลัมน์หลักและรายละเอียดคอลัมน์เข้าด้วยกัน เช่น "ชื่อคอลัมน์ รายละเอียด"
          // โดยถ้าไม่มีรายละเอียดก็ใช้แค่ชื่อคอลัมน์
          const headers = header1.map((h1, i) => `${h1} ${header2[i] || ''}`.trim());

          // หาตำแหน่ง index ของคอลัมน์ที่ชื่อรวมกันนั้นมีคำว่า columnName อยู่ (ค้นหาว่าคอลัมน์นี้อยู่ที่ตำแหน่งไหน)
          const colIndex = headers.findIndex(h => h.includes(columnName));
          if (colIndex === -1) {
            // ถ้าไม่เจอคอลัมน์ที่ต้องการ (colIndex = -1 หมายถึงไม่พบคอลัมน์)
            responseText = `❌ ไม่พบคอลัมน์ "${columnName}"`; // กำหนดข้อความตอบกลับว่าไม่พบคอลัมน์
          } else {
            // ถ้าพบคอลัมน์ที่ต้องการ
            const colLetter = String.fromCharCode(65 + colIndex);
            // แปลงเลข index ของคอลัมน์เป็นตัวอักษร A, B, C,... (ASCII 65 = 'A')

            const range = `${sheetName}!${colLetter}${rowIndex}`;
            // สร้างช่วงของ cell ในรูปแบบ "ชื่อชีต!ตัวอักษรคอลัมน์หมายเลขแถว"
            // เช่น "Sheet1!B5"

            await sheets.spreadsheets.values.update({
              spreadsheetId: GOOGLE_SHEET_FILE_ID, // ไอดีของ Google Sheet ที่จะอัปเดต
              range, // ช่วง cell ที่จะอัปเดต
              valueInputOption: 'USER_ENTERED', // กำหนดให้ Google Sheets แปลค่าเหมือนผู้ใช้พิมพ์ (เช่นสูตร, วันที่)
              requestBody: { values: [[newValue]] } // ข้อมูลใหม่ที่จะเขียนลง cell นั้น (2D array)
            });

            responseText = `✅ แก้ไขแล้ว: ${range} → ${newValue}`;
            // กำหนดข้อความตอบกลับยืนยันการแก้ไขสำเร็จ พร้อมบอกช่วงที่แก้ไขและค่าที่เปลี่ยน
          }
        }
      }
    } else {
      // ถ้าพิมพ์ไม่ถูกสั่ง
      // กรณีที่ข้อความที่ผู้ใช้พิมพ์มาไม่ตรงกับคำสั่งใด ๆ ที่เราเตรียมไว้ในโค้ดนี้
      responseText = '❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "help"';
    }

    // ส่งข้อความตอบกลับ
    // ฟังก์ชัน sendMessageInChunks จะส่งข้อความกลับไปยังแชทในห้องที่ระบุ(roomId) โดยจะแบ่งข้อความยาว ๆ ให้พอดีกับข้อจำกัดของแพลตฟอร์ม
    await sendMessageInChunks(roomId, responseText);

    // ตอบกลับ HTTP status 200 (OK) เพื่อแจ้งว่า request สำเร็จและ bot ส่งข้อความกลับเรียบร้อยแล้ว
    res.status(200).send('OK');
  } catch (err) {
    // ถ้ามีข้อผิดพลาดเกิดขึ้นในส่วนของ try ด้านบน จะเข้ามาที่ catch เพื่อจัดการ error
    // แสดงรายละเอียด error ใน console log เพื่อให้รู้ว่าผิดพลาดอะไร
    console.error('❌ ERROR:', err.stack || err.message);
    // ตอบกลับ HTTP status 500 เพื่อแจ้งว่าเกิดข้อผิดพลาดในเซิร์ฟเวอร์
    res.status(500).send('Error');
  }
});

// === เริ่มให้เซิร์ฟเวอร์ทำงานที่พอร์ตที่กำหนด ===
// สั่งให้แอป Express เริ่มฟังคำขอ (request) ที่พอร์ต PORT และแสดงข้อความว่า bot พร้อมทำงาน
app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`)); 