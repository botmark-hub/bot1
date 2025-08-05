// Load environment variables from .env if not in production
// โหลดตัวแปรสภาพแวดล้อมจากไฟล์ .env หากไม่ได้อยู่ในโหมด production (เช่น ตอนพัฒนา)
if (process.env.NODE_ENV !== 'production') {
    require('dotenv').config(); // เรียกใช้ dotenv เพื่อโหลดไฟล์ .env
}

// Import necessary libraries
// นำเข้าไลบรารีที่จำเป็น
const express = require('express'); // Express.js สำหรับสร้างเว็บเซิร์ฟเวอร์
const bodyParser = require('body-parser'); // Body-parser สำหรับแยกวิเคราะห์ข้อมูลที่ส่งมากับ HTTP request
const axios = require('axios'); // Axios สำหรับส่ง HTTP requests ไปยัง API ภายนอก (เช่น Webex API)
const fs = require('fs'); // File System module สำหรับจัดการไฟล์ในเครื่อง (อ่าน/เขียน/ลบ)
const path = require('path'); // Path module สำหรับจัดการเส้นทางของไฟล์และไดเรกทอรี
const FormData = require('form-data'); // FormData สำหรับสร้างฟอร์มข้อมูลแบบ multipart/form-data (ใช้สำหรับส่งไฟล์)
const { google } = require('googleapis'); // Google APIs client library สำหรับเชื่อมต่อกับบริการของ Google (เช่น Google Sheets)
const ExcelJS = require('exceljs'); // ExcelJS สำหรับสร้างและจัดการไฟล์ Excel (.xlsx)

// Create a web server
// สร้างเว็บเซิร์ฟเวอร์
const app = express(); // สร้าง instance ของ Express app
app.use(bodyParser.json()); // ใช้ middleware ของ body-parser เพื่อให้ Express สามารถอ่าน JSON จาก request body ได้

// Set up important variables
// ตั้งค่าตัวแปรสำคัญต่างๆ
const PORT = process.env.PORT || 3000; // กำหนดพอร์ตที่จะรันเซิร์ฟเวอร์ ใช้ค่าจาก environment variable ชื่อ PORT หรือใช้ 3000 เป็นค่าเริ่มต้น
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN; // โทเค็นสำหรับบอท Webex (ใช้ในการยืนยันตัวตนกับ Webex API)
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID; // ID ของ Google Sheet ที่ต้องการทำงานด้วย
const WEBEX_BOT_NAME = 'bot_small'; // ชื่อที่ตั้งให้กับบอท Webex (ใช้ในการรับรู้คำสั่ง)
const BOT_ID = (process.env.BOT_ID || '').trim(); // ID ของบอท Webex (ใช้เพื่อไม่ให้บอทตอบกลับข้อความที่ตัวเองส่ง)

// Set up Google API connection with Service Account
// ตั้งค่าการเชื่อมต่อ Google API โดยใช้ Service Account (บัญชีบริการ)
const rawCreds = JSON.parse(process.env.GOOGLE_CREDENTIALS); // ดึงข้อมูล Credentials ของ Google จาก environment variable และแปลงจาก JSON string เป็น JavaScript object
rawCreds.private_key = rawCreds.private_key.replace(/\\n/g, '\n'); // แทนที่ '\\n' ด้วย newline จริงๆ ใน private_key (จำเป็นสำหรับ credentials)
const auth = new google.auth.GoogleAuth({ // สร้าง object สำหรับการยืนยันตัวตนกับ Google API
    credentials: rawCreds, // ใช้ credentials ที่ดึงมา
    scopes: [ // กำหนดขอบเขตการเข้าถึง (สิทธิ์) ที่บอทต้องการ
        'https://www.googleapis.com/auth/spreadsheets', // สิทธิ์ในการอ่านและแก้ไข Google Sheets
        'https://www.googleapis.com/auth/drive.readonly' // สิทธิ์ในการอ่านข้อมูล Google Drive (อาจจะใช้เพื่อดูชื่อไฟล์/โฟลเดอร์ แต่หลักๆ คือ sheets)
    ]
});

const sheets = google.sheets({ version: 'v4', auth }); // สร้าง client สำหรับ Google Sheets API (version 4) พร้อมใช้การยืนยันตัวตนที่ตั้งค่าไว้

// Helper function to flatten text (remove newlines, reduce multiple spaces)
// ฟังก์ชันช่วยในการจัดรูปแบบข้อความ: ลบขึ้นบรรทัดใหม่, ลดช่องว่างหลายอันให้เหลือช่องเดียว, ลบช่องว่างหัวท้าย
function flattenText(text) {
    return (text || '') // ถ้า text เป็น null/undefined ให้ใช้ string ว่าง
        .toString() // แปลงให้เป็น string
        .replace(/\n/g, ' ') // แทนที่ newline ด้วยช่องว่าง
        .replace(/\s+/g, ' ') // แทนที่ช่องว่างหลายอันติดกันด้วยช่องว่างอันเดียว
        .trim(); // ลบช่องว่างที่หัวและท้ายของข้อความ
}

// Helper function to get cell value by header name from the second header row
// ฟังก์ชันช่วยในการดึงค่าเซลล์จากแถวข้อมูล โดยใช้ชื่อคอลัมน์จาก headerRow2 (แถวหัวข้อที่สอง)
function getCellByHeader2(rowArray, headerRow2, keyword) {
    const idx = headerRow2.findIndex(h => // ค้นหา index ของคอลัมน์ใน headerRow2
        h.trim().toLowerCase().includes(keyword.toLowerCase()) // โดยที่ชื่อคอลัมน์ (ตัดช่องว่างและแปลงเป็นตัวเล็ก) มี keyword (แปลงเป็นตัวเล็ก) อยู่
    );
    return idx !== -1 ? flattenText(rowArray[idx]) : '-'; // ถ้าเจอ index ให้คืนค่าใน rowArray ที่ index นั้นๆ และจัดรูปแบบข้อความ ถ้าไม่เจอคืนค่า '-'
}

// Helper function to format a row of data into a readable string for Webex message
// ฟังก์ชันช่วยในการจัดรูปแบบข้อมูลหนึ่งแถวให้อยู่ในรูปแบบ string ที่อ่านง่ายสำหรับส่งใน Webex
function formatRow(rowObj, headerRow2, index, sheetName) {
    const rowArray = Object.values(rowObj); // แปลง object ของข้อมูลแถวนั้นให้เป็น array ของค่า
    return `📄 พบข้อมูลในชีต: ${sheetName} (แถว ${index + 3})\n` + // บอกว่าพบข้อมูลในชีตไหน แถวที่เท่าไหร่ (บวก 3 เพราะ header มี 2 แถว และ index เริ่มจาก 0)
        `📝 ชื่องาน: ${flattenText(rowObj['ชื่องาน'])} | 🧾 WBS: ${flattenText(rowObj['WBS'])}\n` + // ดึงค่าจากคอลัมน์ "ชื่องาน" และ "WBS"
        `💰 ชำระเงิน/ลว.: ${flattenText(rowObj['ชำระเงิน/ลว.'])} | ✅ อนุมัติ/ลว.: ${flattenText(rowObj['อนุมัติ/ลว.'])} | 📂 รับแฟ้ม: ${flattenText(rowObj['รับแฟ้ม'])}\n` + // ดึงค่าจากคอลัมน์ "ชำระเงิน/ลว.", "อนุมัติ/ลว.", "รับแฟ้ม"
        `🔌 หม้อแปลง: ${flattenText(rowObj['หม้อแปลง'])} | ⚡ ระยะทาง HT: ${getCellByHeader2(rowArray, headerRow2, 'HT')} | ⚡ ระยะทาง LT: ${getCellByHeader2(rowArray, headerRow2, 'LT')}\n` + // ดึงค่าจากคอลัมน์ "หม้อแปลง" และใช้ getCellByHeader2 เพื่อดึงค่าระยะทาง HT, LT
        `🪵 เสา 8 : ${getCellByHeader2(rowArray, headerRow2, '8')} | 🪵 เสา 9 : ${getCellByHeader2(rowArray, headerRow2, '9')} | 🪵 เสา 12 : ${getCellByHeader2(rowArray, headerRow2, '12')} | 🪵 เสา 12.20 : ${getCellByHeader2(rowArray, headerRow2, '12.20')}\n` + // ดึงค่าจำนวนเสาแต่ละประเภท
        `👷‍♂️ พชง.ควบคุม: ${flattenText(rowObj['พชง.ควบคุม'])}\n` + // ดึงค่าจากคอลัมน์ "พชง.ควบคุม"
        `📌 สถานะงาน: ${flattenText(rowObj['สถานะงาน'])} | 📊 เปอร์เซ็นงาน: ${flattenText(rowObj['เปอร์เซ็นงาน'])}\n` + // ดึงค่าจากคอลัมน์ "สถานะงาน" และ "เปอร์เซ็นงาน"
        `🗒️ หมายเหตุ: ${flattenText(rowObj['หมายเหตุ'])}`; // ดึงค่าจากคอลัมน์ "หมายเหตุ"
}

// Function to get all sheet names from a Google Spreadsheet
// ฟังก์ชันสำหรับดึงชื่อชีตทั้งหมดจาก Google Spreadsheet
async function getAllSheetNames(spreadsheetId) {
    const res = await sheets.spreadsheets.get({ spreadsheetId }); // ส่ง request ไปยัง Google Sheets API เพื่อดึงข้อมูล spreadsheet
    return res.data.sheets.map(sheet => sheet.properties.title); // คืนค่า array ของชื่อชีตทั้งหมด
}

// Function to get sheet data along with combined headers
// ฟังก์ชันสำหรับดึงข้อมูลจากชีตพร้อมกับสร้าง Header ที่รวมกันจาก Header สองแถวแรก
async function getSheetWithHeaders(sheets, spreadsheetId, sheetName) {
    const res = await sheets.spreadsheets.values.get({ // ส่ง request ไปยัง Google Sheets API เพื่อดึงค่าทั้งหมดจากชีตที่ระบุ
        spreadsheetId,
        range: `${sheetName}!A1:Z` // ดึงข้อมูลตั้งแต่ A1 ถึง Z ทั้งหมดในชีตนั้น
    });
    const rows = res.data.values; // ดึงข้อมูลแถวทั้งหมด
    if (!rows || rows.length < 3) // ถ้าไม่มีข้อมูล หรือมีน้อยกว่า 3 แถว (ซึ่งหมายถึงไม่มีแถวข้อมูลจริงหลังจาก 2 แถว header)
        return { data: [], rawHeaders2: [] }; // คืนค่าเป็น array ว่าง
    const headerRow1 = rows[0]; // แถวหัวข้อที่ 1
    const headerRow2 = rows[1]; // แถวหัวข้อที่ 2
    const headers = headerRow1.map((h1, i) => { // สร้าง array ของ header ที่รวมกัน
        const h2 = headerRow2[i] || ''; // ดึง header จากแถวที่ 2, ถ้าไม่มีให้ใช้ string ว่าง
        return h2 ? `${h1} ${h2}`.trim() : h1.trim(); // ถ้ารวมกันแล้วมี header2 ให้รวม h1 กับ h2 ถ้าไม่มีก็ใช้ h1 อย่างเดียว
    });
    const dataRows = rows.slice(2); // ตัดแถวข้อมูลจริงออกมา (ตั้งแต่แถวที่ 3 เป็นต้นไป)
    return {
        data: dataRows.map(row => { // แปลงแต่ละแถวข้อมูลให้เป็น object โดยใช้ headers เป็น key
            const rowData = {};
            headers.forEach((header, i) => {
                rowData[header] = row[i] || ''; // กำหนดค่าให้กับคอลัมน์นั้นๆ ถ้าไม่มีให้ใช้ string ว่าง
            });
            return rowData;
        }),
        rawHeaders2: headerRow2 // คืนค่า headerRow2 ดิบๆ กลับไปด้วย (สำหรับใช้ใน formatRow)
    };
}

// Function to send a message in chunks if it's too long for Webex API limit
// ฟังก์ชันสำหรับส่งข้อความเป็นส่วนๆ หากข้อความยาวเกินขีดจำกัดของ Webex API
async function sendMessageInChunks(roomId, message) {
    const CHUNK_LIMIT = 5500; // กำหนดขีดจำกัดความยาวของข้อความแต่ละส่วน (5500 ตัวอักษร)

    for (let i = 0; i < message.length; i += CHUNK_LIMIT) { // วนลูปแบ่งข้อความเป็นส่วนๆ
        const chunk = message.substring(i, i + CHUNK_LIMIT); // ดึงข้อความส่วนนั้นๆ
        try {
            await axios.post('https://webexapis.com/v1/messages', { // ส่งข้อความไปยัง Webex API
                roomId, // ส่งไปยังห้องที่ระบุ
                text: chunk // ข้อความส่วนนั้น
            }, {
                headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` } // แนบ Authorization token
            });
        } catch (error) {
            console.error('Error sending message in chunk:', error.response ? error.response.data : error.message); // แสดงข้อผิดพลาดหากส่งไม่สำเร็จ
            if (error.response && error.response.data && error.response.data.errors) {
                error.response.data.errors.forEach(err => console.error('   Webex API Error:', err.description)); // แสดงรายละเอียดข้อผิดพลาดจาก Webex API
            }
        }
    }
}

// Function to send a file attachment to a Webex room
// ฟังก์ชันสำหรับส่งไฟล์แนบไปยังห้อง Webex
async function sendFileAttachment(roomId, filename, data) {
    const dirPath = path.join(__dirname, 'tmp'); // สร้าง path สำหรับโฟลเดอร์ชั่วคราว 'tmp'
    if (!fs.existsSync(dirPath)) { // ถ้าโฟลเดอร์ 'tmp' ยังไม่มี
        fs.mkdirSync(dirPath, { recursive: true }); // สร้างโฟลเดอร์ 'tmp'
    }

    if (filename.endsWith('.xlsx')) { // ถ้าเป็นไฟล์ Excel
        const workbook = new ExcelJS.Workbook(); // สร้าง workbook ใหม่
        const worksheet = workbook.addWorksheet('Data'); // เพิ่ม worksheet ชื่อ 'Data'

        if (data && Array.isArray(data) && data.length > 0) { // ถ้ามีข้อมูลและข้อมูลเป็น array ที่ไม่ว่างเปล่า
            const headers = Object.keys(data[0]); // ดึงชื่อ headers จาก key ของ object ในแถวแรกของข้อมูล
            worksheet.addRow(headers); // เพิ่มแถว header ใน worksheet

            headers.forEach((header, index) => {
                worksheet.getColumn(index + 1).width = Math.max(header.length + 5, 15); // กำหนดความกว้างของคอลัมน์ตามความยาวของ header หรืออย่างน้อย 15
            });

            const headerRow = worksheet.getRow(1); // เลือกแถว header (แถวที่ 1)
            headerRow.eachCell((cell) => { // วนลูปผ่านแต่ละเซลล์ในแถว header
                cell.font = { bold: true }; // ทำให้ตัวอักษรเป็นตัวหนา
                cell.fill = { // ใส่สีพื้นหลังเซลล์เป็นสีเหลือง
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFF00' }
                };
                cell.alignment = { vertical: 'middle', horizontal: 'center' }; // จัดตำแหน่งข้อความกลาง
            });

            data.forEach(row => { // วนลูปแต่ละแถวข้อมูล
                const rowData = headers.map(header => row[header]); // จัดเรียงข้อมูลในแถวตามลำดับ header
                const excelRow = worksheet.addRow(rowData); // เพิ่มแถวข้อมูลใน worksheet

                excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => { // วนลูปผ่านแต่ละเซลล์ในแถวข้อมูล (รวมเซลล์ว่าง)
                    const column = worksheet.getColumn(colNumber); // ดึงคอลัมน์นั้นๆ
                    const cellLength = cell.value ? cell.value.toString().length : 10; // กำหนดความยาวของเซลล์ (ถ้ามีค่า)
                    column.width = Math.max(column.width, cellLength + 5); // ปรับความกว้างของคอลัมน์ให้เหมาะสม
                    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; // จัดตำแหน่งข้อความซ้ายและให้ขึ้นบรรทัดใหม่ได้
                });
            });
        } else {
            console.warn('Warning: No data or invalid data format provided for Excel file. Creating an empty Excel file.'); // แจ้งเตือนหากไม่มีข้อมูลหรือข้อมูลไม่ถูกต้องสำหรับสร้าง Excel
        }

        const buffer = await workbook.xlsx.writeBuffer(); // เขียน workbook ลงใน buffer (ข้อมูลไบนารีของไฟล์ Excel)
        const filePath = path.join(dirPath, filename); // สร้าง path สำหรับไฟล์ Excel ชั่วคราว
        fs.writeFileSync(filePath, buffer); // เขียน buffer ลงในไฟล์ Excel

        const form = new FormData(); // สร้าง FormData object
        form.append('roomId', roomId); // เพิ่ม roomId
        form.append('files', fs.createReadStream(filePath)); // เพิ่มไฟล์ที่จะส่ง (อ่านจาก stream)

        try {
            await axios.post('https://webexapis.com/v1/messages', form, { // ส่ง request ไปยัง Webex API เพื่อส่งไฟล์
                headers: {
                    Authorization: `Bearer ${WEBEX_BOT_TOKEN}`, // แนบ Authorization token
                    ...form.getHeaders() // แนบ headers ที่จำเป็นสำหรับ FormData (เช่น Content-Type: multipart/form-data)
                }
            });
        } catch (error) {
            console.error('Error sending file:', error.response ? error.response.data : error.message); // แสดงข้อผิดพลาดหากส่งไม่สำเร็จ
        }

        fs.unlinkSync(filePath); // ลบไฟล์ชั่วคราวหลังจากส่งเสร็จ
    } else { // ถ้าไม่ใช่ไฟล์ Excel (น่าจะเป็นไฟล์ข้อความ)
        const filePath = path.join(dirPath, filename); // สร้าง path สำหรับไฟล์ชั่วคราว
        fs.writeFileSync(filePath, data, 'utf8'); // เขียนข้อมูลลงในไฟล์ (เป็น utf8)
        const form = new FormData(); // สร้าง FormData object
        form.append('roomId', roomId); // เพิ่ม roomId
        form.append('files', fs.createReadStream(filePath)); // เพิ่มไฟล์ที่จะส่ง

        try {
            await axios.post('https://webexapis.com/v1/messages', form, { // ส่ง request ไปยัง Webex API เพื่อส่งไฟล์
                headers: {
                    Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
                    ...form.getHeaders()
                }
            });
        } catch (error) {
            console.error('Error sending file:', error.response ? error.response.data : error.message); // แสดงข้อผิดพลาดหากส่งไม่สำเร็จ
        }

        fs.unlinkSync(filePath); // ลบไฟล์ชั่วคราวหลังจากส่งเสร็จ
    }
}

// Webhook endpoint for Webex messages
// Endpoint สำหรับ Webhook ของ Webex (เมื่อมีข้อความเข้ามา Webex จะส่ง request มาที่นี่)
app.post('/webex', async (req, res) => {
    try {
        const data = req.body.data; // ดึงข้อมูลจาก request body (ข้อมูลเกี่ยวกับข้อความ Webex)
        const personId = (data.personId || '').trim(); // ID ของผู้ส่งข้อความ
        if (personId === BOT_ID) return res.status(200).send('Ignore self-message'); // ถ้าผู้ส่งคือบอทเอง ให้เพิกเฉย (ป้องกันบอทตอบกลับตัวเองเป็นวงวน)
        const messageId = data.id; // ID ของข้อความ
        const roomId = data.roomId; // ID ของห้องที่ส่งข้อความมา

        const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, { // ดึงรายละเอียดข้อความจาก Webex API โดยใช้ messageId
            headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` } // แนบ Authorization token
        });

        let messageText = messageRes.data.text; // ดึงข้อความดิบ
        if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME.toLowerCase())) { // ถ้าข้อความขึ้นต้นด้วยชื่อบอท (เช่น @bot_small ค้นหา...)
            messageText = messageText.substring(WEBEX_BOT_NAME.length).trim(); // ตัดชื่อบอทออกไป เพื่อให้เหลือแต่คำสั่ง
        }

        const [command, ...args] = messageText.split(' '); // แยกคำสั่งและ arguments ออกจากกัน (เช่น "ค้นหา" เป็น command, ส่วนที่เหลือเป็น args)
        const keyword = args.join(' ').trim(); // รวม args ที่เหลือเข้าด้วยกันเป็น keyword
        const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID); // ดึงชื่อชีตทั้งหมดจาก Google Sheet
        let responseText = ''; // ตัวแปรสำหรับเก็บข้อความตอบกลับ

        // Define EXCEL_THRESHOLD for general search. Use a conservative value.
        // กำหนดค่า EXCEL_THRESHOLD สำหรับการค้นหาทั่วไป (ถ้าข้อความผลลัพธ์ยาวเกินกว่านี้จะส่งเป็น Excel แทน)
        const EXCEL_THRESHOLD_GENERAL_SEARCH = 500;

        if (command === 'help') { // ถ้าคำสั่งคือ "help"
            responseText = `📌 คำสั่งที่ใช้ได้:\n` + // สร้างข้อความแสดงรายการคำสั่ง
                `1. @${WEBEX_BOT_NAME} ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
                `2. @${WEBEX_BOT_NAME} ค้นหา <ชื่อชีต> → ดึงข้อมูลทั้งหมดในชีต (แนบไฟล์ถ้ายาว)\n` +
                `3. @${WEBEX_BOT_NAME} แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n` +
                `4. @${WEBEX_BOT_NAME} help → แสดงวิธีใช้ทั้งหมด`;
        } else if (command === 'ค้นหา') { // ถ้าคำสั่งคือ "ค้นหา"
            if (allSheetNames.includes(keyword)) { // ตรวจสอบว่า keyword เป็นชื่อชีตที่มีอยู่หรือไม่ (ค้นหาเฉพาะชีต)
                const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword); // ดึงข้อมูลจากชีตนั้นๆ
                const resultText = data.map((row, idx) => formatRow(row, rawHeaders2, idx, keyword)).join('\n\n'); // จัดรูปแบบข้อมูลที่ได้ให้เป็น string

                // --- DEBUGGING START (specific sheet) ---
                console.log(`DEBUG: resultText length (specific sheet) = ${resultText.length}`); // แสดงความยาวของข้อความผลลัพธ์
                const EXCEL_THRESHOLD_SPECIFIC_SHEET = 500; // กำหนด threshold สำหรับชีตเฉพาะ
                console.log(`DEBUG: EXCEL_THRESHOLD_SPECIFIC_SHEET = ${EXCEL_THRESHOLD_SPECIFIC_SHEET}`);
                console.log(`DEBUG: Is resultText.length > EXCEL_THRESHOLD_SPECIFIC_SHEET? ${resultText.length > EXCEL_THRESHOLD_SPECIFIC_SHEET}`);
                // --- DEBUGGING END ---

                if (resultText.length > EXCEL_THRESHOLD_SPECIFIC_SHEET) { // ถ้าข้อความผลลัพธ์ยาวเกิน threshold
                    console.log('DEBUG: Condition met (specific sheet). Sending Excel file...'); // แสดงข้อความ debug
                    await axios.post('https://webexapis.com/v1/messages', { // ส่งข้อความแจ้งเตือนว่าจะแนบไฟล์ Excel
                        roomId,
                        markdown: '📎 ข้อมูลยาวเกินไป กำลังแนบไฟล์ Excel แทน...'
                    }, {
                        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
                    });

                    const excelData = data.map(row => { // จัดเตรียมข้อมูลสำหรับ Excel (แปลง object ให้อยู่ในรูปแบบที่ ExcelJS ต้องการ)
                        const rowData = {};
                        Object.keys(row).forEach(header => {
                            rowData[header] = row[header];
                        });
                        return rowData;
                    });

                    await sendFileAttachment(roomId, 'ข้อมูล.xlsx', excelData); // ส่งไฟล์ Excel
                    return res.status(200).send('sent file'); // ส่งสถานะกลับไปว่าส่งไฟล์แล้ว
                } else {
                    console.log('DEBUG: Condition NOT met (specific sheet). Sending text message in chunks...'); // แสดงข้อความ debug
                    responseText = resultText; // ถ้าข้อความไม่ยาวเกิน ให้เก็บผลลัพธ์ใน responseText เพื่อส่งเป็นข้อความปกติ
                }
            } else { // Handle 'ค้นหา <คำ>' across all sheets (ค้นหาคำทั่วไปในทุกชีต)
                let results = []; // array สำหรับเก็บผลลัพธ์ที่ตรงกัน
                let rawExcelDataForSearch = []; // array สำหรับเก็บข้อมูลดิบที่ตรงกัน เพื่อสร้างไฟล์ Excel หากจำเป็น

                for (const sheetName of allSheetNames) { // วนลูปผ่านทุกชีต
                    const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName); // ดึงข้อมูลจากแต่ละชีต
                    data.forEach((row, idx) => { // วนลูปแต่ละแถวข้อมูลในชีตนั้นๆ
                        const match = Object.values(row).some(v => flattenText(v).includes(keyword)); // ตรวจสอบว่ามีค่าใดในแถวนั้นตรงกับ keyword หรือไม่
                        if (match) { // ถ้าพบการตรงกัน
                            results.push(formatRow(row, rawHeaders2, idx, sheetName)); // เพิ่มผลลัพธ์ที่จัดรูปแบบแล้วลงใน results
                            // Add raw data to be used for Excel file
                            const rowData = {}; // สร้าง object สำหรับเก็บข้อมูลดิบของแถว
                            Object.keys(row).forEach(header => {
                                rowData[header] = row[header];
                            });
                            rawExcelDataForSearch.push(rowData); // เพิ่มข้อมูลดิบลงใน rawExcelDataForSearch
                        }
                    });
                }
                responseText = results.length ? results.join('\n\n') : '❌ ไม่พบข้อมูลที่ต้องการ'; // ถ้ามีผลลัพธ์ ให้รวมเป็น string ถ้าไม่มีให้แสดงข้อความว่าไม่พบ

                // --- DEBUGGING START (general search) ---
                console.log(`DEBUG: Search across all sheets. responseText length = ${responseText.length}`); // แสดงความยาวของข้อความผลลัพธ์
                console.log(`DEBUG: EXCEL_THRESHOLD_GENERAL_SEARCH = ${EXCEL_THRESHOLD_GENERAL_SEARCH}`);
                console.log(`DEBUG: Is responseText.length > EXCEL_THRESHOLD_GENERAL_SEARCH? ${responseText.length > EXCEL_THRESHOLD_GENERAL_SEARCH}`);
                // --- DEBUGGING END ---

                // *** APPLY EXCEL THRESHOLD CHECK FOR GENERAL SEARCH HERE ***
                // ตรวจสอบว่าผลลัพธ์การค้นหาทั่วไปยาวเกิน threshold หรือไม่
                if (responseText.length > EXCEL_THRESHOLD_GENERAL_SEARCH) {
                    console.log('DEBUG: Condition met (general search). Sending Excel file...');
                    await axios.post('https://webexapis.com/v1/messages', {
                        roomId,
                        markdown: '📎 ข้อมูลยาวเกินไป กำลังแนบไฟล์ Excel แทน (จากการค้นหาหลายชีต)...' // แจ้งเตือนผู้ใช้
                    }, {
                        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
                    });

                    // Use the collected rawExcelDataForSearch to create the Excel file
                    if (rawExcelDataForSearch.length > 0) { // ถ้ามีข้อมูลดิบที่รวบรวมไว้
                        await sendFileAttachment(roomId, 'ผลการค้นหา.xlsx', rawExcelDataForSearch); // ส่งไฟล์ Excel พร้อมข้อมูลดิบ
                        return res.status(200).send('sent file'); // ส่งสถานะกลับไปว่าส่งไฟล์แล้ว
                    } else {
                        // This case handles if responseText was long but somehow rawExcelDataForSearch is empty
                        // (e.g., all matches were on empty rows, or some logic issue)
                        responseText = '❌ ไม่พบข้อมูลที่ต้องการ'; // Fallback to a short message
                        await sendMessageInChunks(roomId, responseText); // ส่งข้อความสั้นๆ แทน
                        return res.status(200).send('OK');
                    }
                }
                // If the responseText is not too long for general search, it will fall through
                // to the final sendMessageInChunks below.
                // ถ้าข้อความไม่ยาวเกิน ก็จะถูกส่งเป็นข้อความปกติใน sendMessageInChunks ด้านล่าง
            }
        } else if (command === 'แก้ไข') { // ถ้าคำสั่งคือ "แก้ไข"
            if (args.length < 3) { // ตรวจสอบว่ามี argument เพียงพอหรือไม่
                responseText = '❗ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>';
            } else {
                let sheetName = args[0]; // ชื่อชีต
                let columnName = args[1]; // ชื่อคอลัมน์
                let rowIndex = parseInt(args[2]); // เลขแถว (แปลงเป็นตัวเลข)
                let newValue = args.slice(3).join(' '); // ค่าใหม่ที่จะใส่

                if (allSheetNames.includes(sheetName)) { // ตรวจสอบว่าชื่อชีตมีอยู่จริง
                    const res = await sheets.spreadsheets.values.get({ // ดึงข้อมูลชีตทั้งหมด (เพื่อหาตำแหน่งคอลัมน์)
                        spreadsheetId: GOOGLE_SHEET_FILE_ID,
                        range: `${sheetName}!A1:Z`
                    });

                    const rows = res.data.values || []; // ดึงข้อมูลแถวทั้งหมด (ถ้าไม่มีให้เป็น array ว่าง)
                    // rowIndex + 2 เพราะว่าใน Google Sheet index เริ่มจาก 1 และมี 2 header row
                    if (rowIndex >= 1 && (rowIndex + 1) < rows.length) { // ตรวจสอบว่า rowIndex ถูกต้องและไม่เกินขอบเขตข้อมูล
                        const header1 = rows[0]; // แถว header ที่ 1
                        const header2 = rows[1]; // แถว header ที่ 2
                        const headers = header1.map((h1, i) => `${h1} ${header2[i] || ''}`.trim()); // สร้าง header รวม
                        const colIndex = headers.findIndex(h => h.includes(columnName)); // หา index ของคอลัมน์ที่ต้องการแก้ไข

                        if (colIndex === -1) { // ถ้าไม่พบคอลัมน์
                            responseText = `❌ ไม่พบคอลัมน์ "${columnName}"`;
                        } else {
                            const colLetter = String.fromCharCode(65 + colIndex); // แปลง index คอลัมน์เป็นตัวอักษร (A, B, C...)
                            const range = `${sheetName}!${colLetter}${rowIndex + 2}`; // สร้าง range สำหรับการอัปเดต (บวก 2 เพราะแถวข้อมูลจริงเริ่มที่ 3)
                            await sheets.spreadsheets.values.update({ // อัปเดตค่าใน Google Sheet
                                spreadsheetId: GOOGLE_SHEET_FILE_ID,
                                range, // range ที่ต้องการอัปเดต
                                valueInputOption: 'USER_ENTERED', // ตัวเลือกการป้อนค่า (เหมือนพิมพ์ด้วยมือ)
                                requestBody: { values: [[newValue]] } // ค่าใหม่ที่จะใส่
                            });
                            responseText = `✅ แก้ไขแล้ว: ${range} → ${newValue}`; // แจ้งผู้ใช้ว่าแก้ไขสำเร็จ
                        }
                    } else {
                        responseText = `❌ ไม่พบแถวที่ ${rowIndex} หรือแถวไม่ถูกต้อง (ควรเป็นแถวข้อมูล เช่น 1, 2, ...)`; // แจ้งว่าแถวไม่ถูกต้อง
                    }
                } else {
                    responseText = `❌ ไม่พบชีตชื่อ "${sheetName}"`; // แจ้งว่าไม่พบชื่อชีต
                }
            }
        } else { // ถ้าเป็นคำสั่งที่ไม่รู้จัก
            responseText = `❓ ไม่เข้าใจคำสั่ง ลองพิมพ์ "@${WEBEX_BOT_NAME} help"`; // แจ้งให้ลองพิมพ์ "help"
        }

        // This is the final step to send responseText if no file was attached earlier
        // (e.g., for 'help', 'แก้ไข', or 'ค้นหา' results that are short enough)
        // ขั้นตอนนี้จะส่ง responseText กลับไปหาผู้ใช้ หากก่อนหน้านี้ไม่มีการแนบไฟล์
        if (!res.headersSent) { // ตรวจสอบว่า header ยังไม่ถูกส่งไปแล้ว (หมายความว่ายังไม่ได้มีการส่งไฟล์)
            await sendMessageInChunks(roomId, responseText); // ส่งข้อความตอบกลับเป็นส่วนๆ
            res.status(200).send('OK'); // ส่งสถานะ 200 OK กลับไปหา Webex
        }
    } catch (err) {
        console.error('❌ ERROR:', err.stack || err.message); // แสดงข้อผิดพลาดที่เกิดขึ้น
        res.status(500).send('Error'); // ส่งสถานะ 500 Internal Server Error กลับไป
    }
});

// Start the server
// เริ่มต้นการทำงานของเซิร์ฟเวอร์
app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`)); // ให้เซิร์ฟเวอร์ฟังที่พอร์ตที่กำหนดและแสดงข้อความเมื่อพร้อมทำงาน