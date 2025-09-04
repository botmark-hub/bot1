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
function flattenText(text) {
    return (text || '')
        .toString()
        .replace(/\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

// Helper function to get cell value by header name from the second header row
function getCellByHeader2(rowArray, headerRow2, keyword) {
    const idx = headerRow2.findIndex(h =>
        h.trim().toLowerCase().includes(keyword.toLowerCase())
    );
    return idx !== -1 ? flattenText(rowArray[idx]) : '-';
}

// Helper function to format a row of data into a readable string for Webex message
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

// Function to get all sheet names from a Google Spreadsheet
async function getAllSheetNames(spreadsheetId) {
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    return res.data.sheets.map(sheet => sheet.properties.title);
}

// Function to get sheet data along with combined headers
async function getSheetWithHeaders(sheets, spreadsheetId, sheetName) {
    const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!A1:Z`
    });
    const rows = res.data.values;
    if (!rows || rows.length < 3)
        return { data: [], rawHeaders2: [] };
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

// Function to send a message in chunks if it's too long for Webex API limit
async function sendMessageInChunks(roomId, message) {
    const CHUNK_LIMIT = 5500;
    for (let i = 0; i < message.length; i += CHUNK_LIMIT) {
        const chunk = message.substring(i, i + CHUNK_LIMIT);
        try {
            await axios.post('https://webexapis.com/v1/messages', {
                roomId,
                text: chunk
            }, {
                headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
            });
        } catch (error) {
            console.error('Error sending message in chunk:', error.response ? error.response.data : error.message);
        }
    }
}

// Function to send a file attachment to a Webex room
async function sendFileAttachment(roomId, filename, data) {
    const dirPath = path.join(__dirname, 'tmp');
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }

    if (filename.endsWith('.xlsx')) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data');

        if (data && Array.isArray(data) && data.length > 0) {
            const headers = Object.keys(data[0]);
            worksheet.addRow(headers);

            headers.forEach((header, index) => {
                worksheet.getColumn(index + 1).width = Math.max(header.length + 5, 15);
            });

            const headerRow = worksheet.getRow(1);
            headerRow.eachCell((cell) => {
                cell.font = { bold: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFF00' }
                };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });

            data.forEach(row => {
                const rowData = headers.map(header => row[header]);
                const excelRow = worksheet.addRow(rowData);

                excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const column = worksheet.getColumn(colNumber);
                    const cellLength = cell.value ? cell.value.toString().length : 10;
                    column.width = Math.max(column.width, cellLength + 5);
                    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                });
            });
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const filePath = path.join(dirPath, filename);
        fs.writeFileSync(filePath, buffer);

        const form = new FormData();
        form.append('roomId', roomId);
        form.append('files', fs.createReadStream(filePath));

        try {
            await axios.post('https://webexapis.com/v1/messages', form, {
                headers: {
                    Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
                    ...form.getHeaders()
                }
            });
        } catch (error) {
            console.error('Error sending file:', error.response ? error.response.data : error.message);
        }

        fs.unlinkSync(filePath);
    } else {
        const filePath = path.join(dirPath, filename);
        fs.writeFileSync(filePath, data, 'utf8');
        const form = new FormData();
        form.append('roomId', roomId);
        form.append('files', fs.createReadStream(filePath));

        try {
            await axios.post('https://webexapis.com/v1/messages', form, {
                headers: {
                    Authorization: `Bearer ${WEBEX_BOT_TOKEN}`,
                    ...form.getHeaders()
                }
            });
        } catch (error) {
            console.error('Error sending file:', error.response ? error.response.data : error.message);
        }

        fs.unlinkSync(filePath);
    }
}

// Webhook endpoint for Webex messages
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
        if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME.toLowerCase())) {
            messageText = messageText.substring(WEBEX_BOT_NAME.length).trim();
        }

        const [command, ...args] = messageText.split(' ');
        const keyword = args.join(' ').trim();
        const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
        let responseText = '';
        const EXCEL_THRESHOLD_GENERAL_SEARCH = 500;

        if (command === 'help') {
            responseText = `📌 คำสั่งที่ใช้ได้:\n` +
                `1. @${WEBEX_BOT_NAME} ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
                                `2. @${WEBEX_BOT_NAME} ค้นหา <ชื่อชีต> → แสดงข้อมูลทั้งหมดในชีตนั้น\n` +
                `3. @${WEBEX_BOT_NAME} ค้นหา <ชื่อชีต> <ชื่อคอลัมน์> → แสดงเฉพาะคอลัมน์\n` +
                `4. @${WEBEX_BOT_NAME} แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ> → แก้ไขค่าใน cell\n`;

            await sendMessageInChunks(roomId, responseText);
        }

        else if (command === 'แก้ไข') {
            // แก้ไขให้รองรับชื่อชีตมีเว้นวรรค
            // รูปแบบ: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>
            if (args.length < 4) {
                await sendMessageInChunks(roomId, '❌ รูปแบบคำสั่งไม่ถูกต้อง: แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>');
                return res.status(200).send('ok');
            }

            // หาชื่อชีต
            let sheetName = '';
            let columnName = '';
            let rowIndex = 0;
            let newValue = '';
            // ลอง match sheet name จาก allSheetNames
            for (let i = allSheetNames.length; i > 0; i--) {
                const possibleSheetName = args.slice(0, i).join(' ');
                if (allSheetNames.includes(possibleSheetName)) {
                    sheetName = possibleSheetName;
                    columnName = args[i];
                    rowIndex = parseInt(args[i + 1]);
                    newValue = args.slice(i + 2).join(' ');
                    break;
                }
            }

            if (!sheetName) {
                await sendMessageInChunks(roomId, '❌ ไม่พบชื่อชีต: ' + args.join(' '));
                return res.status(200).send('ok');
            }

            // ดึงข้อมูลชีต
            const sheetData = await sheets.spreadsheets.values.get({
                spreadsheetId: GOOGLE_SHEET_FILE_ID,
                range: `${sheetName}!A1:Z1000`
            });

            const rows = sheetData.data.values;
            if (!rows || rows.length < 3) {
                await sendMessageInChunks(roomId, '❌ ชีตไม่มีข้อมูล');
                return res.status(200).send('ok');
            }

            // หา column index
            // หา column index จาก header จริง (แถว 1)
const headers = rows[0]; // แถวที่ 1 เป็น header จริง
const colIndex = headers.findIndex(h => h.trim() === columnName);
if (colIndex === -1) {
    await sendMessageInChunks(roomId, '❌ ไม่พบคอลัมน์: ' + columnName);
    return res.status(200).send('ok');
}

// ตรวจสอบแถว
if (rowIndex < 1 || rowIndex > rows.length - 1) { // rows.length -1 เพราะ header แถว 1
    await sendMessageInChunks(roomId, '❌ แถวที่ระบุไม่ถูกต้อง');
    return res.status(200).send('ok');
}

// แก้ไขค่า
const targetRow = rowIndex + 1; // +1 เพราะ Google Sheet index เริ่มที่ 1 และ header อยู่แถวแรก
const updateRange = `${sheetName}!${String.fromCharCode(65 + colIndex)}${targetRow + 1}`;
await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_FILE_ID,
    range: updateRange,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
        values: [[newValue]]
    }
});


            await sendMessageInChunks(roomId, `✅ แก้ไขสำเร็จ: ${sheetName} [${columnName} แถว ${rowIndex}] → ${newValue}`);
        }

        res.status(200).send('ok');
    } catch (error) {
        console.error('Error in /webex webhook:', error.message || error);
        res.status(500).send('error');
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`Webex bot server running on port ${PORT}`);
});

