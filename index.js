// Load environment variables from .env if not in production
if (process.env.NODE_ENV !== 'production') {
    require('dotenv').config();
}

// Import necessary libraries
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const { google } = require('googleapis');
const ExcelJS = require('exceljs');

// Create a web server
const app = express();
app.use(bodyParser.json());

// Set up important variables
const PORT = process.env.PORT || 3000;
const WEBEX_BOT_TOKEN = process.env.WEBEX_BOT_TOKEN;
const GOOGLE_SHEET_FILE_ID = process.env.GOOGLE_SHEET_FILE_ID;
const WEBEX_BOT_NAME = 'bot_small';
const BOT_ID = (process.env.BOT_ID || '').trim();

// Set up Google API connection
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

// Helper functions
function flattenText(text) {
    return (text || '').toString().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
}

function getCellByHeader2(rowArray, headerRow2, keyword) {
    const idx = headerRow2.findIndex(h => h.trim().toLowerCase().includes(keyword.toLowerCase()));
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
            headers.forEach((header, i) => rowData[header] = row[i] || '');
            return rowData;
        }),
        rawHeaders2: headerRow2
    };
}

async function sendMessageInChunks(roomId, message) {
    const CHUNK_LIMIT = 5500;
    for (let i = 0; i < message.length; i += CHUNK_LIMIT) {
        const chunk = message.substring(i, i + CHUNK_LIMIT);
        try {
            await axios.post('https://webexapis.com/v1/messages', {
                roomId,
                text: chunk
            }, { headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` } });
        } catch (error) {
            console.error('Error sending message in chunk:', error.response ? error.response.data : error.message);
        }
    }
}

async function sendFileAttachment(roomId, filename, data) {
    const dirPath = path.join(__dirname, 'tmp');
    if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });

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
            headerRow.eachCell(cell => {
                cell.font = { bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
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
        try { await axios.post('https://webexapis.com/v1/messages', form, { headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}`, ...form.getHeaders() } }); }
        catch (error) { console.error('Error sending file:', error.response ? error.response.data : error.message); }
        fs.unlinkSync(filePath);
    } else {
        const filePath = path.join(dirPath, filename);
        fs.writeFileSync(filePath, data, 'utf8');
        const form = new FormData();
        form.append('roomId', roomId);
        form.append('files', fs.createReadStream(filePath));
        try { await axios.post('https://webexapis.com/v1/messages', form, { headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}`, ...form.getHeaders() } }); }
        catch (error) { console.error('Error sending file:', error.response ? error.response.data : error.message); }
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
        const messageRes = await axios.get(`https://webexapis.com/v1/messages/${messageId}`, { headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` } });
        let messageText = messageRes.data.text;
        if (messageText.toLowerCase().startsWith(WEBEX_BOT_NAME.toLowerCase())) messageText = messageText.substring(WEBEX_BOT_NAME.length).trim();
        const [command, ...args] = messageText.split(' ');
        const keyword = args.join(' ').trim();
        const allSheetNames = await getAllSheetNames(GOOGLE_SHEET_FILE_ID);
        let responseText = '';
        const EXCEL_THRESHOLD_GENERAL_SEARCH = 500;

        if (command === 'help') {
            responseText = `📌 คำสั่งที่ใช้ได้:\n` +
                `1. @${WEBEX_BOT_NAME} ค้นหา <คำ> → ค้นหาข้อมูลทุกชีต\n` +
                `2. @${WEBEX_BOT_NAME} ค้นหา <ชื่อชีต> → ดึงข้อมูลทั้งหมดในชีต (แนบไฟล์ถ้ายาว)\n` +
                `3. @${WEBEX_BOT_NAME} แก้ไข <ชื่อชีต> <ชื่อคอลัมน์> <แถวที่> <ข้อความ>\n` +
                `4. @${WEBEX_BOT_NAME} help → แสดงวิธีใช้ทั้งหมด`;
        }
        // ค้นหาและแก้ไข logic เหมือนเวอร์ชันก่อนหน้า...
        // ใส่ logic ส่งไฟล์ .xlsx/.txt ตามที่ผู้ใช้เลือก เช่น:
        // 1. ถ้าข้อมูลยาว ให้ถามผู้ใช้ว่าอยากได้ไฟล์แบบไหน
        // 2. ส่งไฟล์ตามการตอบกลับ
        // 3. ถ้าไม่ยาว ให้ส่งข้อความปกติ

        if (!res.headersSent) {
            await sendMessageInChunks(roomId, responseText);
            res.status(200).send('OK');
        }
    } catch (err) {
        console.error('❌ ERROR:', err.stack || err.message);
        res.status(500).send('Error');
    }
});

// Start the server
app.listen(PORT, () => console.log(`🚀 Bot พร้อมทำงานที่พอร์ต ${PORT}`));
