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

// Set up Google API connection with Service Account
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

// Helper function to flatten text
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
            const headers = rows[0];
            const colIndex = headers.findIndex(h => h.trim() === columnName);
            if (colIndex === -1) {
                await sendMessageInChunks(roomId, '❌ ไม่พบคอลัมน์: ' + columnName);
                return res.status(200).send('ok');
            }

            // ตรวจสอบแถว
            if (rowIndex < 1 || rowIndex > rows.length - 2) {
                await sendMessageInChunks(roomId, '❌ แถวที่ระบุไม่ถูกต้อง แก้ไขไม่สำเร็จ');
                return res.status(200).send('ok');
            }

            // คำนวณแถวเป้าหมาย
            const targetRow = rowIndex + 2;

            // กำหนดช่วงที่จะแก้ไข
            const updateRange = `${sheetName}!${String.fromCharCode(65 + colIndex)}${targetRow}`;

            // แก้ไขค่าใน Google Sheets
            await sheets.spreadsheets.values.update({
                spreadsheetId: GOOGLE_SHEET_FILE_ID,
                range: updateRange,
                valueInputOption: 'USER_ENTERED',
                requestBody: {
                    values: [[newValue]]
                }
            });

            // ส่งข้อความยืนยันการแก้ไข
            await sendMessageInChunks(roomId, `✅ แก้ไขสำเร็จ: ${sheetName} [${columnName} แถว ${rowIndex}] → ${newValue}`);
        }
        else if (command === 'ค้นหา') {
            if (!keyword) {
                await sendMessageInChunks(roomId, '❌ โปรดระบุคำค้นหา เช่น: ค้นหา เดือน ธันวาคม');
                return res.status(200).send('ok');
            }

            // 1) ถ้า keyword ตรงกับชื่อชีต → แสดงข้อมูลทั้งหมด
            if (allSheetNames.includes(keyword)) {
                const { data } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);
                if (!data.length) {
                    await sendMessageInChunks(roomId, `❌ ไม่พบข้อมูลในชีต ${keyword}`);
                } else if (data.length > 100) {
                    // ถ้าเกิน 100 แถว → ส่งเป็นไฟล์
                    await sendFileAttachment(roomId, `${keyword}.xlsx`, data);
                } else {
                    let msg = `📑 ข้อมูลทั้งหมดในชีต: ${keyword}\n\n`;
                    data.forEach((row, idx) => {
                        msg += formatRow(row, Object.keys(row), idx, keyword) + '\n\n';
                    });
                    await sendMessageInChunks(roomId, msg);
                }
                return res.status(200).send('ok');
            }

            // 2) ถ้า keyword = "<ชื่อชีต> <ชื่อคอลัมน์>"
            for (const sheetName of allSheetNames) {
                if (keyword.startsWith(sheetName + ' ')) {
                    const colName = keyword.slice(sheetName.length).trim();
                    const { data } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
                    if (!data.length) {
                        await sendMessageInChunks(roomId, `❌ ไม่พบข้อมูลในชีต ${sheetName}`);
                    } else {
                        const colExists = Object.keys(data[0]).find(h => h.includes(colName));
                        if (!colExists) {
                            await sendMessageInChunks(roomId, `❌ ไม่พบคอลัมน์ ${colName} ในชีต ${sheetName}`);
                        } else {
                            let msg = `📑 คอลัมน์ ${colName} ในชีต ${sheetName}\n\n`;
                            data.forEach((row, idx) => {
                                msg += `แถว ${idx + 3}: ${flattenText(row[colExists])}\n`;
                            });
                            await sendMessageInChunks(roomId, msg);
                        }
                    }
                    return res.status(200).send('ok');
                }
            }

            // 3) ค้นหาคำทั่วทุกชีต
            let results = [];
            for (const sheetName of allSheetNames) {
                const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
                data.forEach((row, idx) => {
                    const values = Object.values(row).join(' ');
                    if (values.includes(keyword)) {
                        results.push(formatRow(row, rawHeaders2, idx, sheetName));
                    }
                });
            }

            if (!results.length) {
                await sendMessageInChunks(roomId, `❌ ไม่พบ "${keyword}" ในทุกชีต`);
            } else if (results.length > EXCEL_THRESHOLD_GENERAL_SEARCH) {
                await sendFileAttachment(roomId, `ผลการค้นหา_${keyword}.xlsx`, results);
            } else {
                await sendMessageInChunks(roomId, results.join('\n\n'));
            }
            return res.status(200).send('ok');
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