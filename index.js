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

// Helper function to format a row of data into a readable string
function formatRow(rowObj, headerRow2, index, sheetName) {
    const rowArray = Object.values(rowObj);
    return `📄 Found in sheet: ${sheetName} (Row ${index + 3})\n` +
        `📝 Task Name: ${flattenText(rowObj['ชื่องาน'])} | 🧾 WBS: ${flattenText(rowObj['WBS'])}\n` +
        `💰 Payment/Week: ${flattenText(rowObj['ชำระเงิน/ลว.'])} | ✅ Approval/Week: ${flattenText(rowObj['อนุมัติ/ลว.'])} | 📂 File Received: ${flattenText(rowObj['รับแฟ้ม'])}\n` +
        `🔌 Transformer: ${flattenText(rowObj['หม้อแปลง'])} | ⚡ HT Distance: ${getCellByHeader2(rowArray, headerRow2, 'HT')} | ⚡ LT Distance: ${getCellByHeader2(rowArray, headerRow2, 'LT')}\n` +
        `🪵 Pole 8: ${getCellByHeader2(rowArray, headerRow2, '8')} | 🪵 Pole 9: ${getCellByHeader2(rowArray, headerRow2, '9')} | 🪵 Pole 12: ${getCellByHeader2(rowArray, headerRow2, '12')} | 🪵 Pole 12.20: ${getCellByHeader2(rowArray, headerRow2, '12.20')}\n` +
        `👷‍♂️ Supervisor: ${flattenText(rowObj['พชง.ควบคุม'])}\n` +
        `📌 Status: ${flattenText(rowObj['สถานะงาน'])} | 📊 Progress: ${flattenText(rowObj['เปอร์เซ็นงาน'])}\n` +
        `🗒️ Notes: ${flattenText(rowObj['หมายเหตุ'])}`;
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

// Function to send a message in chunks if it's too long
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
            if (error.response && error.response.data && error.response.data.errors) {
                error.response.data.errors.forEach(err => console.error('   Webex API Error:', err.description));
            }
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
        } else {
            console.warn('Warning: No data or invalid data format provided for Excel file. Creating an empty Excel file.');
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
            responseText = `📌 Available Commands:\n` +
                `1. @${WEBEX_BOT_NAME} ค้นหา <keyword> → Search all sheets\n` +
                `2. @${WEBEX_BOT_NAME} ค้นหา <sheet_name> → Get all data in the sheet (attach file if long)\n` +
                `3. @${WEBEX_BOT_NAME} แก้ไข <sheet_name> <column_name> <row_number> <new_value>\n` +
                `4. @${WEBEX_BOT_NAME} help → Show all commands`;
        } else if (command === 'ค้นหา') {
            if (allSheetNames.includes(keyword)) {
                const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, keyword);
                const resultText = data.map((row, idx) => formatRow(row, rawHeaders2, idx, keyword)).join('\n\n');

                const EXCEL_THRESHOLD_SPECIFIC_SHEET = 500;
                if (resultText.length > EXCEL_THRESHOLD_SPECIFIC_SHEET) {
                    await axios.post('https://webexapis.com/v1/messages', {
                        roomId,
                        markdown: '📎 Data too long. Attaching Excel file instead...'
                    }, {
                        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
                    });

                    const excelData = data.map(row => {
                        const rowData = {};
                        Object.keys(row).forEach(header => {
                            rowData[header] = row[header];
                        });
                        return rowData;
                    });

                    await sendFileAttachment(roomId, 'Data.xlsx', excelData);
                    return res.status(200).send('sent file');
                } else {
                    responseText = resultText;
                }
            } else {
                let results = [];
                let rawExcelDataForSearch = [];

                for (const sheetName of allSheetNames) {
                    const { data, rawHeaders2 } = await getSheetWithHeaders(sheets, GOOGLE_SHEET_FILE_ID, sheetName);
                    data.forEach((row, idx) => {
                        const match = Object.values(row).some(v => flattenText(v).includes(keyword));
                        if (match) {
                            results.push(formatRow(row, rawHeaders2, idx, sheetName));
                            const rowData = {};
                            Object.keys(row).forEach(header => {
                                rowData[header] = row[header];
                            });
                            rawExcelDataForSearch.push(rowData);
                        }
                    });
                }

                responseText = results.length ? results.join('\n\n') : '❌ No data found';

                if (responseText.length > EXCEL_THRESHOLD_GENERAL_SEARCH) {
                    await axios.post('https://webexapis.com/v1/messages', {
                        roomId,
                        markdown: '📎 Data too long. Attaching Excel file instead (from multiple sheets)...'
                    }, {
                        headers: { Authorization: `Bearer ${WEBEX_BOT_TOKEN}` }
                    });

                    if (rawExcelDataForSearch.length > 0) {
                        await sendFileAttachment(roomId, 'Search_Results.xlsx', rawExcelDataForSearch);
                        return res.status(200).send('sent file');
                    } else {
                        responseText = '❌ No data found';
                        await sendMessageInChunks(roomId, responseText);
                        return res.status(200).send('OK');
                    }
                }
            }
        } else if (command === 'แก้ไข') {
            if (args.length < 3) {
                responseText = '❗ Incorrect command format: แก้ไข <sheet_name> <column_name> <row_number> <new_value>';
            } else {
                let sheetName = args[0];
                let columnName = args[1];
                let rowIndex = parseInt(args[2]);
                let newValue = args.slice(3).join(' ');

                if (allSheetNames.includes(sheetName)) {
                    const res = await sheets.spreadsheets.values.get({
                        spreadsheetId: GOOGLE_SHEET_FILE_ID,
                        range: `${sheetName}!A1:Z`
                    });

                    const rows = res.data.values || [];
                    if (rowIndex >= 1 && (rowIndex + 1) < rows.length) {
                        const header1 = rows[0];
                        const header2 = rows[1];
                        const headers = header1.map((h1, i) => `${h1} ${header2[i] || ''}`.trim());
                        const colIndex = headers.findIndex(h => h.includes(columnName));

                        if (colIndex === -1) {
                            responseText = `❌ Column "${columnName}" not found`;
                        } else {
                            const colLetter = String.fromCharCode(65 + colIndex);
                            const range = `${sheetName}!${colLetter}${rowIndex + 2}`;

                            await sheets.spreadsheets.values.update({
                                spreadsheetId: GOOGLE_SHEET_FILE_ID,
                                range,
                                valueInputOption: 'USER_ENTERED',
                                requestBody: { values: [[newValue]] }
                            });

                            responseText = `✅ Updated: ${range} → ${newValue}`;
                        }
                    } else {
                        responseText = `❌ Row ${rowIndex} not found or invalid (should be a data row, e.g., 1, 2, ...)`;
                    }
                } else {
                    responseText = `❌ Sheet "${sheetName}" not found`;
                }
            }
        } else {
            responseText = `❓ Unknown command. Try "@${WEBEX_BOT_NAME} help"`;
        }

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
app.listen(PORT, () => console.log(`🚀 Bot is ready on port ${PORT}`));