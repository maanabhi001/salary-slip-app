const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 5000;

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Serve static files from the frontend directory
app.use(express.static(path.join(__dirname, '../frontend')));

// Helper function to convert Excel date serial number to a DD-MM-YYYY string
function excelDateToString(excelDate) {
    if (typeof excelDate !== 'number' || isNaN(excelDate)) {
        return excelDate || '';
    }
    const jsDate = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    const day = String(jsDate.getDate()).padStart(2, '0');
    const month = String(jsDate.getMonth() + 1).padStart(2, '0');
    const year = jsDate.getFullYear();

    if (isNaN(year)) {
        return excelDate;
    }

    return `${day}-${month}-${year}`;
}

// Function to get the next non-zero value from an array starting from a specific index
function getNextNonZeroValue(arr, startIndex) {
    for (let i = startIndex; i < arr.length; i++) {
        const val = Number(arr[i][1]);
        if (val > 0) {
            return arr[i];
        }
    }
    return null;
}

app.post('/upload-excel', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
        return res.status(400).send('No file uploaded');
    }

    try {
        const workbook = XLSX.readFile(file.path);
        const companyInfoSheet = workbook.Sheets[workbook.SheetNames[0]];
        const employeeDataSheet = workbook.Sheets[workbook.SheetNames[1]];

        const companyInfo = XLSX.utils.sheet_to_json(companyInfoSheet)[0];
        const data = XLSX.utils.sheet_to_json(employeeDataSheet);

        const outputFolder = path.join(__dirname, 'slips');
        if (!fs.existsSync(outputFolder)) {
            fs.mkdirSync(outputFolder);
        }

        const formatCurrency = (value) => `₹ ${Number(value || 0).toFixed(2)}`;

        await Promise.all(
            data.map((emp) => {
                return new Promise((resolve, reject) => {
                    const doc = new PDFDocument({ margin: 30, size: 'A4' });
                    try {
                        doc.registerFont('Custom', path.join(__dirname, 'fonts', 'NotoSans-Regular.ttf'));
                        doc.font('Custom');
                    } catch (fontError) {
                        console.error('Font loading error:', fontError);
                        doc.font('Helvetica');
                    }

                    const filename = `${emp['Casper Id']}.pdf`;
                    const filePath = path.join(outputFolder, filename);
                    const writeStream = fs.createWriteStream(filePath);
                    doc.pipe(writeStream);

                    let currentY = 20;

                    // Outer Border
                    doc.rect(20, 20, 555, 800).stroke();

                    // Header
                    currentY += 15;
                    doc.fontSize(16).fillColor('black').text(companyInfo['Company Name'], { align: 'center' });
                    doc.fontSize(10).text(companyInfo['Company Address'], { align: 'center' });
                    doc.moveDown(0.5);
                    doc.fontSize(12).text(`Pay Slip for the Month of ${companyInfo['Salary Month']}`, { align: 'center' });
                    currentY = doc.y + 15;

                    // Employee Details Box
                    const detailsBoxStartX = 20;
                    const detailsBoxWidth = 555;
                    const rowPadding = 4; // space between rows
                    const detailsBoxStartY = currentY;

                    const leftDetails = [
                        `Casper ID : ${emp['Casper Id'] || ''}`,
                        `Employee ID : ${emp['Employee Id'] || ''}`,
                        `Employee Name : ${emp['Employee Name'] || ''}`,
                        `Father Name : ${emp['Father Name'] || ''}`,
                        `ESIC Number : ${emp['ESIC Number'] || ''}`,
                        `UAN Number : ${emp['UAN Number'] || ''}`,
                        `Work Place : ${emp['Work Place'] || ''}`
                    ];

                    const rightDetails = [
                        `Date Of Birth : ${excelDateToString(emp['Date Of Birth'])}`,
                        `Designation : ${emp['Designation'] || ''}`,
                        `Date of Joining : ${excelDateToString(emp['Date of Joining'])}`,
                        `Bank Name : ${emp['Bank Name'] || ''}`,
                        `Account Number : ${emp['Account Number'] || ''}`,
                        `IFSC Code : ${emp['IFSC code'] || ''}`,
                        `Total Days in Month : ${emp['Total Days in Month'] || ''}`,
                        `Pay Days : ${emp['Pay Days'] || ''}`,
                        `Over Time (H) : ${emp['Over Time (H)'] || ''}`,
                        `Arrear Days : ${emp['Arrear Days'] || ''}`
                    ];

                    doc.fontSize(8);
                    let detailsY = detailsBoxStartY + 5;
                    const maxDetailLines = Math.max(leftDetails.length, rightDetails.length);
                    const rowYPositions = [detailsY];

                    for (let i = 0; i < maxDetailLines; i++) {
                        const leftText = leftDetails[i] || '';
                        const rightText = rightDetails[i] || '';

                        const leftColX = detailsBoxStartX + 5;
                        const rightColX = detailsBoxStartX + detailsBoxWidth / 2 + 5;
                        const colWidth = detailsBoxWidth / 2 - 10;

                        const leftHeight = doc.heightOfString(leftText, { width: colWidth });
                        const rightHeight = doc.heightOfString(rightText, { width: colWidth });
                        const rowHeight = Math.max(leftHeight, rightHeight);

                        doc.text(leftText, leftColX, detailsY, { width: colWidth });
                        doc.text(rightText, rightColX, detailsY, { width: colWidth });

                        detailsY += rowHeight + rowPadding;
                        rowYPositions.push(detailsY);
                    }

                    const employeeDetailsBoxBottomY = detailsY;

                    // Draw box and inner lines
                    doc.rect(detailsBoxStartX, detailsBoxStartY, detailsBoxWidth, employeeDetailsBoxBottomY - detailsBoxStartY).stroke();

                    for (let i = 1; i < rowYPositions.length - 1; i++) {
                        const y = rowYPositions[i] - rowPadding / 2;
                        doc.moveTo(detailsBoxStartX, y).lineTo(detailsBoxStartX + detailsBoxWidth, y).strokeColor('#bbb').stroke();
                    }
                    // Middle vertical line
                    doc.moveTo(detailsBoxStartX + detailsBoxWidth / 2, detailsBoxStartY).lineTo(detailsBoxStartX + detailsBoxWidth / 2, employeeDetailsBoxBottomY).strokeColor('#bbb').stroke();

                    currentY = employeeDetailsBoxBottomY + 10;

                    // Earnings & Deductions Table
                    const tableStartX = 20;
                    const tableWidth = 555;
                    const tableColSplit = tableStartX + tableWidth / 2;
                    const colWidthLabel = tableWidth / 2 - 95;
                    const colWidthAmount = 80;

                    const tableStartY = currentY;
                    let tableY = tableStartY;

                    // Headers
                    doc.fontSize(9).fillColor('black');
                    const headerText = "Earnings";
                    const headerHeight = doc.heightOfString(headerText, { width: colWidthLabel });
                    doc.text('Earnings', tableStartX + 5, tableY + 5, { width: colWidthLabel });
                    doc.text('Deductions', tableColSplit + 5, tableY + 5, { width: colWidthLabel });
                    tableY += headerHeight + 10;
                    doc.moveTo(tableStartX, tableY).lineTo(tableStartX + tableWidth, tableY).stroke();

                    // Prepare filtered earnings
                    const earnings = [
                        ['Basic', emp['Basic']],
                        ['Leave Encashment', emp['Leave Encashment']],
                        ['Additional Bonus', emp['Additional Bonus']],
                        ['Over Time', emp['Over Time']],
                        ['Shift Allowance', emp['Shift Allowance']],
                        ['Total Others', emp['OTotal Others']],
                        ['Gross', emp['Gross']]
                    ];

                    // Filter earnings: skip zeros, replace with next non-zero
                   const filteredEarnings = earnings;


                    // Similarly for deductions
                    const deductions = [
                        ['PF', emp['PF']],
                        ['ESIC', emp['ESIC']],
                        ['LWF', emp['LWF']],
                        ['Advance Amount', emp['Advance Amount']],
                        ['Total Deduction', emp['Total Deduction']]
                    ];

                    const filteredDeductions = deductions;


                    // Render filtered earnings
                    const maxTableRows = Math.max(filteredEarnings.length, filteredDeductions.length);
                    for (let i = 0; i < maxTableRows; i++) {
                        let rowHeight = 0;
                        const leftItem = filteredEarnings[i] || [];
                        const rightItem = filteredDeductions[i] || [];

                        const leftText = leftItem[0] || '';
                        const rightText = rightItem[0] || '';

                        const leftH = doc.heightOfString(leftText, { width: colWidthLabel });
                        const rightH = doc.heightOfString(rightText, { width: colWidthLabel });
                        rowHeight = Math.max(leftH, rightH, 12);

                        if (leftItem[1] !== undefined) {
                            doc.text(leftItem[0], tableStartX + 5, tableY + 4, { width: colWidthLabel });
                            doc.text(formatCurrency(leftItem[1]), tableColSplit - colWidthAmount - 5, tableY + 4, { width: colWidthAmount, align: 'right' });
                        }
                        if (rightItem[1] !== undefined) {
                            doc.text(rightItem[0], tableColSplit + 5, tableY + 4, { width: colWidthLabel });
                            doc.text(formatCurrency(rightItem[1]), tableStartX + tableWidth - colWidthAmount - 5, tableY + 4, { width: colWidthAmount, align: 'right' });
                        }

                        tableY += rowHeight + rowPadding;
                        // Horizontal line
                        doc.moveTo(tableStartX, tableY).lineTo(tableStartX + tableWidth, tableY).strokeColor('#bbb').stroke();
                    }

                    const tableBottomY = tableY;
                    // Outer rectangle and middle line
                    doc.rect(tableStartX, tableStartY, tableWidth, tableBottomY - tableStartY).stroke();
                    doc.moveTo(tableColSplit, tableStartY).lineTo(tableColSplit, tableBottomY).stroke();

                    currentY = tableBottomY + 10;

                    // Net Payable
                    const netPayBoxHeight = 25;
                    doc.rect(tableStartX, currentY, tableWidth, netPayBoxHeight).stroke();
                    doc.fontSize(10).fillColor('black').text(`Net Payable : ${formatCurrency(emp['Net Payable'] || 0)}`, tableStartX + 5, currentY + 7);
                    currentY += netPayBoxHeight + 10;

                    // Footer
                    doc.fontSize(8).fillColor('#666').text('This is a software-generated document and does not require signature.', 20, 800, {
                        align: 'center', width: 555
                    });

                    doc.end();
                    writeStream.on('finish', resolve);
                    writeStream.on('error', reject);
                });
            })
        );

        // Create zip archive
        const zipPath = path.join(__dirname, 'salary_slips.zip');
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });

        output.on('close', () => {
            res.download(zipPath, 'salary_slips.zip', (err) => {
                if (err) console.error('Error sending zip file:', err);
                fs.unlink(file.path, (e) => e && console.error('Error unlinking uploaded file:', e));
                fs.unlink(zipPath, (e) => e && console.error('Error unlinking zip file:', e));
                fs.rm(outputFolder, { recursive: true, force: true }, (e) => e && console.error('Error removing output folder:', e));
            });
        });

        archive.on('error', (err) => {
            console.error('Archiving error:', err);
            res.status(500).send('Error creating zip archive.');
        });

        archive.pipe(output);
        archive.directory(outputFolder, false);
        archive.finalize();

    } catch (error) {
        console.error('Server error:', error);
        res.status(500).send('Error processing Excel or generating PDFs.');
        if (file && file.path) {
            fs.unlink(file.path, (err) => err && console.error('Error unlinking uploaded file:', err));
        }
        const outputFolder = path.join(__dirname, 'slips');
        if (fs.existsSync(outputFolder)) {
            fs.rm(outputFolder, { recursive: true, force: true }, (err) => err && console.error('Error removing output folder:', err));
        }
    }
});

app.listen(PORT, () => {
    console.log(`✅ Server started at http://localhost:${PORT}`);
});