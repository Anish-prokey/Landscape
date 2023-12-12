const express = require('express');
const nodemailer = require('nodemailer');
const fs = require('fs');
const pdfkit = require('pdfkit');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const axios = require('axios');
const bodyParser = require('body-parser')
const app = express();
app.use(cors());
app.use(express.json());
const pdf = require('html-pdf');
 
 
 
 
// Generate PDF and Excel files from static table data and send them as attachments via email
app.post('/generate-and-send', async (req, res) => {
  const { recipientEmail, fileFormat, filteredData } = req.body;
  // app.get('https://sapd49.tyson.com/sap/opu/odata/sap/ZAPI_PRDVERS_SRV/LandscapeDetailsSet?sap-client=100')
  console.log(filteredData);
  try {
    // Read table data from JSON file
   
 
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir);
    }
    console.log(tempDir)
    const tableData = filteredData;
 
    let attachmentPath, attachmentType;
    if (fileFormat === 'pdf') {
      const headers = ['SAP System ID','Does It Run On A Hana Database','SAP Product', 'Type'];

      const htmlContent = `
        <html>
          <head>
            <style>
              table {
                width: 100%;
                border-collapse: collapse;
              }
              th, td {
                border: 1px solid black;
                padding: 8px;
                text-align: left;
              }
            </style>
          </head>
          <body>
            <h1>SAP Landscape Table</h1>
            <table>
              <tr>${headers.map(header => `<th>${header}</th>`).join('')}</tr>
              ${filteredData.map(row => `<tr>${Object.values(row).map(value => `<td>${value}</td>`).join('')}</tr>`).join('')}
            </table>
          </body>
        </html>
      `;

      attachmentPath = path.join(tempDir, 'table.pdf');

      pdf.create(htmlContent).toFile(attachmentPath, function (err, res) {
        if (err) throw err;
      });

      attachmentType = 'application/pdf';

    } else if (fileFormat === 'excel') {
      // Generate Excel
      attachmentPath = path.join(tempDir, 'table.xlsx');
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir);
      }
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet 1');
      worksheet.columns = [
        { header: 'SAP Product', key: 'Sapproduct', width: 20 },
        { header: 'SAP System ID', key: 'Sysid', width: 20 },
        
        { header: 'Type', key: 'Systype', width: 20 },
       
        { header: 'Does it run on a Hana database', key: 'Runhdb', width: 30 },
        
      ];
      tableData.forEach((row) => {
        worksheet.addRow(row);
      });
      await workbook.xlsx.writeFile(attachmentPath);
      // await workbook.xlsx.writeFile(attachmentPath);
      attachmentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else {
      throw new Error('Invalid file format');
    }
 
    // Send email with generated file as attachment
    const transporter = nodemailer.createTransport({
      host: 'smtp.gmail.com',
      port: 587,
      auth: {
        user: 'spsr7109@gmail.com',
        pass: 'ozlenukwwdvfqucn',
      },
    });
    console.log("reached here");
    const mailOptions = {
      from: 'spsr7109@gmail.com',
      to: recipientEmail,
      subject: 'Table Data',
      text: 'Please find the attached file with table data.',
      attachments: [
        {
          filename: `table.${fileFormat === "excel" ? "xlsx" : "pdf"}`,
          path: attachmentPath,
          contentType: attachmentType,
        },
      ],
    };
 
    await transporter.sendMail(mailOptions);
 
    // Delete temporary file
    fs.unlinkSync(attachmentPath);
 
    res.json({ message: 'sent' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Failed to send email' });
  }
});

 
app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
 
