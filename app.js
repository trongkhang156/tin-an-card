const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const puppeteer = require('puppeteer-core');
const chromium = require('chromium');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.get('/', (req, res) => {
  res.send(`
  <!DOCTYPE html>
  <html lang="vi">
  <head>
    <meta charset="UTF-8" />
    <title>Upload Excel - Tín An</title>
    <style>
      body {
        font-family: Arial, sans-serif;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        background: #f4f4f4;
        margin: 0;
      }
      .upload-card {
        background: white;
        padding: 40px;
        border-radius: 10px;
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        text-align: center;
        width: 400px;
      }
      .upload-card h2 {
        color: #001f80;
        margin-bottom: 25px;
      }
      input[type="file"] {
        padding: 10px;
        border-radius: 5px;
        border: 1px solid #ccc;
        width: 100%;
        margin-bottom: 20px;
      }
      button {
        background: #0052cc;
        color: white;
        font-weight: bold;
        border: none;
        padding: 12px 25px;
        border-radius: 5px;
        cursor: pointer;
        transition: 0.3s;
      }
      button:hover {
        background: #003d99;
      }
    </style>
  </head>
  <body>
    <div class="upload-card">
      <h2>Upload danh sách nhân viên</h2>
      <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="excel" accept=".xlsx,.xls" required />
        <button type="submit">Tải lên và tải PDF</button>
      </form>
    </div>
  </body>
  </html>
  `);
});

app.post('/upload', upload.single('excel'), async (req, res) => {
  if (!req.file) return res.status(400).send('Không có file được tải lên');
  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.xls', '.xlsx'].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).send('File tải lên không phải định dạng Excel (.xls, .xlsx)');
  }

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const employees = XLSX.utils.sheet_to_json(worksheet);

    let html = `
    <!DOCTYPE html>
    <html lang="vi">
    <head>
      <meta charset="UTF-8" />
      <title>Thẻ nhân viên - Tín An</title>
      <style>
        @page { size: A4; margin: 0mm; }
        body {
          font-family: Arial, sans-serif;
          padding: 0;
          margin: 0;
          background: white;
        }
        .container {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 3mm;
          padding: 5mm;
        }
        .card {
          display: flex;
          width: 90%;
          height: 55mm;
          background: white;
          border: 1px solid black;
          box-sizing: border-box;
          page-break-inside: avoid;
          padding: 5px;
          gap: 5px;
        }
        .left, .right {
          flex: 1;
          border: 1px solid black;
          display: flex;
          flex-direction: column;
          box-sizing: border-box;
        }
        .left {
          display: grid;
          grid-template-columns: 80px auto;
          grid-template-rows: 50px auto 25px;
          padding: 5px;
          gap: 3px;
        }
        .header-left {
          grid-column: 1 / 3;
          display: flex;
          align-items: center;
          gap: 5px;
          padding-left: 3px;
        }
        .logo {
          width: 60px;
          height: 45px;
          object-fit: contain;
        }
        .company-info .name1 {
          color: #001f80;
          font-weight: bold;
          text-align:center;
          font-size: 16px;
        }
        .company-info .name2 {
          color: red;
          font-weight: bold;
          font-size: 22px;
          text-align:center;
          line-height: 1.1;
        }

        /* ✅ Làm cho ảnh và ô thông tin bằng chiều cao */
        .photo-box, .info-box {
          height: 90px;
          display: flex;
          align-items: center;
          justify-content: center;
          box-sizing: border-box;
        }

        .photo-box {
          border: 1px solid black;
          width: 65px;
          font-size: 10px;
        }

        .info-box {
          background: #0052cc;
          color: white;
          text-align: center;
          padding: 5px 3px;
          flex-direction: column;
        }

        .emp-name {
          font-weight: bold;
          font-size: 17px;
          word-break: break-word;
        }
        .emp-title {
          margin-top: 2px;
          font-size: 15px;
        }

        .emp-id {
          background: red;
          color: white;
          font-weight: bold;
          font-size: 20px;
          text-align: center;
          grid-column: 1 / 3;
          display: flex;
          justify-content: center;
          align-items: center;
        }

        .right { font-size: 10px; }
        .rules-header {
          background: red;
          color: white;
          font-weight: bold;
          font-size: 20px;
          text-align: center;
          padding: 2px 0;
        }
        .rules {
          background: #0052cc;
          color: white;
          flex-grow: 1;
          padding: 5px 10px;
          font-size: 15px;
        }
        .rules ol { margin: 0; padding-left: 15px; }
        .contact {
          background: #0073e6;
          color: white;
          text-align: center;
          font-size: 15px;
          padding: 2px 5px;
          line-height: 1.2;
        }
        .footer {
          background: red;
          color: white;
          text-align: center;
          font-weight: bold;
          font-size: 20px;
          padding: 2px 0;
        }
        @media print {
          .card { page-break-inside: avoid; }
          .card:nth-child(5n) { page-break-after: always; }
        }
      </style>
    </head>
    <body>
      <div class="container">
    `;

    employees.forEach(emp => {
      html += `
        <div class="card">
          <div class="left">
            <div class="header-left">
              <img src="https://cdn-new.topcv.vn/unsafe/https://static.topcv.vn/company_logos/gz2QKmW8jIO5zGz1dn0PxojiHlgMRTIX_1702439899____2e4eada971dc1d243b704375033db719.png" class="logo" />
              <div class="company-info">
                <div class="name1">CÔNG TY CỔ PHẦN ĐT-SX-TM</div>
                <div class="name2">TÍN AN</div>
              </div>
            </div>
            <div class="photo-box">ẢNH 3X4</div>
            <div class="info-box">
              <div class="emp-name">${emp["Họ và tên"] || ''}</div>
              <div class="emp-title">${emp["Chức vụ"] || ''}</div>
            </div>
            <div class="emp-id">MSNV: ${emp["MSNV"] || ''}</div>
          </div>
          <div class="right">
            <div class="rules-header">QUY ĐỊNH</div>
            <div class="rules">
              <ol>
                <li>CB-CNV phải đeo thẻ khi làm việc</li>
                <li>Không được cho người khác sử dụng thẻ</li>
                <li>Phải quẹt thẻ khi ra vào cổng</li>
                <li>Khi mất thẻ phải báo ngay về P.HCNS</li>
              </ol>
            </div>
            <div class="contact">
              Lô B1, Đường D3, KCN Đồng An 2, P. Hòa Phú, TP. TDM<br>ĐT: 0274 386 6661
            </div>
            <div class="footer">TIN AN JSC</div>
          </div>
        </div>
      `;
    });

    html += `</div></body></html>`;

    const browser = await puppeteer.launch({
  executablePath: chromium.path,
  headless: true,
  args: ['--no-sandbox', '--disable-setuid-sandbox']
});
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '0mm', right: '0mm', bottom: '0mm', left: '0mm' },
    });
    await browser.close();
    fs.unlinkSync(req.file.path);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': 'attachment; filename="the_nhan_vien.pdf"',
      'Content-Length': pdfBuffer.length,
    });
    res.send(pdfBuffer);

  } catch (err) {
    console.error(err);
    if (req.file?.path) fs.unlinkSync(req.file.path);
    res.status(500).send('Lỗi khi tạo PDF');
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
});
