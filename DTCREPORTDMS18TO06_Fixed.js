const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs'); // เปลี่ยนมาใช้ exceljs
const { JSDOM } = require('jsdom');

// ฟังก์ชันแปลงวันที่
function getFormattedDate(date) {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

// ฟังก์ชันแปลง HTML เป็น Excel แบบสวยงาม (ใช้ ExcelJS)
async function convertHtmlToExcel(sourcePath, destPath) {
    try {
        console.log(`   Converting HTML-XLS to Beautiful XLSX (ExcelJS)...`);
        
        // 1. อ่านและ Parse HTML
        const htmlContent = fs.readFileSync(sourcePath, 'utf-8');
        const dom = new JSDOM(htmlContent);
        const table = dom.window.document.querySelector('table');
        
        if (!table) throw new Error('No table found in downloaded file');

        // 2. สร้าง Workbook และ Worksheet ใหม่
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('DMS Report');

        // 3. ดึงข้อมูลจาก HTML Table
        const rows = Array.from(table.querySelectorAll('tr'));
        
        // วนลูปทีละแถว
        rows.forEach((row, rowIndex) => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            const rowData = cells.map(cell => {
                // Clean data: ลบ html tags และ space ส่วนเกิน
                return cell.textContent.replace(/<[^>]*>/g, '').trim();
            });
            
            // เพิ่มแถวลงใน Excel
            const excelRow = worksheet.addRow(rowData);

            // --- ตกแต่ง (Styling) ---
            excelRow.eachCell((cell, colNumber) => {
                // ใส่เส้นขอบ (Borders) ทุกช่อง
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };

                // จัด Alignment (กึ่งกลางแนวตั้ง, ซ้ายแนวนอน)
                cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

                // ถ้าเป็น Header (แถวแรก) หรือ tag <th>
                if (rowIndex === 0 || cells[colNumber-1].tagName === 'TH') {
                    cell.font = { bold: true, color: { argb: '000000' } }; // ตัวหนา
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFD3D3D3' } // พื้นหลังสีเทาอ่อน
                    };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' }; // Header จัดกึ่งกลาง
                }
            });
        });

        // 4. Auto-fit Column Width (คำนวณความกว้าง)
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            if (column.values) {
                column.values.forEach(val => {
                    const len = val ? String(val).length : 0;
                    if (len > maxLength) maxLength = len;
                });
            }
            // เผื่อความกว้างนิดหน่อย (Limit ไม่ให้เกิน 50)
            column.width = Math.min(Math.max(maxLength + 2, 10), 50);
        });

        // 5. บันทึกไฟล์
        await workbook.xlsx.writeFile(destPath);
        console.log(`   Conversion success: ${destPath}`);
        return true;

    } catch (e) {
        console.error(`   Conversion failed: ${e.message}`);
        return false;
    }
}

(async () => {
    // --- ส่วนการรับค่าจาก Secrets ---
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;

    if (!DTC_USERNAME || !DTC_PASSWORD || !EMAIL_USER || !EMAIL_PASS || !EMAIL_TO) {
        console.error('Error: Missing required secrets.');
        process.exit(1);
    }

    console.log('Launching browser...');
    const downloadPath = path.resolve('./downloads');
    
    // Force clean folder
    if (fs.existsSync(downloadPath)) {
        try { fs.rmSync(downloadPath, { recursive: true, force: true }); } catch(e) {}
    }
    fs.mkdirSync(downloadPath);

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });
    
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);
    await page.emulateTimezone('Asia/Bangkok');

    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // Step 1: Login
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        // Step 2: Navigate
        console.log('2️⃣ Step 2: Go to Report Page...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#date9', { visible: true, timeout: 60000 });

        // Step 2.5: Truck
        console.log('   Selecting Truck "ทั้งหมด"...');
        await page.waitForSelector('#ddl_truck', { visible: true });
        await page.waitForFunction(() => document.getElementById('ddl_truck').options.length > 0);
        await page.evaluate(() => {
            const select = document.getElementById('ddl_truck');
            for (let opt of select.options) {
                if (opt.text.includes('ทั้งหมด') || opt.text.toLowerCase().includes('all')) {
                    select.value = opt.value;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
        });

        // Step 2.6: Report Types
        console.log('   Selecting 3 Report Types...');
        try {
            await page.waitForSelector('#s2id_ddlharsh', { visible: true });
            const keywords = ["ความง่วงระดับ 1", "ความง่วงระดับ 2", "หาว"];
            for (const kw of keywords) {
                await page.click('#s2id_ddlharsh'); 
                await new Promise(r => setTimeout(r, 500));
                const input = await page.$('#s2id_ddlharsh input') || await page.$('.select2-input');
                if (input) {
                    await input.type(kw);
                    await new Promise(r => setTimeout(r, 1000));
                    await page.keyboard.press('Enter');
                    console.log(`      Selected: "${kw}"`);
                    await new Promise(r => setTimeout(r, 500));
                }
            }
        } catch (e) { console.log('❌ Error selecting reports:', e.message); }

        // Step 3: Date Range 18:00 - 06:00
        console.log('3️⃣ Step 3: Setting Date Range...');
        const now = new Date();
        const yesterday = new Date(now);
        yesterday.setDate(yesterday.getDate() - 1);
        const start = `${getFormattedDate(yesterday)} 18:00`;
        const end = `${getFormattedDate(now)} 06:00`;
        
        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', start);
        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', end);

        console.log('   Clicking Search & Waiting 120s...');
        await page.click('td:nth-of-type(5) > span');
        await new Promise(r => setTimeout(r, 120000)); 

        // Step 4: Export
        console.log('4️⃣ Step 4: Clicking Export/Excel...');
        
        // Force Clean again
        try {
            if (fs.existsSync(downloadPath)) {
                fs.rmSync(downloadPath, { recursive: true, force: true });
                fs.mkdirSync(downloadPath);
            }
        } catch(e) {}

        await page.waitForSelector('#btnexport', { visible: true });
        await page.evaluate(() => document.querySelector('#btnexport').click());
        console.log('   Waiting for download (30s)...');
        await new Promise(r => setTimeout(r, 30000));

        // Step 5: Convert & Email
        console.log('5️⃣ Step 5: Preparing Email...');
        const files = fs.readdirSync(downloadPath).filter(f => !f.startsWith('.'));
        if (files.length > 0) {
            const recentFile = files.map(f => ({ 
                name: f, 
                time: fs.statSync(path.join(downloadPath, f)).mtime.getTime() 
            })).sort((a, b) => b.time - a.time)[0];

            const originalPath = path.join(downloadPath, recentFile.name);
            const newFileName = recentFile.name.replace(/\.xls$/, '') + '.xlsx';
            const newFilePath = path.join(downloadPath, newFileName);
            
            // Convert using ExcelJS (ต้อง await)
            const converted = await convertHtmlToExcel(originalPath, newFilePath);
            
            const fileToSend = converted ? newFilePath : originalPath;
            const nameToSend = converted ? newFileName : recentFile.name;

            const subject = `${recentFile.name} ช่วง1800ถึง0600`;

            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            console.log(`   Sending email to: ${EMAIL_TO}`);
            await transporter.sendMail({
                from: `"DTC DMS Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: subject,
                text: 'ถึง ผู้เกี่ยวข้อง\nรายงาน DMS กะกลางคืน (18:00 - 06:00)\nด้วยความนับถือ\nBot Report`,
                attachments: [{ filename: nameToSend, path: fileToSend }]
            });
            console.log('📧 Email Sent!');
            
            try { fs.unlinkSync(originalPath); } catch(e){}
            if(converted) try { fs.unlinkSync(newFilePath); } catch(e){}
        } else {
            console.error('❌ No file found!');
            process.exit(1);
        }

    } catch (err) {
        console.error('❌ Error:', err);
        await page.screenshot({ path: 'error.png' });
        process.exit(1);
    } finally {
        await browser.close();
    }
})();


