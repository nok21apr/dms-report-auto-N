const puppeteer = require('puppeteer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');

// ฟังก์ชันสำหรับแปลงวันที่เป็น YYYY-MM-DD (รับค่า Date Object)
function getFormattedDate(date) {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    const thaiDate = new Intl.DateTimeFormat('en-CA', options).format(date);
    return thaiDate;
}

(async () => {
    // --- ส่วนการรับค่าจาก Secrets ---
    const USERNAME = process.env.DTC_USERNAME;
    const PASSWORD = process.env.DTC_PASSWORD;
    const EMAIL_USER = process.env.EMAIL_USER;
    const EMAIL_PASS = process.env.EMAIL_PASS;
    const EMAIL_TO   = process.env.EMAIL_TO;

    if (!USERNAME || !PASSWORD || !EMAIL_USER || !EMAIL_PASS || !EMAIL_TO) {
        console.error('Error: Missing required secrets.');
        process.exit(1);
    }

    console.log('Launching browser...');
    const downloadPath = path.resolve('./downloads');
    if (!fs.existsSync(downloadPath)) {
        fs.mkdirSync(downloadPath);
    }

    const browser = await puppeteer.launch({
        headless: true, // ตั้งเป็น false เพื่อดูการทำงานตอนเทสได้
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--start-maximized'
        ]
    });
    
    const page = await browser.newPage();
    
    // --- Setup ---
    // Timeout 5 นาที (300000 ms) เพียงพอสำหรับการรอ 120 วินาที
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);

    await page.emulateTimezone('Asia/Bangkok');
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });

    await page.setViewport({ width: 1920, height: 1080 });

    try {
        // ---------------------------------------------------------
        // Step 1: Login
        // ---------------------------------------------------------
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', USERNAME);
        await page.type('#txtpass', PASSWORD);
        
        console.log('   Clicking Login...');
        await Promise.all([
            page.evaluate(() => {
                const btn = document.getElementById('btnLogin');
                if(btn) btn.click();
            }),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        // ---------------------------------------------------------
        // Step 2: Navigate to Report (Direct URL)
        // ---------------------------------------------------------
        console.log('2️⃣ Step 2: Go to Report Page (Direct URL)...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/report_other_status.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true, timeout: 60000 });
        console.log('✅ Report Page Loaded');

        // ---------------------------------------------------------
        // Step 2.5: Select Truck "ทั้งหมด" (Direct DOM Method)
        // ---------------------------------------------------------
        console.log('   Selecting Truck "ทั้งหมด"...');
        await page.waitForSelector('#ddl_truck', { visible: true, timeout: 60000 });

        // รอให้ Option โหลดมา
        await page.waitForFunction(() => {
            const select = document.getElementById('ddl_truck');
            return select && select.options.length > 0;
        }, { timeout: 60000 });

        await page.evaluate(() => {
            var selectElement = document.getElementById('ddl_truck'); 
            if (selectElement) {
                var options = selectElement.options; 
                for (var i = 0; i < options.length; i++) { 
                    if (options[i].text.includes('ทั้งหมด') || options[i].text.toLowerCase().includes('all')) { 
                        selectElement.value = options[i].value; 
                        var event = new Event('change', { bubbles: true });
                        selectElement.dispatchEvent(event);
                        break; 
                    } 
                }
            }
        });
        console.log('✅ Truck "ทั้งหมด" Selected');

        // ---------------------------------------------------------
        // Step 2.6: Select Report Types (Using #s2id_ddlharsh)
        // ---------------------------------------------------------
        console.log('   Selecting 3 Report Types (via #s2id_ddlharsh)...');
        
        const select2ContainerSelector = '#s2id_ddlharsh';
        const select2InputSelector = '#s2id_ddlharsh input'; 
        
        const searchKeywords = [
            "ความง่วงระดับ 1", 
            "ความง่วงระดับ 2",
            "หาว"       
        ];

        try {
            await page.waitForSelector(select2ContainerSelector, { visible: true, timeout: 30000 });
            
            for (const keyword of searchKeywords) {
                console.log(`      Processing "${keyword}"...`);
                await page.click(select2ContainerSelector);
                await new Promise(r => setTimeout(r, 500)); 

                const inputHandle = await page.$(select2InputSelector) || await page.$('.select2-input');
                
                if (inputHandle) {
                    await inputHandle.type(keyword);
                    await new Promise(r => setTimeout(r, 1000));
                    await page.keyboard.press('Enter');
                    console.log(`      Selected: "${keyword}"`);
                    await new Promise(r => setTimeout(r, 500));
                } else {
                    console.log(`      ⚠️ Could not find input field inside ${select2ContainerSelector}`);
                }
            }

        } catch (e) {
            console.log('      ❌ Error Selecting Report Types:', e.message);
        }
        
        console.log('✅ Report Types Selection Finished');

        // ---------------------------------------------------------
        // Step 3: Setting Date Range 18:00 (Yesterday) - 06:00 (Today)
        // ---------------------------------------------------------
        console.log('3️⃣ Step 3: Setting Date Range 18:00 (Yesterday) - 06:00 (Today)...');
        
        const now = new Date();
        const todayStr = getFormattedDate(now); // วันที่ปัจจุบัน (เช่น 15)

        const yesterday = new Date(now);
        yesterday.setDate(yesterday.getDate() - 1); // ย้อนหลัง 1 วัน (เช่น 14)
        const yesterdayStr = getFormattedDate(yesterday);

        // กำหนดเวลาเริ่ม: เมื่อวาน 18:00
        const startDateTime = `${yesterdayStr} 18:00`;
        // กำหนดเวลาจบ: วันนี้ 06:00 (วันรุ่งขึ้นของ 18:00)
        const endDateTime = `${todayStr} 06:00`;
        
        console.log(`      Range: ${startDateTime} to ${endDateTime}`);

        await page.evaluate(() => document.getElementById('date9').value = '');
        await page.type('#date9', startDateTime);

        await page.evaluate(() => document.getElementById('date10').value = '');
        await page.type('#date10', endDateTime);
        
        console.log('   Clicking Search to update report...');
        try {
            // ใช้ Selector จากไฟล์ Recording ที่คุณส่งมา: td:nth-of-type(5) > span
            const searchSelector = 'td:nth-of-type(5) > span';
            
            // รอให้ปุ่มค้นหาพร้อม
            await page.waitForSelector(searchSelector, { visible: true, timeout: 10000 });
            await page.click(searchSelector);
            
            // --- NEW: Wait 120 Seconds ---
            console.log('   ⏳ Waiting 120 seconds for report generation...');
            // รอ 120,000 ms (120 วินาที)
            await new Promise(r => setTimeout(r, 120000)); 
            console.log('   ✅ Wait complete.');

        } catch (e) {
            console.log('⚠️ Warning: Could not click Search button or wait failed.', e.message);
        }

        // ---------------------------------------------------------
        // Step 4: Export Excel
        // ---------------------------------------------------------
        console.log('4️⃣ Step 4: Clicking Export/Excel...');
        
        cleanDownloadFolder(downloadPath);

        const excelBtnSelector = '#btnexport, button[title="Excel"], ::-p-aria(Excel)';
        
        // รอให้ปุ่มปรากฏและคลิกได้
        await page.waitForSelector(excelBtnSelector, { visible: true, timeout: 60000 });
        
        console.log('   Clicking Export Button...');
        await page.evaluate(() => {
            const btn = document.querySelector('#btnexport') || document.querySelector('button[title="Excel"]');
            if(btn) btn.click();
        });
        
        // เพิ่มเวลารอ Download ให้นานขึ้น
        console.log('   ⏳ Waiting for download (30s)...');
        await new Promise(r => setTimeout(r, 30000));

        // ---------------------------------------------------------
        // Step 5: Email & Cleanup
        // ---------------------------------------------------------
        console.log('5️⃣ Step 5: Processing email...');
        
        const recentFile = getMostRecentFile(downloadPath);
        
        if (recentFile) {
            const filePath = path.join(downloadPath, recentFile.file);
            const fileName = recentFile.file;
            
            // เปลี่ยนหัวข้ออีเมลให้ตรงกับช่วงเวลาใหม่
            const subjectLine = `${fileName} ช่วง1800ถึง0600`;

            await sendEmail({
                user: EMAIL_USER,
                pass: EMAIL_PASS,
                to: EMAIL_TO,
                subject: subjectLine,
                attachmentPath: filePath
            });

            console.log('   Deleting downloaded file...');
            try {
                fs.unlinkSync(filePath);
                console.log('✅ File deleted successfully.');
            } catch (err) {
                console.error('⚠️ Error deleting file:', err);
            }

        } else {
            console.log('❌ No file downloaded to send.');
            // ลองถ่ายรูปหน้าจอตอนจบเผื่อดู error
            await page.screenshot({ path: 'final_no_file.png' });
            throw new Error('Download failed or no file found');
        }

        console.log('🎉 Script completed successfully.');

    } catch (error) {
        console.error('❌ Error occurred:', error);
        await page.screenshot({ path: 'error_screenshot.png' });
        process.exit(1);
    } finally {
        await browser.close();
    }
})();

async function sendEmail({ user, pass, to, subject, attachmentPath }) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user, pass }
    });

    const mailOptions = {
        from: user,
        to: to,
        subject: subject,
        text: 'รายงาน DMS ประจำช่วงเวลา 18:00 (เมื่อวาน) - 06:00 (วันนี้)\n\n(Auto-generated email)',
        attachments: attachmentPath ? [{ path: attachmentPath }] : []
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('📧 Email sent: ' + info.response);
}

const getMostRecentFile = (dir) => {
    try {
        const files = fs.readdirSync(dir);
        const validFiles = files.filter(file => fs.lstatSync(path.join(dir, file)).isFile() && !file.startsWith('.'));
        if (validFiles.length === 0) return null;
        return validFiles
            .map(file => ({ file, mtime: fs.lstatSync(path.join(dir, file)).mtime }))
            .sort((a, b) => b.mtime.getTime() - a.mtime.getTime())[0];
    } catch (e) { return null; }
};

const cleanDownloadFolder = (dir) => {
    try {
        if (fs.existsSync(dir)) {
            const files = fs.readdirSync(dir);
            for (const file of files) {
                fs.unlinkSync(path.join(dir, file));
            }
        }
    } catch (e) {}
};