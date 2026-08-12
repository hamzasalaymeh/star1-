const puppeteer = require('puppeteer');
const fs = require('fs').promises;
const path = require('path');

class AdvancedTablePDFExporter {
  constructor(username, password, options = {}) {
    this.username = username;
    this.password = password;
    this.browser = null;
    this.page = null;

    this.config = {
      baseUrl: options.baseUrl || 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10',
      outputDir: options.outputDir || './exports',
      headless: options.headless !== false,
      waitForNavigation: options.waitForNavigation || 'networkidle2',
      navigationTimeout: options.navigationTimeout || 30000,
      maxRetries: options.maxRetries || 3,
      ...options,
    };
  }

  async init() {
    this.browser = await puppeteer.launch({
      headless: this.config.headless,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
      ],
    });

    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(this.config.navigationTimeout);

    // Set viewport for consistent rendering
    await this.page.setViewport({ width: 1920, height: 1080 });

    await fs.mkdir(this.config.outputDir, { recursive: true });
  }

  async login() {
    console.log('🔐 جارٍ تسجيل الدخول...');

    await this.page.goto(this.config.baseUrl, {
      waitUntil: this.config.waitForNavigation,
    });

    // Detect and fill login inputs dynamically
    await this.page.evaluate((username, password) => {
      const inputs = document.querySelectorAll('input');
      const usernameInput = Array.from(inputs).find(
        i => i.type === 'text' || i.name?.includes('user') || i.placeholder?.includes('user')
      );
      const passwordInput = Array.from(inputs).find(
        i => i.type === 'password' || i.name?.includes('pass')
      );

      if (usernameInput) usernameInput.value = username;
      if (passwordInput) passwordInput.value = password;
    }, this.username, this.password);

    // Submit form
    const submitButton = await this.page.$(
      'button[type="submit"], .login-btn, [class*="signin"], [class*="login"]'
    );

    if (submitButton) {
      await submitButton.click();
      await this.page.waitForNavigation({ waitUntil: this.config.waitForNavigation });
    }

    console.log('✅ تم تسجيل الدخول');
  }

  async selectDecisionNumber(decisionNumber) {
    console.log(`📋 جارٍ إدخال رقم قرار التسجيل: ${decisionNumber}`);

    try {
      // استخدام الـ selector الفعلي من الـ HTML
      const input = await this.page.$('input[name="ctl12$ctl28"]');

      if (input) {
        await input.click();
        await this.page.keyboard.press('Control+A');
        await this.page.keyboard.type(decisionNumber.toString());

        // Trigger change event
        await this.page.evaluate(() => {
          const input = document.querySelector('input[name="ctl12$ctl28"]');
          if (input) {
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
          }
        });

        console.log(`✅ تم إدخال رقم القرار: ${decisionNumber}`);
      } else {
        console.error('❌ لم يتم العثور على حقل رقم القرار');
      }

      // انتظر قليلاً للبحث
      await this.page.waitForTimeout(1500);
    } catch (error) {
      console.error(`⚠️ خطأ في إدخال رقم القرار: ${error.message}`);
    }
  }

  async clickSearchButton() {
    console.log('🔍 جارٍ الضغط على زر البحث...');

    try {
      const searchBtn = await this.page.$('#ctl12_btnSearch');
      if (searchBtn) {
        await searchBtn.click();
        // انتظر تحديث الجدول
        await this.page.waitForTimeout(2000);
        console.log('✅ تم الضغط على البحث');
      } else {
        console.error('❌ لم يتم العثور على زر البحث');
      }
    } catch (error) {
      console.error(`⚠️ خطأ في الضغط على البحث: ${error.message}`);
    }
  }

  async openFirstForm() {
    console.log('📂 جارٍ فتح الاستمارة الأولى...');

    try {
      // البحث عن أول رابط استمارة في الجدول (DataEdit/xxx)
      const firstFormLink = await this.page.$('a[href*="DataEdit/"]');

      if (firstFormLink) {
        const href = await this.page.evaluate(el => el.getAttribute('href'), firstFormLink);
        console.log(`   الرابط: ${href}`);

        // ضغط على الرابط
        await firstFormLink.click();

        // انتظر تحميل الصفحة الجديدة
        await this.page.waitForNavigation({ waitUntil: this.config.waitForNavigation });
        console.log('✅ تم فتح الاستمارة');
      } else {
        console.error('❌ لم يتم العثور على استمارة في الجدول');
      }
    } catch (error) {
      console.error(`⚠️ خطأ في فتح الاستمارة: ${error.message}`);
    }
  }

  async navigateToClassificationCriteria() {
    console.log('📊 جارٍ الانتقال إلى معايير التصنيف...');

    try {
      // Look for "معايير التصنيف" or "التصنيف" tab/link
      const criteriaButton = await this.page.evaluate(() => {
        const allElements = document.querySelectorAll('a, button, div, span, li');

        for (const el of allElements) {
          const text = el.innerText || el.textContent || '';
          if (text.includes('التصنيف') || text.includes('معايير')) {
            const isClickable = el.onclick || el.href || el.getAttribute('data-toggle') || el.classList.toString().includes('tab');
            if (isClickable) {
              return true;
            }
          }
        }

        return false;
      });

      if (criteriaButton) {
        // Find and click the criteria button
        const buttons = await this.page.$$('a, button, li');

        for (const btn of buttons) {
          const text = await this.page.evaluate(el => el.innerText || el.textContent, btn);

          if (text.includes('التصنيف') || text.includes('معايير')) {
            await btn.click();
            await this.page.waitForTimeout(1500);
            console.log('✅ تم الانتقال إلى معايير التصنيف');
            return;
          }
        }
      }

      console.log('⚠️ لم يتم العثور على زر معايير التصنيف');
    } catch (error) {
      console.error(`⚠️ خطأ في الانتقال: ${error.message}`);
    }
  }

  async getSiteName() {
    console.log('📍 جارٍ استخراج اسم الموقع...');

    const siteName = await this.page.evaluate(() => {
      // Try to find site name from the page
      const selectors = [
        'h1', // Usually the main heading
        '[class*="title"]',
        '[class*="name"]',
        'span[id*="LabelSiteName"]',
        'label[id*="LabelSiteName"]',
      ];

      for (const selector of selectors) {
        const el = document.querySelector(selector);
        if (el) {
          const text = (el.innerText || el.textContent || '').trim();
          if (text && text.length > 0 && text.length < 200) {
            return text;
          }
        }
      }

      // Fallback: Try to get from any element with Arabic text
      const allElements = document.querySelectorAll('*');
      for (const el of allElements) {
        const text = (el.innerText || el.textContent || '').trim();
        if (text && /[؀-ۿ]/.test(text) && text.length > 5 && text.length < 100) {
          return text;
        }
      }

      return `موقع_${Date.now()}`;
    });

    // Clean the site name for use as filename
    const cleanName = siteName
      .replace(/[\/\\:*?"<>|]/g, '_') // Remove invalid filename characters
      .replace(/\s+/g, '_') // Replace spaces with underscores
      .substring(0, 100); // Limit length

    console.log(`   اسم الموقع: ${siteName}`);
    return cleanName;
  }

  async extractRedBoxTable() {
    console.log('🔍 جارٍ البحث عن جدول معايير التصنيف...');

    // Get the classification criteria table (the one with the red box in the summary)
    const tableInfo = await this.page.evaluate(() => {
      // First, try to find the classification criteria table
      let tableElement = document.querySelector('#ctl12_TemplateRate_GridView1');

      // If not found, search for any table with specific class
      if (!tableElement) {
        tableElement = document.querySelector('table.GridView, table[class*="Grid"]');
      }

      // Fallback: get the first table on page
      if (!tableElement) {
        tableElement = document.querySelector('table');
      }

      if (tableElement) {
        const rect = tableElement.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;

        // Also get the element that comes after the table (summary/total row)
        let nextElement = tableElement.nextElementSibling;
        let totalHeight = rect.height;

        // Look for summary data below the table
        while (nextElement && totalHeight < rect.height + 500) {
          const nextRect = nextElement.getBoundingClientRect();
          const nextText = (nextElement.innerText || nextElement.textContent || '').trim();

          if (nextRect.height > 0 && nextText && nextText.length > 0) {
            totalHeight += nextRect.height + 10; // Add padding
          }

          if (nextElement.tagName === 'TABLE' || nextElement.classList.contains('button') ||
              nextElement.classList.contains('btn') || nextElement.tagName === 'BUTTON') {
            break;
          }

          nextElement = nextElement.nextElementSibling;
        }

        return {
          found: true,
          x: Math.max(0, rect.left + scrollLeft - 5),
          y: Math.max(0, rect.top + scrollTop - 5),
          width: rect.width + 10,
          height: totalHeight + 10,
          tagName: tableElement.tagName,
          classes: tableElement.className,
          id: tableElement.id,
        };
      }

      return { found: false };
    });

    if (!tableInfo.found) {
      throw new Error('لم يتم العثور على جدول معايير التصنيف');
    }

    console.log(`✅ تم العثور على الجدول: ${tableInfo.width}x${tableInfo.height}px`);
    return tableInfo;
  }

  async exportTableAsCleanPDF(filename, tableInfo) {
    console.log(`📄 جارٍ تحويل الجدول إلى PDF: ${filename}`);

    // Export with precise clipping of just the table
    if (tableInfo && tableInfo.width && tableInfo.height) {
      const filepath = `${this.config.outputDir}/${filename}`;

      // Create a clipped PDF that captures only the table and summary below it
      await this.page.pdf({
        path: filepath,
        clip: {
          x: Math.max(0, tableInfo.x),
          y: Math.max(0, tableInfo.y),
          width: tableInfo.width,
          height: tableInfo.height,
        },
        format: 'A4',
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
      });

      console.log(`✅ تم التصدير: ${filepath}`);
      return filepath;
    }

    // Fallback: Full page PDF
    const filepath = `${this.config.outputDir}/${filename}`;
    await this.page.pdf({
      path: filepath,
      format: 'A4',
      printBackground: true,
      margin: { top: '5mm', right: '5mm', bottom: '5mm', left: '5mm' },
    });

    console.log(`✅ تم التصدير: ${filepath}`);
    return filepath;
  }

  async captureTableScreenshot(filename, tableInfo) {
    console.log(`📸 جارٍ حفظ لقطة شاشة الجدول...`);

    const screenshotPath = `${this.config.outputDir}/${filename.replace('.pdf', '.png')}`;

    if (tableInfo && tableInfo.width && tableInfo.height) {
      // Capture just the table and summary
      await this.page.screenshot({
        path: screenshotPath,
        clip: {
          x: Math.max(0, tableInfo.x),
          y: Math.max(0, tableInfo.y),
          width: tableInfo.width,
          height: tableInfo.height,
        },
      });
    } else {
      // Full screenshot
      await this.page.screenshot({ path: screenshotPath, fullPage: true });
    }

    console.log(`✅ تم حفظ لقطة الشاشة: ${screenshotPath}`);
    return screenshotPath;
  }

  async exportBatch(decisions) {
    const results = [];

    for (let i = 0; i < decisions.length; i++) {
      const decision = decisions[i];

      try {
        console.log(`\n[${i + 1}/${decisions.length}] معالجة: ${decision.name}`);

        await this.page.waitForTimeout(800);

        // Select decision number if provided
        if (decision.decisionNumber) {
          await this.selectDecisionNumber(decision.decisionNumber);
          await this.clickSearchButton(); // اضغط على زر البحث
          await this.page.waitForTimeout(2000); // انتظر تحديث الجدول
        }

        // Navigate if needed
        if (decision.url) {
          await this.page.goto(decision.url, {
            waitUntil: this.config.waitForNavigation,
          });
        }

        // Open first form if not already navigated
        if (!decision.url) {
          await this.openFirstForm();
        }

        // Navigate to classification criteria
        await this.navigateToClassificationCriteria();

        // Extract table info
        const tableInfo = await this.extractRedBoxTable();

        // Get site name from the page
        const siteName = await this.getSiteName();

        // Export as PDF with site name
        const timestamp = Date.now();
        const pdfFile = `${siteName}_${timestamp}.pdf`;
        const pngFile = `${siteName}_${timestamp}.png`;

        const pdfPath = await this.exportTableAsCleanPDF(pdfFile, tableInfo);
        const pngPath = await this.captureTableScreenshot(pngFile, tableInfo);

        results.push({
          name: decision.name,
          siteName: siteName,
          decisionNumber: decision.decisionNumber,
          status: 'success',
          pdfPath,
          pngPath,
          tableInfo: tableInfo,
        });
      } catch (error) {
        console.error(`❌ خطأ: ${error.message}`);
        results.push({
          name: decision.name,
          decisionNumber: decision.decisionNumber,
          status: 'failed',
          error: error.message,
        });
      }
    }

    return results;
  }

  async scrollTableAndLoadData(maxRows = null) {
    console.log('📜 جارٍ تحميل كل البيانات...');

    const rowsLoaded = await this.page.evaluate((max) => {
      let lastHeight = 0;
      let scrolls = 0;

      const tableContainer = document.querySelector('table') ||
                           document.querySelector('[role="table"]') ||
                           document.querySelector('.table-container');

      if (!tableContainer) return 0;

      const interval = setInterval(() => {
        tableContainer.scrollTop = tableContainer.scrollHeight;
        scrolls++;

        if (max && scrolls > max) {
          clearInterval(interval);
        }
      }, 300);

      return new Promise((resolve) => {
        setTimeout(() => {
          clearInterval(interval);
          const rows = document.querySelectorAll('tr, [role="row"]');
          resolve(rows.length);
        }, scrolls * 300 + 500);
      });
    }, maxRows);

    console.log(`✅ تم تحميل ${rowsLoaded} صف`);
    return rowsLoaded;
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
      console.log('🔌 تم إغلاق المتصفح');
    }
  }

  async run(decisions) {
    try {
      await this.init();
      await this.login();
      await this.scrollTableAndLoadData();

      const results = await this.exportBatch(decisions);
      await this.close();

      return results;
    } catch (error) {
      await this.close();
      throw error;
    }
  }
}

// CLI usage
async function main() {
  const args = process.argv.slice(2);
  const username = args[0];
  const password = args[1];

  if (!username || !password) {
    console.log('استخدام: node table-pdf-exporter-advanced.js <username> <password>');
    process.exit(1);
  }

  const decisions = [
    { name: 'قرار_ترشيح_1' },
    { name: 'قرار_ترشيح_2' },
    { name: 'قرار_ترشيح_3' },
  ];

  const exporter = new AdvancedTablePDFExporter(username, password);

  try {
    const results = await exporter.run(decisions);

    console.log('\n╔════════════════════════════════════════╗');
    console.log('║             النتائج النهائية           ║');
    console.log('╚════════════════════════════════════════╝');

    results.forEach(r => {
      if (r.status === 'success') {
        console.log(`✅ ${r.siteName || r.name}: ${r.pdfPath}`);
      } else {
        console.log(`❌ ${r.name}: ${r.error}`);
      }
    });

    const successCount = results.filter(r => r.status === 'success').length;
    console.log(`\n📊 النتيجة: ${successCount}/${results.length}`);
  } catch (error) {
    console.error('❌ فشل التنفيذ:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = AdvancedTablePDFExporter;
