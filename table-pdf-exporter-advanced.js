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

  async extractRedBoxTable() {
    console.log('🔍 جارٍ البحث عن الجدول المحدد بمربع أحمر...');

    // Get red-boxed element coordinates
    const redBoxInfo = await this.page.evaluate(() => {
      // Find element with red border
      let redBoxElement = null;

      // Try multiple selectors
      const selectors = [
        'div[style*="border: 2px solid red"]',
        'div[style*="border-color: red"]',
        '.red-box',
        '[class*="red"]',
        'table',
      ];

      for (const selector of selectors) {
        const el = document.querySelector(selector);
        if (el) {
          const style = window.getComputedStyle(el);
          if (style.borderColor.includes('red') || style.borderColor === 'rgb(255, 0, 0)') {
            redBoxElement = el;
            break;
          }
        }
      }

      if (!redBoxElement) {
        redBoxElement = document.querySelector('table') || document.querySelector('[role="table"]');
      }

      if (redBoxElement) {
        const rect = redBoxElement.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;

        return {
          found: true,
          x: rect.left + scrollLeft,
          y: rect.top + scrollTop,
          width: rect.width,
          height: rect.height,
          tagName: redBoxElement.tagName,
          classes: redBoxElement.className,
        };
      }

      return { found: false };
    });

    if (!redBoxInfo.found) {
      throw new Error('لم يتم العثور على الجدول');
    }

    console.log(`✅ تم العثور على الجدول: ${redBoxInfo.width}x${redBoxInfo.height}px`);
    return redBoxInfo;
  }

  async exportTableAsCleanPDF(filename, redBoxInfo) {
    console.log(`📄 جارٍ تحويل الجدول إلى PDF: ${filename}`);

    // Option 1: Export with precise clipping
    if (redBoxInfo && redBoxInfo.width && redBoxInfo.height) {
      const clipPath = `${this.config.outputDir}/${filename.replace('.pdf', '_clip.pdf')}`;

      // Create a clipped PDF
      await this.page.pdf({
        path: clipPath,
        clip: {
          x: Math.max(0, redBoxInfo.x - 10),
          y: Math.max(0, redBoxInfo.y - 10),
          width: redBoxInfo.width + 20,
          height: redBoxInfo.height + 20,
        },
        format: 'A4',
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
      });

      console.log(`✅ تم التصدير: ${clipPath}`);
      return clipPath;
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

  async captureTableScreenshot(filename, redBoxInfo) {
    console.log(`📸 جارٍ حفظ لقطة شاشة الجدول...`);

    const screenshotPath = `${this.config.outputDir}/${filename.replace('.pdf', '.png')}`;

    if (redBoxInfo && redBoxInfo.width && redBoxInfo.height) {
      // Capture just the red box
      await this.page.screenshot({
        path: screenshotPath,
        clip: {
          x: Math.max(0, redBoxInfo.x - 5),
          y: Math.max(0, redBoxInfo.y - 5),
          width: redBoxInfo.width + 10,
          height: redBoxInfo.height + 10,
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

        // Navigate if needed
        if (decision.url) {
          await this.page.goto(decision.url, {
            waitUntil: this.config.waitForNavigation,
          });
        }

        // Extract table info
        const redBoxInfo = await this.extractRedBoxTable();

        // Export as PDF
        const timestamp = Date.now();
        const pdfFile = `${decision.name}_${timestamp}.pdf`;
        const pngFile = `${decision.name}_${timestamp}.png`;

        const pdfPath = await this.exportTableAsCleanPDF(pdfFile, redBoxInfo);
        const pngPath = await this.captureTableScreenshot(pngFile, redBoxInfo);

        results.push({
          name: decision.name,
          status: 'success',
          pdfPath,
          pngPath,
          tableInfo: redBoxInfo,
        });
      } catch (error) {
        console.error(`❌ خطأ: ${error.message}`);
        results.push({
          name: decision.name,
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
        console.log(`✅ ${r.name}: ${r.pdfPath}`);
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
