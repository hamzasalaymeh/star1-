const puppeteer = require('puppeteer');
const fs = require('fs').promises;
const path = require('path');

// Configuration
const CONFIG = {
  baseUrl: 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10',
  browserHeadless: true,
  outputDir: './exports',
  maxRetries: 3,
  navigationTimeout: 30000,
};

class TablePDFExporter {
  constructor(username, password) {
    this.username = username;
    this.password = password;
    this.browser = null;
    this.page = null;
  }

  async init() {
    this.browser = await puppeteer.launch({
      headless: CONFIG.browserHeadless,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
      ],
    });
    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(CONFIG.navigationTimeout);

    // Create output directory
    await fs.mkdir(CONFIG.outputDir, { recursive: true });
  }

  async login() {
    console.log('🔐 جارٍ تسجيل الدخول...');

    try {
      await this.page.goto(CONFIG.baseUrl, { waitUntil: 'networkidle2' });

      // Find and fill login form
      const usernameSelector = 'input[type="text"]';
      const passwordSelector = 'input[type="password"]';

      await this.page.type(usernameSelector, this.username);
      await this.page.type(passwordSelector, this.password);

      // Submit login form
      const loginButton = await this.page.$('button[type="submit"]') ||
                         await this.page.$('button:has-text("دخول")') ||
                         await this.page.$('a.login-btn');

      if (loginButton) {
        await loginButton.click();
        await this.page.waitForNavigation({ waitUntil: 'networkidle2' });
        console.log('✅ تم تسجيل الدخول بنجاح');
      }
    } catch (error) {
      console.error('❌ فشل تسجيل الدخول:', error.message);
      throw error;
    }
  }

  async waitForTableLoad() {
    console.log('⏳ جارٍ انتظار تحميل الجدول...');

    // Wait for main table container
    await this.page.waitForSelector('[role="table"], table, .table, .grid', {
      timeout: 15000,
    });

    // Wait for table rows to be populated
    await this.page.waitForFunction(
      () => {
        const rows = document.querySelectorAll('table tr, [role="row"]');
        return rows.length > 1;
      },
      { timeout: 15000 }
    );

    console.log('✅ تم تحميل الجدول');
  }

  async exportToPDF(outputFilename) {
    console.log(`📄 جارٍ تصدير الجدول إلى PDF: ${outputFilename}`);

    try {
      // Find the red-boxed table container
      const tableSelector = '.red-box, [style*="border: 2px solid red"], .export-table, table';

      // Wait for the specific element
      await this.page.waitForSelector(tableSelector, { timeout: 10000 }).catch(() => {
        console.warn('⚠️ لم يتم العثور على الجدول المحدد، جاري البحث عن أي جدول...');
      });

      // Get the red box element bounds for PDF capture
      const elementHandle = await this.page.$(tableSelector);

      if (elementHandle) {
        const boundingBox = await elementHandle.boundingBox();

        if (boundingBox) {
          console.log(`📐 أبعاد الجدول: ${boundingBox.width}x${boundingBox.height}`);

          // Take screenshot of just the table area
          await elementHandle.screenshot({
            path: `${CONFIG.outputDir}/${outputFilename}.png`,
          });

          // Alternative: Use Puppeteer's PDF with proper margins
          await this.page.pdf({
            path: `${CONFIG.outputDir}/${outputFilename}`,
            scale: 1.2,
            printBackground: true,
            margin: {
              top: '10mm',
              right: '10mm',
              bottom: '10mm',
              left: '10mm',
            },
            // Capture specific clip region if needed
            ...(boundingBox && {
              width: `${boundingBox.width + 20}px`,
              height: `${boundingBox.height + 20}px`,
            }),
          });

          console.log(`✅ تم التصدير بنجاح: ${path.join(CONFIG.outputDir, outputFilename)}`);
          return path.join(CONFIG.outputDir, outputFilename);
        }
      } else {
        // Fallback: export full page
        await this.page.pdf({
          path: `${CONFIG.outputDir}/${outputFilename}`,
          scale: 1,
          printBackground: true,
        });
        console.log(`✅ تم التصدير بنجاح (كامل الصفحة): ${path.join(CONFIG.outputDir, outputFilename)}`);
      }
    } catch (error) {
      console.error('❌ خطأ في التصدير:', error.message);
      throw error;
    }
  }

  async scrollAndLoadAll() {
    console.log('🔄 جارٍ تحميل جميع البيانات...');

    let lastHeight = 0;
    let scrollCount = 0;
    const maxScrolls = 50; // Prevent infinite loops

    while (scrollCount < maxScrolls) {
      const newHeight = await this.page.evaluate(
        () => document.documentElement.scrollHeight
      );

      if (newHeight === lastHeight) break;

      lastHeight = newHeight;
      await this.page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
      await this.page.waitForTimeout(500);
      scrollCount++;
    }

    console.log(`✅ تم تحميل البيانات (${scrollCount} حركة تمرير)`);
  }

  async exportMultipleTables(decisions) {
    /**
     * decisions: Array of { name: string, selector?: string }
     * Example: [
     *   { name: 'decision_1', selector: '.table-red-box' },
     *   { name: 'decision_2', selector: '#main-table' }
     * ]
     */
    const results = [];

    for (const decision of decisions) {
      try {
        console.log(`\n📋 معالجة القرار: ${decision.name}`);
        await this.page.waitForTimeout(1000);

        if (decision.selector) {
          await this.page.waitForSelector(decision.selector, { timeout: 5000 });
        }

        const filename = `${decision.name}_${Date.now()}.pdf`;
        const filepath = await this.exportToPDF(filename);
        results.push({
          name: decision.name,
          status: 'success',
          filepath,
        });
      } catch (error) {
        console.error(`❌ فشل تصدير ${decision.name}:`, error.message);
        results.push({
          name: decision.name,
          status: 'failed',
          error: error.message,
        });
      }
    }

    return results;
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
      console.log('🔌 تم إغلاق المتصفح');
    }
  }

  async run(decisions) {
    let retries = 0;

    while (retries < CONFIG.maxRetries) {
      try {
        await this.init();
        await this.login();
        await this.waitForTableLoad();
        await this.scrollAndLoadAll();

        const results = await this.exportMultipleTables(decisions);
        return results;
      } catch (error) {
        retries++;
        console.error(`❌ محاولة ${retries} فشلت:`, error.message);

        if (retries < CONFIG.maxRetries) {
          console.log(`⏳ إعادة محاولة بعد 5 ثواني...`);
          await this.page?.close().catch(() => {});
          await this.browser?.close().catch(() => {});
          await new Promise(resolve => setTimeout(resolve, 5000));
        }
      }
    }

    throw new Error(`فشل التصدير بعد ${CONFIG.maxRetries} محاولات`);
  }
}

// Main execution
async function main() {
  const username = process.argv[2] || 'username';
  const password = process.argv[3] || 'password';
  const decisions = [
    { name: 'decision_1', selector: '.red-box' },
    { name: 'decision_2', selector: '.red-box' },
    { name: 'decision_3', selector: '.red-box' },
  ];

  console.log('╔════════════════════════════════════════╗');
  console.log('║   جهاز تصدير جداول منصة STAR1         ║');
  console.log('╚════════════════════════════════════════╝\n');

  const exporter = new TablePDFExporter(username, password);

  try {
    const results = await exporter.run(decisions);

    console.log('\n📊 ملخص النتائج:');
    console.table(results);

    const successful = results.filter(r => r.status === 'success').length;
    console.log(`\n✅ تم تصدير ${successful}/${results.length} ملفات بنجاح`);
  } catch (error) {
    console.error('\n❌ خطأ عام:', error.message);
    process.exit(1);
  } finally {
    await exporter.close();
  }
}

if (require.main === module) {
  main();
}

module.exports = TablePDFExporter;
