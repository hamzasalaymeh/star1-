#!/usr/bin/env node

const puppeteer = require('puppeteer');
const fs = require('fs').promises;

class DebugExporter {
  constructor(username, password, options = {}) {
    this.username = username;
    this.password = password;
    this.browser = null;
    this.page = null;
    this.config = {
      baseUrl: options.baseUrl || 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10',
      headless: false, // Show browser
      outputDir: './debug-exports',
      navigationTimeout: 30000,
      ...options,
    };
  }

  log(icon, message) {
    console.log(`${icon} [${new Date().toLocaleTimeString()}] ${message}`);
  }

  async init() {
    this.log('⚙️', 'جارٍ تهيئة المتصفح...');
    this.browser = await puppeteer.launch({
      headless: this.config.headless,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
      ],
      dumpio: true,
    });

    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(this.config.navigationTimeout);

    // Log all console messages
    this.page.on('console', msg => {
      this.log('📝', `[Browser] ${msg.text()}`);
    });

    // Log errors
    this.page.on('error', err => {
      this.log('❌', `[Page Error] ${err.message}`);
    });

    // Log requests
    this.page.on('request', req => {
      this.log('🔗', `[Request] ${req.method()} ${req.url().substring(0, 80)}`);
    });

    // Log responses
    this.page.on('response', res => {
      this.log('📡', `[Response] ${res.status()} ${res.url().substring(0, 80)}`);
    });

    await fs.mkdir(this.config.outputDir, { recursive: true });
    this.log('✅', 'تم تهيئة المتصفح');
  }

  async login() {
    this.log('🔐', 'جارٍ تسجيل الدخول...');

    try {
      await this.page.goto(this.config.baseUrl, { waitUntil: 'networkidle2' });
      this.log('✅', 'تم تحميل الصفحة');

      // Take screenshot before login
      await this.page.screenshot({
        path: `${this.config.outputDir}/01-before-login.png`,
      });
      this.log('📸', 'تم حفظ لقطة قبل تسجيل الدخول');

      // Find all inputs
      const inputs = await this.page.$$('input');
      this.log('📊', `تم العثور على ${inputs.length} حقل إدخال`);

      // Log input details
      for (let i = 0; i < inputs.length; i++) {
        const type = await this.page.evaluate(
          el => el.type,
          inputs[i]
        );
        const name = await this.page.evaluate(
          el => el.name,
          inputs[i]
        );
        const placeholder = await this.page.evaluate(
          el => el.placeholder,
          inputs[i]
        );
        this.log('📝', `Input ${i}: type=${type}, name=${name}, placeholder=${placeholder}`);
      }

      // Fill first text input with username
      const textInputs = await this.page.$$('input[type="text"]');
      if (textInputs.length > 0) {
        await this.page.type('input[type="text"]', this.username);
        this.log('✅', 'تم إدخال اسم المستخدم');
      }

      // Fill password input
      const passwordInputs = await this.page.$$('input[type="password"]');
      if (passwordInputs.length > 0) {
        await this.page.type('input[type="password"]', this.password);
        this.log('✅', 'تم إدخال كلمة المرور');
      }

      // Take screenshot after filling
      await this.page.screenshot({
        path: `${this.config.outputDir}/02-filled-inputs.png`,
      });

      // Find and click submit button
      const buttons = await this.page.$$('button');
      this.log('📊', `تم العثور على ${buttons.length} زر`);

      for (let i = 0; i < buttons.length; i++) {
        const text = await this.page.evaluate(
          el => el.innerText,
          buttons[i]
        );
        const type = await this.page.evaluate(
          el => el.type,
          buttons[i]
        );
        this.log('🔘', `Button ${i}: type=${type}, text="${text}"`);
      }

      // Click submit button
      const submitButton = await this.page.$('button[type="submit"]');
      if (submitButton) {
        this.log('🔘', 'جارٍ النقر على زر تسجيل الدخول...');
        await submitButton.click();
        await this.page.waitForNavigation({ waitUntil: 'networkidle2' });
        this.log('✅', 'تم تسجيل الدخول بنجاح');
      } else {
        this.log('⚠️', 'لم يتم العثور على زر التقديم');
      }

      // Take screenshot after login
      await this.page.screenshot({
        path: `${this.config.outputDir}/03-after-login.png`,
      });
    } catch (error) {
      this.log('❌', `خطأ في تسجيل الدخول: ${error.message}`);
      throw error;
    }
  }

  async inspectPageStructure() {
    this.log('🔍', 'جارٍ فحص بنية الصفحة...');

    const structure = await this.page.evaluate(() => {
      return {
        title: document.title,
        url: window.location.href,
        tables: document.querySelectorAll('table').length,
        divs: document.querySelectorAll('div').length,
        buttons: document.querySelectorAll('button').length,
        inputs: document.querySelectorAll('input').length,
        redElements: Array.from(document.querySelectorAll('*')).filter(el => {
          const style = window.getComputedStyle(el);
          return style.borderColor.includes('red') || style.borderColor === 'rgb(255, 0, 0)';
        }).length,
      };
    });

    this.log('📊', `Page Title: ${structure.title}`);
    this.log('📊', `URL: ${structure.url}`);
    this.log('📊', `Tables: ${structure.tables}`);
    this.log('📊', `Red Elements: ${structure.redElements}`);

    return structure;
  }

  async findRedBoxTable() {
    this.log('🔍', 'جارٍ البحث عن الجدول الأحمر...');

    const redBoxInfo = await this.page.evaluate(() => {
      const results = [];

      // Find by red border
      const allElements = document.querySelectorAll('*');
      for (const el of allElements) {
        const style = window.getComputedStyle(el);
        if (style.borderColor === 'rgb(255, 0, 0)' || style.borderColor.includes('red')) {
          const rect = el.getBoundingClientRect();
          results.push({
            tag: el.tagName,
            class: el.className,
            id: el.id,
            width: rect.width,
            height: rect.height,
            found: 'red-border',
          });
        }
      }

      // Find tables
      const tables = document.querySelectorAll('table');
      tables.forEach((table, i) => {
        const rect = table.getBoundingClientRect();
        results.push({
          tag: 'table',
          index: i,
          class: table.className,
          width: rect.width,
          height: rect.height,
          rows: table.querySelectorAll('tr').length,
          found: 'table',
        });
      });

      return results;
    });

    this.log('📋', `تم العثور على ${redBoxInfo.length} عنصر مناسب`);
    redBoxInfo.forEach((info, i) => {
      this.log(
        '📊',
        `[${i}] ${info.tag} - ${info.width}x${info.height} - ${info.found}`
      );
    });

    return redBoxInfo;
  }

  async exportDebugPDF() {
    this.log('📄', 'جارٍ تصدير PDF...');

    const pdfPath = `${this.config.outputDir}/debug-page.pdf`;
    await this.page.pdf({
      path: pdfPath,
      format: 'A4',
      printBackground: true,
    });

    this.log('✅', `تم تصدير PDF: ${pdfPath}`);
  }

  async exportDebugScreenshot() {
    this.log('📸', 'جارٍ تصدير لقطة شاشة...');

    const screenshotPath = `${this.config.outputDir}/debug-fullpage.png`;
    await this.page.screenshot({
      path: screenshotPath,
      fullPage: true,
    });

    this.log('✅', `تم تصدير الصورة: ${screenshotPath}`);
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
      this.log('🔌', 'تم إغلاق المتصفح');
    }
  }

  async run() {
    try {
      await this.init();
      await this.login();

      // Wait a bit for page to fully load
      await this.page.waitForTimeout(2000);

      await this.inspectPageStructure();
      await this.findRedBoxTable();
      await this.exportDebugPDF();
      await this.exportDebugScreenshot();

      this.log('✅', 'تم الانتهاء من جلسة التصحيح');
      this.log('📁', `الملفات محفوظة في: ${this.config.outputDir}`);
    } catch (error) {
      this.log('❌', `خطأ: ${error.message}`);
      throw error;
    } finally {
      await this.close();
    }
  }
}

// CLI
async function main() {
  const args = process.argv.slice(2);
  const username = args[0];
  const password = args[1];

  if (!username || !password) {
    console.log('استخدام: node debug-exporter.js <username> <password>');
    process.exit(1);
  }

  const debugger = new DebugExporter(username, password);

  try {
    await debugger.run();
  } catch (error) {
    console.error('❌ فشل التصحيح:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = DebugExporter;
