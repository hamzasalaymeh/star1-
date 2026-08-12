#!/usr/bin/env node

const puppeteer = require('puppeteer');
const fs = require('fs').promises;

class PageInspector {
  constructor(username, password, options = {}) {
    this.username = username;
    this.password = password;
    this.browser = null;
    this.page = null;
    this.baseUrl = 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10';
  }

  log(message) {
    console.log(`[INSPECTOR] ${message}`);
  }

  async init() {
    this.browser = await puppeteer.launch({
      headless: false,
      args: ['--no-sandbox'],
    });
    this.page = await this.browser.newPage();
    await this.page.setViewport({ width: 1920, height: 1080 });
  }

  async goToLoginPage() {
    this.log('جارٍ فتح صفحة تسجيل الدخول...');
    await this.page.goto(this.baseUrl, { waitUntil: 'networkidle2' });
    this.log('✅ تم فتح الصفحة');
  }

  async inspectLoginPage() {
    this.log('🔍 فحص عناصر تسجيل الدخول...\n');

    const elements = await this.page.evaluate(() => {
      const data = {
        inputs: [],
        buttons: [],
        links: [],
        allElements: {}
      };

      // Inputs
      const inputs = document.querySelectorAll('input');
      inputs.forEach((input, i) => {
        data.inputs.push({
          index: i,
          type: input.type,
          name: input.name,
          id: input.id,
          placeholder: input.placeholder,
          classes: input.className,
          selector: input.id ? `#${input.id}` : input.name ? `input[name="${input.name}"]` : `input[type="${input.type}"]:nth-of-type(${i})`,
          visible: input.offsetParent !== null
        });
      });

      // Buttons
      const buttons = document.querySelectorAll('button, input[type="submit"], a[class*="btn"], a[class*="login"]');
      buttons.forEach((btn, i) => {
        data.buttons.push({
          index: i,
          tag: btn.tagName,
          text: (btn.innerText || btn.textContent || '').trim().substring(0, 50),
          type: btn.type,
          id: btn.id,
          classes: btn.className,
          selector: btn.id ? `#${btn.id}` : `button:nth-of-type(${i})`,
          visible: btn.offsetParent !== null
        });
      });

      // Links
      const links = document.querySelectorAll('a');
      links.forEach((link, i) => {
        const text = (link.innerText || link.textContent || '').trim();
        if (text && text.length < 100) {
          data.links.push({
            index: i,
            text: text,
            href: link.href,
            id: link.id,
            classes: link.className,
            visible: link.offsetParent !== null
          });
        }
      });

      return data;
    });

    console.log('\n📋 INPUTS (حقول الإدخال):');
    console.log('═════════════════════════════════════════');
    elements.inputs.forEach(inp => {
      console.log(`\n[${inp.index}] ${inp.type.toUpperCase()}`);
      console.log(`   Name: ${inp.name || 'N/A'}`);
      console.log(`   ID: ${inp.id || 'N/A'}`);
      console.log(`   Placeholder: ${inp.placeholder || 'N/A'}`);
      console.log(`   Classes: ${inp.classes || 'N/A'}`);
      console.log(`   Selector: ${inp.selector}`);
      console.log(`   Visible: ${inp.visible ? '✓ نعم' : '✗ لا'}`);
    });

    console.log('\n\n🔘 BUTTONS (الأزرار):');
    console.log('═════════════════════════════════════════');
    elements.buttons.forEach(btn => {
      console.log(`\n[${btn.index}] ${btn.tag}`);
      console.log(`   Text: ${btn.text}`);
      console.log(`   Type: ${btn.type || 'N/A'}`);
      console.log(`   ID: ${btn.id || 'N/A'}`);
      console.log(`   Classes: ${btn.classes || 'N/A'}`);
      console.log(`   Selector: ${btn.selector}`);
      console.log(`   Visible: ${btn.visible ? '✓ نعم' : '✗ لا'}`);
    });

    console.log('\n\n🔗 LINKS (الروابط):');
    console.log('═════════════════════════════════════════');
    if (elements.links.length > 0) {
      elements.links.slice(0, 10).forEach(link => {
        console.log(`\n   Text: ${link.text}`);
        console.log(`   Href: ${link.href}`);
        console.log(`   Visible: ${link.visible ? '✓ نعم' : '✗ لا'}`);
      });
      if (elements.links.length > 10) {
        console.log(`\n   ... و ${elements.links.length - 10} روابط أخرى`);
      }
    } else {
      console.log('   لا توجد روابط');
    }

    return elements;
  }

  async inspectAfterLogin() {
    this.log('\n\n🔍 فحص عناصر بعد تسجيل الدخول...\n');

    const pageInfo = await this.page.evaluate(() => {
      const info = {
        url: window.location.href,
        title: document.title,
        tables: [],
        filterInputs: [],
        redElements: [],
        tabs: [],
        forms: []
      };

      // Tables
      document.querySelectorAll('table').forEach((table, i) => {
        const rows = table.querySelectorAll('tr');
        info.tables.push({
          index: i,
          rows: rows.length,
          columns: rows[0]?.querySelectorAll('td, th').length || 0,
          id: table.id,
          classes: table.className
        });
      });

      // Filter inputs (البحث/الفلترة)
      document.querySelectorAll('input').forEach((input, i) => {
        if (input.offsetParent !== null) {
          const style = window.getComputedStyle(input);
          info.filterInputs.push({
            index: i,
            type: input.type,
            name: input.name,
            value: input.value,
            placeholder: input.placeholder,
            isBordered: style.borderColor === 'rgb(255, 0, 0)' ||
                       style.borderColor.includes('red') ? '🔴 أحمر' : '⚪ عادي'
          });
        }
      });

      // Red bordered elements
      document.querySelectorAll('*').forEach(el => {
        const style = window.getComputedStyle(el);
        if (style.borderColor === 'rgb(255, 0, 0)') {
          const rect = el.getBoundingClientRect();
          info.redElements.push({
            tag: el.tagName,
            classes: el.className,
            id: el.id,
            width: Math.round(rect.width),
            height: Math.round(rect.height),
            isTable: el.tagName === 'TABLE' || el.querySelector('table') ? '✓' : '✗'
          });
        }
      });

      // Tabs/Navigation
      document.querySelectorAll('[role="tab"], .nav-link, .tab, [class*="tab"]').forEach(tab => {
        const text = (tab.innerText || tab.textContent || '').trim();
        if (text && text.length < 50) {
          info.tabs.push({
            text: text,
            id: tab.id,
            classes: tab.className,
            selected: tab.getAttribute('aria-selected') || tab.classList.contains('active')
          });
        }
      });

      // Forms
      document.querySelectorAll('form').forEach((form, i) => {
        info.forms.push({
          index: i,
          id: form.id,
          classes: form.className,
          method: form.method,
          action: form.action
        });
      });

      return info;
    });

    console.log(`URL: ${pageInfo.url}`);
    console.log(`Title: ${pageInfo.title}`);

    console.log('\n\n📊 TABLES (الجداول):');
    console.log('═════════════════════════════════════════');
    if (pageInfo.tables.length > 0) {
      pageInfo.tables.forEach(table => {
        console.log(`\n[جدول ${table.index}]`);
        console.log(`   الصفوف: ${table.rows}`);
        console.log(`   الأعمدة: ${table.columns}`);
        console.log(`   ID: ${table.id || 'N/A'}`);
        console.log(`   Classes: ${table.classes || 'N/A'}`);
      });
    } else {
      console.log('   لا توجد جداول');
    }

    console.log('\n\n🔴 RED BORDERED ELEMENTS (عناصر بحدود حمراء):');
    console.log('═════════════════════════════════════════');
    pageInfo.redElements.forEach((el, i) => {
      console.log(`\n[${i}] ${el.tag}`);
      console.log(`   الحجم: ${el.width}x${el.height}px`);
      console.log(`   ID: ${el.id || 'N/A'}`);
      console.log(`   Classes: ${el.classes || 'N/A'}`);
      console.log(`   يحتوي على جدول: ${el.isTable}`);
    });

    console.log('\n\n📋 FILTER INPUTS (حقول الفلترة/البحث):');
    console.log('═════════════════════════════════════════');
    pageInfo.filterInputs.forEach(inp => {
      console.log(`\n[${inp.index}]`);
      console.log(`   Type: ${inp.type}`);
      console.log(`   Name: ${inp.name || 'N/A'}`);
      console.log(`   Placeholder: ${inp.placeholder || 'N/A'}`);
      console.log(`   Border: ${inp.isBordered}`);
      console.log(`   Value: ${inp.value || 'فارغ'}`);
    });

    console.log('\n\n📑 TABS/NAVIGATION (التنقل):');
    console.log('═════════════════════════════════════════');
    if (pageInfo.tabs.length > 0) {
      pageInfo.tabs.forEach(tab => {
        console.log(`\n   ${tab.text} ${tab.selected ? '✓ مختار' : '⚪'}`);
        console.log(`   ID: ${tab.id || 'N/A'}`);
        console.log(`   Classes: ${tab.classes}`);
      });
    } else {
      console.log('   لا توجد tabs');
    }

    console.log('\n\n📝 FORMS (النماذج):');
    console.log('═════════════════════════════════════════');
    pageInfo.forms.forEach(form => {
      console.log(`\n[نموذج ${form.index}]`);
      console.log(`   ID: ${form.id || 'N/A'}`);
      console.log(`   Method: ${form.method || 'GET'}`);
      console.log(`   Action: ${form.action || 'N/A'}`);
    });

    return pageInfo;
  }

  async close() {
    if (this.browser) {
      await this.browser.close();
    }
  }

  async run() {
    try {
      await this.init();
      await this.goToLoginPage();
      await this.inspectLoginPage();

      console.log('\n\n⏳ اضغط Enter بعد تسجيل الدخول يدوياً...');

      await new Promise(resolve => {
        process.stdin.once('data', resolve);
      });

      await this.inspectAfterLogin();

      console.log('\n\n✅ انتهى الفحص!');
      console.log('\n📝 الآن، قل لي:');
      console.log('   1️⃣ أين حقل إدخال رقم القرار؟');
      console.log('   2️⃣ أين زر الإدخال/البحث؟');
      console.log('   3️⃣ ما اسم التبويب "معايير التصنيف"؟');
      console.log('   4️⃣ هل الجدول الأحمر موجود في صفحة منفصلة أم نفس الصفحة؟');
    } catch (error) {
      console.error('❌ خطأ:', error.message);
    } finally {
      await this.close();
    }
  }
}

async function main() {
  const args = process.argv.slice(2);
  const username = args[0] || 'admin';
  const password = args[1] || 'password';

  console.log('╔═══════════════════════════════════════════╗');
  console.log('║     🔍 أداة فحص الواجهة الاحترافية 🔍     ║');
  console.log('╚═══════════════════════════════════════════╝\n');

  const inspector = new PageInspector(username, password);
  await inspector.run();
}

main();
