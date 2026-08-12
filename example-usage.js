/**
 * مثال على استخدام أداة تصدير جداول STAR1
 *
 * هذا المثال يوضح كيفية:
 * 1. تسجيل الدخول تلقائياً
 * 2. إدخال رقم قرار التسجيل
 * 3. فتح الاستمارات
 * 4. الذهاب إلى معايير التصنيف
 * 5. تصدير الجداول إلى PDF
 */

const AdvancedTablePDFExporter = require('./table-pdf-exporter-advanced');

async function example1_BasicUsage() {
  console.log('═══════════════════════════════════════════');
  console.log('  مثال 1: الاستخدام الأساسي');
  console.log('═══════════════════════════════════════════\n');

  const exporter = new AdvancedTablePDFExporter('username', 'password');

  const decisions = [
    { decisionNumber: '21', name: 'قرار_21' },
    { decisionNumber: '22', name: 'قرار_22' },
  ];

  try {
    const results = await exporter.run(decisions);

    console.log('\n✅ النتائج:');
    results.forEach(r => {
      if (r.status === 'success') {
        console.log(`  ✓ ${r.decisionNumber}: ${r.pdfPath}`);
      } else {
        console.log(`  ✗ ${r.decisionNumber}: ${r.error}`);
      }
    });
  } catch (error) {
    console.error('❌ خطأ:', error.message);
  }
}

async function example2_CustomSettings() {
  console.log('\n═══════════════════════════════════════════');
  console.log('  مثال 2: إعدادات مخصصة');
  console.log('═══════════════════════════════════════════\n');

  const exporter = new AdvancedTablePDFExporter('username', 'password', {
    baseUrl: 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10',
    outputDir: './my-exports',
    headless: false, // عرض المتصفح للمراقبة
    navigationTimeout: 45000,
    maxRetries: 5,
  });

  const decisions = [
    {
      decisionNumber: '21',
      name: 'قرار_ترشيح_رئيسي',
      url: 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/19112/10' // URL اختياري
    },
  ];

  try {
    const results = await exporter.run(decisions);
    console.log('✅ تم التصدير بنجاح!');
  } catch (error) {
    console.error('❌ خطأ:', error.message);
  }
}

async function example3_ManualSteps() {
  console.log('\n═══════════════════════════════════════════');
  console.log('  مثال 3: خطوات يدوية للتحكم الكامل');
  console.log('═══════════════════════════════════════════\n');

  const exporter = new AdvancedTablePDFExporter('username', 'password');

  try {
    // الخطوة 1: تهيئة المتصفح
    console.log('1. تهيئة المتصفح...');
    await exporter.init();

    // الخطوة 2: تسجيل الدخول
    console.log('2. تسجيل الدخول...');
    await exporter.login();

    // الخطوة 3: إدخال رقم القرار
    console.log('3. إدخال رقم القرار 21...');
    await exporter.selectDecisionNumber('21');

    // الخطوة 4: فتح الاستمارة الأولى
    console.log('4. فتح الاستمارة...');
    await exporter.openFirstForm();

    // الخطوة 5: الانتقال إلى معايير التصنيف
    console.log('5. الانتقال إلى معايير التصنيف...');
    await exporter.navigateToClassificationCriteria();

    // الخطوة 6: تحميل كل البيانات
    console.log('6. تحميل البيانات...');
    await exporter.scrollTableAndLoadData();

    // الخطوة 7: استخراج الجدول
    console.log('7. استخراج الجدول...');
    const redBoxInfo = await exporter.extractRedBoxTable();

    // الخطوة 8: تصدير PDF
    console.log('8. تصدير PDF...');
    const pdfPath = await exporter.exportTableAsCleanPDF('قرار_21.pdf', redBoxInfo);

    // الخطوة 9: تصدير لقطة شاشة
    console.log('9. تصدير لقطة شاشة...');
    const pngPath = await exporter.captureTableScreenshot('قرار_21.png', redBoxInfo);

    console.log(`\n✅ تم الانتهاء!`);
    console.log(`   PDF: ${pdfPath}`);
    console.log(`   PNG: ${pngPath}`);

    // الخطوة 10: إغلاق المتصفح
    console.log('10. إغلاق المتصفح...');
    await exporter.close();
  } catch (error) {
    console.error('❌ خطأ:', error.message);
    await exporter.close();
  }
}

async function example4_BatchProcessing() {
  console.log('\n═══════════════════════════════════════════');
  console.log('  مثال 4: معالجة دفعة كبيرة من القرارات');
  console.log('═══════════════════════════════════════════\n');

  const exporter = new AdvancedTablePDFExporter('username', 'password', {
    outputDir: './batch-exports',
    headless: true,
  });

  // قرارات متعددة
  const decisions = Array.from({ length: 10 }, (_, i) => ({
    decisionNumber: String(21 + i),
    name: `قرار_${21 + i}`
  }));

  try {
    console.log(`📋 جاري معالجة ${decisions.length} قرار...`);
    const results = await exporter.run(decisions);

    const successful = results.filter(r => r.status === 'success').length;
    console.log(`\n✅ النتيجة: ${successful}/${results.length}`);

    // عرض ملخص
    results.forEach((r, i) => {
      const icon = r.status === 'success' ? '✓' : '✗';
      console.log(`  ${icon} قرار ${r.decisionNumber}`);
    });
  } catch (error) {
    console.error('❌ خطأ:', error.message);
  }
}

// اختيار أي مثال تريد تشغيله
const exampleNumber = process.argv[2] || '1';

(async () => {
  switch (exampleNumber) {
    case '1':
      await example1_BasicUsage();
      break;
    case '2':
      await example2_CustomSettings();
      break;
    case '3':
      await example3_ManualSteps();
      break;
    case '4':
      await example4_BatchProcessing();
      break;
    default:
      console.log('استخدام: node example-usage.js [1|2|3|4]');
      console.log('  1 - الاستخدام الأساسي');
      console.log('  2 - إعدادات مخصصة');
      console.log('  3 - خطوات يدوية');
      console.log('  4 - معالجة دفعة');
  }
})();
