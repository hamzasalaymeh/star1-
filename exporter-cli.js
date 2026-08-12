#!/usr/bin/env node

const readline = require('readline');
const AdvancedTablePDFExporter = require('./table-pdf-exporter-advanced');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

function question(prompt) {
  return new Promise(resolve => {
    rl.question(prompt, resolve);
  });
}

async function main() {
  console.clear();
  console.log('╔════════════════════════════════════════════════════════╗');
  console.log('║        🌟 جهاز تصدير جداول منصة STAR1 الأوتوماتيكي 🌟 ║');
  console.log('╚════════════════════════════════════════════════════════╝\n');

  // Get credentials
  const username = await question('👤 اسم المستخدم: ');
  const password = await question('🔑 كلمة المرور: ');

  // Get decision names
  console.log('\n📋 أدخل أسماء القرارات المراد تصديرها (واحد في السطر)');
  console.log('   (اضغط Enter مرتين للانتهاء)\n');

  const decisions = [];
  let decisionName = await question('اسم القرار #1: ');

  while (decisionName.trim()) {
    decisions.push({ name: decisionName.trim() });
    decisionName = await question(`اسم القرار #${decisions.length + 1}: `);
  }

  if (decisions.length === 0) {
    console.log('❌ لم تدخل أي قرارات!');
    rl.close();
    return;
  }

  console.log(`\n✅ سيتم تصدير ${decisions.length} قرار\n`);

  // Ask for confirmation
  const confirm = await question('هل تريد المتابعة؟ (نعم/لا): ');

  if (!confirm.toLowerCase().startsWith('ن')) {
    console.log('❌ تم الإلغاء');
    rl.close();
    return;
  }

  rl.close();

  // Start export
  console.log('\n🚀 جارٍ البدء بالتصدير...\n');

  const exporter = new AdvancedTablePDFExporter(username, password, {
    headless: true,
    outputDir: './exports',
  });

  try {
    const results = await exporter.run(decisions);

    console.log('\n╔════════════════════════════════════════════════════════╗');
    console.log('║                  النتائج النهائية                    ║');
    console.log('╚════════════════════════════════════════════════════════╝\n');

    results.forEach((r, i) => {
      if (r.status === 'success') {
        console.log(`✅ [${i + 1}] ${r.name}`);
        console.log(`   📄 PDF: ${r.pdfPath}`);
        console.log(`   📸 لقطة: ${r.pngPath}\n`);
      } else {
        console.log(`❌ [${i + 1}] ${r.name}`);
        console.log(`   الخطأ: ${r.error}\n`);
      }
    });

    const successCount = results.filter(r => r.status === 'success').length;
    console.log(`📊 النتيجة النهائية: ${successCount}/${results.length}`);

    if (successCount === results.length) {
      console.log('🎉 تم التصدير بنجاح!');
    }
  } catch (error) {
    console.error('\n❌ خطأ:', error.message);
    process.exit(1);
  }
}

main();
