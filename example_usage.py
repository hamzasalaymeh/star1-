#!/usr/bin/env python3
"""
أمثلة لاستخدام أداة تصدير معايير التصنيف
"""

import asyncio
from table_pdf_exporter_advanced import AdvancedTablePDFExporter


async def example_1_basic_usage():
    """المثال 1️⃣: الاستخدام الأساسي"""
    print('\n╔════════════════════════════════════════╗')
    print('║       المثال 1️⃣: الاستخدام الأساسي       ║')
    print('╚════════════════════════════════════════╝\n')

    decisions = [
        {'name': 'قرار_21', 'decisionNumber': '21'},
    ]

    exporter = AdvancedTablePDFExporter('admin', 'password123')

    try:
        results = await exporter.run(decisions)
        for r in results:
            if r['status'] == 'success':
                print(f"✅ {r['siteName']}: {r['pdfPath']}")
            else:
                print(f"❌ {r['error']}")
    except Exception as e:
        print(f'❌ خطأ: {e}')


async def example_2_multiple_decisions():
    """المثال 2️⃣: معالجة عدة قرارات"""
    print('\n╔════════════════════════════════════════╗')
    print('║      المثال 2️⃣: عدة قرارات في دفعة        ║')
    print('╚════════════════════════════════════════╝\n')

    decisions = [
        {'name': 'قرار_1', 'decisionNumber': '21'},
        {'name': 'قرار_2', 'decisionNumber': '22'},
        {'name': 'قرار_3', 'decisionNumber': '23'},
    ]

    exporter = AdvancedTablePDFExporter('admin', 'password123')

    try:
        results = await exporter.run(decisions)

        print('\n╔════════════════════════════════════════╗')
        print('║             النتائج النهائية           ║')
        print('╚════════════════════════════════════════╝')

        for i, r in enumerate(results, 1):
            if r['status'] == 'success':
                print(f"{i}. ✅ {r['siteName']}")
                print(f"   📄 PDF: {r['pdfPath']}")
                print(f"   🖼️  PNG: {r['pngPath']}\n")
            else:
                print(f"{i}. ❌ {r['name']}: {r['error']}\n")

        success_count = sum(1 for r in results if r['status'] == 'success')
        print(f'📊 النتيجة: {success_count}/{len(results)} بنجاح')

    except Exception as e:
        print(f'❌ خطأ: {e}')


async def example_3_custom_settings():
    """المثال 3️⃣: إعدادات مخصصة"""
    print('\n╔════════════════════════════════════════╗')
    print('║      المثال 3️⃣: إعدادات مخصصة            ║')
    print('╚════════════════════════════════════════╝\n')

    decisions = [
        {'name': 'قرار_21', 'decisionNumber': '21'},
    ]

    exporter = AdvancedTablePDFExporter(
        username='admin',
        password='password123',
        options={
            'outputDir': './custom-exports',  # مجلد مخصص
            'headless': False,  # عرض المتصفح
            'navigationTimeout': 60000,  # 60 ثانية
        }
    )

    try:
        results = await exporter.run(decisions)
        for r in results:
            if r['status'] == 'success':
                print(f"✅ تم التصدير: {r['pdfPath']}")
            else:
                print(f"❌ خطأ: {r['error']}")
    except Exception as e:
        print(f'❌ خطأ: {e}')


async def example_4_manual_steps():
    """المثال 4️⃣: خطوات يدوية للتحكم الكامل"""
    print('\n╔════════════════════════════════════════╗')
    print('║      المثال 4️⃣: خطوات يدوية مفصلة         ║')
    print('╚════════════════════════════════════════╝\n')

    exporter = AdvancedTablePDFExporter('admin', 'password123')

    try:
        # 1. تهيئة
        print('1️⃣ تهيئة المتصفح...')
        await exporter.init()

        # 2. تسجيل دخول
        print('2️⃣ تسجيل الدخول...')
        await exporter.login()

        # 3. تحميل البيانات
        print('3️⃣ تحميل البيانات...')
        await exporter.scroll_table_and_load_data()

        # 4. إدخال رقم القرار
        print('4️⃣ إدخال رقم القرار...')
        await exporter.select_decision_number('21')

        # 5. البحث
        print('5️⃣ الضغط على البحث...')
        await exporter.click_search_button()

        # 6. فتح الاستمارة
        print('6️⃣ فتح الاستمارة...')
        await exporter.open_first_form()

        # 7. الانتقال لمعايير التصنيف
        print('7️⃣ الانتقال لمعايير التصنيف...')
        await exporter.navigate_to_classification_criteria()

        # 8. استخراج الجدول
        print('8️⃣ استخراج معلومات الجدول...')
        table_info = await exporter.extract_red_box_table()
        print(f'   حجم الجدول: {table_info["width"]}x{table_info["height"]}px')

        # 9. استخراج اسم الموقع
        print('9️⃣ استخراج اسم الموقع...')
        site_name = await exporter.get_site_name()
        print(f'   الموقع: {site_name}')

        # 10. تصدير PDF
        print('🔟 تصدير PDF...')
        pdf_path = await exporter.export_table_as_clean_pdf(f'{site_name}.pdf', table_info)
        print(f'    ✅ {pdf_path}')

        # 11. حفظ لقطة شاشة
        print('1️⃣1️⃣ حفظ لقطة شاشة...')
        png_path = await exporter.capture_table_screenshot(f'{site_name}.png', table_info)
        print(f'    ✅ {png_path}')

        # 12. إغلاق
        print('1️⃣2️⃣ إغلاق المتصفح...')
        await exporter.close()

        print('\n✅ انتهت جميع الخطوات بنجاح!')

    except Exception as e:
        await exporter.close()
        print(f'❌ خطأ: {e}')


async def example_5_batch_processing():
    """المثال 5️⃣: معالجة دفعية كبيرة"""
    print('\n╔════════════════════════════════════════╗')
    print('║      المثال 5️⃣: معالجة دفعية كبيرة        ║')
    print('╚════════════════════════════════════════╝\n')

    # إنشاء قائمة كبيرة من القرارات
    decisions = []
    for i in range(1, 11):  # 10 قرارات
        decisions.append({
            'name': f'قرار_{i}',
            'decisionNumber': str(20 + i),
        })

    exporter = AdvancedTablePDFExporter('admin', 'password123')

    try:
        results = await exporter.run(decisions)

        print('\n╔════════════════════════════════════════╗')
        print('║             الملخص النهائي             ║')
        print('╚════════════════════════════════════════╝')

        success_results = [r for r in results if r['status'] == 'success']
        failed_results = [r for r in results if r['status'] == 'failed']

        print(f'\n✅ نجح: {len(success_results)}/{len(results)}')
        for r in success_results:
            print(f'   • {r["siteName"]}')

        if failed_results:
            print(f'\n❌ فشل: {len(failed_results)}/{len(results)}')
            for r in failed_results:
                print(f'   • {r["name"]}: {r["error"]}')

        print(f'\n📊 إجمالي الملفات: {len(success_results) * 2}')
        print(f'   (PDF و PNG لكل استمارة)')

    except Exception as e:
        print(f'❌ خطأ: {e}')


async def main():
    """تشغيل الأمثلة"""
    import sys

    if len(sys.argv) < 2:
        print('\nأمثلة لاستخدام أداة التصدير:')
        print('  python example_usage.py 1  - استخدام أساسي')
        print('  python example_usage.py 2  - عدة قرارات')
        print('  python example_usage.py 3  - إعدادات مخصصة')
        print('  python example_usage.py 4  - خطوات يدوية')
        print('  python example_usage.py 5  - معالجة دفعية\n')
        sys.exit(1)

    example_num = sys.argv[1]

    examples = {
        '1': example_1_basic_usage,
        '2': example_2_multiple_decisions,
        '3': example_3_custom_settings,
        '4': example_4_manual_steps,
        '5': example_5_batch_processing,
    }

    if example_num in examples:
        await examples[example_num]()
    else:
        print(f'❌ المثال {example_num} غير موجود')
        print('   اختر من: 1, 2, 3, 4, 5')


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print('\n\n⚠️ تم الإيقاف بواسطة المستخدم')
