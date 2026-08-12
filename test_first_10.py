#!/usr/bin/env python3
"""
تجربة سريعة: قرار رقم 21 - أول 10 استمارات فقط
"""

import asyncio
import sys
from table_pdf_exporter_advanced import AdvancedTablePDFExporter


async def main():
    username = sys.argv[1] if len(sys.argv) > 1 else input('👤 اسم المستخدم: ').strip()
    password = sys.argv[2] if len(sys.argv) > 2 else input('🔑 كلمة المرور: ').strip()

    decision_number = '21'
    forms_limit = 10

    exporter = AdvancedTablePDFExporter(
        username,
        password,
        options={
            'headless': False,  # لمشاهدة المتصفح أثناء التجربة
            'outputDir': './test-exports',
        }
    )

    results = []

    try:
        await exporter.init()
        await exporter.login()
        await exporter.goto_search_page()

        # إدخال رقم القرار والبحث
        await exporter.select_decision_number(decision_number)
        await exporter.click_search_button()

        # جمع أول 10 روابط استمارات من نتائج البحث
        form_links = await exporter.get_form_links(forms_limit)

        if not form_links:
            print('❌ لم يتم العثور على أي استمارات لهذا القرار')
            await exporter.close()
            return

        for i, link in enumerate(form_links, 1):
            try:
                # رابط تبويب "معايير التصنيف" مباشرة: DataEdit/{id}/10
                site_id = link.rstrip('/').split('/')[-1]
                classification_link = link.rstrip('/') + '/10'
                print(f'\n[{i}/{len(form_links)}] فتح: {classification_link}')

                await exporter.page.goto(classification_link, wait_until=exporter.config['waitForNavigation'])

                # انتظار تحميل الجدول فعلياً (وليس شاشة "جاري التحميل")
                try:
                    await exporter.page.wait_for_selector('#ctl12_TemplateRate_GridView1', timeout=15000)
                except Exception:
                    pass  # extract_red_box_table() has its own fallback selectors

                # استخراج الجدول
                table_info = await exporter.extract_red_box_table()

                # اسم الموقع
                site_name = await exporter.get_site_name(site_id=site_id)

                # التصدير
                pdf_path = await exporter.export_table_as_clean_pdf(f'{site_name}.pdf', table_info)
                png_path = await exporter.capture_table_screenshot(f'{site_name}.png', table_info)

                results.append({'name': site_name, 'status': 'success', 'pdfPath': pdf_path})

            except Exception as e:
                print(f'❌ خطأ في الاستمارة {i}: {e}')
                results.append({'name': f'استمارة_{i}', 'status': 'failed', 'error': str(e)})

        await exporter.close()

        print('\n╔════════════════════════════════════════╗')
        print('║              نتيجة التجربة              ║')
        print('╚════════════════════════════════════════╝')

        for r in results:
            if r['status'] == 'success':
                print(f"✅ {r['name']}: {r['pdfPath']}")
            else:
                print(f"❌ {r['name']}: {r['error']}")

        success_count = sum(1 for r in results if r['status'] == 'success')
        print(f'\n📊 النتيجة: {success_count}/{len(results)}')

    except Exception as e:
        await exporter.close()
        print(f'❌ خطأ عام: {e}')


if __name__ == '__main__':
    asyncio.run(main())
