#!/usr/bin/env python3
"""
تشغيل كامل لقرار تسجيل واحد:
1. تسجيل دخول + الوصول لصفحة البحث
2. إدخال رقم القرار والبحث
3. جمع كل روابط الاستمارات من كل صفحات النتائج (بالتنقل بين الصفحات)
4. فتح كل استمارة مباشرة عبر رابطها وتصدير الجدول - بدون الرجوع للقائمة أو إعادة البحث
"""

import asyncio
import os
import sys
from table_pdf_exporter_advanced import AdvancedTablePDFExporter


async def main():
    if len(sys.argv) < 2:
        print('استخدام: python run_decision_export.py <رقم_القرار> [عدد الاستمارات المطلوب] [username] [password]')
        print('مثال: python run_decision_export.py 21          → يصدر كل الاستمارات')
        print('مثال: python run_decision_export.py 21 10        → يصدر أول 10 استمارات فقط')
        sys.exit(1)

    decision_number = sys.argv[1]
    forms_limit = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else None

    remaining_args = [a for i, a in enumerate(sys.argv[1:], 1) if not (i == 2 and forms_limit is not None)]
    username = sys.argv[3] if len(sys.argv) > 3 else input('👤 اسم المستخدم: ').strip()
    password = sys.argv[4] if len(sys.argv) > 4 else input('🔑 كلمة المرور: ').strip()

    exporter = AdvancedTablePDFExporter(
        username,
        password,
        options={
            'headless': False,  # غيّرها لـ True لما تتأكد كل شي شغال تمام
            'outputDir': f'./exports/قرار_{decision_number}',
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
        await asyncio.sleep(2)

        # جمع روابط الاستمارات - لو فيه حد محدد، يوقف التنقل بمجرد جمع عدد كافي
        form_links = await exporter.collect_all_form_links(min_links=forms_limit)

        if forms_limit:
            form_links = form_links[:forms_limit]
            print(f'ℹ️ تم تحديد المعالجة لأول {forms_limit} استمارة فقط')

        if not form_links:
            print('❌ لم يتم العثور على أي استمارات لهذا القرار')
            await exporter.close()
            return

        total = len(form_links)

        for i, link in enumerate(form_links, 1):
            try:
                # رابط تبويب "معايير التصنيف" مباشرة: DataEdit/{id}/10
                site_id = link.rstrip('/').split('/')[-1]
                classification_link = link.rstrip('/') + '/10'
                print(f'\n[{i}/{total}] فتح: {classification_link}')

                # فتح تبويب معايير التصنيف مباشرة - بدون فتح الاستمارة الأساسية أو الرجوع لقائمة البحث
                await exporter.page.goto(classification_link, wait_until=exporter.config['waitForNavigation'])

                # انتظار تحميل بطاقة "التصنيف" بالكامل (نستخدم رقم المجموع بأسفلها
                # كمؤشر تأكيد لأن الجدول العلوي وبطاقة التصنيف بيرندروا بشكل غير متزامن)
                try:
                    await exporter.page.wait_for_selector('#ctl12_TemplateRate_lblTotl', timeout=15000)
                except Exception:
                    pass  # extract_red_box_table() has its own fallback selectors
                await asyncio.sleep(0.5)  # small buffer for the card's rows to finish laying out

                table_info = await exporter.extract_red_box_table()
                site_name = await exporter.get_site_name(site_id=site_id)

                # مجلد خاص بكل موقع يحتوي على الـ PDF والصورة بنفس اسم الموقع
                site_folder = os.path.join(exporter.config['outputDir'], site_name)
                os.makedirs(site_folder, exist_ok=True)

                pdf_path = await exporter.export_table_as_clean_pdf(
                    os.path.join(site_name, f'{site_name}.pdf'), table_info)
                png_path = await exporter.capture_table_screenshot(
                    os.path.join(site_name, f'{site_name}.png'), table_info)

                results.append({'name': site_name, 'status': 'success', 'pdfPath': pdf_path})

            except Exception as e:
                print(f'❌ خطأ في الاستمارة {i}: {e}')
                results.append({'name': f'استمارة_{i}', 'status': 'failed', 'error': str(e)})

        await exporter.close()

        print('\n╔════════════════════════════════════════╗')
        print('║              النتيجة النهائية           ║')
        print('╚════════════════════════════════════════╝')

        success_count = sum(1 for r in results if r['status'] == 'success')
        failed = [r for r in results if r['status'] == 'failed']

        print(f'✅ نجح: {success_count}/{len(results)}')
        if failed:
            print(f'❌ فشل: {len(failed)}/{len(results)}')
            for r in failed:
                print(f'   • {r["name"]}: {r["error"]}')

    except Exception as e:
        await exporter.close()
        print(f'❌ خطأ عام: {e}')


if __name__ == '__main__':
    asyncio.run(main())
