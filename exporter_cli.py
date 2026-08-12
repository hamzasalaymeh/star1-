#!/usr/bin/env python3

import asyncio
import sys
from getpass import getpass
from table_pdf_exporter_advanced import AdvancedTablePDFExporter


async def main():
    """Interactive CLI for the exporter"""
    print('\n╔════════════════════════════════════════════╗')
    print('║  🚀 أداة تصدير جداول معايير التصنيف - PDF   ║')
    print('╚════════════════════════════════════════════╝\n')

    # Get credentials
    username = input('👤 اسم المستخدم: ').strip()
    password = getpass('🔑 كلمة المرور: ')

    if not username or not password:
        print('❌ يجب إدخال اسم المستخدم وكلمة المرور')
        sys.exit(1)

    # Get decision numbers
    print('\n📋 أدخل أرقام قرارات التسجيل (واحد في السطر):')
    print('   (اضغط Enter مرتين للانتهاء)\n')

    decision_numbers = []
    counter = 1

    while True:
        decision = input(f'رقم قرار التسجيل #{counter}: ').strip()

        if not decision:
            if decision_numbers:
                break
            print('   ⚠️ يجب إدخال رقم قرار واحد على الأقل')
            continue

        decision_numbers.append(decision)
        counter += 1

    # Confirm
    print(f'\n✅ سيتم تصدير {len(decision_numbers)} قرار')
    confirm = input('هل تريد المتابعة؟ (نعم/لا): ').strip().lower()

    if confirm not in ['نعم', 'yes', 'y']:
        print('❌ تم الإلغاء')
        sys.exit(1)

    # Prepare decisions
    decisions = [
        {'name': f'قرار_{i + 1}', 'decisionNumber': num}
        for i, num in enumerate(decision_numbers)
    ]

    # Run exporter
    print('\n' + '=' * 44)
    exporter = AdvancedTablePDFExporter(username, password)

    try:
        results = await exporter.run(decisions)

        print('\n╔════════════════════════════════════════╗')
        print('║             النتائج النهائية           ║')
        print('╚════════════════════════════════════════╝')

        for r in results:
            if r['status'] == 'success':
                print(f"✅ {r.get('siteName') or r['name']}: {r['pdfPath']}")
            else:
                print(f"❌ {r['name']}: {r['error']}")

        success_count = sum(1 for r in results if r['status'] == 'success')
        print(f'\n📊 النتيجة: {success_count}/{len(results)}')

        if success_count == len(results):
            print('\n🎉 تم التصدير بنجاح!')
        else:
            print(f'\n⚠️ تم تصدير {success_count} من {len(results)} قرار')

    except Exception as e:
        print(f'\n❌ فشل التنفيذ: {str(e)}')
        sys.exit(1)


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print('\n\n⚠️ تم الإيقاف بواسطة المستخدم')
        sys.exit(0)
