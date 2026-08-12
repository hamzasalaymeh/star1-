#!/usr/bin/env python3
"""
تجربة سريعة: قرار رقم 21 فقط - أول استمارة (index 0)
"""

import asyncio
import sys
from table_pdf_exporter_advanced import AdvancedTablePDFExporter


async def main():
    username = sys.argv[1] if len(sys.argv) > 1 else input('👤 اسم المستخدم: ').strip()
    password = sys.argv[2] if len(sys.argv) > 2 else input('🔑 كلمة المرور: ').strip()

    # قرار واحد فقط - سيفتح أول استمارة تلقائياً (index 0)
    decisions = [
        {'name': 'تجربة_قرار_21', 'decisionNumber': '21'},
    ]

    exporter = AdvancedTablePDFExporter(
        username,
        password,
        options={
            'headless': False,   # لمشاهدة المتصفح أثناء التجربة
            'outputDir': './test-exports',
        }
    )

    try:
        results = await exporter.run(decisions)

        print('\n╔════════════════════════════════════════╗')
        print('║              نتيجة التجربة              ║')
        print('╚════════════════════════════════════════╝')

        for r in results:
            if r['status'] == 'success':
                print(f"✅ نجحت التجربة: {r['siteName']}")
                print(f"   📄 PDF: {r['pdfPath']}")
                print(f"   🖼️  PNG: {r['pngPath']}")
            else:
                print(f"❌ فشلت التجربة: {r['error']}")

    except Exception as e:
        print(f'❌ خطأ: {e}')


if __name__ == '__main__':
    asyncio.run(main())
