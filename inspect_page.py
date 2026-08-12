#!/usr/bin/env python3
"""
أداة فحص: تسجل دخول، تروح لصفحة البحث، وتطبع كل الحقول والأزرار الموجودة فعلياً
"""

import asyncio
import sys
from playwright.async_api import async_playwright


async def main():
    username = sys.argv[1] if len(sys.argv) > 1 else input('👤 اسم المستخدم: ').strip()
    password = sys.argv[2] if len(sys.argv) > 2 else input('🔑 كلمة المرور: ').strip()

    base_url = 'https://heritage2.itqan-consultant.com/Web/App/Home/Request/37528/1'
    search_url = 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataView'

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        await page.set_viewport_size({"width": 1920, "height": 1080})

        print('🔐 تسجيل الدخول...')
        await page.goto(base_url, wait_until='networkidle')

        username_input = await page.query_selector('input[name="ctl12$ctl03"]')
        password_input = await page.query_selector('input[name="ctl12$ctl07"]')

        if not username_input or not password_input:
            print('❌ لم يتم العثور على حقول تسجيل الدخول')
            await browser.close()
            return

        await username_input.fill(username)
        await password_input.fill(password)

        submit_button = await page.query_selector(
            '#ctl12_btnLogin, input[type="submit"], button[type="submit"]'
        )
        if submit_button:
            await submit_button.click()
            await page.wait_for_load_state('networkidle')

        if '/Login' in page.url:
            print('❌ فشل تسجيل الدخول - تأكد من صحة اسم المستخدم وكلمة المرور')
            await browser.close()
            return

        print('✅ تم تسجيل الدخول')
        print(f'🔎 الانتقال إلى: {search_url}')
        await page.goto(search_url, wait_until='networkidle')
        await asyncio.sleep(2)

        print(f'\n📍 الرابط الحالي: {page.url}')
        print(f'📍 عنوان الصفحة: {await page.title()}')

        # فحص كل الحقول
        info = await page.evaluate("""
            () => {
                const data = { inputs: [], buttons: [], selects: [], tables: [], iframes: [] };

                document.querySelectorAll('input').forEach((input) => {
                    if (input.offsetParent !== null) {
                        data.inputs.push({
                            type: input.type,
                            name: input.name,
                            id: input.id,
                            placeholder: input.placeholder,
                            value: input.value,
                            classes: input.className,
                        });
                    }
                });

                document.querySelectorAll('button, input[type="submit"], input[type="button"]').forEach((btn) => {
                    if (btn.offsetParent !== null) {
                        data.buttons.push({
                            tag: btn.tagName,
                            text: (btn.innerText || btn.textContent || btn.value || '').trim().substring(0, 50),
                            name: btn.name,
                            id: btn.id,
                            classes: btn.className,
                        });
                    }
                });

                document.querySelectorAll('select').forEach((sel) => {
                    if (sel.offsetParent !== null) {
                        data.selects.push({
                            name: sel.name,
                            id: sel.id,
                            classes: sel.className,
                        });
                    }
                });

                document.querySelectorAll('table').forEach((table) => {
                    data.tables.push({
                        id: table.id,
                        classes: table.className,
                        rows: table.querySelectorAll('tr').length,
                    });
                });

                document.querySelectorAll('iframe').forEach((f) => {
                    data.iframes.push({ src: f.src, id: f.id, name: f.name });
                });

                return data;
            }
        """)

        print('\n═══════════ IFRAMES ═══════════')
        if info['iframes']:
            for f in info['iframes']:
                print(f"   src={f['src']} id={f['id']} name={f['name']}")
            print('\n⚠️ يوجد iframe! الحقول ممكن تكون جوّاه، هذا مهم جداً.')
        else:
            print('   لا يوجد iframes')

        print('\n═══════════ INPUTS (الحقول) ═══════════')
        for i, inp in enumerate(info['inputs']):
            print(f"[{i}] type={inp['type']} name={inp['name']} id={inp['id']} "
                  f"placeholder={inp['placeholder']} classes={inp['classes']}")

        print('\n═══════════ BUTTONS (الأزرار) ═══════════')
        for i, btn in enumerate(info['buttons']):
            print(f"[{i}] tag={btn['tag']} text=\"{btn['text']}\" name={btn['name']} "
                  f"id={btn['id']} classes={btn['classes']}")

        print('\n═══════════ SELECTS (القوائم المنسدلة) ═══════════')
        for i, sel in enumerate(info['selects']):
            print(f"[{i}] name={sel['name']} id={sel['id']} classes={sel['classes']}")

        print('\n═══════════ TABLES (الجداول) ═══════════')
        for i, t in enumerate(info['tables']):
            print(f"[{i}] id={t['id']} classes={t['classes']} rows={t['rows']}")

        # لقطة شاشة للتأكد
        await page.screenshot(path='inspect_screenshot.png', full_page=True)
        print('\n📸 تم حفظ لقطة شاشة: inspect_screenshot.png')

        print('\n\n⏳ المتصفح رح يضل مفتوح 60 ثانية عشان تتأكد بعينك من الصفحة...')
        await asyncio.sleep(60)

        await browser.close()


if __name__ == '__main__':
    asyncio.run(main())
