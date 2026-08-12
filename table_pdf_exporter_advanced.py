#!/usr/bin/env python3

import asyncio
import os
import re
from pathlib import Path
from datetime import datetime
from playwright.async_api import async_playwright, Browser, Page


class AdvancedTablePDFExporter:
    def __init__(self, username: str, password: str, options: dict = None):
        self.username = username
        self.password = password
        self.browser = None
        self.page = None

        options = options or {}
        self.config = {
            'baseUrl': options.get('baseUrl', 'https://heritage2.itqan-consultant.com/Web/App/Home/Login'),
            'searchUrl': options.get('searchUrl', 'https://heritage2.itqan-consultant.com/Web/App/Pools/DataView'),
            'outputDir': options.get('outputDir', './exports'),
            'headless': options.get('headless', True),
            'waitForNavigation': options.get('waitForNavigation', 'networkidle'),
            'navigationTimeout': options.get('navigationTimeout', 30000),
            'maxRetries': options.get('maxRetries', 3),
        }
        self.config.update(options)

    async def init(self):
        """Initialize browser and page"""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=self.config['headless'],
            args=[
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
            ]
        )

        self.page = await self.browser.new_page()
        self.page.set_default_navigation_timeout(self.config['navigationTimeout'])
        await self.page.set_viewport_size({"width": 1920, "height": 1080})

        Path(self.config['outputDir']).mkdir(parents=True, exist_ok=True)

    async def login(self):
        """Login to the platform"""
        print('🔐 جارٍ تسجيل الدخول...')

        await self.page.goto(self.config['baseUrl'],
                            wait_until=self.config['waitForNavigation'])

        # Wait for the login form to actually render before interacting with it
        await self.page.wait_for_selector('input[name="ctl12$ctl03"]', timeout=15000)

        # Fill username/password using the confirmed selectors
        username_input = await self.page.query_selector('input[name="ctl12$ctl03"]')
        password_input = await self.page.query_selector('input[name="ctl12$ctl07"]')

        if not username_input or not password_input:
            raise Exception('لم يتم العثور على حقول تسجيل الدخول')

        await username_input.fill(self.username)
        await password_input.fill(self.password)

        # Submit form - the login button is an <input type="submit">, not a <button>
        submit_button = await self.page.query_selector(
            '#ctl12_btnLogin, input[type="submit"], button[type="submit"]'
        )

        if not submit_button:
            raise Exception('لم يتم العثور على زر تسجيل الدخول')

        await submit_button.click()
        await self.page.wait_for_load_state(self.config['waitForNavigation'])
        await asyncio.sleep(1.5)  # allow for a possible AJAX postback / JS redirect

        # Verify login actually succeeded (still on login page = failed)
        if '/Login' in self.page.url:
            # Grab any visible error message on the page for diagnosis
            error_text = await self.page.evaluate("""
                () => {
                    const selectors = [
                        '[id*="lblError"]', '[id*="lblMsg"]', '[class*="error"]',
                        '[class*="alert"]', '[class*="danger"]', '[id*="Error"]',
                    ];
                    for (const sel of selectors) {
                        const el = document.querySelector(sel);
                        if (el) {
                            const text = (el.innerText || el.textContent || '').trim();
                            if (text) return text;
                        }
                    }
                    return null;
                }
            """)
            await self.page.screenshot(path='login_debug.png', full_page=True)
            print(f'   الرابط بعد محاولة الدخول: {self.page.url}')
            if error_text:
                print(f'   رسالة من الصفحة: {error_text}')
            print('   تم حفظ لقطة شاشة: login_debug.png')
            raise Exception('فشل تسجيل الدخول - تأكد من صحة اسم المستخدم وكلمة المرور')

        print('✅ تم تسجيل الدخول')

    async def goto_search_page(self):
        """Navigate to the search/listing page (Pools/DataView)"""
        print('🔎 جارٍ الانتقال إلى صفحة البحث...')

        await self.page.goto(self.config['searchUrl'],
                            wait_until=self.config['waitForNavigation'])

        # Wait for the search form to actually render before interacting with it
        await self.page.wait_for_selector('input[name="ctl12$ctl28"]', timeout=15000)

        print('✅ تم الوصول إلى صفحة البحث')

    async def select_decision_number(self, decision_number: str):
        """Select decision number"""
        print(f'📋 جارٍ إدخال رقم قرار التسجيل: {decision_number}')

        try:
            input_field = await self.page.query_selector('input[name="ctl12$ctl28"]')

            if input_field:
                await input_field.click()
                await self.page.keyboard.press('Control+A')
                await self.page.keyboard.type(str(decision_number))

                # Trigger change event
                await self.page.evaluate("""
                    () => {
                        const input = document.querySelector('input[name="ctl12$ctl28"]');
                        if (input) {
                            input.dispatchEvent(new Event('input', { bubbles: true }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                    }
                """)

                print(f'✅ تم إدخال رقم القرار: {decision_number}')
            else:
                print('❌ لم يتم العثور على حقل رقم القرار')

            await asyncio.sleep(1.5)
        except Exception as e:
            print(f'⚠️ خطأ في إدخال رقم القرار: {str(e)}')

    async def click_search_button(self):
        """Click search button"""
        print('🔍 جارٍ الضغط على زر البحث...')

        try:
            search_btn = await self.page.query_selector('#ctl12_btnSearch')
            if search_btn:
                await search_btn.click()
                await asyncio.sleep(2)
                print('✅ تم الضغط على البحث')
            else:
                print('❌ لم يتم العثور على زر البحث')
        except Exception as e:
            print(f'⚠️ خطأ في الضغط على البحث: {str(e)}')

    async def open_first_form(self):
        """Open first form from the list"""
        print('📂 جارٍ فتح الاستمارة الأولى...')

        try:
            first_form_link = await self.page.query_selector('a[href*="DataEdit/"]')

            if first_form_link:
                href = await self.page.evaluate('el => el.getAttribute("href")', first_form_link)
                print(f'   الرابط: {href}')

                await first_form_link.click()
                await self.page.wait_for_load_state(self.config['waitForNavigation'])
                print('✅ تم فتح الاستمارة')
            else:
                print('❌ لم يتم العثور على استمارة في الجدول')
        except Exception as e:
            print(f'⚠️ خطأ في فتح الاستمارة: {str(e)}')

    async def get_form_links(self, count: int = None) -> list:
        """Get form links from the results grid (optionally limited to first `count`)"""
        print(f'📑 جارٍ جمع روابط الاستمارات{f" (أول {count})" if count else ""}...')

        links = await self.page.evaluate("""
            () => {
                const anchors = document.querySelectorAll('a[href*="DataEdit/"]');
                return Array.from(anchors).map(a => a.href);
            }
        """)

        if count:
            links = links[:count]

        print(f'✅ تم جمع {len(links)} رابط استمارة')
        return links

    async def get_current_page_form_links(self) -> list:
        """Extract unique DataEdit links from the current results grid page only"""
        links = await self.page.evaluate("""
            () => {
                const grid = document.querySelector('#ctl12_GridView1');
                if (!grid) return [];
                const anchors = grid.querySelectorAll('a[href*="DataEdit/"]');
                const seen = new Set();
                const result = [];
                anchors.forEach(a => {
                    if (!seen.has(a.href)) {
                        seen.add(a.href);
                        result.push(a.href);
                    }
                });
                return result;
            }
        """)
        return links

    async def goto_next_grid_page(self) -> bool:
        """
        Click the next page number in the results grid pager.
        This is an ASP.NET UpdatePanel AJAX postback (__doPostBack), so the URL
        does not change - only the grid content updates in place.

        Because the click itself always "succeeds" from the DOM's point of view
        even if the postback silently fails to update anything, we verify the
        move actually happened by comparing the grid's content before and after
        instead of trusting the click result alone - this avoids looping forever
        on a page that never actually changes.

        Returns False if there is no next page, or the grid content never changes.
        """
        before_links = await self.get_current_page_form_links()
        before_fingerprint = ','.join(before_links[:5])

        click_info = await self.page.evaluate("""
            () => {
                const pagerRow = document.querySelector('#ctl12_GridView1 tr.Pagger');
                if (!pagerRow) return { clicked: false, reason: 'no-pager' };

                let currentPage = null;
                pagerRow.querySelectorAll('td').forEach(td => {
                    const span = td.querySelector('span');
                    if (span) currentPage = parseInt(span.textContent.trim(), 10);
                });
                if (currentPage === null) return { clicked: false, reason: 'no-current-page-span' };

                const links = Array.from(pagerRow.querySelectorAll('a'));
                const target = String(currentPage + 1);
                let nextLink = links.find(a => a.textContent.trim() === target);

                // Fallback: "..." link at the end jumps to the next block of pages
                if (!nextLink && links.length > 0) {
                    const last = links[links.length - 1];
                    if (last.textContent.trim() === '...') {
                        nextLink = last;
                    }
                }

                if (!nextLink) return { clicked: false, reason: 'no-next-link', currentPage };

                nextLink.click();
                return { clicked: true, currentPage, clickedText: nextLink.textContent.trim() };
            }
        """)

        if not click_info.get('clicked'):
            print(f'   ⏹️ لا يوجد صفحة تالية ({click_info.get("reason")})')
            return False

        # Poll for the grid content to actually change (up to ~8 seconds)
        for _ in range(8):
            await asyncio.sleep(1)
            after_links = await self.get_current_page_form_links()
            after_fingerprint = ','.join(after_links[:5])
            if after_fingerprint != before_fingerprint:
                return True

        print(f'   ⚠️ الضغط على "{click_info.get("clickedText")}" ما غيّر محتوى الجدول - توقف')
        return False

    async def collect_all_form_links(self, max_pages: int = 500, min_links: int = None) -> list:
        """
        Crawl the current search results (via AJAX pagination) and collect
        unique DataEdit form links, without opening any of them. This lets the
        caller iterate forms directly afterwards (page.goto per link) instead
        of re-running the search for every single form.

        max_pages is a safety cap to guarantee this always terminates even if
        pagination detection misbehaves; default 500 comfortably covers the
        ~120 pages expected for a 3000-form decision.

        min_links, if given, stops crawling as soon as at least that many
        links have been collected (e.g. testing with 10 forms doesn't need to
        paginate at all if the first page already has 50).
        """
        print('📑 جارٍ جمع روابط الاستمارات من نتائج البحث...')

        all_links = []
        seen = set()
        page_num = 1

        while True:
            page_links = await self.get_current_page_form_links()
            new_count = 0
            for link in page_links:
                if link not in seen:
                    seen.add(link)
                    all_links.append(link)
                    new_count += 1

            print(f'   صفحة {page_num}: {new_count} رابط جديد (الإجمالي: {len(all_links)})')

            if min_links and len(all_links) >= min_links:
                print(f'   ⏹️ تم جمع العدد الكافي ({len(all_links)} >= {min_links})')
                break

            if max_pages and page_num >= max_pages:
                print(f'   ⏹️ تم الوصول للحد الأقصى ({max_pages} صفحة)')
                break

            moved = await self.goto_next_grid_page()
            if not moved:
                break

            page_num += 1

        print(f'✅ تم جمع {len(all_links)} رابط استمارة من {page_num} صفحة نتائج')
        return all_links

    async def navigate_to_classification_criteria(self):
        """Navigate to classification criteria tab"""
        print('📊 جارٍ الانتقال إلى معايير التصنيف...')

        try:
            # Look for "معايير التصنيف" or "التصنيف" tab/link
            criteria_button = await self.page.evaluate("""
                () => {
                    const allElements = document.querySelectorAll('a, button, div, span, li');

                    for (const el of allElements) {
                        const text = el.innerText || el.textContent || '';
                        if (text.includes('التصنيف') || text.includes('معايير')) {
                            const isClickable = el.onclick || el.href || el.getAttribute('data-toggle') || el.classList.toString().includes('tab');
                            if (isClickable) {
                                return true;
                            }
                        }
                    }

                    return false;
                }
            """)

            if criteria_button:
                # Find and click the criteria button
                buttons = await self.page.query_selector_all('a, button, li')

                for btn in buttons:
                    text = await self.page.evaluate('el => el.innerText || el.textContent', btn)

                    if 'التصنيف' in text or 'معايير' in text:
                        await btn.click()
                        await asyncio.sleep(1.5)
                        print('✅ تم الانتقال إلى معايير التصنيف')
                        return

            print('⚠️ لم يتم العثور على زر معايير التصنيف')
        except Exception as e:
            print(f'⚠️ خطأ في الانتقال: {str(e)}')

    async def get_site_name(self) -> str:
        """Extract site name from the page"""
        print('📍 جارٍ استخراج اسم الموقع...')

        site_name = await self.page.evaluate("""
            () => {
                const selectors = [
                    'h1',
                    '[class*="title"]',
                    '[class*="name"]',
                    'span[id*="LabelSiteName"]',
                    'label[id*="LabelSiteName"]',
                ];

                for (const selector of selectors) {
                    const el = document.querySelector(selector);
                    if (el) {
                        const text = (el.innerText || el.textContent || '').trim();
                        if (text && text.length > 0 && text.length < 200) {
                            return text;
                        }
                    }
                }

                // Fallback: Try to get from any element with Arabic text
                const allElements = document.querySelectorAll('*');
                for (const el of allElements) {
                    const text = (el.innerText || el.textContent || '').trim();
                    if (text && /[؀-ۿ]/.test(text) && text.length > 5 && text.length < 100) {
                        return text;
                    }
                }

                return `موقع_${Date.now()}`;
            }
        """)

        # Clean the site name for use as filename
        clean_name = site_name
        clean_name = re.sub(r'[\/\\:*?"<>|]', '_', clean_name)  # Remove invalid filename characters
        clean_name = re.sub(r'\s+', '_', clean_name)  # Replace spaces with underscores
        clean_name = clean_name[:100]  # Limit length

        print(f'   اسم الموقع: {site_name}')
        return clean_name

    async def extract_red_box_table(self) -> dict:
        """Extract classification criteria table info"""
        print('🔍 جارٍ البحث عن جدول معايير التصنيف...')

        table_info = await self.page.evaluate("""
            () => {
                let tableElement = document.querySelector('#ctl12_TemplateRate_GridView1');

                if (!tableElement) {
                    tableElement = document.querySelector('table.GridView, table[class*="Grid"]');
                }

                if (!tableElement) {
                    tableElement = document.querySelector('table');
                }

                if (tableElement) {
                    const rect = tableElement.getBoundingClientRect();
                    const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
                    const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;

                    let nextElement = tableElement.nextElementSibling;
                    let totalHeight = rect.height;

                    while (nextElement && totalHeight < rect.height + 500) {
                        const nextRect = nextElement.getBoundingClientRect();
                        const nextText = (nextElement.innerText || nextElement.textContent || '').trim();

                        if (nextRect.height > 0 && nextText && nextText.length > 0) {
                            totalHeight += nextRect.height + 10;
                        }

                        if (nextElement.tagName === 'TABLE' || nextElement.classList.contains('button') ||
                            nextElement.classList.contains('btn') || nextElement.tagName === 'BUTTON') {
                            break;
                        }

                        nextElement = nextElement.nextElementSibling;
                    }

                    return {
                        found: true,
                        x: Math.max(0, rect.left + scrollLeft - 5),
                        y: Math.max(0, rect.top + scrollTop - 5),
                        width: rect.width + 10,
                        height: totalHeight + 10,
                        tagName: tableElement.tagName,
                        classes: tableElement.className,
                        id: tableElement.id,
                    };
                }

                return { found: false };
            }
        """)

        if not table_info.get('found'):
            raise Exception('لم يتم العثور على جدول معايير التصنيف')

        print(f"✅ تم العثور على الجدول: {table_info['width']}x{table_info['height']}px")
        return table_info

    async def export_table_as_clean_pdf(self, filename: str, table_info: dict) -> str:
        """Export table as PDF with clipping"""
        print(f'📄 جارٍ تحويل الجدول إلى PDF: {filename}')

        filepath = os.path.join(self.config['outputDir'], filename)

        if table_info and table_info.get('width') and table_info.get('height'):
            # Create a clipped PDF that captures only the table and summary below it
            await self.page.pdf(
                path=filepath,
                clip={
                    'x': max(0, table_info['x']),
                    'y': max(0, table_info['y']),
                    'width': table_info['width'],
                    'height': table_info['height'],
                },
                format='A4',
                margin={'top': 0, 'right': 0, 'bottom': 0, 'left': 0},
            )
        else:
            # Fallback: Full page PDF
            await self.page.pdf(
                path=filepath,
                format='A4',
                print_background=True,
                margin={'top': '5mm', 'right': '5mm', 'bottom': '5mm', 'left': '5mm'},
            )

        print(f'✅ تم التصدير: {filepath}')
        return filepath

    async def capture_table_screenshot(self, filename: str, table_info: dict) -> str:
        """Capture table screenshot"""
        print('📸 جارٍ حفظ لقطة شاشة الجدول...')

        screenshot_path = os.path.join(self.config['outputDir'], filename.replace('.pdf', '.png'))

        if table_info and table_info.get('width') and table_info.get('height'):
            # Capture just the table and summary
            await self.page.screenshot(
                path=screenshot_path,
                clip={
                    'x': max(0, table_info['x']),
                    'y': max(0, table_info['y']),
                    'width': table_info['width'],
                    'height': table_info['height'],
                },
            )
        else:
            # Full screenshot
            await self.page.screenshot(path=screenshot_path, full_page=True)

        print(f'✅ تم حفظ لقطة الشاشة: {screenshot_path}')
        return screenshot_path

    async def export_batch(self, decisions: list) -> list:
        """Export batch of decisions"""
        results = []

        for i, decision in enumerate(decisions):
            try:
                print(f'\n[{i + 1}/{len(decisions)}] معالجة: {decision["name"]}')

                await asyncio.sleep(0.8)

                # Select decision number if provided
                if decision.get('decisionNumber'):
                    await self.select_decision_number(decision['decisionNumber'])
                    await self.click_search_button()
                    await asyncio.sleep(2)

                # Navigate if needed
                if decision.get('url'):
                    await self.page.goto(decision['url'],
                                        wait_until=self.config['waitForNavigation'])

                # Open first form if not already navigated
                if not decision.get('url'):
                    await self.open_first_form()

                # Navigate to classification criteria
                await self.navigate_to_classification_criteria()

                # Extract table info
                table_info = await self.extract_red_box_table()

                # Get site name from the page
                site_name = await self.get_site_name()

                # Export as PDF with site name
                timestamp = int(datetime.now().timestamp() * 1000)
                pdf_file = f'{site_name}_{timestamp}.pdf'
                png_file = f'{site_name}_{timestamp}.png'

                pdf_path = await self.export_table_as_clean_pdf(pdf_file, table_info)
                png_path = await self.capture_table_screenshot(png_file, table_info)

                results.append({
                    'name': decision['name'],
                    'siteName': site_name,
                    'decisionNumber': decision.get('decisionNumber'),
                    'status': 'success',
                    'pdfPath': pdf_path,
                    'pngPath': png_path,
                    'tableInfo': table_info,
                })
            except Exception as e:
                print(f'❌ خطأ: {str(e)}')
                results.append({
                    'name': decision['name'],
                    'decisionNumber': decision.get('decisionNumber'),
                    'status': 'failed',
                    'error': str(e),
                })

        return results

    async def scroll_table_and_load_data(self, max_rows: int = None) -> int:
        """Scroll table and load data"""
        print('📜 جارٍ تحميل كل البيانات...')

        rows_loaded = await self.page.evaluate("""
            (max) => {
                let lastHeight = 0;
                let scrolls = 0;

                const tableContainer = document.querySelector('table') ||
                                     document.querySelector('[role="table"]') ||
                                     document.querySelector('.table-container');

                if (!tableContainer) return 0;

                const interval = setInterval(() => {
                    tableContainer.scrollTop = tableContainer.scrollHeight;
                    scrolls++;

                    if (max && scrolls > max) {
                        clearInterval(interval);
                    }
                }, 300);

                return new Promise((resolve) => {
                    setTimeout(() => {
                        clearInterval(interval);
                        const rows = document.querySelectorAll('tr, [role="row"]');
                        resolve(rows.length);
                    }, scrolls * 300 + 500);
                });
            }
        """, max_rows)

        print(f'✅ تم تحميل {rows_loaded} صف')
        return rows_loaded

    async def close(self):
        """Close browser"""
        if self.browser:
            await self.browser.close()
            await self.playwright.stop()
            print('🔌 تم إغلاق المتصفح')

    async def run(self, decisions: list) -> list:
        """Run the exporter"""
        try:
            await self.init()
            await self.login()
            await self.goto_search_page()
            await self.scroll_table_and_load_data()

            results = await self.export_batch(decisions)
            await self.close()

            return results
        except Exception as e:
            await self.close()
            raise e


async def main():
    """Main function"""
    import sys

    if len(sys.argv) < 3:
        print('استخدام: python table_pdf_exporter_advanced.py <username> <password>')
        sys.exit(1)

    username = sys.argv[1]
    password = sys.argv[2]

    decisions = [
        {'name': 'قرار_ترشيح_1'},
        {'name': 'قرار_ترشيح_2'},
        {'name': 'قرار_ترشيح_3'},
    ]

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
    except Exception as e:
        print(f'❌ فشل التنفيذ: {str(e)}')
        sys.exit(1)


if __name__ == '__main__':
    asyncio.run(main())
