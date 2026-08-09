#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
منصة اكتشاف وتسجيل أصول التراث العمراني - أتمتة التعديلات
Automation for Heritage Platform - Mass Update Script
"""

import os
import re
import json
import sqlite3
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# ==================== إعدادات ====================
BASE_URL = "https://heritage2.itqan-consultant.com/Web/"
LOGIN_URL = "https://heritage2.itqan-consultant.com/Web/App/Home/Login"
DATAVIEW_URL = "https://heritage2.itqan-consultant.com/Web/App/Pools/DataView"
# صفحة "بيانات المعماري" لأي موقع تُبنى مباشرة من كوده (رقم) - تم تأكيدها من
# HTML خام حقيقي (form action="./7"، وتاب "بيانات المعماري" برابط href="7")
ARCHITECT_DATA_URL_TEMPLATE = "https://heritage2.itqan-consultant.com/Web/App/Pools/DataEdit/{site_id}/7"

# ملاحظة: جميع صفحات هذا النظام (تسجيل الدخول، دليل المواقع، بيانات المعماري)
# مؤكدة من HTML خام أنها تستخدم بادئة عناصر ASP.NET "ctl12" (بنفس القيمة عبر
# الصفحات الثلاث، على الأرجح بسبب Master Page مشتركة بنفس عدد العناصر السابقة).
LOGIN_USERNAME_FIELD_NAME = "ctl12$ctl03"       # مؤكد من HTML صفحة /Home/Login
LOGIN_PASSWORD_FIELD_NAME = "ctl12$ctl07"       # مؤكد من HTML صفحة /Home/Login
LOGIN_BUTTON_ID = "ctl12_btnLogin"              # مؤكد من HTML صفحة /Home/Login
GRIDVIEW_ID = "ctl12_GridView1"                 # مؤكد من HTML صفحة DataView
GRIDVIEW_POSTBACK_TARGET = "ctl12$GridView1"    # مستخدم مع __doPostBack للترقيم
RESULTS_COUNT_LABEL_ID = "ctl12_lblCount"       # نص: "نتيجة البحث: N موقع"
SEARCH_BUTTON_ID = "ctl12_btnSearch"
REGION_FILTER_SELECT_NAME = "ctl12$ctl08"       # فلتر "المنطقة" الجغرافية (محافظات المملكة)
STEP_FILTER_SELECT_NAME = "ctl12$ctl32"         # فلتر "الخطوة"
REGION_FILTER_VALUE_ALBAHA = "12"               # قيمة خيار "الباحة" بقائمة المنطقة
STEP_FILTER_VALUE_INITIAL_REG = "700"           # قيمة خيار "التسجيل المبدئي" بقائمة الخطوة
RESULTS_PER_PAGE = 25
SAVE_BUTTON_ID = "ctl12_btnSave"                # مؤكد من HTML صفحة DataEdit/{id}/7

USERNAME = os.environ.get("HERITAGE_USERNAME")
PASSWORD = os.environ.get("HERITAGE_PASSWORD")

# "المنطقة الجنوبية" هنا تخص حقل "الطراز المعماري" (radio button، قسم 4)
# داخل نموذج بيانات الموقع بتبويب بيانات المعماري - لا علاقة لها بفلتر
# "المنطقة" الجغرافي أعلاه رغم تشابه الاسم.
TARGET_REGION = "المنطقة الجنوبية"

if not USERNAME or not PASSWORD:
    raise SystemExit(
        "يجب ضبط بيانات الدخول عبر متغيرات البيئة قبل التشغيل:\n"
        "  export HERITAGE_USERNAME='...'\n"
        "  export HERITAGE_PASSWORD='...'\n"
        "أو استخدم ملف .env (راجع README_AUTOMATION.md)"
    )

# مسارات الملفات
DB_PATH = Path("heritage_automation.db")
CHECKPOINT_PATH = Path("checkpoint.json")
EXCEL_OUTPUT = Path("heritage_report.xlsx")
LOG_FILE = Path("heritage_automation.log")

# إعدادات Selenium
HEADLESS = True
IMPLICIT_WAIT = 10
EXPLICIT_WAIT = 20
LOGIN_WAIT_SECONDS = 60   # المنصة أحياناً تتعلق/تبطئ أثناء تسجيل الدخول تحديداً
LOGIN_MAX_RETRIES = 3     # عدد محاولات إعادة تسجيل الدخول قبل الفشل النهائي

# ==================== إعداد Logging ====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ==================== Database Functions ====================
class HeritageDB:
    """إدارة قاعدة البيانات"""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.init_db()

    def init_db(self):
        """إنشاء جداول البيانات"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                site_id TEXT UNIQUE NOT NULL,
                site_name TEXT,
                site_url TEXT,
                status TEXT DEFAULT 'pending',
                steps_completed TEXT DEFAULT '',
                error_message TEXT,
                processed_time REAL DEFAULT 0,
                processed_at TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS checkpoint (
                id INTEGER PRIMARY KEY,
                last_site_id TEXT,
                total_sites INTEGER,
                processed_count INTEGER,
                failed_count INTEGER,
                skipped_count INTEGER,
                start_time TIMESTAMP,
                last_update TIMESTAMP
            )
        """)

        conn.commit()
        conn.close()

    def insert_site(self, site_id: str, site_name: str, site_url: str):
        """إضافة موقع جديد"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT OR IGNORE INTO sites (site_id, site_name, site_url)
                VALUES (?, ?, ?)
            """, (site_id, site_name, site_url))
            conn.commit()
        finally:
            conn.close()

    def update_site_status(self, site_id: str, status: str, steps: str = "", error: str = "", duration: float = 0):
        """تحديث حالة الموقع"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("""
                UPDATE sites
                SET status = ?, steps_completed = ?, error_message = ?, processed_time = ?, processed_at = CURRENT_TIMESTAMP
                WHERE site_id = ?
            """, (status, steps, error, duration, site_id))
            conn.commit()
        finally:
            conn.close()

    def get_pending_sites(self, limit: int = 0) -> List[Tuple]:
        """الحصول على المواقع المعلقة"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            if limit > 0:
                cursor.execute("SELECT site_id, site_name, site_url FROM sites WHERE status = 'pending' LIMIT ?", (limit,))
            else:
                cursor.execute("SELECT site_id, site_name, site_url FROM sites WHERE status = 'pending'")
            return cursor.fetchall()
        finally:
            conn.close()

    def get_statistics(self) -> Dict:
        """الحصول على إحصائيات"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT status, COUNT(*) FROM sites GROUP BY status")
            stats = {row[0]: row[1] for row in cursor.fetchall()}

            cursor.execute("SELECT COUNT(*) FROM sites")
            total = cursor.fetchone()[0]

            return {
                'total': total,
                'success': stats.get('success', 0),
                'failed': stats.get('failed', 0),
                'skipped': stats.get('skipped', 0),
                'pending': stats.get('pending', 0)
            }
        finally:
            conn.close()

    def get_all_sites_report(self) -> List[Dict]:
        """الحصول على تقرير كامل"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT site_id, site_name, status, steps_completed, error_message, processed_time, processed_at
                FROM sites
                ORDER BY id
            """)

            columns = ['site_id', 'site_name', 'status', 'steps_completed', 'error_message', 'processed_time', 'processed_at']
            return [dict(zip(columns, row)) for row in cursor.fetchall()]
        finally:
            conn.close()

# ==================== Checkpoint Functions ====================
class CheckpointManager:
    """إدارة نقاط التوقف"""

    @staticmethod
    def save(last_site_id: str, total: int, processed: int, failed: int, skipped: int):
        """حفظ نقطة توقف"""
        data = {
            'last_site_id': last_site_id,
            'total_sites': total,
            'processed_count': processed,
            'failed_count': failed,
            'skipped_count': skipped,
            'timestamp': datetime.now().isoformat()
        }
        with open(CHECKPOINT_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logger.info(f"✓ تم حفظ نقطة توقف: {last_site_id}")

    @staticmethod
    def load() -> Optional[Dict]:
        """تحميل نقطة توقف"""
        if CHECKPOINT_PATH.exists():
            with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None

    @staticmethod
    def clear():
        """حذف نقطة توقف"""
        if CHECKPOINT_PATH.exists():
            CHECKPOINT_PATH.unlink()

# ==================== Selenium Driver ====================
class HeritageDriver:
    """إدارة متصفح Selenium"""

    def __init__(self):
        self.driver = None
        self.wait = None
        self.init_driver()

    def init_driver(self):
        """تهيئة المتصفح"""
        options = Options()

        # Headless Mode
        if HEADLESS:
            options.add_argument('--headless=new')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')

        # Profile منفصل (اختياري)
        profile_path = Path.home() / ".wdm" / "heritage_automation"
        profile_path.mkdir(parents=True, exist_ok=True)
        options.add_argument(f'user-data-dir={profile_path}')

        # مسار متصفح مخصص (اختياري) - فقط عند الحاجة لتجاوز Chrome الافتراضي بالنظام
        chrome_binary = os.environ.get("CHROME_BINARY_PATH")
        if chrome_binary:
            options.binary_location = chrome_binary

        # إعدادات أخرى
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('start-maximized')
        # 'start-maximized' لا يعمل بوضع headless (لا توجد نافذة حقيقية تتكبر)،
        # فنجبر حجم نافذة كبير صراحةً - يمنع تراكب عناصر ثابتة الموضع (position:
        # absolute) في أسفل الصفحة فوق أزرار الصفحة بسبب viewport افتراضي صغير
        options.add_argument('--window-size=1920,1080')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        try:
            self.driver = webdriver.Chrome(options=options)
            self.driver.implicitly_wait(IMPLICIT_WAIT)
            self.wait = WebDriverWait(self.driver, EXPLICIT_WAIT)
            logger.info("✓ تم فتح المتصفح بنجاح")
        except Exception as e:
            logger.error(f"✗ خطأ في فتح المتصفح: {e}")
            raise

    def _safe_click(self, element):
        """كليك آمن: يجرب الكليك العادي، وإن اعترضه عنصر آخر (تراكب عناصر
        ثابتة الموضع مثلاً) يلجأ لكليك عبر JavaScript الذي لا يتأثر بذلك"""
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)

    def close(self):
        """إغلاق المتصفح"""
        if self.driver:
            self.driver.quit()
            logger.info("✓ تم إغلاق المتصفح")

    def save_debug_snapshot(self, name: str):
        """حفظ صورة شاشة + مصدر الصفحة الحالية للتشخيص (يُستخدم عند فشل خطوة
        حرجة مثل تسجيل الدخول، لأن المتصفح يعمل بالخلفية بدون نافذة مرئية)"""
        try:
            screenshot_path = f"{name}.png"
            html_path = f"{name}.html"
            self.driver.save_screenshot(screenshot_path)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
            logger.info(f"🖼️  تم حفظ لقطة تشخيصية: {screenshot_path} و {html_path}")
            logger.info(f"🔗 الرابط الحالي وقت الفشل: {self.driver.current_url}")
        except Exception as e:
            logger.warning(f"⚠️ تعذر حفظ اللقطة التشخيصية: {e}")

    def login(self, retries: int = LOGIN_MAX_RETRIES) -> bool:
        """تسجيل الدخول (selectors مؤكدة 100% من HTML خام لصفحة /Home/Login).
        المنصة أحياناً بتتعلق/تبطئ، فهذي الخطوة تحديداً تنتظر مدة أطول
        (LOGIN_WAIT_SECONDS) وتعيد المحاولة عدة مرات قبل ما تفشل نهائياً.

        ⚠️ بما أن المتصفح يستخدم بروفايل Chrome محفوظ بين التشغيلات (لتبقى
        الجلسة قائمة)، فإذا كانت الجلسة مسجلة دخول أصلاً (من محاولة سابقة أو
        تشغيلة سابقة)، خادم المنصة لا يعرض نموذج اليوزر/الباسورد إطلاقاً عند
        فتح /Login - فنتحقق من ذلك أولاً ونتخطى تعبئة النموذج إن كانت الجلسة
        فعّالة أصلاً، بدل ما ننتظر عبثاً حقلاً غير موجود."""
        for attempt in range(1, retries + 1):
            try:
                logger.info(f"🔐 جاري تسجيل الدخول... (محاولة {attempt}/{retries})")
                self.driver.get(LOGIN_URL)

                # تحقق سريع: هل الجلسة مسجلة دخول أصلاً؟ (بروفايل محفوظ من
                # تشغيلة/محاولة سابقة) - إن كانت كذلك، لا داعي لتعبئة النموذج
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.ID, "TopBar_lblName"))
                    )
                    logger.info("✓ الجلسة مسجلة دخول أصلاً (من تشغيلة سابقة) - تخطي نموذج الدخول")
                    return True
                except TimeoutException:
                    pass  # لسا صفحة الدخول الفعلية ظاهرة، كمل بالتعبئة العادية أدناه

                # إدخال اسم المستخدم (name="ctl12$ctl03")
                username_field = self.wait.until(
                    EC.presence_of_element_located((By.NAME, LOGIN_USERNAME_FIELD_NAME))
                )
                username_field.clear()
                username_field.send_keys(USERNAME)
                time.sleep(0.5)

                # إدخال كلمة المرور (name="ctl12$ctl07")
                password_field = self.driver.find_element(By.NAME, LOGIN_PASSWORD_FIELD_NAME)
                password_field.clear()
                password_field.send_keys(PASSWORD)
                time.sleep(0.5)

                # النقر على زر "تسجيل الدخول" (id=ctl12_btnLogin)
                login_button = self.driver.find_element(By.ID, LOGIN_BUTTON_ID)
                self._safe_click(login_button)

                # ⚠️ هذه المنصة لا تُغيّر رابط الصفحة بعد تسجيل الدخول (تحدّث
                # المحتوى عبر AJAX/UpdatePanel وتبقى على /Login في شريط
                # العنوان). لذلك العلامة الصحيحة على نجاح الدخول هي ظهور
                # عنصر لا يظهر إلا بعد المصادقة: اسم المستخدم بأعلى الصفحة
                # (id=TopBar_lblName) - تم تأكيدها من لقطة تشخيصية حقيقية
                # أظهرت "Mohammed sobhy" رغم بقاء الرابط على /Login.
                WebDriverWait(self.driver, LOGIN_WAIT_SECONDS).until(
                    EC.presence_of_element_located((By.ID, "TopBar_lblName"))
                )
                time.sleep(2)

                logger.info(f"✓ تم تسجيل الدخول بنجاح - الرابط الحالي: {self.driver.current_url}")
                return True
            except TimeoutException as e:
                logger.warning(
                    f"⚠️ انتهت مهلة انتظار عنصر خلال محاولة تسجيل الدخول {attempt}/{retries} "
                    f"(يُحتمل بطء بالمنصة أو خطأ باليوزر/الباسورد أو تغيّر بهيكل الصفحة): {e}"
                )
                if attempt == retries:
                    self.save_debug_snapshot("login_failure_debug")
            except Exception as e:
                logger.warning(f"⚠️ خطأ بمحاولة تسجيل الدخول {attempt}/{retries}: {e}")
                if attempt == retries:
                    self.save_debug_snapshot("login_failure_debug")

        logger.error(f"✗ فشل تسجيل الدخول نهائياً بعد {retries} محاولات")
        return False

    def navigate_to_guide_list(self) -> bool:
        """الانتقال إلى دليل المواقع (رابط مباشر مؤكد)"""
        try:
            logger.info("📍 جاري الانتقال إلى دليل المواقع...")
            self.driver.get(DATAVIEW_URL)

            self.wait.until(EC.presence_of_element_located((By.ID, GRIDVIEW_ID)))

            logger.info("✓ تم الانتقال إلى دليل المواقع")
            return True
        except Exception as e:
            logger.error(f"✗ خطأ في الانتقال إلى دليل المواقع: {e}")
            return False

    def goto_gridview_page(self, page_number: int) -> bool:
        """الانتقال مباشرة إلى رقم صفحة معين داخل الجدول عبر __doPostBack
        (يعمل حتى لو رقم الصفحة غير ظاهر في شريط الترقيم المرئي)"""
        try:
            # نحتفظ بمرجع للجدول القديم للتأكد من إعادة تحميله (postback) قبل المتابعة
            old_table = self.driver.find_element(By.ID, GRIDVIEW_ID)

            self.driver.execute_script(
                f"__doPostBack('{GRIDVIEW_POSTBACK_TARGET}','Page${page_number}')"
            )

            # ننتظر أن يصبح الجدول القديم "stale" (تم استبداله فعلياً بعد الـ postback)
            self.wait.until(EC.staleness_of(old_table))
            self.wait.until(EC.presence_of_element_located((By.ID, GRIDVIEW_ID)))
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"⚠️ خطأ في الانتقال إلى صفحة {page_number}: {e}")
            return False

    def apply_filters(self) -> bool:
        """تطبيق الفلاتر: المنطقة=الباحة (value=12) والخطوة=التسجيل المبدئي (value=700)
        باستخدام القيم الرقمية الثابتة (value) بدل نص الخيار - أضمن من مطابقة النص"""
        try:
            logger.info("🔍 جاري تطبيق الفلاتر...")

            old_count_label = self.driver.find_element(By.ID, RESULTS_COUNT_LABEL_ID)

            # اختيار المنطقة: الباحة (value=12)
            region_select = Select(self.driver.find_element(By.NAME, REGION_FILTER_SELECT_NAME))
            region_select.select_by_value(REGION_FILTER_VALUE_ALBAHA)
            time.sleep(1.5)  # تحديث تلقائي (AJAX) لقوائم الفرع/المحافظة المرتبطة

            # اختيار الخطوة: التسجيل المبدئي (value=700)
            step_select = Select(self.driver.find_element(By.NAME, STEP_FILTER_SELECT_NAME))
            step_select.select_by_value(STEP_FILTER_VALUE_INITIAL_REG)
            time.sleep(1)

            # الضغط على زر "بحث" لتطبيق الفلترة فعلياً على الجدول
            search_button = self.driver.find_element(By.ID, SEARCH_BUTTON_ID)
            self._safe_click(search_button)

            # ننتظر تحديث تسمية عدد النتائج (يثبت أن نتائج الفلترة وصلت)
            self.wait.until(EC.staleness_of(old_count_label))
            self.wait.until(EC.presence_of_element_located((By.ID, RESULTS_COUNT_LABEL_ID)))
            time.sleep(1)

            logger.info("✓ تم تطبيق الفلاتر بنجاح")
            return True
        except Exception as e:
            logger.error(f"✗ خطأ في تطبيق الفلاتر: {e}")
            return False

    def get_total_results_count(self) -> int:
        """قراءة عدد النتائج الإجمالي من ctl12_lblCount، نصه مثل: 'نتيجة البحث: 5256  موقع'"""
        try:
            label_text = self.driver.find_element(By.ID, RESULTS_COUNT_LABEL_ID).text
            match = re.search(r'(\d+)', label_text.replace(',', ''))
            return int(match.group(1)) if match else 0
        except Exception as e:
            logger.warning(f"⚠️ تعذر قراءة إجمالي النتائج: {e}")
            return 0

    def extract_all_site_links(self) -> List[Tuple[str, str, str]]:
        """استخراج أكواد وأسماء المواقع من الصفحة الحالية (بدون الحاجة لأي رابط
        بالجدول - الرابط يُبنى لاحقاً مباشرة من الكود عبر ARCHITECT_DATA_URL_TEMPLATE)"""
        sites = []
        try:
            logger.info("📥 جاري استخراج المواقع من الصفحة الحالية...")

            # ملاحظة: صف الرأس يستخدم <th> (يُستبعد تلقائياً لعدم وجود <td> فيه).
            # صف الترقيم (Pagger) يحتوي <td><table><tr>...</tr></table></td> -
            # جدول متداخل بداخل الجدول الرئيسي. استخدام ".//tbody/tr" (بحث
            # متكرر بكل المستويات) كان يلتقط أيضاً صفوف الجدول الداخلي هذا
            # (أرقام الصفحات 1، 2، 3...) رغم استبعاد class='Pagger'، لأن
            # الاستبعاد ينطبق فقط على الصف الخارجي وليس صفوف الجدول المتداخل
            # بداخله. الحل: "./tbody/tr" (مستوى واحد فقط، بدون تكرار) يقتصر
            # على صفوف الجدول الرئيسي المباشرة ولا ينزل للجداول المتداخلة.
            table = self.driver.find_element(By.ID, GRIDVIEW_ID)
            rows = table.find_elements(
                By.XPATH, "./tbody/tr[not(contains(@class, 'Pagger'))]"
            )
            logger.info(f"📌 عدد الصفوف في الصفحة الحالية: {len(rows)}")

            for row in rows:
                try:
                    cells = row.find_elements(By.XPATH, "./td")
                    if not cells:
                        continue

                    # الكود هو أول عمود (أقصى اليمين في RTL = أول td في الـ DOM)
                    site_id = cells[0].text.strip()
                    if not site_id.isdigit():
                        continue

                    site_name = cells[1].text.strip() if len(cells) > 1 else site_id
                    site_url = ARCHITECT_DATA_URL_TEMPLATE.format(site_id=site_id)

                    sites.append((site_id, site_name, site_url))
                    logger.debug(f"  └─ {site_id}: {site_name}")
                except Exception as e:
                    logger.warning(f"⚠️ خطأ في استخراج صف: {e}")
                    continue

            logger.info(f"✓ تم استخراج {len(sites)} موقع من الصفحة الحالية")
            return sites
        except Exception as e:
            logger.error(f"✗ خطأ في استخراج الروابط: {e}")
            return sites

    def open_site(self, site_url: str) -> bool:
        """فتح صفحة الموقع (تفتح مباشرة على تاب "بيانات المعماري" بفضل الرابط المباشر)"""
        try:
            self.driver.get(site_url)

            # ننتظر عنصراً مؤكداً من هذه الصفحة تحديداً: زر الحفظ (id=ctl12_btnSave)
            self.wait.until(EC.presence_of_element_located((By.ID, SAVE_BUTTON_ID)))
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"⚠️ خطأ في فتح الموقع: {e}")
            return False

    def update_region(self) -> Tuple[bool, str]:
        """تحديث حقل 'الطراز المعماري' (قسم 4) لاختيار 'المنطقة الجنوبية'.

        ⚠️ مؤكد من HTML حقيقي: هذا الحقل عبارة عن radio buttons منفصلة عن
        بعضها (وليس <select> كما كان مفترضاً سابقاً)، كل واحد له <label>
        مستقل يحمل خاصية for تشير لمعرّف الـ <input type="radio">."""
        try:
            logger.info("  ↳ جاري تحديث المنطقة...")

            labels = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, f"//label[normalize-space(text())='{TARGET_REGION}']")
                )
            )
            if len(labels) > 1:
                logger.warning(f"  ⚠️ عثر على {len(labels)} تطابقات لِـ '{TARGET_REGION}' - سيُستخدم الأول")
            label = labels[0]

            radio_id = label.get_attribute('for')
            radio = self.driver.find_element(By.ID, radio_id)

            if radio.is_selected():
                logger.info("  ⏭️  المنطقة مختارة بالفعل (المنطقة الجنوبية)")
                return True, "skipped"

            # النقر على الـ label (أضمن من input مباشرة، فقد يكون مغطى بصرياً)
            self._safe_click(label)
            time.sleep(0.5)

            if radio.is_selected():
                logger.info("  ✓ تم تحديث المنطقة بنجاح")
                return True, "updated"
            else:
                logger.warning("  ⚠️ فشل التحديث - الخيار لم يُحدد بعد النقر")
                return False, "failed"
        except Exception as e:
            logger.warning(f"  ✗ خطأ في تحديث المنطقة: {e}")
            return False, f"error: {str(e)}"

    def save_form(self) -> bool:
        """حفظ الاستمارة (زر مؤكد من HTML خام: id=ctl12_btnSave،
        value='حـفــظ بيانــات الاستمـــارة')"""
        try:
            logger.info("  ↳ جاري حفظ الاستمارة...")

            save_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, SAVE_BUTTON_ID))
            )
            self._safe_click(save_button)

            # انتظر اكتمال الـ postback بعد الحفظ
            time.sleep(2)

            logger.info("  ✓ تم إرسال الحفظ")
            return True
        except Exception as e:
            logger.warning(f"  ✗ خطأ في الحفظ: {e}")
            return False

# ==================== Main Automation ====================
class HeritageAutomation:
    """فئة الأتمتة الرئيسية"""

    def __init__(self):
        self.db = HeritageDB(DB_PATH)
        self.driver_manager = HeritageDriver()
        self.total_processed = 0
        self.total_failed = 0
        self.total_skipped = 0

    def run(self, resume_from: Optional[str] = None, limit: Optional[int] = None):
        """تشغيل الأتمتة. limit: إن حُدد، يعالج هذا العدد فقط من المواقع
        (مفيد لتجربة سريعة قبل التشغيل الكامل على كل المواقع)"""
        start_time = datetime.now()
        logger.info("=" * 50)
        logger.info("🚀 بدء أتمتة منصة التراث العمراني")
        if limit:
            logger.info(f"🧪 وضع تجربة: سيُعالَج {limit} موقع فقط")
        logger.info(f"⏰ وقت البداية: {start_time}")
        logger.info("=" * 50)

        try:
            # تسجيل الدخول
            if not self.driver_manager.login():
                logger.error("❌ فشل تسجيل الدخول")
                return False

            time.sleep(2)

            # الانتقال إلى دليل المواقع
            if not self.driver_manager.navigate_to_guide_list():
                logger.error("❌ فشل الانتقال إلى دليل المواقع")
                return False

            time.sleep(2)

            # تطبيق الفلاتر
            if not self.driver_manager.apply_filters():
                logger.error("❌ فشل تطبيق الفلاتر")
                return False

            time.sleep(2)

            # استخراج المواقع (من جميع الصفحات - هذا جزء مهم يحتاج iteration)
            all_sites = self._extract_all_paginated_sites(limit=limit)

            if not all_sites:
                logger.error("❌ لم يتم استخراج أي مواقع")
                return False

            logger.info(f"📊 إجمالي المواقع المستخرجة: {len(all_sites)}")

            # إضافة المواقع للقاعدة
            for site_id, site_name, site_url in all_sites:
                self.db.insert_site(site_id, site_name, site_url)

            # معالجة كل موقع
            self._process_all_sites(resume_from, limit)

            # إنشاء التقرير
            self._generate_report()

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()

            logger.info("=" * 50)
            logger.info(f"✅ انتهت الأتمتة بنجاح")
            logger.info(f"⏰ وقت الانتهاء: {end_time}")
            logger.info(f"⏱️ المدة الإجمالية: {duration:.2f} ثانية ({duration/60:.2f} دقيقة)")
            logger.info("=" * 50)

            return True
        except Exception as e:
            logger.error(f"❌ خطأ عام: {e}", exc_info=True)
            return False
        finally:
            self.driver_manager.close()

    def _extract_all_paginated_sites(self, limit: Optional[int] = None) -> List[Tuple[str, str, str]]:
        """استخراج المواقع من جميع الصفحات باستخدام __doPostBack المباشر
        (لا يعتمد على النقر على أزرار الصفحات المرئية، يدعم 200+ صفحة بأمان).
        إن حُدد limit، يتوقف الاستخراج فور جمع عدد كافٍ من المواقع (توفيراً
        للوقت أثناء التجربة على عينة صغيرة)."""
        all_sites = []

        try:
            total_sites = self.driver_manager.get_total_results_count()
            total_pages = (total_sites + RESULTS_PER_PAGE - 1) // RESULTS_PER_PAGE if total_sites else 1
            logger.info(f"📊 إجمالي المواقع المعلن: {total_sites} — إجمالي الصفحات المتوقع: {total_pages}")

            page = 1
            while page <= total_pages:
                logger.info(f"📄 جاري استخراج الصفحة {page}/{total_pages}...")

                sites_in_page = self.driver_manager.extract_all_site_links()
                if not sites_in_page:
                    logger.warning(f"⚠️ لا توجد مواقع بالصفحة {page} — توقف الاستخراج")
                    break

                all_sites.extend(sites_in_page)

                if limit and len(all_sites) >= limit:
                    logger.info(f"🧪 تم جمع {len(all_sites)} موقع (>= limit={limit}) — إيقاف الاستخراج المبكر")
                    break

                if page >= total_pages:
                    break

                page += 1
                if not self.driver_manager.goto_gridview_page(page):
                    logger.error(f"❌ فشل الانتقال إلى الصفحة {page} — توقف الاستخراج")
                    break
        except Exception as e:
            logger.warning(f"⚠️ خطأ في استخراج الصفحات: {e}")

        return all_sites

    def _process_all_sites(self, resume_from: Optional[str] = None, limit: Optional[int] = None):
        """معالجة جميع المواقع (أو أول limit موقع فقط إن حُدد)"""
        pending_sites = self.db.get_pending_sites()

        start_index = 0
        if resume_from:
            logger.info(f"🔄 استئناف من الموقع: {resume_from}")
            for i, (site_id, _, _) in enumerate(pending_sites):
                if site_id == resume_from:
                    start_index = i
                    break

        batch = pending_sites[start_index:]
        if limit:
            batch = batch[:limit]

        for index, (site_id, site_name, site_url) in enumerate(batch, 1):
            logger.info(f"[{index}/{len(batch)}] 🔄 معالجة الموقع: {site_id} - {site_name}")

            process_start = time.time()
            steps = []

            try:
                # فتح الموقع (يفتح مباشرة على تاب بيانات المعماري عبر الرابط المباشر)
                steps.append("فتح صفحة بيانات المعماري")
                if not self.driver_manager.open_site(site_url):
                    raise Exception("فشل فتح الموقع")

                # تحديث المنطقة
                steps.append("تحديث المنطقة")
                updated, update_status = self.driver_manager.update_region()

                if update_status == "skipped":
                    steps.append("تخطي (مختار بالفعل)")
                    status = "skipped"
                    self.total_skipped += 1
                elif updated:
                    steps.append("حفظ البيانات")
                    if self.driver_manager.save_form():
                        steps.append("تأكيد الحفظ")
                        status = "success"
                        self.total_processed += 1
                    else:
                        raise Exception("فشل الحفظ")
                else:
                    raise Exception("فشل التحديث")

                process_duration = time.time() - process_start
                self.db.update_site_status(site_id, status, " → ".join(steps), "", process_duration)
                logger.info(f"  ✅ {status.upper()}")

            except Exception as e:
                process_duration = time.time() - process_start
                error_msg = str(e)
                steps.append(f"خطأ: {error_msg}")
                self.db.update_site_status(site_id, "failed", " → ".join(steps[:-1]), error_msg, process_duration)
                self.total_failed += 1
                logger.error(f"  ❌ خطأ: {error_msg}")

            # حفظ checkpoint كل 20 موقع
            if (index) % 20 == 0:
                stats = self.db.get_statistics()
                CheckpointManager.save(site_id, stats['total'], stats['success'], stats['failed'], stats['skipped'])

    def _generate_report(self):
        """إنشاء تقرير Excel"""
        try:
            logger.info("📊 جاري إنشاء التقرير...")

            # الحصول على البيانات
            report_data = self.db.get_all_sites_report()
            stats = self.db.get_statistics()

            # إنشاء DataFrame
            df = pd.DataFrame(report_data)

            # تحويل الأعمدة
            df.columns = [
                'رقم الموقع',
                'اسم الموقع',
                'الحالة',
                'الخطوات المتمة',
                'الخطأ (إن وجد)',
                'المدة (ثانية)',
                'تاريخ المعالجة'
            ]

            # ترجمة الحالات
            df['الحالة'] = df['الحالة'].map({
                'success': '✓ نجح',
                'failed': '✗ فشل',
                'skipped': '⏭️ تخطي',
                'pending': '⏳ معلق'
            })

            # إنشاء Excel مع multiple sheets
            with pd.ExcelWriter(EXCEL_OUTPUT, engine='openpyxl') as writer:
                # ورقة التفاصيل
                df.to_excel(writer, sheet_name='التقرير المفصل', index=False)

                # ورقة الإحصائيات
                stats_df = pd.DataFrame({
                    'المؤشر': ['الإجمالي', 'نجح', 'فشل', 'تخطي', 'معلق'],
                    'العدد': [
                        stats['total'],
                        stats['success'],
                        stats['failed'],
                        stats['skipped'],
                        stats['pending']
                    ]
                })
                stats_df.to_excel(writer, sheet_name='الإحصائيات', index=False)

            logger.info(f"✓ تم إنشاء التقرير: {EXCEL_OUTPUT}")
            logger.info(f"  ├─ الإجمالي: {stats['total']}")
            logger.info(f"  ├─ نجح: {stats['success']}")
            logger.info(f"  ├─ فشل: {stats['failed']}")
            logger.info(f"  └─ تخطي: {stats['skipped']}")
        except Exception as e:
            logger.error(f"✗ خطأ في إنشاء التقرير: {e}")

# ==================== Main ====================
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="أتمتة تحديث مواقع منصة التراث العمراني")
    parser.add_argument(
        "--resume", metavar="SITE_ID", default=None,
        help="استئناف يدوي بدءاً من رقم موقع معين ضمن المواقع المعلقة (اختياري - "
             "غير ضروري عادة، لأن إعادة التشغيل تتخطى المواقع المُنجزة تلقائياً)"
    )
    parser.add_argument(
        "--limit", type=int, default=None,
        help="معالجة هذا العدد من المواقع فقط (للتجربة على عينة صغيرة قبل التشغيل الكامل)"
    )
    args = parser.parse_args()

    if args.resume:
        logger.info(f"🔄 استئناف من الموقع: {args.resume}")

    automation = HeritageAutomation()
    automation.run(resume_from=args.resume, limit=args.limit)
