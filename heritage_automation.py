#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
منصة اكتشاف وتسجيل أصول التراث العمراني - أتمتة التعديلات
Automation for Heritage Platform - Mass Update Script
"""

import os
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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# ==================== إعدادات ====================
BASE_URL = "https://heritage2.itqan-consultant.com/Web/"
LOGIN_URL = "https://heritage2.itqan-consultant.com/Web/App/Home/Login"
USERNAME = os.environ.get("HERITAGE_USERNAME")
PASSWORD = os.environ.get("HERITAGE_PASSWORD")
TARGET_REGION = "المنطقة الجنوبية"
REGION_FIELD_NUMBER = 4

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

        # إعدادات أخرى
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('start-maximized')
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

    def close(self):
        """إغلاق المتصفح"""
        if self.driver:
            self.driver.quit()
            logger.info("✓ تم إغلاق المتصفح")

    def login(self) -> bool:
        """تسجيل الدخول"""
        try:
            logger.info("🔐 جاري تسجيل الدخول...")
            self.driver.get(LOGIN_URL)

            # إدخال اسم المستخدم (name="ctl112$ct103")
            username_field = self.wait.until(
                EC.presence_of_element_located((By.NAME, "ctl112$ct103"))
            )
            username_field.clear()
            username_field.send_keys(USERNAME)
            time.sleep(0.5)

            # إدخال كلمة المرور (name="ctl112$ct107")
            password_field = self.driver.find_element(By.NAME, "ctl112$ct107")
            password_field.clear()
            password_field.send_keys(PASSWORD)
            time.sleep(0.5)

            # النقر على زر "تسجيل الدخول"
            login_button = self.driver.find_element(
                By.XPATH, "//*[self::button or self::input][contains(., 'تسجيل الدخول') or @value='تسجيل الدخول']"
            )
            login_button.click()

            # انتظر تحميل الصفحة بعد تسجيل الدخول (تأكد أننا خرجنا من صفحة /Login)
            self.wait.until(lambda d: "/Login" not in d.current_url)
            time.sleep(2)

            logger.info(f"✓ تم تسجيل الدخول بنجاح - الرابط الحالي: {self.driver.current_url}")
            return True
        except TimeoutException:
            logger.error("✗ انتهت مهلة الانتظار - فشل تسجيل الدخول (تحقق من صحة اليوزر/الباسورد)")
            return False
        except Exception as e:
            logger.error(f"✗ خطأ في تسجيل الدخول: {e}")
            return False

    def navigate_to_guide_list(self) -> bool:
        """الانتقال إلى دليل المواقع"""
        try:
            logger.info("📍 جاري الانتقال إلى دليل المواقع...")

            # البحث عن الرابط (دليل المواقع)
            guide_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'دليل المواقع')]"))
            )
            guide_link.click()
            time.sleep(2)

            logger.info("✓ تم الانتقال إلى دليل المواقع")
            return True
        except Exception as e:
            logger.error(f"✗ خطأ في الانتقال إلى دليل المواقع: {e}")
            return False

    def apply_filters(self) -> bool:
        """تطبيق الفلاتر (المنطقة والخطوة)"""
        try:
            logger.info("🔍 جاري تطبيق الفلاتر...")

            # اختيار المنطقة: الباحة
            region_dropdown = self.wait.until(
                EC.presence_of_element_located((By.NAME, "ddlRegion"))
            )
            region_dropdown.click()
            time.sleep(1)

            # البحث عن خيار الباحة
            region_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), 'الباحة')]"))
            )
            region_option.click()
            time.sleep(1)

            # اختيار الخطوة: التسجيل المبدأي
            step_dropdown = self.wait.until(
                EC.presence_of_element_located((By.NAME, "ddlStep"))
            )
            step_dropdown.click()
            time.sleep(1)

            step_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), 'التسجيل المبدأي')]"))
            )
            step_option.click()
            time.sleep(2)

            # انتظر تحميل البيانات
            logger.info("⏳ جاري تحميل البيانات...")
            time.sleep(3)

            logger.info("✓ تم تطبيق الفلاتر بنجاح")
            return True
        except Exception as e:
            logger.error(f"✗ خطأ في تطبيق الفلاتر: {e}")
            return False

    def extract_all_site_links(self) -> List[Tuple[str, str, str]]:
        """استخراج جميع روابط المواقع (مع pagination)"""
        sites = []
        try:
            logger.info("📥 جاري استخراج روابط المواقع...")

            # الحصول على عدد الصفحات
            pagination_info = self.driver.find_element(By.CSS_SELECTOR, ".pagination-info")
            pagination_text = pagination_info.text
            logger.info(f"📊 معلومات الترقيم: {pagination_text}")

            # استخراج المواقع من الجدول
            rows = self.driver.find_elements(By.XPATH, "//table[@class='data-table']//tbody//tr")
            logger.info(f"📌 عدد الصفوف في الصفحة الحالية: {len(rows)}")

            for row in rows:
                try:
                    # الحصول على رقم الموقع (ID)
                    site_id = row.find_element(By.XPATH, ".//td[1]").text.strip()

                    # الحصول على اسم الموقع
                    site_link = row.find_element(By.XPATH, ".//a")
                    site_name = site_link.text.strip()
                    site_url = site_link.get_attribute('href')

                    if site_id and site_name and site_url:
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
        """فتح صفحة الموقع"""
        try:
            self.driver.get(site_url)
            time.sleep(2)

            # التحقق من أن الصفحة تحملت
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "site-details")))
            return True
        except Exception as e:
            logger.warning(f"⚠️ خطأ في فتح الموقع: {e}")
            return False

    def navigate_to_architect_data(self) -> bool:
        """الانتقال إلى صفحة بيانات المعماري"""
        try:
            logger.info("  ↳ جاري الانتقال إلى بيانات المعماري...")

            # البحث عن التاب "بيانات المعماري"
            architect_tab = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'بيانات المعماري')]"))
            )
            architect_tab.click()
            time.sleep(1)

            logger.info("  ✓ تم الانتقال إلى بيانات المعماري")
            return True
        except Exception as e:
            logger.warning(f"  ✗ خطأ في الانتقال: {e}")
            return False

    def update_region(self) -> Tuple[bool, str]:
        """تحديث المنطقة إلى "المنطقة الجنوبية" """
        try:
            logger.info("  ↳ جاري تحديث المنطقة...")

            # الوصول إلى حقل المنطقة رقم 4
            region_field = self.wait.until(
                EC.presence_of_element_located((By.NAME, f"Region{REGION_FIELD_NUMBER}"))
            )

            # الحصول على القيمة الحالية
            current_value = region_field.get_attribute('value')

            # إذا كانت مختارة بالفعل
            if TARGET_REGION in current_value or "جنوبية" in current_value:
                logger.info(f"  ⏭️  المنطقة مختارة بالفعل: {current_value}")
                return True, "skipped"

            # النقر على dropdown
            region_field.click()
            time.sleep(0.5)

            # اختيار المنطقة الجنوبية
            option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//option[contains(text(), 'المنطقة الجنوبية')]"))
            )
            option.click()
            time.sleep(1)

            # التحقق من التحديث
            updated_value = region_field.get_attribute('value')
            if TARGET_REGION in updated_value or "جنوبية" in updated_value:
                logger.info(f"  ✓ تم تحديث المنطقة بنجاح")
                return True, "updated"
            else:
                logger.warning(f"  ⚠️ فشل التحديث - القيمة الحالية: {updated_value}")
                return False, "failed"
        except Exception as e:
            logger.warning(f"  ✗ خطأ في تحديث المنطقة: {e}")
            return False, f"error: {str(e)}"

    def save_form(self) -> bool:
        """حفظ الاستمارة"""
        try:
            logger.info("  ↳ جاري حفظ الاستمارة...")

            # البحث عن زر الحفظ
            save_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'حفظ')]"))
            )
            save_button.click()

            # انتظر الرسالة (نجاح أو خطأ)
            time.sleep(2)

            # التحقق من رسالة النجاح
            try:
                success_msg = self.driver.find_element(By.CLASS_NAME, "alert-success")
                logger.info(f"  ✓ تم الحفظ بنجاح: {success_msg.text}")
                return True
            except:
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

    def run(self, resume_from: Optional[str] = None):
        """تشغيل الأتمتة"""
        start_time = datetime.now()
        logger.info("=" * 50)
        logger.info("🚀 بدء أتمتة منصة التراث العمراني")
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
            all_sites = self._extract_all_paginated_sites()

            if not all_sites:
                logger.error("❌ لم يتم استخراج أي مواقع")
                return False

            logger.info(f"📊 إجمالي المواقع المستخرجة: {len(all_sites)}")

            # إضافة المواقع للقاعدة
            for site_id, site_name, site_url in all_sites:
                self.db.insert_site(site_id, site_name, site_url)

            # معالجة كل موقع
            self._process_all_sites(resume_from)

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

    def _extract_all_paginated_sites(self) -> List[Tuple[str, str, str]]:
        """استخراج المواقع من جميع الصفحات"""
        all_sites = []
        page = 1

        try:
            while True:
                logger.info(f"📄 جاري استخراج الصفحة {page}...")

                sites_in_page = self.driver_manager.extract_all_site_links()
                if not sites_in_page:
                    break

                all_sites.extend(sites_in_page)

                # محاولة الانتقال للصفحة التالية
                try:
                    next_button = self.driver_manager.driver.find_element(
                        By.XPATH, "//a[@class='next-page' or contains(text(), 'التالي')]"
                    )
                    next_button.click()
                    time.sleep(2)
                    page += 1
                except:
                    logger.info(f"✓ انتهت جميع الصفحات (إجمالي: {page} صفحات)")
                    break
        except Exception as e:
            logger.warning(f"⚠️ خطأ في استخراج الصفحات: {e}")

        return all_sites

    def _process_all_sites(self, resume_from: Optional[str] = None):
        """معالجة جميع المواقع"""
        pending_sites = self.db.get_pending_sites()

        start_index = 0
        if resume_from:
            logger.info(f"🔄 استئناف من الموقع: {resume_from}")
            for i, (site_id, _, _) in enumerate(pending_sites):
                if site_id == resume_from:
                    start_index = i
                    break

        for index, (site_id, site_name, site_url) in enumerate(pending_sites[start_index:], start_index + 1):
            logger.info(f"[{index}/{len(pending_sites)}] 🔄 معالجة الموقع: {site_id} - {site_name}")

            process_start = time.time()
            steps = []

            try:
                # فتح الموقع
                steps.append("فتح الموقع")
                if not self.driver_manager.open_site(site_url):
                    raise Exception("فشل فتح الموقع")

                # الانتقال إلى بيانات المعماري
                steps.append("الانتقال إلى بيانات المعماري")
                if not self.driver_manager.navigate_to_architect_data():
                    raise Exception("فشل الانتقال")

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
    # التحقق من المعاملات
    import sys

    resume_site_id = None
    if len(sys.argv) > 1:
        resume_site_id = sys.argv[1]
        logger.info(f"🔄 استئناف من الموقع: {resume_site_id}")

    # تشغيل الأتمتة
    automation = HeritageAutomation()
    automation.run(resume_from=resume_site_id)
