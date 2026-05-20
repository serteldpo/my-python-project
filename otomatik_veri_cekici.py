# -*- coding: utf-8 -*-
"""
Ekosepeti Otomatik Veri Çekici
===============================
DKS (ekosepeti.net) web sitesinden tüm ürün kategorilerinin
Excel dosyalarını otomatik indirip, Google Drive'a yükler.

Kullanım:
    python otomatik_veri_cekici.py                  # Normal çalıştır
    python otomatik_veri_cekici.py --sadece-indir   # Sadece indir, Drive'a yükleme
    python otomatik_veri_cekici.py --sadece-yukle   # Sadece Drive'a yükle (bugünün dosyaları)
"""

import os
import sys
import json
import time
import re
import logging
import argparse
from datetime import datetime

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)
from webdriver_manager.chrome import ChromeDriverManager

# Excel yazma
import openpyxl

# ─── Yapılandırma ────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

# Loglama
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            os.path.join(SCRIPT_DIR, "otomatik_veri_cekici.log"),
            encoding="utf-8"
        ),
    ],
)
log = logging.getLogger("ekosepeti")


def load_config():
    """config.json dosyasını yükle."""
    if not os.path.exists(CONFIG_PATH):
        log.error(f"config.json bulunamadı: {CONFIG_PATH}")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ─── Selenium: Web Sitesinden İndirme ────────────────────────────────────────
# Sayfa yapısı (keşifle doğrulanmış):
#   - Login: #userno, #username, #password, #btnlogin
#   - Ana iframe: #mainiFrame
#   - Menü: #menu1 ("Market Ürün Listesi" tıklanır)
#   - Pencere: #jspdpo_marurunfiyat (adminWindow)
#   - Grup1 select: #dpo_marurunfiyat_grup1 (Select2 widget)
#     Option values: 10000000, 11000000, 12000000, ... (kategori_no * 1000000)
#   - Butonlar:
#     #dpo_marurunfiyat_btnquery   → "Getir"
#     #dpo_marurunfiyat_btngrupexcel → "Grup Excel Çıktısı AL"
#     #dpo_marurunfiyat_btncommit   → "Excel Olarak Çıktı AL"
#   - Excel export: client-side (xlsx.full.min.js / SheetJS)

class EkosepetiScraper:
    """DKS ekosepeti.net sitesinden API ile Excel verilerini çeker."""

    GRUP1_SELECT_ID = "dpo_marurunfiyat_grup1"
    MENU_ID = "menu1"

    def __init__(self, config):
        self.config = config
        self.web = config["website"]
        self.paths = config["paths"]
        self.categories = config["categories"]
        self.page_wait = config.get("page_load_wait_seconds", 15)
        self.headless = config.get("headless", False)
        self.driver = None

        # Bugünün tarihli klasör: dd.MM.yyyy
        self.today_str = datetime.now().strftime("%d.%m.%Y")
        self.target_folder = os.path.join(self.paths["base_folder"], self.today_str)

    def setup_driver(self):
        """Chrome WebDriver'ı yapılandır."""
        options = ChromeOptions()
        if self.headless:
            options.add_argument("--headless=new")

        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--lang=tr")

        try:
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.set_script_timeout(120)
            log.info("Chrome WebDriver başarıyla başlatıldı.")
        except WebDriverException as e:
            log.error(f"Chrome başlatılamadı: {e}")
            raise

    def login(self):
        """ekosepeti.net'e giriş yap."""
        log.info(f"Giriş yapılıyor: {self.web['url']}")
        self.driver.get(self.web["url"])

        wait = WebDriverWait(self.driver, self.page_wait)
        wait.until(EC.presence_of_element_located((By.ID, "userno")))

        self.driver.find_element(By.ID, "userno").send_keys(self.web["firma_no"])
        self.driver.find_element(By.ID, "username").send_keys(self.web["username"])
        self.driver.find_element(By.ID, "password").send_keys(self.web["password"])
        self.driver.find_element(By.ID, "btnlogin").click()
        log.info("Giriş bilgileri gönderildi...")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "mainiFrame")))
        log.info("Ana iframe'e geçildi.")
        time.sleep(3)

    def open_market_page(self):
        """Menüden Market Ürün Listesi penceresini aç."""
        wait = WebDriverWait(self.driver, self.page_wait)
        menu = wait.until(EC.element_to_be_clickable((By.ID, self.MENU_ID)))
        menu.click()
        log.info("Market Ürün Listesi menüsüne tıklandı.")
        wait.until(EC.presence_of_element_located((By.ID, self.GRUP1_SELECT_ID)))
        time.sleep(3)
        log.info("Market Ürün Listesi penceresi açıldı.")

    def _get_category_names(self):
        """Dropdown'dan kategori isimlerini al. {category_num: "Ad"} sözlüğü döner."""
        options_list = self.driver.execute_script("""
            var sel = document.getElementById('""" + self.GRUP1_SELECT_ID + """');
            var result = [];
            for (var i = 0; i < sel.options.length; i++) {
                var v = sel.options[i].value;
                var t = sel.options[i].text;
                if (v && v !== '0') result.push({value: v, text: t});
            }
            return result;
        """)
        name_map = {}
        for opt in options_list:
            try:
                cat_num = int(opt["value"]) // 1000000
                # Dropdown text örnek: "10 - Temel ve Kuru Gıdalar (1618)"
                # Parantez içi sayıyı kaldır, text'i olduğu gibi kullan
                clean_text = re.sub(r'\s*\(\d+\)\s*$', '', opt["text"].strip())
                name_map[cat_num] = clean_text
            except (ValueError, KeyError):
                pass
        return name_map

    def download_category(self, category_num, filename):
        """Tek bir kategori için API çağrısı yapıp Excel olarak kaydet."""
        option_value = str(category_num * 1000000)
        log.info(f"  API çağrısı: kategori {category_num} (value={option_value})...")

        response = self.driver.execute_async_script("""
            var grup1 = arguments[0];
            var callback = arguments[arguments.length - 1];
            var data = {
                page: "dpo_marurunfiyat",
                func: "URUNGRUPEXCEL",
                type: "cursor",
                hlse: JSON.stringify({grup1: grup1, grup2: "0", grup3: "0"})
            };
            window.util.runAjax(data, "/getdata", function(response) {
                callback(response);
            });
        """, option_value)

        if not response or "names" not in response or "data" not in response:
            log.error(f"  Kategori {category_num}: API boş yanıt döndü")
            return False

        names = response["names"]
        data = response["data"]
        row_count = len(data)

        if row_count == 0:
            log.warning(f"  Kategori {category_num}: 0 satır (boş dosya oluşturuluyor)")

        # openpyxl ile Excel kaydet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "sheet1"
        ws.append(names)
        for row in data:
            ws.append(row)

        filepath = os.path.join(self.target_folder, filename)
        wb.save(filepath)
        file_size = os.path.getsize(filepath)
        log.info(f"  Kaydedildi: {filename} ({row_count} satır, {file_size:,} bytes)")
        return True

    def download_all(self):
        """Tüm kategorileri indir."""
        os.makedirs(self.target_folder, exist_ok=True)
        log.info(f"Hedef klasör: {self.target_folder}")

        # Kategori isimlerini dropdown'dan al
        name_map = self._get_category_names()
        log.info(f"Toplam {len(name_map)} kategori bulundu")

        success_count = 0
        fail_count = 0

        for cat in self.categories:
            filename = name_map.get(cat, f"{cat} - Kategori {cat}") + ".xlsx"
            log.info(f"Kategori {cat}: {filename}")
            try:
                if self.download_category(cat, filename):
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                log.error(f"Kategori {cat} hatası: {e}")
                fail_count += 1
            time.sleep(1)

        log.info(f"İndirme tamamlandı: {success_count} başarılı, {fail_count} başarısız")
        return success_count > 0

    def close(self):
        """Tarayıcıyı kapat."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            log.info("Tarayıcı kapatıldı.")

    def run(self):
        """Tam indirme sürecini çalıştır."""
        try:
            self.setup_driver()
            self.login()
            self.open_market_page()
            return self.download_all()
        except Exception as e:
            log.error(f"İndirme hatası: {e}", exc_info=True)
            return False
        finally:
            self.close()


# ─── Google Drive Yükleme ────────────────────────────────────────────────────

class GoogleDriveUploader:
    """Google Drive'a OAuth2 ile dosya yükler."""

    SCOPES = ["https://www.googleapis.com/auth/drive"]

    def __init__(self, config):
        self.config = config["google_drive"]
        self.folder_id = self.config["folder_id"]
        self.creds_file = os.path.join(SCRIPT_DIR, self.config["credentials_file"])
        self.token_file = os.path.join(SCRIPT_DIR, self.config["token_file"])
        self.service = None

    def authenticate(self):
        """OAuth2 ile Google Drive API bağlantısı kur."""
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build

        creds = None
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.creds_file):
                    log.error(f"Credentials dosyası bulunamadı: {self.creds_file}")
                    return False
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.creds_file, self.SCOPES
                )
                creds = flow.run_local_server(port=0)

            with open(self.token_file, "w") as f:
                f.write(creds.to_json())

        self.service = build("drive", "v3", credentials=creds)
        log.info("Google Drive API bağlantısı kuruldu.")
        return True

    def _find_or_create_folder(self, folder_name, parent_id):
        """Drive'da klasör bul veya oluştur."""
        query = (
            f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' "
            f"and '{parent_id}' in parents and trashed=false"
        )
        results = self.service.files().list(
            q=query, spaces="drive", fields="files(id, name)"
        ).execute()

        files = results.get("files", [])
        if files:
            return files[0]["id"]

        # Klasör yoksa oluştur
        file_metadata = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = self.service.files().create(
            body=file_metadata, fields="id"
        ).execute()
        log.info(f"Drive klasörü oluşturuldu: {folder_name}")
        return folder["id"]

    def upload_folder(self, local_folder_path):
        """Yerel klasördeki tüm Excel dosyalarını Drive'a yükle."""
        if not self.service:
            if not self.authenticate():
                return False

        folder_name = os.path.basename(local_folder_path)
        drive_folder_id = self._find_or_create_folder(folder_name, self.folder_id)

        files = [f for f in os.listdir(local_folder_path)
                 if f.endswith((".xlsx", ".xls"))]

        if not files:
            log.warning(f"Yüklenecek dosya bulunamadı: {local_folder_path}")
            return False

        from googleapiclient.http import MediaFileUpload

        uploaded = 0
        for filename in files:
            filepath = os.path.join(local_folder_path, filename)
            try:
                # Aynı isimde dosya varsa güncelle
                query = (
                    f"name='{filename}' and '{drive_folder_id}' in parents "
                    f"and trashed=false"
                )
                existing = self.service.files().list(
                    q=query, spaces="drive", fields="files(id)"
                ).execute().get("files", [])

                media = MediaFileUpload(
                    filepath,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    resumable=True,
                )

                if existing:
                    # Güncelle
                    self.service.files().update(
                        fileId=existing[0]["id"],
                        media_body=media,
                    ).execute()
                    log.info(f"  Güncellendi: {filename}")
                else:
                    # Yeni yükle
                    file_metadata = {
                        "name": filename,
                        "parents": [drive_folder_id],
                    }
                    self.service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields="id",
                    ).execute()
                    log.info(f"  Yüklendi: {filename}")

                uploaded += 1
            except Exception as e:
                log.error(f"  {filename} yüklenemedi: {e}")

        log.info(f"Google Drive'a yükleme: {uploaded}/{len(files)} başarılı")
        return uploaded > 0


# ─── Ana Çalıştırma ──────────────────────────────────────────────────────────

def main():
    os.chdir(SCRIPT_DIR)

    parser = argparse.ArgumentParser(description="Ekosepeti Otomatik Veri Çekici")
    parser.add_argument("--sadece-indir", action="store_true",
                        help="Sadece web sitesinden indir, Drive'a yükleme")
    parser.add_argument("--sadece-yukle", action="store_true",
                        help="Sadece bugünün dosyalarını Drive'a yükle")
    parser.add_argument("--tarih", type=str, default=None,
                        help="Belirli bir tarih klasörünü yükle (dd.mm.yyyy)")
    args = parser.parse_args()

    config = load_config()
    today_str = args.tarih or datetime.now().strftime("%d.%m.%Y")
    target_folder = os.path.join(config["paths"]["base_folder"], today_str)

    log.info("=" * 60)
    log.info(f"Ekosepeti Otomatik Veri Çekici - {today_str}")
    log.info("=" * 60)

    # 1. Web sitesinden indir
    if not args.sadece_yukle:
        scraper = EkosepetiScraper(config)
        download_ok = scraper.run()
        if not download_ok:
            log.error("İndirme işlemi başarısız oldu.")
            if not args.sadece_indir:
                sys.exit(1)

    # 2. Google Drive'a yükle
    if not args.sadece_indir:
        if config["google_drive"].get("enabled", True):
            if os.path.exists(target_folder):
                uploader = GoogleDriveUploader(config)
                uploader.upload_folder(target_folder)
            else:
                log.warning(f"Yüklenecek klasör bulunamadı: {target_folder}")
        else:
            log.info("Google Drive yükleme devre dışı.")

    log.info("İşlem tamamlandı.")


if __name__ == "__main__":
    main()
