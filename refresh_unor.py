import os
import pandas as pd
import time
import pyautogui
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
import json
import shutil
from functools import wraps

# ==================== KONFIGURASI ====================
EXCEL_FILE = 'Book1.xlsx'  # Ubah nama file Excel di sini
PROGRESS_FILE = 'progress_log.json'
LOG_FILE = 'process_log.txt'
FAILED_NIPS_FILE = 'nip_gagal.txt'

# Timeouts
DEFAULT_TIMEOUT = 10
SHORT_TIMEOUT = 5
LONG_TIMEOUT = 15

# Wait times
WAIT_AFTER_CLICK = 1
WAIT_AFTER_SEARCH = 2
WAIT_BETWEEN_RECORDS = 1.5

# XPath Locators - Konstanta untuk menghindari repetisi
class Locators:
    NIP_INPUT = (By.XPATH, "//input[@placeholder='Masukan NIP Baru']")
    SEARCH_BUTTON = (By.XPATH, "//button[@type='submit']//img[@src='/img/magnify-scan.png']/parent::button")
    ERROR_NOTIFICATION = (By.XPATH, "//div[contains(@class, 'swal2-container')]//h2[@id='swal2-title' and contains(text(), 'PNS Tidak Ditemukan')]")
    POSISI_JABATAN_TAB = (By.XPATH, "//label[@for='Posisi & Jabatan']")
    REFRESH_BUTTON = (By.XPATH, "//button[contains(., 'Refresh Data Unor')]")
    MODAL_PROSES_BUTTON = (By.XPATH, "//button[@class='swal2-confirm swal2-styled' and contains(., 'Proses')]")

# ==================== DECORATOR UNTUK RETRY ====================
def retry_on_exception(max_retries=2, delay=2):
    """Decorator untuk retry operasi yang gagal"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except (TimeoutException, NoSuchElementException) as e:
                    if attempt < max_retries - 1:
                        time.sleep(delay)
                        continue
                    raise
            return None
        return wrapper
    return decorator

# ==================== GLOBAL VARIABLES ====================
_log_buffer = []
_buffer_size = 10

def log_message(message, level="INFO", flush=False):
    """Fungsi untuk logging dengan timestamp dan buffering untuk efisiensi I/O"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}"
    print(log_entry)
    
    _log_buffer.append(log_entry)
    
    # Flush buffer jika sudah penuh atau diminta
    if len(_log_buffer) >= _buffer_size or flush:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write('\n'.join(_log_buffer) + '\n')
        _log_buffer.clear()

def flush_logs():
    """Flush semua log yang masih di buffer"""
    if _log_buffer:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write('\n'.join(_log_buffer) + '\n')
        _log_buffer.clear()

# Cache untuk progress
_progress_cache = None
_cache_dirty = False

def load_progress(force_reload=False):
    """Load progress dari file dengan caching untuk mengurangi disk I/O"""
    global _progress_cache
    
    if _progress_cache is not None and not force_reload:
        return _progress_cache
    
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                _progress_cache = json.load(f)
                return _progress_cache
        except:
            _progress_cache = {'source_file': '', 'total_rows': 0, 'processed_nips': [], 'last_index': -1}
            return _progress_cache
    
    _progress_cache = {'source_file': '', 'total_rows': 0, 'processed_nips': [], 'last_index': -1}
    return _progress_cache

def save_progress(nip, index, status="success", excel_file=None, total_rows=None, force_write=False):
    """Simpan progress dengan caching - hanya write ke disk jika perlu"""
    global _progress_cache, _cache_dirty
    
    if _progress_cache is None:
        _progress_cache = load_progress()
    
    _progress_cache['last_index'] = index
    if excel_file:
        _progress_cache['source_file'] = excel_file
    if total_rows is not None:
        _progress_cache['total_rows'] = total_rows
    
    # Simpan semua status (berhasil maupun gagal)
    _progress_cache['processed_nips'].append({
        'nip': nip,
        'index': index,
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'status': status
    })
    
    _cache_dirty = True
    
    # Write ke disk setiap 5 record atau jika dipaksa (untuk keamanan data)
    if len(_progress_cache['processed_nips']) % 5 == 0 or force_write:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(_progress_cache, f, indent=2, ensure_ascii=False)
        _cache_dirty = False
    
    # Jika status gagal, tambahkan ke file NIP gagal
    if status != "success":
        save_failed_nip(nip, status)

def flush_progress():
    """Force write progress ke disk"""
    global _cache_dirty
    if _cache_dirty and _progress_cache:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(_progress_cache, f, indent=2, ensure_ascii=False)
        _cache_dirty = False

def clear_progress_cache():
    """Clear progress cache untuk reset"""
    global _progress_cache, _cache_dirty
    _progress_cache = None
    _cache_dirty = False

def save_failed_nip(nip, status):
    """Simpan NIP yang gagal ke file txt"""
    with open(FAILED_NIPS_FILE, 'a', encoding='utf-8') as f:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"{nip} | Status: {status} | Waktu: {timestamp}\n")

def check_connection(driver):
    """Cek koneksi ke browser masih aktif"""
    try:
        driver.current_url
        return True
    except WebDriverException:
        return False

# ==================== HELPER FUNCTIONS ====================
def safe_click(driver, locator, timeout=DEFAULT_TIMEOUT, wait_after=WAIT_AFTER_CLICK):
    """Klik element dengan wait dan error handling"""
    try:
        wait = WebDriverWait(driver, timeout)
        element = wait.until(EC.element_to_be_clickable(locator))
        element.click()
        if wait_after > 0:
            time.sleep(wait_after)
        return True
    except TimeoutException:
        return False

def safe_input(driver, locator, text, clear_first=True, timeout=DEFAULT_TIMEOUT):
    """Input text ke field dengan wait dan error handling"""
    try:
        wait = WebDriverWait(driver, timeout)
        element = wait.until(EC.presence_of_element_located(locator))
        if clear_first:
            element.clear()
        element.send_keys(text)
        return True
    except TimeoutException:
        return False

def check_element_exists(driver, locator, timeout=SHORT_TIMEOUT):
    """Cek apakah element ada tanpa throw exception"""
    try:
        elements = driver.find_elements(*locator)
        return len(elements) > 0
    except:
        return False

def reset_page_state(driver, wait_time=3):
    """Reset halaman dengan refresh"""
    driver.refresh()
    time.sleep(wait_time)

# Membaca file Excel
log_message("="*60)
log_message("MEMULAI PROSES RPA REFRESH UNOR")
log_message("="*60)

try:
    df = pd.read_excel(EXCEL_FILE)
    log_message(f"File Excel '{EXCEL_FILE}' ditemukan, melanjutkan proses...")
except FileNotFoundError:
    log_message(f"Error: file Excel '{EXCEL_FILE}' tidak ditemukan.", "ERROR")
    exit()

# Load progress untuk resume jika ada
progress = load_progress()
last_index = progress.get('last_index', -1)
source_file = progress.get('source_file', '')
total_rows_saved = progress.get('total_rows', 0)
processed_nips = [item['nip'] for item in progress.get('processed_nips', []) if item['status'] == 'success']

# Deteksi perubahan file Excel atau data baru
file_changed = False
change_reason = ""

if source_file and source_file != EXCEL_FILE:
    file_changed = True
    change_reason = f"Nama file berbeda: {source_file} → {EXCEL_FILE}"
elif source_file == EXCEL_FILE and total_rows_saved > 0 and total_rows_saved != len(df):
    file_changed = True
    change_reason = f"Jumlah baris berbeda: {total_rows_saved} → {len(df)} (file sama tapi isinya berbeda)"
elif not source_file and last_index >= 0:
    # Progress ada tapi source_file kosong (dari versi script lama)
    file_changed = True
    change_reason = "Progress lama dari versi script sebelumnya"

# Jika file berubah, backup dan reset
if file_changed:
    log_message("="*60)
    log_message(f"DETEKSI PERUBAHAN DATA!", "WARNING")
    log_message(f"   Alasan: {change_reason}")
    if source_file:
        log_message(f"   File sebelumnya: {source_file} ({total_rows_saved} baris)")
    log_message(f"   File sekarang: {EXCEL_FILE} ({len(df)} baris)")
    log_message("="*60)
    
    # Backup file progress lama
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_progress = f"progress_log_backup_{timestamp}.json"
    backup_log = f"process_log_backup_{timestamp}.txt"
    backup_failed = f"nip_gagal_backup_{timestamp}.txt"
    
    if os.path.exists(PROGRESS_FILE):
        shutil.copy2(PROGRESS_FILE, backup_progress)
        log_message(f"Progress log lama di-backup ke: {backup_progress}")
    
    if os.path.exists(LOG_FILE):
        shutil.copy2(LOG_FILE, backup_log)
        log_message(f"Process log lama di-backup ke: {backup_log}")
    
    if os.path.exists(FAILED_NIPS_FILE):
        shutil.copy2(FAILED_NIPS_FILE, backup_failed)
        log_message(f"Failed NIPs log lama di-backup ke: {backup_failed}")
    
    # Reset progress untuk file baru
    log_message("Memulai proses baru untuk data yang berbeda...")
    last_index = -1
    processed_nips = []
    start_index = 0
    
    # Clear cache dan reload progress
    clear_progress_cache()
    
    # Hapus file log lama untuk mulai fresh
    if os.path.exists(PROGRESS_FILE):
        os.remove(PROGRESS_FILE)
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
    if os.path.exists(FAILED_NIPS_FILE):
        os.remove(FAILED_NIPS_FILE)
    
    log_message("="*60)
    log_message("MEMULAI PROSES BARU DARI AWAL")
    log_message("="*60)
elif last_index >= 0:
    log_message(f"Ditemukan progress sebelumnya. Terakhir diproses: baris ke-{last_index + 1}")
    log_message(f"Total NIP yang sudah berhasil diproses: {len(processed_nips)}")
    resume = input("Lanjutkan dari posisi terakhir? (y/n): ").lower()
    if resume == 'y':
        start_index = last_index + 1
        log_message(f"Melanjutkan dari baris ke-{start_index + 1}")
    else:
        start_index = 0
        log_message("Memulai dari awal")
else:
    start_index = 0
    log_message("Tidak ada progress sebelumnya, memulai dari awal")

# Koneksi ke Chrome yang sudah berjalan dengan remote debugging
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

try:
    driver = webdriver.Chrome(options=chrome_options)
    log_message("Berhasil terhubung ke Chrome")
except Exception as e:
    log_message(f"Error koneksi ke Chrome: {e}", "ERROR")
    log_message("Pastikan Chrome sudah berjalan dengan: chrome.exe --remote-debugging-port=9222 --user-data-dir=\"C:\\ChromeTemp\"", "ERROR")
    exit()

# Proses setiap NIP dari Excel
for index, row in df.iterrows():
    # Skip jika sudah diproses atau belum sampai start_index
    if index < start_index:
        continue
    
    nip_pppk = str(row['nip'])
    
    # Skip jika NIP sudah pernah berhasil diproses
    if nip_pppk in processed_nips:
        log_message(f"NIP {nip_pppk} sudah pernah diproses, skip...")
        continue
    
    log_message(f"\n{'='*60}")
    log_message(f"Memulai proses untuk NIP: {nip_pppk} (Baris ke-{index + 1} dari {len(df)})")
    log_message(f"{'='*60}")
    
    # Cek koneksi browser
    if not check_connection(driver):
        log_message("Koneksi ke browser terputus!", "ERROR", flush=True)
        log_message(f"Posisi terakhir: Baris ke-{index + 1}, NIP: {nip_pppk}", "ERROR", flush=True)
        flush_progress()
        flush_logs()
        exit()
    
    try:
        wait = WebDriverWait(driver, DEFAULT_TIMEOUT)
        
        # Step 1: Isi NIP
        log_message("1. Mengisi NIP BARU...")
        if not safe_input(driver, Locators.NIP_INPUT, nip_pppk):
            raise TimeoutException("Input NIP gagal")
        log_message(f"   ✓ NIP {nip_pppk} berhasil diisi")
        
        # Step 2: Klik tombol search
        log_message("2. Menekan tombol search...")
        if not safe_click(driver, Locators.SEARCH_BUTTON, wait_after=WAIT_AFTER_SEARCH):
            raise TimeoutException("Tombol search tidak ditemukan")
        log_message("   ✓ Tombol search diklik")
        
        # Step 2.5: Cek notifikasi error
        log_message("2.5. Memeriksa hasil pencarian...")
        if check_element_exists(driver, Locators.ERROR_NOTIFICATION):
            log_message(f"❌ PNS dengan NIP {nip_pppk} tidak ditemukan di database!", "WARNING")
            save_progress(nip_pppk, index, "not_found", EXCEL_FILE, len(df))
            continue
        log_message("   ✓ PNS ditemukan, melanjutkan proses...")
        
        # Step 3: Buka tab Posisi & Jabatan
        log_message("3. Membuka tab Posisi & Jabatan...")
        if not safe_click(driver, Locators.POSISI_JABATAN_TAB, wait_after=1.5):
            log_message("❌ Error: Tidak dapat membuka tab Posisi & Jabatan", "ERROR")
            save_progress(nip_pppk, index, "network_error", EXCEL_FILE, len(df))
            reset_page_state(driver)
            continue
        log_message("   ✓ Tab Posisi & Jabatan dibuka")
        
        # Step 4: Scroll untuk menemukan tombol
        log_message("4. Mencari tombol Refresh Data Unor...")
        driver.execute_script("window.scrollBy(0, 500);")
        time.sleep(0.5)
        
        # Step 5: Klik tombol Refresh Data Unor
        log_message("5. Menekan tombol Refresh Data Unor...")
        if not safe_click(driver, Locators.REFRESH_BUTTON, wait_after=1):
            log_message("❌ Error: Tombol Refresh Data Unor tidak ditemukan", "ERROR")
            save_progress(nip_pppk, index, "element_not_found", EXCEL_FILE, len(df))
            reset_page_state(driver)
            continue
        log_message("   ✓ Tombol Refresh Data Unor diklik")
        
        # Step 6: Konfirmasi di modal
        log_message("6. Mengkonfirmasi proses di modal...")
        if not safe_click(driver, Locators.MODAL_PROSES_BUTTON, wait_after=2):
            log_message("❌ Error: Modal konfirmasi tidak muncul", "ERROR")
            save_progress(nip_pppk, index, "modal_timeout", EXCEL_FILE, len(df))
            reset_page_state(driver)
            continue
        log_message("   ✓ Tombol Proses diklik")
        
        # Tunggu proses selesai
        log_message("7. Menunggu proses refresh selesai...")
        time.sleep(1)
        
        # Simpan progress setelah berhasil
        save_progress(nip_pppk, index, "success", EXCEL_FILE, len(df))
        log_message(f"✅ Proses untuk NIP {nip_pppk} selesai!")
        
    except TimeoutException as e:
        log_message(f"❌ Error: Timeout saat memproses NIP {nip_pppk}", "ERROR")
        log_message(f"   Detail: {str(e)}", "ERROR")
        save_progress(nip_pppk, index, "timeout_error", EXCEL_FILE, len(df))
        reset_page_state(driver)
        continue
    except WebDriverException as e:
        log_message(f"❌ Error koneksi browser: {e}", "ERROR", flush=True)
        save_progress(nip_pppk, index - 1, "connection_error", EXCEL_FILE, len(df), force_write=True)
        flush_progress()
        flush_logs()
        exit()
    except Exception as e:
        log_message(f"❌ Error saat memproses NIP {nip_pppk}: {e}", "ERROR")
        save_progress(nip_pppk, index, "general_error", EXCEL_FILE, len(df))
        continue
    
    # Jeda sebelum proses NIP berikutnya
    time.sleep(WAIT_BETWEEN_RECORDS)

# Flush buffer sebelum summary
flush_progress()
flush_logs()

# Summary akhir
progress = load_progress(force_reload=True)
log_message(f"\n{'='*60}")
log_message(f"✅ SEMUA PROSES SELESAI!")
log_message(f"   Total data di Excel: {len(df)} pegawai")

# Hitung status
success_count = len([p for p in progress['processed_nips'] if p['status'] == 'success'])
not_found_count = len([p for p in progress['processed_nips'] if p['status'] == 'not_found'])
error_count = len([p for p in progress['processed_nips'] if p['status'] not in ['success', 'not_found']])

log_message(f"   Total berhasil diproses: {success_count} pegawai")
log_message(f"   PNS tidak ditemukan: {not_found_count} pegawai")
log_message(f"   Gagal karena error: {error_count} pegawai")
log_message(f"   Log detail tersimpan di: {LOG_FILE}")
log_message(f"   Progress tersimpan di: {PROGRESS_FILE}")

# Buat file summary NIP gagal
if not_found_count + error_count > 0:
    failed_summary_file = 'daftar_nip_gagal.txt'
    with open(failed_summary_file, 'w', encoding='utf-8') as f:
        f.write(f"DAFTAR NIP YANG GAGAL DIPROSES\n")
        f.write(f"Tanggal: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"{'='*60}\n\n")
        
        # NIP tidak ditemukan
        not_found_nips = [p for p in progress['processed_nips'] if p['status'] == 'not_found']
        if not_found_nips:
            f.write(f"PNS TIDAK DITEMUKAN ({len(not_found_nips)} pegawai):\n")
            f.write("-" * 60 + "\n")
            for item in not_found_nips:
                f.write(f"{item['nip']}\n")
            f.write("\n")
        
        # NIP gagal karena error
        error_nips = [p for p in progress['processed_nips'] if p['status'] not in ['success', 'not_found']]
        if error_nips:
            f.write(f"GAGAL KARENA ERROR ({len(error_nips)} pegawai):\n")
            f.write("-" * 60 + "\n")
            for item in error_nips:
                f.write(f"{item['nip']} | Error: {item['status']} | Waktu: {item['timestamp']}\n")
    
    log_message(f"   Daftar NIP gagal tersimpan di: {failed_summary_file}")

log_message(f"{'='*60}")

# Tutup browser (optional, bisa di-comment jika ingin tetap buka)
# driver.quit()