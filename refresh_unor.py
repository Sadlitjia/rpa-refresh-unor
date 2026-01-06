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
from selenium.common.exceptions import TimeoutException, WebDriverException
import json

# File untuk menyimpan progress
PROGRESS_FILE = 'progress_log.json'
LOG_FILE = 'process_log.txt'

def log_message(message, level="INFO"):
    """Fungsi untuk logging dengan timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}"
    print(log_entry)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_entry + "\n")

def load_progress():
    """Load progress dari file jika ada"""
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {'processed_nips': [], 'last_index': -1}
    return {'processed_nips': [], 'last_index': -1}

def save_progress(nip, index, status="success"):
    """Simpan progress ke file"""
    progress = load_progress()
    progress['last_index'] = index
    if status == "success":
        progress['processed_nips'].append({
            'nip': nip,
            'index': index,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'status': status
        })
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(progress, f, indent=2, ensure_ascii=False)

def check_connection(driver):
    """Cek koneksi ke browser masih aktif"""
    try:
        driver.current_url
        return True
    except WebDriverException:
        return False

# Membaca file Excel
log_message("="*60)
log_message("MEMULAI PROSES RPA REFRESH UNOR")
log_message("="*60)

try:
    df = pd.read_excel('test_list_pppk.xlsx')
    log_message("File Excel ditemukan, melanjutkan proses...")
except FileNotFoundError:
    log_message("Error: file Excel tidak ditemukan.", "ERROR")
    exit()

# Load progress untuk resume jika ada
progress = load_progress()
last_index = progress['last_index']
processed_nips = [item['nip'] for item in progress['processed_nips']]

if last_index >= 0:
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
        log_message(f"NIP {nip_pppk} sudah pernah diproses, skip...", "INFO")
        continue
    
    log_message(f"\n{'='*60}")
    log_message(f"Memulai proses untuk NIP: {nip_pppk} (Baris ke-{index + 1} dari {len(df)})")
    log_message(f"{'='*60}")
    
    # Cek koneksi browser
    if not check_connection(driver):
        log_message("Koneksi ke browser terputus!", "ERROR")
        log_message(f"Posisi terakhir: Baris ke-{index + 1}, NIP: {nip_pppk}", "ERROR")
        log_message("Silakan restart script untuk melanjutkan.", "ERROR")
        exit()
    
    try:
        # Step 1: Isi NIP BARU di input field
        log_message("1. Mengisi NIP BARU...")
        wait = WebDriverWait(driver, 10)
        
        # Cari input field dengan placeholder "Masukan NIP Baru"
        nip_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@placeholder='Masukan NIP Baru']")
        ))
        nip_input.clear()
        nip_input.send_keys(nip_pppk)
        log_message(f"   ✓ NIP {nip_pppk} berhasil diisi")
        time.sleep(1)
        
        # Step 2: Klik tombol search (button dengan img magnify-scan.png)
        log_message("2. Menekan tombol search...")
        search_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@type='submit']//img[@src='/img/magnify-scan.png']/parent::button")
        ))
        search_button.click()
        log_message("   ✓ Tombol search diklik")
        time.sleep(3)  # Tunggu hasil pencarian
        
        # Step 3: Klik checkbox/label "Posisi & Jabatan"
        log_message("3. Membuka tab Posisi & Jabatan...")
        posisi_jabatan = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//label[@for='Posisi & Jabatan']")
        ))
        posisi_jabatan.click()
        log_message("   ✓ Tab Posisi & Jabatan dibuka")
        time.sleep(2)
        
        # Step 4: Scroll ke bawah untuk menemukan tombol refresh
        log_message("4. Mencari tombol Refresh Data Unor...")
        driver.execute_script("window.scrollBy(0, 500);")
        time.sleep(1)
        
        # Step 5: Klik tombol "Refresh Data Unor"
        log_message("5. Menekan tombol Refresh Data Unor...")
        refresh_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Refresh Data Unor')]")
        ))
        refresh_button.click()
        log_message("   ✓ Tombol Refresh Data Unor diklik")
        time.sleep(2)
        
        # Step 6: Klik tombol "Proses" di modal
        log_message("6. Mengkonfirmasi proses di modal...")
        proses_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@class='swal2-confirm swal2-styled' and contains(., 'Proses')]")
        ))
        proses_button.click()
        log_message("   ✓ Tombol Proses diklik")
        
        # Tunggu 3 detik untuk proses refresh selesai
        log_message("7. Menunggu proses refresh selesai...")
        time.sleep(3)
        
        # Simpan progress setelah berhasil
        save_progress(nip_pppk, index, "success")
        log_message(f"✅ Proses untuk NIP {nip_pppk} selesai dan tersimpan!")
        
    except TimeoutException as e:
        log_message(f"❌ Error: Timeout saat memproses NIP {nip_pppk}", "ERROR")
        log_message(f"   Detail: Element tidak ditemukan dalam waktu yang ditentukan", "ERROR")
        save_progress(nip_pppk, index, "timeout_error")
        continue
    except WebDriverException as e:
        log_message(f"❌ Error koneksi browser saat memproses NIP {nip_pppk}: {e}", "ERROR")
        log_message(f"Posisi terakhir berhasil: Baris ke-{index}, NIP sebelumnya", "ERROR")
        log_message("Kemungkinan terjadi masalah jaringan atau browser crash", "ERROR")
        log_message("Restart script untuk melanjutkan dari posisi ini", "ERROR")
        save_progress(nip_pppk, index - 1, "connection_error")
        exit()
    except Exception as e:
        log_message(f"❌ Error saat memproses NIP {nip_pppk}: {e}", "ERROR")
        save_progress(nip_pppk, index, "general_error")
        continue
    
    # Jeda sebelum proses NIP berikutnya
    time.sleep(2)

# Summary akhir
progress = load_progress()
log_message(f"\n{'='*60}")
log_message(f"✅ SEMUA PROSES SELESAI!")
log_message(f"   Total data di Excel: {len(df)} pegawai")
log_message(f"   Total berhasil diproses: {len(progress['processed_nips'])} pegawai")
log_message(f"   Log detail tersimpan di: {LOG_FILE}")
log_message(f"   Progress tersimpan di: {PROGRESS_FILE}")
log_message(f"{'='*60}")

# Tutup browser (optional, bisa di-comment jika ingin tetap buka)
# driver.quit()