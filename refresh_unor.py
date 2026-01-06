import os
import pandas as pd
import time
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Membaca file Excel
try:
    df = pd.read_excel('test_list_pppk.xlsx')
    print("File ditemukan, melanjutkan proses...")
except FileNotFoundError:
    print(f"Error: file tidak ditemukan.")
    exit()

# Koneksi ke Chrome yang sudah berjalan dengan remote debugging
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

try:
    driver = webdriver.Chrome(options=chrome_options)
    print("Berhasil terhubung ke Chrome")
except Exception as e:
    print(f"Error koneksi ke Chrome: {e}")
    print("Pastikan Chrome sudah berjalan dengan: chrome.exe --remote-debugging-port=9222 --user-data-dir=\"C:\\ChromeTemp\"")
    exit()

# Proses setiap NIP dari Excel
for index, row in df.iterrows():
    nip_pppk = str(row['nip'])
    print(f"\n{'='*60}")
    print(f"Memulai proses untuk NIP: {nip_pppk} (Baris ke-{index + 1} dari {len(df)})")
    print(f"{'='*60}")
    
    try:
        # Step 1: Isi NIP BARU di input field
        print("1. Mengisi NIP BARU...")
        wait = WebDriverWait(driver, 10)
        
        # Cari input field dengan placeholder "Masukan NIP Baru"
        nip_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@placeholder='Masukan NIP Baru']")
        ))
        nip_input.clear()
        nip_input.send_keys(nip_pppk)
        print(f"   ✓ NIP {nip_pppk} berhasil diisi")
        time.sleep(1)
        
        # Step 2: Klik tombol search (button dengan img magnify-scan.png)
        print("2. Menekan tombol search...")
        search_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@type='submit']//img[@src='/img/magnify-scan.png']/parent::button")
        ))
        search_button.click()
        print("   ✓ Tombol search diklik")
        time.sleep(3)  # Tunggu hasil pencarian
        
        # Step 3: Klik checkbox/label "Posisi & Jabatan"
        print("3. Membuka tab Posisi & Jabatan...")
        posisi_jabatan = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//label[@for='Posisi & Jabatan']")
        ))
        posisi_jabatan.click()
        print("   ✓ Tab Posisi & Jabatan dibuka")
        time.sleep(2)
        
        # Step 4: Scroll ke bawah untuk menemukan tombol refresh
        print("4. Mencari tombol Refresh Data Unor...")
        driver.execute_script("window.scrollBy(0, 500);")
        time.sleep(1)
        
        # Step 5: Klik tombol "Refresh Data Unor"
        print("5. Menekan tombol Refresh Data Unor...")
        refresh_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Refresh Data Unor')]")
        ))
        refresh_button.click()
        print("   ✓ Tombol Refresh Data Unor diklik")
        time.sleep(2)
        
        # Step 6: Klik tombol "Proses" di modal
        print("6. Mengkonfirmasi proses di modal...")
        proses_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@class='swal2-confirm swal2-styled' and contains(., 'Proses')]")
        ))
        proses_button.click()
        print("   ✓ Tombol Proses diklik")
        
        # Tunggu 3 detik untuk proses refresh selesai
        print("7. Menunggu proses refresh selesai...")
        time.sleep(3)
        
        print(f"✅ Proses untuk NIP {nip_pppk} selesai!")
        
    except TimeoutException as e:
        print(f"❌ Error: Timeout saat memproses NIP {nip_pppk}")
        print(f"   Detail: Element tidak ditemukan dalam waktu yang ditentukan")
        continue
    except Exception as e:
        print(f"❌ Error saat memproses NIP {nip_pppk}: {e}")
        continue
    
    # Jeda sebelum proses NIP berikutnya
    time.sleep(2)

print(f"\n{'='*60}")
print(f"✅ SEMUA PROSES SELESAI!")
print(f"   Total data diproses: {len(df)} pegawai")
print(f"{'='*60}")

# Tutup browser (optional, bisa di-comment jika ingin tetap buka)
# driver.quit()