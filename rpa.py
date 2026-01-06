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


# --- PENGATURAN KONEKSI KE BROWSER ---
# "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeTemp"
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)
print("Berhasil terhubung ke browser Chrome yang sudah aktif.")
wait = WebDriverWait(driver, 20)

# --- MEMBACA DATA DARI CSV ---
try:
    df = pd.read_csv('data.csv')
    print(f"Data berhasil dibaca. Ditemukan {len(df)} baris data untuk diproses.")
except FileNotFoundError:
    print("Error: File 'data.csv' tidak ditemukan. Pastikan file ada di folder yang sama.")
    exit()

# --- PROSES OTOMASI PER BARIS DATA ---
for index, row in df.iterrows():
    nip_pegawai = str(row['nip'])
    print(f"\nMemulai proses untuk NIP: {nip_pegawai} (Baris ke-{index + 1})")

    try:
        
        # =====================================================================
        # LANGKAH 1: CARI PEGAWAI
        # =====================================================================
       
        print("   - Tahap 1: Mencari Pegawai...")
        nip_input = wait.until(EC.presence_of_element_located((By.NAME, 'nip_baru'))) 
        nip_input.send_keys(nip_pegawai)
        cari_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Cari Pegawai']")))
        cari_button.click()
        print("   - Data pegawai berhasil ditampilkan.")

        print("   - Tahap 2: Memilih Unit Verifikasi...")
        unit_verifikasi_dropdown = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'ant-select-selector')]")
        ))
        unit_verifikasi_dropdown.click()
        
        opsi_biro = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'ant-select-item-option-content') and text()='Biro Sumber Daya Manusia dan Organisasi']")
        ))
        opsi_biro.click()
        print("   - 'Biro Sumber Daya Manusia dan Organisasi' telah dipilih.")

        print("   - Tahap 3: Mengklik tombol 'Berikutnya'...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Berikutnya') and not(@disabled)]"))).click()
        
        print("Langkah 1 berhasil diselesaikan.")
        time.sleep(2)
       
        # =====================================================================
        # LANGKAH 2: PILIH PROSEDUR (DENGAN PERBAIKAN)
        # =====================================================================
        print("--- Memulai Langkah 2: Pilih Prosedur ---")
        
        label_unit_organisasi = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[text()='Unit Organisasi']")
        ))
        label_unit_organisasi.click()
        print("   - Label 'Unit Organisasi' telah diklik.")
        
        # 2. Tunggu tombol "Berikutnya" muncul dan aktif, lalu klik
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='Berikutnya']")
        )).click()
        
        print("Langkah 2 berhasil diselesaikan.")

        time.sleep(2)

        # =====================================================================
        # LANGKAH 2: INPUT FORM DATA & UPLOAD FILE 
        # =====================================================================
        print("--- Memulai Langkah 3 ")
        label_unit_organisasi = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[text()='Unit Organisasi']")
        ))
        label_unit_organisasi.click()
        time.sleep(2)
        
        print("   - Mengisi form 'Ubah Data'...")
        
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[text()='Unit Organisasi']")
        )).click()

       
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='value' and text()='Pilih unit organisasi']"))).click()
        
        search_unor_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Cari unit organiasi']")))
        unit_organisasi_target = row['unit_organisasi_baru']
        search_unor_input.send_keys(unit_organisasi_target)

        wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[@class='dropdown-item' and normalize-space()='{unit_organisasi_target}']"))).click()
        print(f"     - Unit Organisasi Baru: '{unit_organisasi_target}' dipilih.")

        time.sleep(2)
        # 2. Isi No SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='Masukkan no SK pindah unor']").send_keys(row['no_sk_pindah_unor'])
        time.sleep(2)
        # 3. Isi Tanggal SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='Pilih Tanggal SK Pindah Unor']").send_keys(row['tanggal_sk_pindah_unor'])
        time.sleep(2)
        # 4. Isi TMT SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='TMT SK Pindah Unor']").send_keys(row['tmt_sk_pindah_unor'])
        time.sleep(2)
        # 5. Isi No Surat Persetujuan PAN RB
        driver.find_element(By.XPATH, "//input[@placeholder='Masukkan no surat persetujuan Pan RB']").send_keys(row['no_pan_rb'])
        time.sleep(2)
        # 6. Isi Tanggal Surat Persetujuan PAN RB
        driver.find_element(By.XPATH, "//input[@placeholder='Pilih Tanggal surat persetujuan Pan RB']").send_keys(row['tgl_surat_pan_rb'])
        time.sleep(2)
        # Klik tombol Simpan
        driver.find_element(By.XPATH, "//button[text()='Simpan']").click()
        print("   - Data pada form 'Ubah Data' telah disimpan.")
        time.sleep(2)

        driver.refresh()
        time.sleep(5)

        print("--- Memulai Langkah 3 ")
        label_unit_organisasi = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[text()='Unit Organisasi']")
        ))
        label_unit_organisasi.click()
        time.sleep(2)
        

        print("   - Mengisi form 'Ubah Data'...")
       

        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='value' and text()='Pilih unit organisasi']"))).click()
  
        search_unor_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Cari unit organiasi']")))
        unit_organisasi_target = row['unit_organisasi_baru']
        search_unor_input.send_keys(unit_organisasi_target)
 
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[@class='dropdown-item' and normalize-space()='{unit_organisasi_target}']"))).click()
        print(f"     - Unit Organisasi Baru: '{unit_organisasi_target}' dipilih.")

        time.sleep(2)
        # 2. Isi No SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='Masukkan no SK pindah unor']").send_keys(row['no_sk_pindah_unor'])
        time.sleep(2)
        # 3. Isi Tanggal SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='Pilih Tanggal SK Pindah Unor']").send_keys(row['tanggal_sk_pindah_unor'])
        time.sleep(2)
        # 4. Isi TMT SK Pindah Unor
        driver.find_element(By.XPATH, "//input[@placeholder='TMT SK Pindah Unor']").send_keys(row['tmt_sk_pindah_unor'])
        time.sleep(2)
        # 5. Isi No Surat Persetujuan PAN RB
        driver.find_element(By.XPATH, "//input[@placeholder='Masukkan no surat persetujuan Pan RB']").send_keys(row['no_pan_rb'])
        time.sleep(2)
        # 6. Isi Tanggal Surat Persetujuan PAN RB
        driver.find_element(By.XPATH, "//input[@placeholder='Pilih Tanggal surat persetujuan Pan RB']").send_keys(row['tgl_surat_pan_rb'])
        time.sleep(2)
        # Klik tombol Simpan
        driver.find_element(By.XPATH, "//button[text()='Simpan']").click()
        print("   - Data pada form 'Ubah Data' telah disimpan.")
        time.sleep(2)


        dokumen_tab = wait.until(EC.element_to_be_clickable((
        By.XPATH,"//ul[@class='tab-list']/li[normalize-space(.)='Dokumen Pendukung']")))

        dokumen_tab.click()
        assert "active" in dokumen_tab.get_attribute("class")

        upload_paths = [       
        f"{nip_pegawai}.pdf",
        r"sk_panrb.pdf"  
        ]

        # upload_folder = "dokumen_upload"
        # filenames = [
        #     f"{nip_pegawai}.pdf",
        #     "sk_panrb.pdf"
        # ]
        # upload_paths = [os.path.abspath(os.path.join(upload_folder, fname)) for fname in filenames]

        browse_buttons = wait.until(EC.presence_of_all_elements_located((  
            By.CSS_SELECTOR, "a.upload"  
        )))
 
        for button, path in zip(browse_buttons, upload_paths):
 
            button.click()
  
            time.sleep(4)
  
            pyautogui.write(path, interval=0.02)
            pyautogui.press("enter")

            time.sleep(5)

        print(" Semua file telah berhasil di‐upload.")   

        driver.find_element(By.XPATH, "//button[text()='Simpan']").click()
        print("   - Data pada form 'Ubah Data' telah disimpan.")
        time.sleep(2)

        # =====================================================================
        # LANGKAH 4: VERIFIKASI DATA
        # =====================================================================
        confirm_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "button.swal2-confirm.swal2-styled"
            ))
        )
        confirm_btn.click()

        langkah4 = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='Berikutnya']")
        ))
        
        langkah4.click()

        time.sleep(2)

        # =====================================================================
        # LANGKAH 5: SIMPAN BERKAS
        # =====================================================================

        langkah5 = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='Berikutnya']")
        ))
        langkah5.click()

        time.sleep(2)
        
        print("--- Memulai Langkah Final: Simpan Berkas ---")
        
        langkah6 = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[text()='SIMPAN BERKAS']")
        ))
        langkah6.click()
        print("   - Tombol 'SIMPAN BERKAS' berhasil diklik.")

        wait.until(EC.visibility_of_element_located(
            (By.CLASS_NAME, "swal2-popup")
        ))
        print("   - Dialog konfirmasi muncul.")
        
        final_btn = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, "button.swal2-confirm.swal2-styled"
        )))
        final_btn.click()
        print("   - Tombol konfirmasi 'Simpan' berhasil diklik.")

        print("Proses untuk NIP ini selesai, menunggu sebelum lanjut...")

        # =====================================================================
        # KEMBALI KE HALAMAN UTAMA PPPK
        # =====================================================================
        time.sleep(2)

        pppk_card = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[@class='card-title h5' and text()='PPPK']")
        ))

        pppk_card.click()
        print("Tombol/kartu PPPK berhasil diklik.")

    except TimeoutException:
        print(f"Error: Gagal menemukan elemen atau waktu tunggu habis pada NIP {nip_pegawai}. Melanjutkan ke data berikutnya.")
        driver.refresh()
        time.sleep(3)
        continue
    except Exception as e:
        print(f"Terjadi error yang tidak terduga pada NIP {nip_pegawai}: {e}")
        continue

print("\nSemua proses telah selesai.")