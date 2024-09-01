import pyautogui
import os
import ctypes
import keyboard
import pygetwindow as gw
import subprocess
import time

# Lokasi direktori input gambar
image_dir = r'C:\python\remove_text_\detected_text_frames'
output_dir = r'C:\python\remove_text_\cleaned_text_frames'

# Path lengkap ke executable GIMP
gimp_executable = r'C:\Program Files\GIMP 2\bin\gimp-2.10.exe'

# Waktu tunggu maksimal untuk operasi
max_wait_time = 60  # Dalam detik
operation_delay = 10  # Waktu tunggu antara operasi (dalam detik)
heal_delay = 25  # Waktu tunggu untuk Heal Selection (dalam detik)
rectangle_select_tool_delay = 10  # Waktu tunggu setelah Rectangle Select Tool (dalam detik)
export_delay = 20  # Waktu tunggu setelah ekspor (dalam detik)
satu_detik_delay = 1  # Waktu tunggu 1 detik

# Nama jendela GIMP
gimp_window_title = "GIMP"

def minimize_command_prompt():
    """Minimalkan jendela Command Prompt"""
    try:
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
        print("Command Prompt minimized.")
    except Exception as e:
        print(f"Error minimizing Command Prompt: {e}")

def check_abort():
    """Cek apakah tombol ESC ditekan untuk membatalkan proses"""
    if keyboard.is_pressed('esc'):
        print("Proses dibatalkan oleh pengguna.")
        exit()

def print_debug_info(message):
    """Print informasi debug"""
    print(message)

def bring_gimp_to_foreground(title):
    """Pastikan jendela GIMP dalam fokus"""
    try:
        print("Memeriksa jendela GIMP...")
        windows = gw.getWindowsWithTitle(title)
        if windows:
            gimp_window = windows[0]
            print(f"Menemukan jendela GIMP: {gimp_window.title}")
            if not gimp_window.isActive:
                print("Mengaktifkan jendela GIMP...")
                gimp_window.activate()
            if not gimp_window.isMaximized:
                print("Memaksimalkan jendela GIMP...")
                gimp_window.maximize()
            return True
        else:
            print(f"GIMP window with title '{title}' not found.")
            return False
    except Exception as e:
        print(f"Error bringing GIMP to foreground: {e}")
        return False

def wait_for_gimp_to_open(title):
    """Tunggu sampai jendela GIMP terbuka"""
    try:
        print("Menunggu GIMP terbuka...")
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            if bring_gimp_to_foreground(title):
                print("GIMP sudah terbuka.")
                return True
            time.sleep(5)
        print("Timeout: GIMP tidak terbuka dalam waktu yang ditentukan.")
        return False
    except Exception as e:
        print(f"Error waiting for GIMP to open: {e}")
        return False

def wait_for_image_to_load():
    """Tunggu hingga gambar benar-benar dimuat di GIMP"""
    try:
        print("Menunggu gambar dimuat di GIMP...")
        time.sleep(5)  # Waktu tunda tambahan untuk pemuatan gambar
        print("Gambar mungkin telah dimuat.")
        return True
    except Exception as e:
        print(f"Error waiting for image to load: {e}")
        return False

def open_gimp(filepath):
    """Buka GIMP dengan gambar yang ditentukan"""
    try:
        print(f"Membuka GIMP dengan gambar: {filepath}")
        subprocess.Popen([gimp_executable, filepath])
        print("GIMP berhasil dibuka dengan gambar.")
        return True
    except Exception as e:
        print(f"Terjadi kesalahan saat membuka GIMP: {e}")
        return False

def perform_image_editing(x1=400, y1=400, x2=600, y2=600):
    """Lakukan proses pengeditan gambar di GIMP"""
    try:
        print_debug_info("GIMP dalam fokus")

        # Pilih Rectangle Select Tool
        try:
            print("Memilih Rectangle Select Tool...")
            pyautogui.hotkey('r')
            time.sleep(rectangle_select_tool_delay)
        except Exception as e:
            print(f"Error selecting Rectangle Select Tool: {e}")

        # Klik dan drag untuk memilih area dengan Rectangle Select
        try:
            print("Memilih area dengan Rectangle Select...")
            pyautogui.moveTo(x=x1, y=y1)
            pyautogui.dragTo(x=x2, y=y2, button='left')
            time.sleep(rectangle_select_tool_delay)
        except Exception as e:
            print(f"Error selecting area: {e}")

        check_abort()  # Periksa pembatalan sebelum menjalankan Heal Selection

        # Jalankan Heal Selection
        try:
            print_debug_info("Menjalankan Heal Selection...")
            pyautogui.hotkey('ctrl', 'shift', 'h')
            time.sleep(2)
            pyautogui.press('tab')  # Pindah 1
            pyautogui.press('tab')  # Pindah 2
            pyautogui.press('tab')  # Pindah 3
            pyautogui.press('tab')  # Pindah Help
            pyautogui.press('enter')  # pilih OK
            time.sleep(heal_delay)
        except Exception as e:
            print(f"Error running Heal Selection: {e}")

        check_abort()  # Periksa pembatalan sebelum menyimpan gambar

        # Simpan gambar dengan ekspor
        try:
            print("Menyimpan gambar...")
            pyautogui.hotkey('ctrl', 'shift', 'e')
            time.sleep(3)  # Tunggu sebelum memulai dialog ekspor

            # Masukkan nama file dan lokasi simpan di dialog ekspor
            export_filepath = os.path.join(output_dir, f'processed_{filename}')
            print(f"Menulis nama file ekspor: {export_filepath}")
            pyautogui.write(export_filepath)
            pyautogui.press('enter')

            # Menyelesaikan dialog ekspor
            print("Menyelesaikan dialog ekspor...")
            time.sleep(2)  # Tunggu setelah menyelesaikan dialog ekspor
            pyautogui.press('enter')
        except Exception as e:
            print(f"Error saving image: {e}")

        # Tutup gambar tanpa menyimpan ulang (save changes)
        try:
            print("Menutup gambar tanpa menyimpan...")
            time.sleep(6)
            # Cek jendela gambar dengan getWindowsWithTitle
            windows = gw.getWindowsWithTitle(filename)
            if windows:
                image_window = windows[0]
                if image_window.isActive:
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(satu_detik_delay)

                    # Mengatasi prompt "Save Changes"
                    pyautogui.hotkey('ctrl', 'd')
                    time.sleep(operation_delay)
                else:
                    print(f"Window for image '{filename}' is not active.")
            else:
                print(f"Window for image '{filename}' not found.")
        except Exception as e:
            print(f"Error closing image: {e}")

    except Exception as e:
        print(f"Terjadi kesalahan saat melakukan pengeditan gambar: {e}")

# Minimize Command Prompt sebelum membuka GIMP
minimize_command_prompt()

# Loop melalui setiap gambar di direktori
for filename in os.listdir(image_dir):
    check_abort()  # Periksa pembatalan sebelum memulai proses

    if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        filepath = os.path.join(image_dir, filename)
        if open_gimp(filepath):
            # Tunggu sampai GIMP benar-benar terbuka dan gambar dimuat
            if wait_for_gimp_to_open(gimp_window_title):
                if wait_for_image_to_load():
                    perform_image_editing()
                else:
                    print("Gambar tidak dimuat. Lanjut ke gambar berikutnya.")
            else:
                print("GIMP tidak dapat dibuka atau tidak dapat diakses.")
        else:
            print("GIMP gagal dibuka dengan gambar. Lanjut ke gambar berikutnya.")

# Selesai
print("Batch processing selesai.")
