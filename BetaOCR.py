import ctypes
import os
import subprocess
import logging
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pikepdf

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def install_missing_packages(package):
    try:
        __import__(package)
    except ImportError:
        logging.info(f"Package '{package}' not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        logging.info(f"Package '{package}' installed successfully.")

# Check and install necessary Python packages
required_packages = ["PyPDF2", "pikepdf"]
for package in required_packages:
    install_missing_packages(package)

import PyPDF2

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logging.error("Error checking admin status: %s", e)
        return False

def run_powershell_command(command):
    try:
        subprocess.check_call(["powershell", "-Command", command], shell=True)
    except subprocess.CalledProcessError as e:
        logging.error("PowerShell command failed: %s", e)
        raise

def select_files():
    root = tk.Tk()
    root.withdraw()
    file_types = [('PDF files', '*.pdf')]
    file_paths = filedialog.askopenfilenames(title="Select files", filetypes=file_types)
    return file_paths

def select_directory():
    root = tk.Tk()
    root.withdraw()
    directory_path = filedialog.askdirectory(title="Select directory")
    return directory_path

def show_progress_window(file_count):
    progress_win = tk.Tk()
    progress_win.title("Processing Files")
    ttk.Label(progress_win, text="Processing, please wait...").pack(pady=10)
    progress_bar = ttk.Progressbar(progress_win, orient=tk.HORIZONTAL, length=300, mode='determinate', maximum=file_count)
    progress_bar.pack(pady=10)
    progress_win.update()
    return progress_win, progress_bar

def decrypt_pdf(file_path, password):
    try:
        with pikepdf.open(file_path, password=password) as pdf:
            temp_filename = f"{file_path}_temp.pdf"
            pdf.save(temp_filename)
        return temp_filename
    except pikepdf._qpdf.PasswordError as e:
        logging.error(f"Incorrect password for file {file_path}: {e}")
        return None

def check_pdf_password(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        if reader.is_encrypted:
            root = tk.Tk()
            root.withdraw()
            password = simpledialog.askstring("Password", f"Enter password for {os.path.basename(file_path)}:", show='*')
            root.destroy()
            return password
        return None

def get_installed_ghostscript_version():
    try:
        result = subprocess.check_output(["choco", "list", "--local-only", "ghostscript"], text=True)
        versions = re.findall(r'ghostscript (\d+\.\d+\.\d+)', result)
        return versions[0] if versions else None
    except subprocess.CalledProcessError:
        return None

def upgrade_ghostscript():
    installed_version = get_installed_ghostscript_version()
    if installed_version:
        logging.info(f"Currently installed Ghostscript version: {installed_version}")
        major, minor, _ = map(int, installed_version.split('.'))
        if major < 10 or (major == 10 and minor <= 2):
            logging.info("Upgrading Ghostscript...")
            run_powershell_command("choco upgrade ghostscript -y")
            logging.info("Ghostscript upgraded successfully.")
        else:
            logging.info("Ghostscript version is up to date.")
    else:
        logging.info("Ghostscript is not currently installed. Installing it...")
        run_powershell_command("choco install ghostscript -y")
        logging.info("Ghostscript installed successfully.")

def process_pdf(file_paths, output_directory, progress_bar):
    processed_files = []
    skipped_files = []
    errors = []

    for index, file_path in enumerate(file_paths):
        filename = os.path.basename(file_path)
        output_filename = f"ocr_{filename}"
        output_path = os.path.join(output_directory, output_filename)

        if filename.lower().startswith('ocr_') or os.path.exists(output_path):
            logging.info(f"Skipping file (already processed or OCR version exists): {file_path}")
            skipped_files.append(file_path)
            continue

        password = check_pdf_password(file_path)
        temp_file_path = file_path
        if password:
            temp_file_path = decrypt_pdf(file_path, password)
            if temp_file_path is None:
                errors.append(f"Failed to decrypt {file_path}")
                continue

        ocrmypdf_args = ['ocrmypdf', '--force-ocr', '--output-type', 'pdf', temp_file_path, output_path]

        try:
            subprocess.run(ocrmypdf_args, check=True, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            logging.info(f"Output saved to {output_path}")
            processed_files.append(file_path)
        except subprocess.CalledProcessError as e:
            errors.append(f"{file_path}: {e.stderr}")

        if temp_file_path != file_path:
            os.remove(temp_file_path)  # Clean up the temporary decrypted file

        progress_bar['value'] = (index + 1) / len(file_paths) * 100
        progress_bar.update()

    return processed_files, skipped_files, errors

def show_summary(processed_files, skipped_files, errors, input_dir, output_dir):
    summary_win = tk.Tk()
    summary_win.title("Processing Summary")
    ttk.Label(summary_win, text="Processing Complete! Summary:", font=("Arial", 14)).pack(pady=10)

    ttk.Label(summary_win, text=f"Input Directory: {input_dir}").pack()
    ttk.Label(summary_win, text=f"Output Directory: {output_dir}").pack()

    ttk.Label(summary_win, text="Processed Files:").pack()
    for file in processed_files:
        ttk.Label(summary_win, text=file).pack()

    if skipped_files:
        ttk.Label(summary_win, text="Skipped Files:").pack()
        for file in skipped_files:
            ttk.Label(summary_win, text=file).pack()

    error_text = "Errors: None" if not errors else "Errors:"
    ttk.Label(summary_win, text=error_text).pack()
    for error in errors:
        ttk.Label(summary_win, text=error).pack()

    ttk.Button(summary_win, text="Close", command=lambda: close_summary(summary_win)).pack(pady=10)
    summary_win.protocol("WM_DELETE_WINDOW", lambda: close_summary(summary_win))
    summary_win.mainloop()

def close_summary(window):
    window.destroy()
    sys.exit()

def is_chocolatey_installed():
    try:
        subprocess.check_call(["choco", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False
    except FileNotFoundError:
        return False

def install_chocolatey():
    logging.info("Installing Chocolatey...")
    powershell_command = "Set-ExecutionPolicy Bypass -Scope Process -Force; " \
                         "[System.Net.ServicePointManager]::SecurityProtocol = " \
                         "[System.Net.ServicePointManager]::SecurityProtocol -bor 3072; " \
                         "iwr https://chocolatey.org/install.ps1 -UseBasicParsing | iex"
    run_powershell_command(powershell_command)
    logging.info("Chocolatey installed successfully.")

def add_choco_to_path():
    choco_path = "C:\\ProgramData\\chocolatey\\bin"
    current_path = os.environ.get("PATH", "")
    if choco_path not in current_path:
        updated_path = current_path + ";" + choco_path
        os.environ["PATH"] = updated_path
        logging.info("Added Chocolatey to PATH")

def install_ghostscript():
    if not get_installed_ghostscript_version():
        logging.info("Installing Ghostscript...")
        run_powershell_command("choco install ghostscript -y")
        logging.info("Ghostscript installed successfully.")

def is_tesseract_installed():
    try:
        subprocess.check_call(["tesseract", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False
    except FileNotFoundError:
        return False

def install_tesseract():
    logging.info("Installing Tesseract...")
    run_powershell_command("choco install tesseract -y")
    logging.info("Tesseract installed successfully.")

def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
        logging.info(f"Successfully installed/updated package: {package}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error installing package {package}: {e}")

def automated_setup():
    install_package("ocrmypdf")
    if not is_tesseract_installed():
        install_tesseract()
    if not is_chocolatey_installed():
        install_chocolatey()
        add_choco_to_path()
    install_ghostscript()

if __name__ == "__main__":
    if is_admin():
        automated_setup()

        messagebox.showinfo("Directory Selection", "Select directory to scan PDFs")
        input_directory = select_directory()

        messagebox.showinfo("Directory Selection", "Select a directory to save processed files")
        output_directory = select_directory()

        file_paths = [os.path.join(input_directory, f) for f in os.listdir(input_directory) if f.lower().endswith('.pdf')]
        progress_win, progress_bar = show_progress_window(len(file_paths))

        processed_files, skipped_files, errors = process_pdf(file_paths, output_directory, progress_bar)
        progress_win.destroy()

        show_summary(processed_files, skipped_files, errors, input_directory, output_directory)
    else:
        logging.error("Administrative privileges not detected. Please run the script as an administrator.")
