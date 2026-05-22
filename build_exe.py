"""
Build script to create a standalone .exe for the Auth Extractor application.
Run this script once to generate the executable.

Usage: python build_exe.py
"""

import subprocess
import sys
import os
import shutil
import pathlib

def main():
    print("=" * 60)
    print("Auth Extractor - Build Executable")
    print("=" * 60)
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Step 1: Install PyInstaller if not already installed
    print("\n[1/4] Checking PyInstaller...")
    try:
        import PyInstaller
        print("      PyInstaller is already installed.")
    except ImportError:
        print("      Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("      PyInstaller installed.")
    
    # Step 2: Check for required files
    print("\n[2/4] Checking required files...")
    main_script = os.path.join(script_dir, "auth_extractor.py")
    if not os.path.exists(main_script):
        print("      ERROR: auth_extractor.py not found!")
        return
    print("      auth_extractor.py found.")
    
    # Check for poppler
    poppler_path = os.path.join(script_dir, "poppler", "poppler-24.08.0", "Library", "bin")
    if os.path.exists(poppler_path):
        print(f"      Poppler found at: {poppler_path}")
    else:
        print("      WARNING: Poppler not found. PDF extraction may not work.")
        poppler_path = None
    
    # Check for credentials.json
    creds_file = os.path.join(script_dir, "credentials.json")
    if os.path.exists(creds_file):
        print("      credentials.json found (will be included).")
    else:
        print("      NOTE: credentials.json not found. Users will need to add it for Gmail.")
    
    # Step 2b: Close any running AuthExtractor.exe so PyInstaller can overwrite dist/
    print("\n      Checking for running AuthExtractor.exe...")
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "AuthExtractor.exe"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print("      Closed running AuthExtractor.exe.")
            import time; time.sleep(1)  # brief pause so OS releases file handles
        else:
            print("      AuthExtractor.exe not running (OK).")
    except Exception:
        pass  # taskkill not available or failed — proceed anyway

    # Step 3: Build the executable
    print("\n[3/4] Building executable (this may take a few minutes)...")
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=AuthExtractor",
        "--onedir",  # Create a folder with exe + dependencies (more reliable than onefile)
        "--windowed",  # No console window
        "--noconfirm",  # Overwrite without asking
    ]
    
    # Add icon if it exists
    icon_path = os.path.join(script_dir, "icon.ico")
    if os.path.exists(icon_path):
        cmd.append(f"--icon={icon_path}")

    # Add logo image
    logo_path = os.path.join(script_dir, "Auth Radar Logo.png")
    if os.path.exists(logo_path):
        cmd.append(f"--add-data={logo_path};.")

    # Add .env file
    env_file = os.path.join(script_dir, ".env")
    if os.path.exists(env_file):
        cmd.append(f"--add-data={env_file};.")
        print("      .env found (will be included).")

    # Add new module packages
    for pkg_dir in ["integrations", "extraction", "audit", "services", "review", "db"]:
        pkg_path = os.path.join(script_dir, pkg_dir)
        if os.path.exists(pkg_path):
            cmd.append(f"--add-data={pkg_path};{pkg_dir}")
            print(f"      {pkg_dir}/ package found (will be included).")

    # Add config.py
    config_path = os.path.join(script_dir, "config.py")
    if os.path.exists(config_path):
        cmd.append(f"--add-data={config_path};.")
        print("      config.py found (will be included).")

    # Add poppler binaries
    if poppler_path:
        cmd.append(f"--add-data={poppler_path};poppler")
    
    # Add credentials.json if it exists
    if os.path.exists(creds_file):
        cmd.append(f"--add-data={creds_file};.")
    
    # Add hidden imports for libraries that PyInstaller might miss
    hidden_imports = [
        "pdfplumber",
        "pdfplumber.page",
        "pdfminer",
        "pdfminer.pdfparser",
        "pdfminer.pdfdocument",
        "pdfminer.pdfpage",
        "PIL",
        "PIL.Image",
        "pdf2image",
        "pdf2image.pdf2image",
        "pytesseract",
        "pandas",
        "openpyxl",
        "google.oauth2.credentials",
        "google_auth_oauthlib.flow",
        "google.auth.transport.requests",
        "googleapiclient.discovery",
        "googleapiclient.errors",
        "dotenv",
        "dropbox",
        "dropbox.files",
        "dropbox.oauth",
        "integrations",
        "integrations.dropbox_service",
        "extraction",
        "extraction.schema",
        "extraction.router",
        "extraction.pdf_text_extractor",
        "extraction.ocr_extractor",
        "extraction.structured_extractor",
        "extraction.excel_extractor",
        "audit",
        "audit.logger",
        "services",
        "services.pdf_unlock_service",
        "services.excel_export_service",
        "services.page_detection_service",
        "fitz",
        "fitz._fitz",
        "config",
    ]
    
    for imp in hidden_imports:
        cmd.append(f"--hidden-import={imp}")

    # PyMuPDF (fitz) ships native DLLs that --hidden-import alone won't collect.
    # --collect-all ensures the compiled extension and all data files are bundled.
    cmd.append("--collect-all=fitz")

    # Add the main script
    cmd.append(main_script)
    
    print(f"      Running: {' '.join(cmd[:5])}...")
    
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as e:
        print(f"\n      ERROR: Build failed with error: {e}")
        return
    
    # Step 4: Post-build setup
    print("\n[4/4] Finalizing...")
    
    dist_dir = os.path.join(script_dir, "dist", "AuthExtractor")
    
    if os.path.exists(dist_dir):
        # Copy poppler to the dist folder if not already there
        if poppler_path:
            dest_poppler = os.path.join(dist_dir, "poppler")
            if not os.path.exists(dest_poppler):
                print("      Copying Poppler binaries...")
                shutil.copytree(poppler_path, dest_poppler)
        
        # Copy gmail_accounts.json if it exists
        gmail_config = os.path.join(script_dir, "data", "gmail_accounts.json")
        if os.path.exists(gmail_config):
            print("      Copying gmail_accounts.json...")
            shutil.copy2(gmail_config, dist_dir)
        
        # Copy any existing token files
        for token_file in pathlib.Path(script_dir, "data").glob("token_*.json"):
            print(f"      Copying {token_file.name}...")
            shutil.copy2(token_file, dist_dir)
        
        # Copy patient_names.json if it exists
        patient_names = os.path.join(script_dir, "data", "patient_names.json")
        if os.path.exists(patient_names):
            print("      Copying patient_names.json...")
            shutil.copy2(patient_names, dist_dir)

        # Copy logo if it exists
        logo_src = os.path.join(script_dir, "Auth Radar Logo.png")
        if os.path.exists(logo_src):
            print("      Copying Auth Radar Logo.png...")
            shutil.copy2(logo_src, dist_dir)

        # Copy .env
        env_src = os.path.join(script_dir, ".env")
        if os.path.exists(env_src):
            print("      Copying .env...")
            shutil.copy2(env_src, dist_dir)

        # Copy config.py
        config_src = os.path.join(script_dir, "config.py")
        if os.path.exists(config_src):
            print("      Copying config.py...")
            shutil.copy2(config_src, dist_dir)

        # Ensure data/ directory exists in dist
        dist_data = os.path.join(dist_dir, "data")
        os.makedirs(dist_data, exist_ok=True)

        print("\n" + "=" * 60)
        print("BUILD SUCCESSFUL!")
        print("=" * 60)
        print(f"\nExecutable created at:")
        print(f"  {os.path.join(dist_dir, 'AuthExtractor.exe')}")
        print(f"\nTo distribute:")
        print(f"  1. Copy the entire 'dist/AuthExtractor' folder to the shared location")
        print(f"  2. Users run 'AuthExtractor.exe' from that folder")
        print(f"\nNote: If using Gmail, place 'credentials.json' in the same folder as the .exe")
    else:
        print("\n      ERROR: Build output not found!")

if __name__ == "__main__":
    main()
    input("\nPress Enter to exit...")
