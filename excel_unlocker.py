import zipfile
import os
import re
import shutil
import sys

def decrypt_excel(fn: str) -> None:
    if not os.path.exists(fn):
        print(f"Error: File '{fn}' not found.")
        return

    bn, ext = os.path.splitext(fn)
    if ext.lower() != ".xlsx":
        print("This script is designed for .xlsx files. Proceeding, but unexpected results may occur.")

    bak_fn = bn + ".bak.xlsx"
    try:
        shutil.copy2(fn, bak_fn)
    except Exception as e:
        print(f"Error creating backup: {e}")
        return

    zip_fn = bn + ".zip"
    try:
        os.rename(fn, zip_fn)
    except Exception as e:
        print(f"Error renaming file to .zip: {e}")
        shutil.copy2(bak_fn, fn)
        return

    try:
        with zipfile.ZipFile(zip_fn, 'r') as zip_ref:
            tmp_dir = "temp_extraction"
            os.makedirs(tmp_dir, exist_ok=True)
            zip_ref.extractall(tmp_dir)
            for name in zip_ref.namelist():
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    extracted_file_path = os.path.join(tmp_dir, name)
                    with open(extracted_file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    try:
                        modified_content = re.sub(r'<sheetProtection[^>]*?\s?scenarios="[0-9]+"[^>]*?\s?/>', "", content)
                    except re.error as re_err:
                        continue
                    if modified_content != content:
                        with open(extracted_file_path, 'w', encoding='utf-8') as f:
                            f.write(modified_content)
            outzip_fn = bn + "_modified.zip"
            with zipfile.ZipFile(outzip_fn, 'w', zipfile.ZIP_DEFLATED) as n_zip:
                for root, _, files in os.walk(tmp_dir):
                    for file in files:
                        f_path = os.path.join(root, file)
                        arcname = os.path.relpath(f_path, tmp_dir)
                        n_zip.write(f_path, arcname)

    except zipfile.BadZipFile:
        print(f"Error: '{zip_fn}' is not a valid zip file.")
        os.rename(zip_fn, fn)
        return
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        os.rename(zip_fn, fn)
        return
    finally:
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)

    fin_fn = bn + "_removed" + ext
    try:
        os.rename(outzip_fn, fin_fn)
        os.remove(zip_fn)
    except Exception as e:
        print(f"Error renaming/removing the final file {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <excel_filename>")
    else:
        main_fn = sys.argv[1]
        decrypt_excel(main_fn)
