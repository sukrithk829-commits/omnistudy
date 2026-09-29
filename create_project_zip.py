"""Create a clean, complete, lossless zip archive of the OmniStudy project for migration."""
import os
import sys
import zipfile
import shutil

SOURCE_DIR = r"D:\BCA"
OUTPUT_ZIP_D = r"D:\omnistudy_project_complete.zip"
OUTPUT_ZIP_LOCAL = r"D:\BCA\omnistudy_project_complete.zip"

EXCLUDE_DIRS = {
    ".venv",
    "__pycache__",
    ".pytest_cache",
    ".idea",
    ".vscode",
}

EXCLUDE_FILES = {
    "omnistudy_project_complete.zip",
    "omnistudy-server.error.log",
    "omnistudy-server.log",
}

EXCLUDE_EXTS = {
    ".pyc",
    ".pyo",
    ".tmp",
    ".log",
}

def should_exclude(rel_path):
    parts = rel_path.split(os.sep)
    for p in parts:
        if p in EXCLUDE_DIRS:
            return True
        if p.startswith(".") and p not in {".git", ".env.example", ".gitignore"}:
            return True

    filename = os.path.basename(rel_path)
    if filename in EXCLUDE_FILES:
        return True
    
    _, ext = os.path.splitext(filename)
    if ext.lower() in EXCLUDE_EXTS:
        return True
        
    return False

def build_zip(zip_dest):
    print(f"Creating zip at: {zip_dest}")
    file_count = 0
    total_bytes = 0

    with zipfile.ZipFile(zip_dest, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, dirs, files in os.walk(SOURCE_DIR):
            # Prune excluded directories in-place to avoid descending
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and (not d.startswith(".") or d == ".git")]

            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, SOURCE_DIR)

                if should_exclude(rel_path):
                    continue

                # Add file to zip with relative path
                zf.write(full_path, arcname=rel_path)
                file_count += 1
                total_bytes += os.path.getsize(full_path)

    print(f"Successfully packaged {file_count} files ({total_bytes / (1024*1024):.2f} MB uncompressed).")
    
    # Integrity Verification
    print("Verifying zip archive CRC32 checksums...")
    with zipfile.ZipFile(zip_dest, "r") as zf:
        bad_file = zf.testzip()
        if bad_file is not None:
            raise RuntimeError(f"Zip verification failed on file: {bad_file}")
        print("CRC32 Verification PASSED: All files are intact with zero corruption!")
        
    compressed_size = os.path.getsize(zip_dest)
    print(f"Compressed archive size: {compressed_size / (1024*1024):.2f} MB")
    return file_count, total_bytes, compressed_size

if __name__ == "__main__":
    count, uncompressed, compressed = build_zip(OUTPUT_ZIP_D)
    
    # Also place a copy inside D:\BCA for easy direct access
    shutil.copy2(OUTPUT_ZIP_D, OUTPUT_ZIP_LOCAL)
    print(f"Copied to: {OUTPUT_ZIP_LOCAL}")
    print("\n[SUCCESS] Migration zip package created and verified successfully!")
