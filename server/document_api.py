from flask import Blueprint, request, jsonify
from datetime import datetime
import json
import base64
import re
import mimetypes
import traceback
import config
from google_services import GoogleServiceProvider

# Inisialisasi Blueprint
doc_bp = Blueprint('document_api', __name__)

# Kita akan menggunakan instance google_provider yang nanti di-pass atau di-import
# Untuk simplifikasi, kita asumsikan google_provider dibuat di app.py dan kita import helpernya
# Tapi cara terbaik di Flask adalah membuat instance baru atau menggunakan 'current_app'
# Di sini kita buat instance baru khusus blueprint ini agar mandiri
provider = GoogleServiceProvider()

# --- HELPERS DARI MAIN.PY LAMA ---
CUSTOM_MIME_MAP = {
    ".dwg": "application/acad",
    ".dxf": "application/dxf",
    ".heic": "image/heic",
}

def guess_mime(filename: str, provided: str = None) -> str:
    if provided:
        return provided
    # Simple logic
    import os
    ext = os.path.splitext(filename.lower())[1]
    if ext in CUSTOM_MIME_MAP:
        return CUSTOM_MIME_MAP[ext]
    return mimetypes.guess_type(filename)[0] or "application/octet-stream"

_DATA_URL_RE = re.compile(r"^data:.*?;base64,", re.IGNORECASE)

def decode_base64_maybe_with_prefix(b64_str: str) -> bytes:
    if not isinstance(b64_str, str):
        raise ValueError("base64 data bukan string")
    cleaned = _DATA_URL_RE.sub("", b64_str.strip())
    return base64.b64decode(cleaned, validate=False)


# --- ROUTES ---

@doc_bp.route('/api/doc/login', methods=['POST'])
def login_doc():
    """Login khusus Penyimpanan Dokumen"""
    data = request.get_json()
    username = data.get("username", "").strip().lower()
    password = data.get("password", "").strip().upper()

    if not username or not password:
        return jsonify({"detail": "Username dan password wajib diisi"}), 400

    try:
        # Akses sheet Cabang via provider
        ws = provider.sheet.worksheet("Cabang")
        records = ws.get_all_records()

        allowed_roles = [
            "BRANCH BUILDING SUPPORT",
            "BRANCH BUILDING COORDINATOR",
            "BRANCH BUILDING & MAINTENANCE MANAGER",
            "BRANCH BUILDING & MAINTENANCE ADMINISTRATOR",
        ]

        for row in records:
            email = str(row.get("EMAIL_SAT", "")).strip().lower()
            jabatan = str(row.get("JABATAN", "")).strip().upper()
            cabang = str(row.get("CABANG", "")).strip().upper()
            nama = str(row.get("NAMA LENGKAP", "")).strip()

            if email == username and password == cabang:
                if jabatan in allowed_roles:
                    return jsonify({
                        "ok": True,
                        "user": {
                            "email": email,
                            "nama": nama,
                            "jabatan": jabatan,
                            "cabang": cabang,
                        },
                    })
                else:
                    return jsonify({"detail": "Jabatan tidak diizinkan"}), 403

        return jsonify({"detail": "Username atau password salah"}), 401

    except Exception as e:
        traceback.print_exc()
        return jsonify({"detail": str(e)}), 500


@doc_bp.route('/api/doc/list', methods=['GET'])
def list_documents():
    cabang = request.args.get('cabang')
    try:
        # Buka spreadsheet penyimpanan
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)
        data = ws.get_all_records()

        # Normalisasi kolom
        normalized_data = []
        for row in data:
            new_row = {}
            for k, v in row.items():
                new_row[k.lower()] = v
            normalized_data.append(new_row)

        if cabang:
            cabang_lower = cabang.strip().lower()
            filtered = [r for r in normalized_data if r.get("cabang", "").strip().lower() == cabang_lower]
        else:
            filtered = normalized_data

        return jsonify({"ok": True, "items": filtered})

    except Exception as e:
        return jsonify({"detail": f"Gagal membaca spreadsheet: {e}"}), 500


@doc_bp.route('/api/doc/save', methods=['POST'])
def save_document_base64():
    try:
        payload = request.get_json()
        
        kode_toko = payload.get("kode_toko")
        nama_toko = payload.get("nama_toko")
        cabang = payload.get("cabang")
        luas_sales = payload.get("luas_sales", "")
        luas_parkir = payload.get("luas_parkir", "")
        luas_gudang = payload.get("luas_gudang", "")
        files = payload.get("files", [])
        email = payload.get("email", "")

        if not all([kode_toko, nama_toko, cabang]):
            return jsonify({"detail": "Data toko belum lengkap."}), 400

        # 1. Buka Sheet
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)

        # 2. Validasi Duplikat
        existing_records = ws.get_all_records()
        for row in existing_records:
            existing_code = str(row.get("kode_toko") or row.get("KodeToko") or "").strip().upper()
            if existing_code == kode_toko.strip().upper():
                return jsonify({"detail": f"Kode toko '{kode_toko}' sudah terdaftar."}), 400

        # 3. Upload ke Drive
        cabang_folder = provider.get_or_create_folder(cabang, config.DOC_DRIVE_ROOT_ID)
        toko_folder_name = f"{kode_toko}_{nama_toko}".replace("/", "-")
        toko_folder = provider.get_or_create_folder(toko_folder_name, cabang_folder)

        category_folders = {}
        file_links = []
        
        for idx, f in enumerate(files, start=1):
            category = (f.get("category") or "lainnya").strip() or "lainnya"
            if category not in category_folders:
                category_folders[category] = provider.get_or_create_folder(category, toko_folder)
            
            filename = f.get("filename") or f"file_{idx}"
            mime_type = guess_mime(filename, f.get("type"))
            
            try:
                raw = decode_base64_maybe_with_prefix(f.get("data") or "")
                uploaded = provider.upload_file_simple(
                    folder_id=category_folders[category],
                    filename=filename,
                    mime_type=mime_type,
                    raw_bytes=raw
                )
                
                link = uploaded.get("webViewLink")
                thumb = uploaded.get("thumbnailLink")
                
                direct_link = ""
                if link:
                    fid = link.split("/d/")[-1].split("/")[0]
                    direct_link = f"https://drive.google.com/uc?export=view&id={fid}"
                elif thumb:
                    direct_link = thumb
                
                if direct_link:
                    file_links.append(f"{category}|{filename}|{direct_link}")

            except Exception as e:
                print(f"Gagal upload {filename}: {e}")

        # 4. Simpan ke Sheet
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([
            kode_toko, nama_toko, cabang, luas_sales, luas_parkir, luas_gudang,
            f"https://drive.google.com/drive/folders/{toko_folder}",
            ", ".join(file_links),
            now,
            email  # last_edit
        ])

        return jsonify({
            "ok": True,
            "message": f"{len(file_links)} file berhasil diunggah",
            "folder_link": f"https://drive.google.com/drive/folders/{toko_folder}",
            "last_edit": email
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"detail": f"Gagal menyimpan: {e}"}), 500


@doc_bp.route('/api/doc/update/<kode_toko>', methods=['PUT'])
def update_document(kode_toko):
    try:
        data = request.get_json()
        files = data.get("files", [])
        email = data.get("email", "")

        # Buka Sheet
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()

        # Cari baris
        row_index = next((i + 2 for i, r in enumerate(records) 
                          if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not row_index:
            return jsonify({"detail": "Data tidak ditemukan"}), 404

        # Ambil data lama
        old_data = records[row_index - 2]
        old_folder_link = old_data.get("folder_link")
        if not old_folder_link or "folders/" not in old_folder_link:
            return jsonify({"detail": "Folder Drive toko tidak valid"}), 400
        
        toko_folder_id = old_folder_link.split("folders/")[-1]

        # Logic Hapus File Lama (Simplified for Flask)
        # Ambil file lama dari links
        old_file_links_str = old_data.get("file_links", "")
        old_files_list = []
        if old_file_links_str:
            for entry in old_file_links_str.split(","):
                parts = [p.strip() for p in entry.split("|")]
                if len(parts) >= 3:
                    old_files_list.append({"category": parts[0], "filename": parts[1], "link": parts[2]})

        # ==============================================================================
        # MULAI IMPLEMENTASI DELETE LOGIC (EXPLICIT DELETE FLAG)
        # ==============================================================================
        
        # Ambil daftar folder kategori yang ada di Drive
        category_folders_map = {}
        existing_files_drive = []
        try:
            subfolders = provider.doc_drive_service.files().list(
                q=f"'{toko_folder_id}' in parents and trashed = false and mimeType = 'application/vnd.google-apps.folder'",
                fields="files(id, name)"
            ).execute().get("files", [])
            
            category_folders_map = {sf["name"]: sf["id"] for sf in subfolders}

            # Ambil semua file fisik yang ada di dalam folder-folder tersebut
            for cat_name, folder_id in category_folders_map.items():
                res_files = provider.doc_drive_service.files().list(
                    q=f"'{folder_id}' in parents and trashed = false",
                    fields="files(id, name)"
                ).execute()
                for f in res_files.get("files", []):
                    existing_files_drive.append({
                        "id": f["id"],
                        "name": f["name"],
                        "category": cat_name
                    })
        except Exception as e:
            print(f"⚠️ Error membaca folder Drive: {e}")

        # Pisahkan file berdasarkan tipe:
        # 1. files_to_delete: file dengan flag deleted=true
        # 2. files_to_upload: file baru dengan data base64
        # 3. files_to_keep: file lama tanpa perubahan (tidak ada data, tidak deleted)
        
        files_to_delete = []
        files_to_upload = []
        files_to_keep_keys = set()  # (category, filename) yang harus dipertahankan
        
        for f in files:
            category = (f.get("category") or "pendukung").strip()
            filename = f.get("filename")
            
            if f.get("deleted") == True:
                # File ditandai untuk dihapus
                files_to_delete.append({"category": category, "filename": filename})
            elif f.get("data"):
                # File baru dengan data base64
                files_to_upload.append(f)
                files_to_keep_keys.add((category, filename))
            else:
                # File lama yang harus dipertahankan
                files_to_keep_keys.add((category, filename))

        # Eksekusi hapus file yang ditandai deleted=true
        for del_file in files_to_delete:
            # Cari file di Drive berdasarkan category dan filename
            drive_file = next(
                (f for f in existing_files_drive 
                 if f["category"] == del_file["category"] and f["name"] == del_file["filename"]), 
                None
            )
            if drive_file:
                try:
                    provider.doc_drive_service.files().delete(fileId=drive_file["id"]).execute()
                    print(f"🗑️ Hapus file di Drive: {drive_file['name']} (Kategori: {drive_file['category']})")
                except Exception as del_err:
                    print(f"⚠️ Gagal hapus file {drive_file['name']}: {del_err}")
        
        # ==============================================================================
        # SELESAI IMPLEMENTASI DELETE LOGIC
        # ==============================================================================

        # Logic Upload Baru & Pertahankan File Lama
        new_file_links = []
        category_cache = {}

        # 1. Pertahankan semua file lama yang TIDAK dihapus
        deleted_keys = set((f["category"], f["filename"]) for f in files_to_delete)
        for old_file in old_files_list:
            key = (old_file["category"], old_file["filename"])
            # Pertahankan jika tidak di-delete DAN tidak di-replace dengan file baru
            # (jika ada file baru dengan nama sama, jangan duplikat)
            new_upload_filenames = set((f.get("category", "pendukung").strip(), f.get("filename")) for f in files_to_upload)
            if key not in deleted_keys and key not in new_upload_filenames:
                new_file_links.append(f"{old_file['category']}|{old_file['filename']}|{old_file['link']}")

        # 2. Upload file baru
        for f in files_to_upload:
            category = (f.get("category") or "pendukung").strip()
            filename = f.get("filename")
            
            # Pastikan folder kategori ada
            if category not in category_cache:
                category_cache[category] = provider.get_or_create_folder(category, toko_folder_id)
            
            # Hapus file lama dengan nama sama jika ada (replace)
            old_drive_file = next(
                (df for df in existing_files_drive if df["category"] == category and df["name"] == filename),
                None
            )
            if old_drive_file:
                try:
                    provider.doc_drive_service.files().delete(fileId=old_drive_file["id"]).execute()
                    print(f"🔄 Replace file di Drive: {filename}")
                except Exception as e:
                    print(f"⚠️ Gagal hapus file lama untuk replace: {e}")
            
            raw = decode_base64_maybe_with_prefix(f["data"])
            mime = guess_mime(filename, f.get("type"))
            
            uploaded = provider.upload_file_simple(category_cache[category], filename, mime, raw)
            
            # Buat Link
            link = uploaded.get("webViewLink")
            fid = link.split("/d/")[-1].split("/")[0] if link else ""
            direct = f"https://drive.google.com/uc?export=view&id={fid}" if fid else ""
            
            new_file_links.append(f"{category}|{filename}|{direct}")

        # Update Sheet
        # Hati-hati urutan kolom harus sama persis dengan sheet
        cell_range = f"A{row_index}:J{row_index}"
        ws.update(cell_range, [[
            old_data.get("kode_toko"),
            old_data.get("nama_toko"),
            old_data.get("cabang"),
            data.get("luas_sales", old_data.get("luas_sales")), # Update luas jika ada
            data.get("luas_parkir", old_data.get("luas_parkir")),
            data.get("luas_gudang", old_data.get("luas_gudang")),
            old_folder_link,
            ", ".join(new_file_links),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            email  # last_edit
        ]])

        return jsonify({"ok": True, "message": "Berhasil update", "last_edit": email})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"detail": str(e)}), 500


@doc_bp.route('/api/doc/delete/<kode_toko>', methods=['DELETE'])
def delete_document(kode_toko):
    try:
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()
        
        row_index = next((i + 2 for i, r in enumerate(records) if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not row_index:
            return jsonify({"detail": "Data tidak ditemukan"}), 404
            
        # Hapus folder di Drive
        folder_link = records[row_index - 2].get("folder_link")
        if folder_link and "folders/" in folder_link:
            folder_id = folder_link.split("folders/")[-1]
            provider.delete_drive_file(folder_id)

        ws.delete_rows(row_index)
        return jsonify({"ok": True, "message": "Dokumen dihapus"})
    except Exception as e:
        return jsonify({"detail": str(e)}), 500

@doc_bp.route('/api/doc/detail/<kode_toko>', methods=['GET'])
def get_document_detail(kode_toko):
    try:
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()
        
        found = next((r for r in records if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not found:
            return jsonify({"detail": "Data tidak ditemukan"}), 404
            
        return jsonify({"ok": True, "data": found})
    except Exception as e:
        return jsonify({"detail": str(e)}), 500