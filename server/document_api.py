from flask import Blueprint, request, jsonify
from datetime import datetime
import pytz
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


# --- LOGGING HELPER ---
def log_doc(func: str, message: str, **kwargs):
    ts = datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d %H:%M:%S")
    extra = " | ".join(f"{k}={v}" for k, v in kwargs.items()) if kwargs else ""
    print(f"[DOC][{func}] {ts} - {message}{' | ' + extra if extra else ''}")


# --- ROUTES ---

@doc_bp.route('/api/doc/login', methods=['POST'])
def login_doc():
    """Login khusus Penyimpanan Dokumen"""
    data = request.get_json()
    username = data.get("username", "").strip().lower()
    password = data.get("password", "").strip().upper()

    log_doc("login_doc", "request received", username=username)

    if not username or not password:
        log_doc("login_doc", "missing credentials", has_username=bool(username), has_password=bool(password))
        return jsonify({"detail": "Username dan password wajib diisi"}), 400

    try:
        provider = GoogleServiceProvider()
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
                    log_doc("login_doc", "authenticated", email=email, cabang=cabang, jabatan=jabatan)
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
                    log_doc("login_doc", "forbidden role", email=email, cabang=cabang, jabatan=jabatan)
                    return jsonify({"detail": "Jabatan tidak diizinkan"}), 403

        log_doc("login_doc", "invalid credentials", username=username)
        return jsonify({"detail": "Username atau password salah"}), 401

    except Exception as e:
        log_doc("login_doc", "error", error=str(e))
        traceback.print_exc()
        return jsonify({"detail": str(e)}), 500


@doc_bp.route('/api/doc/list', methods=['GET'])
def list_documents():
    cabang = request.args.get('cabang')
    try:
        provider = GoogleServiceProvider()
        # Buka spreadsheet penyimpanan
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)
        data = ws.get_all_records()

        log_doc("list_documents", "fetched records", total=len(data), cabang_filter=cabang or "-")

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

        log_doc("list_documents", "returning items", count=len(filtered))
        return jsonify({"ok": True, "items": filtered})

    except Exception as e:
        log_doc("list_documents", "error", error=str(e))
        return jsonify({"detail": f"Gagal membaca spreadsheet: {e}"}), 500


@doc_bp.route('/api/doc/save', methods=['POST'])
def save_document_base64():
    try:
        provider = GoogleServiceProvider()
        payload = request.get_json()
        
        kode_toko = payload.get("kode_toko")
        nama_toko = payload.get("nama_toko")
        cabang = payload.get("cabang")
        luas_sales = payload.get("luas_sales", "")
        luas_parkir = payload.get("luas_parkir", "")
        luas_gudang = payload.get("luas_gudang", "")
        luas_bangunan_lantai_1 = payload.get("luas_bangunan_lantai_1", "")
        luas_bangunan_lantai_2 = payload.get("luas_bangunan_lantai_2", "")
        luas_bangunan_lantai_3 = payload.get("luas_bangunan_lantai_3", "")
        total_luas_bangunan = payload.get("total_luas_bangunan", "")
        luas_area_terbuka = payload.get("luas_area_terbuka", "")
        tinggi_plafon = payload.get("tinggi_plafon", "")
        files = payload.get("files", [])
        email = payload.get("email", "")

        log_doc(
            "save_document_base64",
            "request received",
            kode_toko=kode_toko,
            nama_toko=nama_toko,
            cabang=cabang,
            files=len(files),
            email=email,
        )

        if not all([kode_toko, nama_toko, cabang]):
            log_doc("save_document_base64", "missing required fields")
            return jsonify({"detail": "Data toko belum lengkap."}), 400

        # 1. Buka Sheet
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)

        # 2. Validasi Duplikat
        existing_records = ws.get_all_records()
        log_doc("save_document_base64", "existing records fetched", count=len(existing_records))
        for row in existing_records:
            existing_code = str(row.get("kode_toko") or row.get("KodeToko") or "").strip().upper()
            if existing_code == kode_toko.strip().upper():
                log_doc("save_document_base64", "duplicate kode_toko", kode_toko=kode_toko)
                return jsonify({"detail": f"Kode toko '{kode_toko}' sudah terdaftar."}), 400

        # 3. Upload ke Drive
        cabang_folder = provider.get_or_create_folder(cabang, config.DOC_DRIVE_ROOT_ID)
        toko_folder_name = f"{kode_toko}_{nama_toko}".replace("/", "-")
        toko_folder = provider.get_or_create_folder(toko_folder_name, cabang_folder)

        log_doc("save_document_base64", "folders prepared", cabang_folder=cabang_folder, toko_folder=toko_folder)

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
                log_doc("save_document_base64", "upload failed", filename=filename, category=category, error=str(e))

        # 4. Simpan ke Sheet
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        now = datetime.now(jakarta_tz).strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([
            kode_toko, nama_toko, cabang, luas_sales, luas_parkir, luas_gudang, luas_bangunan_lantai_1, luas_bangunan_lantai_2, luas_bangunan_lantai_3, total_luas_bangunan, luas_area_terbuka, tinggi_plafon,
            f"https://drive.google.com/drive/folders/{toko_folder}",
            ", ".join(file_links),
            now,
            email  # last_edit
        ])

        log_doc(
            "save_document_base64",
            "saved",
            uploaded=len(file_links),
            folder_link=f"https://drive.google.com/drive/folders/{toko_folder}",
            last_edit=email,
        )

        return jsonify({
            "ok": True,
            "message": f"{len(file_links)} file berhasil diunggah",
            "folder_link": f"https://drive.google.com/drive/folders/{toko_folder}",
            "last_edit": email
        })

    except Exception as e:
        log_doc("save_document_base64", "error", error=str(e))
        traceback.print_exc()
        return jsonify({"detail": f"Gagal menyimpan: {e}"}), 500


@doc_bp.route('/api/doc/update/<kode_toko>', methods=['PUT'])
def update_document(kode_toko):
    try:
        provider = GoogleServiceProvider()
        data = request.get_json()
        files = data.get("files", [])
        email = data.get("email", "")

        log_doc(
            "update_document",
            "request received",
            kode_toko=kode_toko,
            files=len(files),
            email=email,
        )
        
        # Capture timestamp saat update dimulai (Jakarta timezone)
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        update_timestamp = datetime.now(jakarta_tz).strftime("%Y-%m-%d %H:%M:%S")

        # Buka Sheet
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = provider.doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()

        log_doc("update_document", "records fetched", total=len(records))

        # Cari baris
        row_index = next((i + 2 for i, r in enumerate(records) 
                          if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not row_index:
            log_doc("update_document", "row not found", kode_toko=kode_toko)
            return jsonify({"detail": "Data tidak ditemukan"}), 404

        # Ambil data lama
        old_data = records[row_index - 2]
        old_folder_link = old_data.get("folder_link")
        if not old_folder_link or "folders/" not in old_folder_link:
            log_doc("update_document", "invalid drive folder", folder_link=old_folder_link)
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
            log_doc("update_document", "error reading drive folders", error=str(e))

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

        log_doc(
            "update_document",
            "files classified",
            to_delete=len(files_to_delete),
            to_upload=len(files_to_upload),
            to_keep=len(files_to_keep_keys),
        )

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
                    log_doc(
                        "update_document",
                        "deleted from drive",
                        filename=drive_file['name'],
                        category=drive_file['category'],
                    )
                except Exception as del_err:
                    log_doc(
                        "update_document",
                        "delete failed",
                        filename=drive_file['name'],
                        category=drive_file['category'],
                        error=str(del_err),
                    )
        
       
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
                    log_doc("update_document", "replace existing file", filename=filename, category=category)
                except Exception as e:
                    log_doc("update_document", "replace delete failed", filename=filename, category=category, error=str(e))
            
            raw = decode_base64_maybe_with_prefix(f["data"])
            mime = guess_mime(filename, f.get("type"))
            
            uploaded = provider.upload_file_simple(category_cache[category], filename, mime, raw)
            
            # Buat Link
            link = uploaded.get("webViewLink")
            fid = link.split("/d/")[-1].split("/")[0] if link else ""
            direct = f"https://drive.google.com/uc?export=view&id={fid}" if fid else ""
            
            new_file_links.append(f"{category}|{filename}|{direct}")

        log_doc(
            "update_document",
            "upload completed",
            uploaded=len(files_to_upload),
            kept=len(new_file_links) - len(files_to_upload),
        )

        # Update Sheet
        # Hati-hati urutan kolom harus sama persis dengan sheet
        cell_range = f"A{row_index}:P{row_index}"
        ws.update(cell_range, [[
            old_data.get("kode_toko"),
            old_data.get("nama_toko"),
            old_data.get("cabang"),
            data.get("luas_sales", old_data.get("luas_sales")), # Update luas jika ada
            data.get("luas_parkir", old_data.get("luas_parkir")),
            data.get("luas_gudang", old_data.get("luas_gudang")),
            data.get("luas_bangunan_lantai_1", old_data.get("luas_bangunan_lantai_1")),
            data.get("luas_bangunan_lantai_2", old_data.get("luas_bangunan_lantai_2")),
            data.get("luas_bangunan_lantai_3", old_data.get("luas_bangunan_lantai_3")),
            data.get("total_luas_bangunan", old_data.get("total_luas_bangunan")),
            data.get("luas_area_terbuka", old_data.get("luas_area_terbuka")),
            data.get("tinggi_plafon", old_data.get("tinggi_plafon")),
            old_folder_link,
            ", ".join(new_file_links),
            update_timestamp,  # timestamp diupdate saat ada perubahan
            email  # last_edit
        ]])

        log_doc(
            "update_document",
            "sheet updated",
            kode_toko=kode_toko,
            row=row_index,
            updated_at=update_timestamp,
            last_edit=email,
        )

        return jsonify({"ok": True, "message": "Berhasil update", "last_edit": email, "updated_at": update_timestamp})

    except Exception as e:
        log_doc("update_document", "error", error=str(e))
        traceback.print_exc()
        return jsonify({"detail": str(e)}), 500


@doc_bp.route('/api/doc/delete/<kode_toko>', methods=['DELETE'])
def delete_document(kode_toko):
    try:
        provider = GoogleServiceProvider()
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()

        log_doc("delete_document", "request received", kode_toko=kode_toko, total_records=len(records))
        
        row_index = next((i + 2 for i, r in enumerate(records) if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not row_index:
            log_doc("delete_document", "row not found", kode_toko=kode_toko)
            return jsonify({"detail": "Data tidak ditemukan"}), 404
            
        # Hapus folder di Drive
        folder_link = records[row_index - 2].get("folder_link")
        if folder_link and "folders/" in folder_link:
            folder_id = folder_link.split("folders/")[-1]
            provider.delete_drive_file(folder_id)
            log_doc("delete_document", "drive folder deleted", folder_id=folder_id)

        ws.delete_rows(row_index)
        log_doc("delete_document", "row deleted", row=row_index)
        return jsonify({"ok": True, "message": "Dokumen dihapus"})
    except Exception as e:
        log_doc("delete_document", "error", error=str(e))
        return jsonify({"detail": str(e)}), 500

@doc_bp.route('/api/doc/detail/<kode_toko>', methods=['GET'])
def get_document_detail(kode_toko):
    try:
        provider = GoogleServiceProvider()
        doc_sheet = provider.gspread_client.open_by_key(config.SPREADSHEET_ID)
        ws = doc_sheet.worksheet(config.DOC_SHEET_NAME)
        records = ws.get_all_records()

        log_doc("get_document_detail", "records fetched", total=len(records), kode_toko=kode_toko)
        
        found = next((r for r in records if str(r.get("kode_toko", "")).strip() == str(kode_toko).strip()), None)
        
        if not found:
            log_doc("get_document_detail", "not found", kode_toko=kode_toko)
            return jsonify({"detail": "Data tidak ditemukan"}), 404
        log_doc("get_document_detail", "found", kode_toko=kode_toko)
        return jsonify({"ok": True, "data": found})
    except Exception as e:
        log_doc("get_document_detail", "error", error=str(e))
        return jsonify({"detail": str(e)}), 500