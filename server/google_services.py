import os.path
import io
import re
import gspread
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timezone, timedelta

import config

class GoogleServiceProvider:
    def __init__(self):
        self.scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/gmail.send',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/calendar'
        ]
        self.creds = None
        
        secret_dir = '/etc/secrets/'
        token_path = os.path.join(secret_dir, 'token.json')

        if not os.path.exists(secret_dir):
            token_path = 'token.json'

        if os.path.exists(token_path):
            self.creds = Credentials.from_authorized_user_file(token_path, self.scopes)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                raise Exception("CRITICAL: token.json not found or invalid. Please re-authenticate locally using generate_token.py.")

        self.gspread_client = gspread.authorize(self.creds)
        self.sheet = self.gspread_client.open_by_key(config.SPREADSHEET_ID)
        self.data_entry_sheet = self.sheet.worksheet(config.DATA_ENTRY_SHEET_NAME)
        self.gmail_service = build('gmail', 'v1', credentials=self.creds)
        self.drive_service = build('drive', 'v3', credentials=self.creds)
        self.calendar_service = build('calendar', 'v3', credentials=self.creds)

    def get_cabang_code(self, cabang_name):
        branch_to_ulok_map = {
            "WHC IMAM BONJOL": "7AZ1", "LUWU": "2VZ1", "KARAWANG": "1JZ1", "REMBANG": "2AZ1",
            "BANJARMASIN": "1GZ1", "PARUNG": "1MZ1", "TEGAL": "2PZ1", "GORONTALO": "2SZ1",
            "PONTIANAK": "1PZ1", "LOMBOK": "1SZ1", "KOTABUMI": "1VZ1", "SERANG": "2GZ1",
            "CIANJUR": "2JZ1", "BALARAJA": "TZ01", "SIDOARJO": "UZ01", "MEDAN": "WZ01",
            "BOGOR": "XZ01", "JEMBER": "YZ01", "BALI": "QZ01", "PALEMBANG": "PZ01",
            "KLATEN": "OZ01", "MAKASSAR": "RZ01", "PLUMBON": "VZ01", "PEKANBARU": "1AZ1",
            "JAMBI": "1DZ1", "HEAD OFFICE": "Z001", "BANDUNG 1": "BZ01", "BANDUNG 2": "NZ01",
            "BEKASI": "CZ01", "CILACAP": "IZ01", "CILEUNGSI": "JZ01", "SEMARANG": "HZ01",
            "CIKOKOL": "KZ01", "LAMPUNG": "LZ01", "MALANG": "MZ01", "MANADO": "1YZ1",
            "BATAM": "2DZ1", "MADIUN": "2MZ1"
        }
        return branch_to_ulok_map.get(cabang_name.upper(), cabang_name)

    def get_next_spk_sequence(self, cabang, year, month):
        try:
            spk_sheet = self.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
            all_records = spk_sheet.get_all_records()
            count = 0
            for record in all_records:
                if str(record.get('Cabang', '')).strip().lower() == cabang.lower():
                    timestamp_str = record.get('Timestamp', '')
                    if timestamp_str:
                        try:
                            record_time = datetime.fromisoformat(timestamp_str)
                            if record_time.year == year and record_time.month == month:
                                count += 1
                        except ValueError:
                            continue
            return count + 1
        except Exception as e:
            print(f"Error getting SPK sequence: {e}")
            return 1
            
    def append_to_dynamic_sheet(self, spreadsheet_id, sheet_name, data_dict):
        try:
            spreadsheet = self.gspread_client.open_by_key(spreadsheet_id)
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
            
            headers = worksheet.row_values(1)
            if not headers: 
                headers = list(data_dict.keys())
                worksheet.append_row(headers)

            row_data = [data_dict.get(header, "") for header in headers]
            
            # Gunakan append_row dengan value_input_option dan dapatkan response
            # untuk mendapatkan row index yang akurat (menghindari race condition)
            # insert_data_option='INSERT_ROWS' memastikan data ditambahkan sebagai baris baru
            result = worksheet.append_row(
                row_data, 
                value_input_option='USER_ENTERED',
                insert_data_option='INSERT_ROWS'
            )
            
            # Parse response untuk mendapatkan row number yang sebenarnya
            try:
                updated_range = result.get('updates', {}).get('updatedRange', '')
                if updated_range:
                    range_part = updated_range.split('!')[-1]
                    match = re.search(r'[A-Z]+(\d+)', range_part)
                    if match:
                        return int(match.group(1))
            except Exception as e:
                print(f"Warning: Could not parse row index from append response: {e}")
            
            # Fallback ke method lama jika parsing gagal
            return len(worksheet.get_all_values())
        except Exception as e:
            print(f"Error saat menyimpan ke sheet '{sheet_name}': {e}")
            raise

    def get_rab_url_by_ulok(self, kode_ulok):
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            all_records = worksheet.get_all_records()
            for record in reversed(all_records):
                if str(record.get('Nomor Ulok', '')).strip().upper() == kode_ulok.strip().upper():
                    return record.get('Link PDF Non-SBO') or record.get('Link PDF')
            return None
        except Exception as e:
            print(f"Error saat mencari RAB URL: {e}")
            return None

    def get_rab_url_by_ulok_kedua(self, kode_ulok):
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME_RAB_2)
            all_records = worksheet.get_all_records()
            for record in reversed(all_records):
                if str(record.get('Nomor Ulok', '')).strip().upper() == kode_ulok.strip().upper():
                    return record.get('Link PDF Non-SBO') or record.get('Link PDF')
            return None
        except Exception as e:
            print(f"Error saat mencari RAB URL: {e}")
            return None

    def get_spk_url_by_ulok(self, kode_ulok):
        try:
            spk_sheet = self.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
            all_records = spk_sheet.get_all_records()
            for record in reversed(all_records):
                if str(record.get('Nomor Ulok', '')).strip().upper() == kode_ulok.strip().upper() and \
                   str(record.get('Status', '')).strip() == config.STATUS.SPK_APPROVED:
                    return record.get('Link PDF')
            return None
        except Exception as e:
            print(f"Error saat mencari SPK URL: {e}")
            return None

    def get_spk_data_by_cabang(self, cabang):
        spk_list = []
        try:
            worksheet = self.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
            records = worksheet.get_all_records()
            for record in records:
                if str(record.get('Cabang', '')).strip().lower() == cabang.strip().lower() and \
                   str(record.get('Status', '')).strip() == config.STATUS.SPK_APPROVED:
                    spk_list.append({
                        "Nomor Ulok": record.get("Nomor Ulok"),
                        "Lingkup Pekerjaan": record.get("Lingkup Pekerjaan"),
                        "Link PDF": record.get("Link PDF")
                    })
            unique_spk = {item['Nomor Ulok'] + item['Lingkup Pekerjaan']: item for item in reversed(spk_list)}
            return list(reversed(list(unique_spk.values())))
        except Exception as e:
            print(f"Error saat mengambil data SPK by cabang: {e}")
            return []

    # Mengambil semua data rab untuk dropdown (HO)

    def get_all_rab_ulok(self):
        """
        Mengambil daftar unik Nomor Ulok beserta Proyek, Nama Toko, dan Lingkup Pekerjaan 
        dari sheet APPROVED_DATA_SHEET_NAME untuk dropdown (bukan dari RAB/Form2).
        """
        try:
            # UBAH DISINI: Gunakan worksheet yang mengarah ke APPROVED_DATA_SHEET_NAME
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            records = worksheet.get_all_records()

            ulok_list = []
            seen = set()

            for record in records:
                # Ambil Nomor Ulok (Pastikan nama kolom di Sheet SPK sesuai, biasanya "Nomor Ulok")
                ulok = str(record.get('Nomor Ulok', '')).strip()
                
                # Jika Nomor Ulok kosong, skip
                if not ulok:
                    continue

                # Ambil field lain dari Sheet SPK
                # Note: Di Sheet SPK biasanya header "Lingkup Pekerjaan" pakai spasi, bukan underscore
                nama_toko = str(record.get('Nama_Toko', record.get('nama_toko', ''))).strip()
                lingkup = str(record.get('Lingkup Pekerjaan', record.get('Lingkup_Pekerjaan', ''))).strip()
                proyek = str(record.get('Proyek', '')).strip()

                # Buat kunci unik kombinasi Ulok + Lingkup
                unique_key = f"{ulok}-{lingkup}"

                if unique_key not in seen:
                    seen.add(unique_key)
                    
                    # Format Label: "Z001... - Reguler (ME) - Nama Toko"
                    label = f"{ulok} - {proyek} ({lingkup}) - {nama_toko}"

                    ulok_list.append({
                        "label": label,          # Teks yang tampil di dropdown
                        "value": ulok + "-" + lingkup,           # Nilai yang akan dikirim saat submit
                        
                    })

            # Sortir berdasarkan Nomor Ulok
            ulok_list.sort(key=lambda x: x['value'])
            
            return ulok_list

        except Exception as e:
            print(f"Error saat mengambil daftar data dari Sheet SPK: {e}")
            return [] # Kembalikan list kosong agar tidak error 500

    # Mengambil salah satu data rab berdasarkan nomor ulok
    def get_gantt_data_by_ulok(self, kode_ulok, lingkup_pekerjaan):
        try:
            target_ulok = str(kode_ulok).strip().upper()
            target_lingkup = str(lingkup_pekerjaan).strip()
            
            # --- AMBIL DATA RAB FORM 2 ---
            rab_sheet = self.sheet.worksheet(config.DATA_ENTRY_SHEET_NAME) # Form2
            all_rab_values = rab_sheet.get_all_values()
            filtered_categories = []
            rab_data = None
            
            # Field yang akan diambil dari RAB
            rab_fields = [
                config.COLUMN_NAMES.STATUS,
                config.COLUMN_NAMES.TIMESTAMP,
                config.COLUMN_NAMES.LINK_PDF_NONSBO,
                config.COLUMN_NAMES.KOORDINATOR_APPROVER,
                config.COLUMN_NAMES.KOORDINATOR_APPROVAL_TIME,
                config.COLUMN_NAMES.MANAGER_APPROVER,
                config.COLUMN_NAMES.MANAGER_APPROVAL_TIME,
                config.COLUMN_NAMES.EMAIL_PEMBUAT,
                config.COLUMN_NAMES.LOKASI,
                config.COLUMN_NAMES.PROYEK,
                config.COLUMN_NAMES.ALAMAT,
                config.COLUMN_NAMES.CABANG,
                config.COLUMN_NAMES.LINGKUP_PEKERJAAN,
                config.COLUMN_NAMES.NAMA_TOKO
            ]
            
            if all_rab_values:
                headers = all_rab_values[0]
                data_rows = all_rab_values[1:]
                
                # Cari index kolom penting
                try:
                    ulok_idx = headers.index(config.COLUMN_NAMES.LOKASI)
                    lingkup_idx = headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN) # Biasanya "Lingkup_Pekerjaan"
                except ValueError:
                    print("Kolom Nomor Ulok atau Lingkup Pekerjaan tidak ditemukan di RAB Form 3")
                    ulok_idx = -1
                    lingkup_idx = -1

                if ulok_idx != -1 and lingkup_idx != -1:
                    matched_row_dict = None
                    
                    for row_vals in reversed(data_rows):
                        # Safety check length
                        if len(row_vals) <= max(ulok_idx, lingkup_idx):
                            continue

                        current_ulok = str(row_vals[ulok_idx]).strip().upper()
                        current_lingkup = str(row_vals[lingkup_idx]).strip()
                        
                        # Filter berdasarkan Nomor Ulok DAN Lingkup
                        if current_ulok == target_ulok and current_lingkup.lower() == target_lingkup.lower():
                            # Mapping ke dictionary untuk memudahkan pengecekan kategori
                            if len(row_vals) < len(headers):
                                row_vals += [''] * (len(headers) - len(row_vals))
                            matched_row_dict = dict(zip(headers, row_vals))
                            break
                    
                    # --- PROSES DATA RAB DAN KATEGORI PEKERJAAN (Hanya jika data ditemukan) ---
                    if matched_row_dict:
                        # Ambil hanya field yang diperlukan untuk rab_data
                        rab_data = {}
                        for field in rab_fields:
                            rab_data[field] = matched_row_dict.get(field, "")
                        
                        # Tentukan master list berdasarkan input parameter lingkup
                        if "SIPIL" in target_lingkup.upper():
                            master_list = config.KATEGORI_SIPIL
                        else:
                            master_list = config.KATEGORI_ME
                        
                        found_cats_set = set()

                        # Cek semua kolom yang berawalan "Kategori_Pekerjaan_" di baris yang ditemukan
                        for key, val in matched_row_dict.items():
                            if key.startswith("Kategori_Pekerjaan_"):
                                val_clean = str(val).strip().upper()
                                # Cek apakah nilai ada di master list
                                if val_clean in master_list:
                                    found_cats_set.add(val_clean)
                        
                        # Masukkan ke array hasil urut berdasarkan urutan Master List
                        for cat in master_list:
                            if cat in found_cats_set:
                                filtered_categories.append(cat)

            # --- AMBIL DATA GANTT CHART (Jika ada) ---
            gantt_data = None
            try:
                gantt_sheet = self.sheet.worksheet(config.GANTT_CHART_SHEET_NAME)
                all_gantt_values = gantt_sheet.get_all_values()
                
                if all_gantt_values:
                    gantt_headers = all_gantt_values[0]
                    gantt_data_rows = all_gantt_values[1:]
                    
                    # Cari index kolom Nomor Ulok dan Lingkup_Pekerjaan di sheet gantt
                    try:
                        gantt_ulok_idx = gantt_headers.index(config.COLUMN_NAMES.LOKASI)
                        gantt_lingkup_idx = gantt_headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
                    except ValueError:
                        print("Kolom Nomor Ulok atau Lingkup Pekerjaan tidak ditemukan di Gantt Chart sheet")
                        gantt_ulok_idx = -1
                        gantt_lingkup_idx = -1
                    
                    if gantt_ulok_idx != -1 and gantt_lingkup_idx != -1:
                        # Cari row yang cocok (dari bawah untuk mendapatkan data terbaru)
                        for row_vals in reversed(gantt_data_rows):
                            if len(row_vals) <= max(gantt_ulok_idx, gantt_lingkup_idx):
                                continue
                            
                            current_ulok = str(row_vals[gantt_ulok_idx]).strip().upper()
                            current_lingkup = str(row_vals[gantt_lingkup_idx]).strip()
                            
                            if current_ulok == target_ulok and current_lingkup.lower() == target_lingkup.lower():
                                # Mapping ke dictionary
                                if len(row_vals) < len(gantt_headers):
                                    row_vals += [''] * (len(gantt_headers) - len(row_vals))
                                gantt_data = dict(zip(gantt_headers, row_vals))
                                break
            except gspread.exceptions.WorksheetNotFound:
                print(f"Worksheet '{config.GANTT_CHART_SHEET_NAME}' tidak ditemukan")
            except Exception as gantt_error:
                print(f"Error getting gantt chart data: {gantt_error}")

            # --- AMBIL DATA DAY GANTT CHART (Jika ada) ---
            day_gantt_data = []
            try:
                day_gantt_sheet = self.sheet.worksheet(config.DAY_GANTT_CHART_SHEET_NAME)
                all_day_gantt_values = day_gantt_sheet.get_all_values()
                
                if all_day_gantt_values:
                    day_gantt_headers = all_day_gantt_values[0]
                    day_gantt_data_rows = all_day_gantt_values[1:]
                    
                    # Cari index kolom Nomor Ulok dan Lingkup_Pekerjaan di sheet day_gantt
                    try:
                        day_gantt_ulok_idx = day_gantt_headers.index(config.COLUMN_NAMES.LOKASI)
                        day_gantt_lingkup_idx = day_gantt_headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
                    except ValueError:
                        print("Kolom Nomor Ulok atau Lingkup Pekerjaan tidak ditemukan di Day Gantt Chart sheet")
                        day_gantt_ulok_idx = -1
                        day_gantt_lingkup_idx = -1
                    
                    if day_gantt_ulok_idx != -1 and day_gantt_lingkup_idx != -1:
                        # Cari semua row yang cocok (satu ulok+lingkup bisa punya banyak kategori/range)
                        for row_vals in day_gantt_data_rows:
                            if len(row_vals) <= max(day_gantt_ulok_idx, day_gantt_lingkup_idx):
                                continue
                            
                            current_ulok = str(row_vals[day_gantt_ulok_idx]).strip().upper()
                            current_lingkup = str(row_vals[day_gantt_lingkup_idx]).strip()
                            
                            if current_ulok == target_ulok and current_lingkup.lower() == target_lingkup.lower():
                                # Mapping ke dictionary
                                if len(row_vals) < len(day_gantt_headers):
                                    row_vals += [''] * (len(day_gantt_headers) - len(row_vals))
                                day_gantt_data.append(dict(zip(day_gantt_headers, row_vals)))
            except gspread.exceptions.WorksheetNotFound:
                print(f"Worksheet '{config.DAY_GANTT_CHART_SHEET_NAME}' tidak ditemukan")
            except Exception as day_gantt_error:
                print(f"Error getting day gantt chart data: {day_gantt_error}")

            return {
                "rab": rab_data,
                "filtered_categories": filtered_categories,
                "gantt": gantt_data,
                "day_gantt": day_gantt_data
            }

        except Exception as e:
            print(f"Error getting Gantt data: {e}")
            return {"rab": None, "filtered_categories": [], "gantt": None, "day_gantt": []}

    # get ulok by email pembuat (Buat kontraktor)
    def get_ulok_by_email(self, email):
        """
        Mengambil daftar unik Nomor Ulok beserta Proyek, Nama Toko, dan Lingkup Pekerjaan 
        dari sheet RAB Form3.
        """
        ulok_list = []
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            records = worksheet.get_all_records()
            for record in records:
                if str(record.get('Email_Pembuat', '')).strip().lower() == email.strip().lower():
                    ulok = str(record.get('Nomor Ulok', '')).strip()
                    proyek = str(record.get('Proyek', '')).strip()
                    nama_toko = str(record.get('Nama_Toko', record.get('nama_toko', ''))).strip()
                    lingkup = str(record.get('Lingkup_Pekerjaan', record.get('Lingkup Pekerjaan', ''))).strip()
                    if ulok:
                        label = f"{ulok} - {proyek} ({lingkup}) - {nama_toko}"
                        ulok_list.append({
                            "label": label,
                            "value": ulok + "-" + lingkup,
                        })
            unique_ulok = {item['value']: item for item in reversed(ulok_list)}
            sorted_ulok = sorted(unique_ulok.values(), key=lambda x: x['value'])
            return sorted_ulok
        except Exception as e:
            print(f"Error saat mengambil daftar ulok by email: {e}")
            return []

    # Get ulok by Cabang (Buat PIC)
    def get_ulok_by_cabang_pic(self, cabang):
        ulok_list = []
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            records = worksheet.get_all_records()
            for record in records:
                if str(record.get('Cabang', '')).strip().lower() == cabang.strip().lower():
                    ulok = str(record.get('Nomor Ulok', '')).strip()
                    proyek = str(record.get('Proyek', '')).strip()
                    nama_toko = str(record.get('Nama_Toko', record.get('nama_toko', ''))).strip()
                    lingkup = str(record.get('Lingkup_Pekerjaan', record.get('Lingkup Pekerjaan', ''))).strip()
                    if ulok:
                        label = f"{ulok} - {proyek} ({lingkup}) - {nama_toko}"
                        ulok_list.append({
                            "label": label,
                            "value": ulok + "-" + lingkup,
                        })
            unique_ulok = {item['value']: item for item in reversed(ulok_list)}
            sorted_ulok = sorted(unique_ulok.values(), key=lambda x: x['value'])
            return sorted_ulok
        except Exception as e:
            print(f"Error saat mengambil daftar ulok by cabang PIC: {e}")
            return []

    # Insert atau Update Gantt Chart Data
    def insert_gantt_chart_data(self, data_dict):
        """
        Insert atau Update data ke sheet GANTT_CHART_SHEET_NAME.
        Cek dulu apakah data sudah ada berdasarkan Nomor Ulok dan Lingkup_Pekerjaan.
        Jika ada, update field yang dikirim. Jika belum ada, insert baris baru.
        """
        try:
            worksheet = self.sheet.worksheet(config.GANTT_CHART_SHEET_NAME)
            all_values = worksheet.get_all_values()
            
            if not all_values:
                return {"success": False, "message": "Sheet kosong atau header tidak ditemukan"}
            
            headers = all_values[0]
            data_rows = all_values[1:]
            
            if not headers:
                return {"success": False, "message": "Header sheet tidak ditemukan"}
            
            # Ambil Nomor Ulok dan Lingkup dari data_dict
            target_ulok = str(data_dict.get(config.COLUMN_NAMES.LOKASI, "")).strip().upper()
            target_lingkup = str(data_dict.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, "")).strip().lower()
            
            # Validasi: Ulok dan Lingkup harus ada untuk operasi yang benar
            if not target_ulok:
                return {"success": False, "message": "Nomor Ulok tidak boleh kosong dalam data yang dikirim"}
            
            if not target_lingkup:
                return {"success": False, "message": "Lingkup Pekerjaan tidak boleh kosong dalam data yang dikirim"}
            
            # Cari index kolom Nomor Ulok dan Lingkup_Pekerjaan
            try:
                ulok_idx = headers.index(config.COLUMN_NAMES.LOKASI)
                lingkup_idx = headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            except ValueError:
                return {"success": False, "message": "Kolom Nomor Ulok atau Lingkup_Pekerjaan tidak ditemukan di header"}
            
            # Cari apakah data sudah ada
            found_row_index = None
            for idx, row in enumerate(data_rows):
                if len(row) <= max(ulok_idx, lingkup_idx):
                    continue
                
                current_ulok = str(row[ulok_idx]).strip().upper()
                current_lingkup = str(row[lingkup_idx]).strip().lower()
                
                # Skip row yang ulok atau lingkup-nya kosong (data tidak valid untuk matching)
                if not current_ulok or not current_lingkup:
                    continue
                
                if current_ulok == target_ulok and current_lingkup == target_lingkup:
                    found_row_index = idx + 2  # +2 karena header di row 1 dan index mulai dari 0
                    break
            
            if found_row_index:
                # --- UPDATE: Data sudah ada, update field yang dikirim ---
                updates = []
                for key, value in data_dict.items():
                    if key in headers and value is not None and str(value).strip() != "":
                        col_idx = headers.index(key) + 1  # gspread menggunakan 1-based index
                        updates.append({
                            "range": gspread.utils.rowcol_to_a1(found_row_index, col_idx),
                            "values": [[value]]
                        })
                
                if updates:
                    worksheet.batch_update(updates)
                
                return {
                    "success": True, 
                    "message": "Data berhasil diperbarui",
                    "row_index": found_row_index,
                    "action": "update",
                    "fields_updated": len(updates)
                }
            else:
                # --- INSERT: Data belum ada, tambahkan baris baru ---
                row_data = [data_dict.get(header, "") for header in headers]
                
                # Gunakan append_row dengan value_input_option dan dapatkan response
                # untuk mendapatkan row index yang akurat (menghindari race condition)
                # insert_data_option='INSERT_ROWS' memastikan data ditambahkan sebagai baris baru
                result = worksheet.append_row(
                    row_data,
                    value_input_option='USER_ENTERED',
                    insert_data_option='INSERT_ROWS'
                )
                
                # Parse response untuk mendapatkan row number yang sebenarnya
                new_row_index = None
                try:
                    updated_range = result.get('updates', {}).get('updatedRange', '')
                    if updated_range:
                        range_part = updated_range.split('!')[-1]
                        match = re.search(r'[A-Z]+(\d+)', range_part)
                        if match:
                            new_row_index = int(match.group(1))
                except Exception as e:
                    print(f"Warning: Could not parse row index from append response: {e}")
                
                # Fallback jika parsing gagal
                if not new_row_index:
                    new_row_index = len(worksheet.get_all_values())
                
                return {
                    "success": True, 
                    "message": "Data berhasil ditambahkan",
                    "row_index": new_row_index,
                    "action": "insert"
                }
                
        except Exception as e:
            print(f"Error insert/update gantt chart data: {e}")
            return {"success": False, "message": str(e)}

    def insert_pengawasan_to_gantt_chart(self, nomor_ulok, lingkup_pekerjaan, pengawasan_day):
        """
        Insert data Pengawasan (angka hari) ke kolom Pengawasan_1 s/d Pengawasan_10 secara berurutan.
        Cari baris berdasarkan Nomor Ulok dan Lingkup_Pekerjaan, lalu insert ke kolom Pengawasan 
        yang pertama masih kosong.
        
        Parameters:
        -----------
        nomor_ulok : str
            Nomor Ulok untuk mencari baris
        lingkup_pekerjaan : str
            Lingkup Pekerjaan untuk mencari baris
        pengawasan_day : int/str
            Angka hari yang akan diinsert ke kolom Pengawasan
        
        Returns:
        --------
        dict: Hasil operasi dengan status dan detail
        """
        try:
            worksheet = self.sheet.worksheet(config.GANTT_CHART_SHEET_NAME)
            all_values = worksheet.get_all_values()
            
            if not all_values:
                return {"success": False, "message": "Sheet kosong atau header tidak ditemukan"}
            
            headers = all_values[0]
            data_rows = all_values[1:]
            
            if not headers:
                return {"success": False, "message": "Header sheet tidak ditemukan"}
            
            # Normalisasi input
            target_ulok = str(nomor_ulok).strip().upper()
            target_lingkup = str(lingkup_pekerjaan).strip().lower()
            pengawasan_value = str(pengawasan_day).strip()
            
            # Validasi input
            if not target_ulok:
                return {"success": False, "message": "Nomor Ulok tidak boleh kosong"}
            
            if not target_lingkup:
                return {"success": False, "message": "Lingkup Pekerjaan tidak boleh kosong"}
            
            if not pengawasan_value:
                return {"success": False, "message": "Pengawasan day tidak boleh kosong"}
            
            # Cari index kolom Nomor Ulok dan Lingkup_Pekerjaan
            try:
                ulok_idx = headers.index(config.COLUMN_NAMES.LOKASI)
            except ValueError:
                return {"success": False, "message": "Kolom Nomor Ulok tidak ditemukan di header"}
            
            try:
                lingkup_idx = headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            except ValueError:
                return {"success": False, "message": "Kolom Lingkup_Pekerjaan tidak ditemukan di header"}
            
            # Cari kolom Pengawasan_1 sampai Pengawasan_10
            pengawasan_columns = []
            for i in range(1, 11):  # Pengawasan_1 sampai Pengawasan_10
                col_name = f"Pengawasan_{i}"
                if col_name in headers:
                    pengawasan_columns.append({
                        "name": col_name,
                        "index": headers.index(col_name)
                    })
            
            if not pengawasan_columns:
                return {"success": False, "message": "Kolom Pengawasan_1 sampai Pengawasan_10 tidak ditemukan di header"}
            
            # Cari baris berdasarkan Nomor Ulok dan Lingkup_Pekerjaan
            found_row_index = None
            found_row_data = None
            for idx, row in enumerate(data_rows):
                if len(row) <= max(ulok_idx, lingkup_idx):
                    continue
                
                current_ulok = str(row[ulok_idx]).strip().upper()
                current_lingkup = str(row[lingkup_idx]).strip().lower()
                
                if current_ulok == target_ulok and current_lingkup == target_lingkup:
                    found_row_index = idx + 2  # +2 karena header di row 1 dan index mulai dari 0
                    found_row_data = row
                    break
            
            if not found_row_index:
                return {
                    "success": False, 
                    "message": f"Data dengan Nomor Ulok '{target_ulok}' dan Lingkup '{target_lingkup}' tidak ditemukan"
                }
            
            # Cari kolom Pengawasan pertama yang masih kosong
            target_pengawasan_col = None
            for peng_col in pengawasan_columns:
                col_idx = peng_col["index"]
                # Cek apakah cell kosong
                cell_value = ""
                if len(found_row_data) > col_idx:
                    cell_value = str(found_row_data[col_idx]).strip()
                
                if not cell_value:
                    target_pengawasan_col = peng_col
                    break
            
            if not target_pengawasan_col:
                return {
                    "success": False, 
                    "message": "Semua kolom Pengawasan_1 sampai Pengawasan_10 sudah terisi"
                }
            
            # Update cell Pengawasan yang kosong
            col_idx_1based = target_pengawasan_col["index"] + 1
            cell_range = gspread.utils.rowcol_to_a1(found_row_index, col_idx_1based)
            
            worksheet.update(cell_range, [[pengawasan_value]], value_input_option='USER_ENTERED')
            
            return {
                "success": True,
                "message": f"Data berhasil diinsert ke {target_pengawasan_col['name']}",
                "row_index": found_row_index,
                "column_name": target_pengawasan_col["name"],
                "value": pengawasan_value
            }
            
        except Exception as e:
            print(f"Error insert pengawasan to gantt chart: {e}")
            return {"success": False, "message": str(e)}

    def update_keterlambatan_day_gantt(self, nomor_ulok, lingkup_pekerjaan, kategori, h_awal, h_akhir, keterlambatan):
        """
        Insert atau Update kolom keterlambatan di sheet DAY_GANTT_CHART_SHEET_NAME.
        Cari baris berdasarkan Nomor Ulok, Lingkup_Pekerjaan, Kategori, h_awal, dan h_akhir.
        Jika ditemukan, update kolom keterlambatan. Jika tidak, insert baris baru.
        
        Parameters:
        -----------
        nomor_ulok : str
            Nomor Ulok untuk mencari baris
        lingkup_pekerjaan : str
            Lingkup Pekerjaan untuk mencari baris
        kategori : str
            Kategori pekerjaan untuk mencari baris
        h_awal : str
            Hari awal untuk mencari baris
        h_akhir : str
            Hari akhir untuk mencari baris
        keterlambatan : str/int
            Nilai keterlambatan yang akan diinsert/update
        
        Returns:
        --------
        dict: Hasil operasi dengan status dan detail
        """
        try:
            worksheet = self.sheet.worksheet(config.DAY_GANTT_CHART_SHEET_NAME)
            all_values = worksheet.get_all_values()
            
            # Normalisasi input
            target_ulok = str(nomor_ulok).strip().upper()
            target_lingkup = str(lingkup_pekerjaan).strip().lower()
            target_kategori = str(kategori).strip().lower()
            target_h_awal = str(h_awal).strip()
            target_h_akhir = str(h_akhir).strip()
            keterlambatan_value = str(keterlambatan).strip()
            
            # Validasi input
            if not target_ulok:
                return {"success": False, "message": "Nomor Ulok tidak boleh kosong"}
            if not target_lingkup:
                return {"success": False, "message": "Lingkup Pekerjaan tidak boleh kosong"}
            if not target_kategori:
                return {"success": False, "message": "Kategori tidak boleh kosong"}
            
            if not all_values:
                # Sheet kosong, buat header default dengan kolom keterlambatan
                default_headers = [
                    config.COLUMN_NAMES.LOKASI,
                    config.COLUMN_NAMES.LINGKUP_PEKERJAAN,
                    "Kategori",
                    config.COLUMN_NAMES.HARI_AWAL,
                    config.COLUMN_NAMES.HARI_AKHIR,
                    "keterlambatan"
                ]
                worksheet.append_row(default_headers, value_input_option='USER_ENTERED')
                headers = default_headers
                data_rows = []
            else:
                headers = all_values[0]
                data_rows = all_values[1:]
            
            if not headers:
                return {"success": False, "message": "Header sheet tidak ditemukan"}
            
            # Pastikan kolom keterlambatan ada di header
            if "keterlambatan" not in headers:
                # Tambahkan kolom keterlambatan ke header
                new_col_idx = len(headers) + 1
                worksheet.update_cell(1, new_col_idx, "keterlambatan")
                headers.append("keterlambatan")
            
            # Cari index kolom yang diperlukan
            try:
                ulok_idx = headers.index(config.COLUMN_NAMES.LOKASI)
            except ValueError:
                ulok_idx = headers.index("Nomor Ulok") if "Nomor Ulok" in headers else None
                if ulok_idx is None:
                    return {"success": False, "message": "Kolom Nomor Ulok tidak ditemukan di header"}
            
            try:
                lingkup_idx = headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            except ValueError:
                lingkup_idx = headers.index("Lingkup_Pekerjaan") if "Lingkup_Pekerjaan" in headers else None
                if lingkup_idx is None:
                    return {"success": False, "message": "Kolom Lingkup_Pekerjaan tidak ditemukan di header"}
            
            try:
                kategori_idx = headers.index("Kategori")
            except ValueError:
                return {"success": False, "message": "Kolom Kategori tidak ditemukan di header"}
            
            try:
                h_awal_idx = headers.index(config.COLUMN_NAMES.HARI_AWAL)
            except ValueError:
                h_awal_idx = headers.index("h_awal") if "h_awal" in headers else None
                if h_awal_idx is None:
                    return {"success": False, "message": "Kolom h_awal tidak ditemukan di header"}
            
            try:
                h_akhir_idx = headers.index(config.COLUMN_NAMES.HARI_AKHIR)
            except ValueError:
                h_akhir_idx = headers.index("h_akhir") if "h_akhir" in headers else None
                if h_akhir_idx is None:
                    return {"success": False, "message": "Kolom h_akhir tidak ditemukan di header"}
            
            keterlambatan_idx = headers.index("keterlambatan")
            
            # Cari baris yang cocok dengan semua kriteria
            found_row_index = None
            for idx, row in enumerate(data_rows):
                if len(row) <= max(ulok_idx, lingkup_idx, kategori_idx, h_awal_idx, h_akhir_idx):
                    continue
                
                current_ulok = str(row[ulok_idx]).strip().upper()
                current_lingkup = str(row[lingkup_idx]).strip().lower()
                current_kategori = str(row[kategori_idx]).strip().lower()
                current_h_awal = str(row[h_awal_idx]).strip()
                current_h_akhir = str(row[h_akhir_idx]).strip()
                
                # Match semua kriteria
                if (current_ulok == target_ulok and 
                    current_lingkup == target_lingkup and 
                    current_kategori == target_kategori and
                    current_h_awal == target_h_awal and
                    current_h_akhir == target_h_akhir):
                    found_row_index = idx + 2  # +2 karena header di row 1 dan index mulai dari 0
                    break
            
            if found_row_index:
                # UPDATE: Data ditemukan, update kolom keterlambatan
                cell_range = gspread.utils.rowcol_to_a1(found_row_index, keterlambatan_idx + 1)
                worksheet.update(cell_range, [[keterlambatan_value]], value_input_option='USER_ENTERED')
                
                return {
                    "success": True,
                    "message": "Keterlambatan berhasil diperbarui",
                    "row_index": found_row_index,
                    "action": "update",
                    "value": keterlambatan_value
                }
            else:
                # INSERT: Data tidak ditemukan, tambahkan baris baru
                row_data = []
                for header in headers:
                    if header == config.COLUMN_NAMES.LOKASI or header == "Nomor Ulok":
                        row_data.append(nomor_ulok)
                    elif header == config.COLUMN_NAMES.LINGKUP_PEKERJAAN or header == "Lingkup_Pekerjaan":
                        row_data.append(lingkup_pekerjaan)
                    elif header == "Kategori":
                        row_data.append(kategori)
                    elif header == config.COLUMN_NAMES.HARI_AWAL or header == "h_awal":
                        row_data.append(h_awal)
                    elif header == config.COLUMN_NAMES.HARI_AKHIR or header == "h_akhir":
                        row_data.append(h_akhir)
                    elif header == "keterlambatan":
                        row_data.append(keterlambatan_value)
                    else:
                        row_data.append("")
                
                result = worksheet.append_row(
                    row_data,
                    value_input_option='USER_ENTERED',
                    insert_data_option='INSERT_ROWS'
                )
                
                # Parse response untuk mendapatkan row number
                new_row_index = None
                try:
                    updated_range = result.get('updates', {}).get('updatedRange', '')
                    if updated_range:
                        range_part = updated_range.split('!')[-1]
                        match = re.search(r'[A-Z]+(\d+)', range_part)
                        if match:
                            new_row_index = int(match.group(1))
                except Exception as e:
                    print(f"Warning: Could not parse row index from append response: {e}")
                
                if not new_row_index:
                    new_row_index = len(worksheet.get_all_values())
                
                return {
                    "success": True,
                    "message": "Data baru dengan keterlambatan berhasil ditambahkan",
                    "row_index": new_row_index,
                    "action": "insert",
                    "value": keterlambatan_value
                }
                
        except Exception as e:
            print(f"Error update keterlambatan day gantt: {e}")
            return {"success": False, "message": str(e)}

    def insert_day_gantt_chart_data(self, data_list):
        """
        FIXED: Support Multi-Range per Kategori.
        Logic Update: Menggunakan list of indices untuk existing_data, bukan single index.
        Ini memungkinkan satu kategori memiliki banyak baris (range tanggal berbeda).
        """
        try:
            worksheet = self.sheet.worksheet(config.DAY_GANTT_CHART_SHEET_NAME)
            all_values = worksheet.get_all_values()
            
            if not all_values:
                # Sheet kosong, buat header default
                default_headers = [
                    config.COLUMN_NAMES.LOKASI,
                    config.COLUMN_NAMES.LINGKUP_PEKERJAAN,
                    "Kategori",
                    config.COLUMN_NAMES.HARI_AWAL,
                    config.COLUMN_NAMES.HARI_AKHIR
                ]
                worksheet.append_row(default_headers, value_input_option='USER_ENTERED')
                headers = default_headers
                data_rows = []
            else:
                headers = all_values[0]
                data_rows = all_values[1:]
            
            if not headers:
                return {"success": False, "message": "Header sheet tidak ditemukan"}
            
            # Validasi data_list
            if not data_list or not isinstance(data_list, list):
                return {"success": False, "message": "data_list harus berupa list yang tidak kosong"}
            
            # Cari index kolom yang diperlukan
            try:
                ulok_idx = headers.index(config.COLUMN_NAMES.LOKASI)
            except ValueError:
                ulok_idx = headers.index("Nomor Ulok") if "Nomor Ulok" in headers else None
                if ulok_idx is None:
                    return {"success": False, "message": "Kolom Nomor Ulok tidak ditemukan di header"}
            
            try:
                lingkup_idx = headers.index(config.COLUMN_NAMES.LINGKUP_PEKERJAAN)
            except ValueError:
                lingkup_idx = headers.index("Lingkup_Pekerjaan") if "Lingkup_Pekerjaan" in headers else None
                if lingkup_idx is None:
                    return {"success": False, "message": "Kolom Lingkup_Pekerjaan tidak ditemukan di header"}
            
            try:
                kategori_idx = headers.index("Kategori")
            except ValueError:
                return {"success": False, "message": "Kolom Kategori tidak ditemukan di header"}
            
            # --- PERBAIKAN DI SINI ---
            # Buat dictionary map Key -> LIST of Row Indices (bukan single int)
            # Contoh: "ULOK1|ME|Instalasi" -> [5, 8, 12] (Instalasi ada di baris 5, 8, dan 12)
            existing_data = {}
            for idx, row in enumerate(data_rows):
                if len(row) > max(ulok_idx, lingkup_idx, kategori_idx):
                    key = f"{str(row[ulok_idx]).strip().upper()}|{str(row[lingkup_idx]).strip().lower()}|{str(row[kategori_idx]).strip().lower()}"
                    
                    if key not in existing_data:
                        existing_data[key] = []
                    existing_data[key].append(idx + 2)  # +2 (Header + 0-index)
            
            # Proses setiap item dalam data_list
            updates = []
            inserts = []
            results = {
                "updated": 0,
                "inserted": 0,
                "errors": []
            }
            
            for item in data_list:
                item_ulok = str(item.get(config.COLUMN_NAMES.LOKASI, item.get("Nomor Ulok", ""))).strip().upper()
                item_lingkup = str(item.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, item.get("Lingkup_Pekerjaan", ""))).strip().lower()
                item_kategori = str(item.get("Kategori", "")).strip().lower()
                item_h_awal = str(item.get(config.COLUMN_NAMES.HARI_AWAL, item.get("h_awal", ""))).strip()
                item_h_akhir = str(item.get(config.COLUMN_NAMES.HARI_AKHIR, item.get("h_akhir", ""))).strip()
                
                if not item_ulok or not item_lingkup or not item_kategori:
                    continue
                
                lookup_key = f"{item_ulok}|{item_lingkup}|{item_kategori}"
                
                # Cek apakah ada slot baris tersedia untuk update
                # Kita gunakan .pop(0) untuk mengambil baris pertama yang tersedia, lalu menghapusnya dari list
                # sehingga input range kedua akan mengambil baris berikutnya (jika ada) atau insert baru.
                if lookup_key in existing_data and len(existing_data[lookup_key]) > 0:
                    # UPDATE: Pakai slot baris yang ada
                    row_index = existing_data[lookup_key].pop(0)
                    
                    if item_h_awal:
                        try:
                            h_awal_idx = headers.index(config.COLUMN_NAMES.HARI_AWAL)
                        except ValueError:
                            h_awal_idx = headers.index("h_awal") if "h_awal" in headers else None
                        
                        if h_awal_idx is not None:
                            updates.append({
                                "range": gspread.utils.rowcol_to_a1(row_index, h_awal_idx + 1),
                                "values": [[item_h_awal]]
                            })
                    
                    if item_h_akhir:
                        try:
                            h_akhir_idx = headers.index(config.COLUMN_NAMES.HARI_AKHIR)
                        except ValueError:
                            h_akhir_idx = headers.index("h_akhir") if "h_akhir" in headers else None
                        
                        if h_akhir_idx is not None:
                            updates.append({
                                "range": gspread.utils.rowcol_to_a1(row_index, h_akhir_idx + 1),
                                "values": [[item_h_akhir]]
                            })
                    
                    results["updated"] += 1
                else:
                    # INSERT: Data baru atau slot update habis
                    row_data = []
                    for header in headers:
                        if header == config.COLUMN_NAMES.LOKASI or header == "Nomor Ulok":
                            row_data.append(item_ulok)
                        elif header == config.COLUMN_NAMES.LINGKUP_PEKERJAAN or header == "Lingkup_Pekerjaan":
                            row_data.append(item.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, item.get("Lingkup_Pekerjaan", "")))
                        elif header == "Kategori":
                            row_data.append(item.get("Kategori", ""))
                        elif header == config.COLUMN_NAMES.HARI_AWAL or header == "h_awal":
                            row_data.append(item_h_awal)
                        elif header == config.COLUMN_NAMES.HARI_AKHIR or header == "h_akhir":
                            row_data.append(item_h_akhir)
                        else:
                            row_data.append(item.get(header, ""))
                    
                    inserts.append(row_data)
                    results["inserted"] += 1
                    
                    # Kita TIDAK menambahkan ke existing_data di sini
                    # Karena insert baru akan dilakukan di akhir (append), bukan update di tempat.
            
            if updates:
                worksheet.batch_update(updates)
            
            if inserts:
                worksheet.append_rows(inserts, value_input_option='USER_ENTERED')
            
            return {
                "success": True,
                "message": f"Berhasil: {results['updated']} update, {results['inserted']} insert baru.",
                "details": results
            }
            
        except Exception as e:
            print(f"Error insert/update day gantt chart data: {e}")
            return {"success": False, "message": str(e)}
        
    def insert_day_gantt_chart_single(self, nomor_ulok, lingkup_pekerjaan, kategori_data):
        """
        Insert atau Update data untuk satu Nomor Ulok dengan multiple kategori.
        
        Parameters:
        -----------
        nomor_ulok : str
            Nomor Ulok (contoh: "asa-asa-sas")
        lingkup_pekerjaan : str
            Lingkup Pekerjaan (contoh: "ME" atau "SIPIL")
        kategori_data : list of dict
            List kategori dengan h_awal dan h_akhir:
            [
                {"Kategori": "Persiapan", "h_awal": "19/12/2025", "h_akhir": "21/12/2025"},
                {"Kategori": "Pembersihan", "h_awal": "23/12/2025", "h_akhir": "25/12/2025"},
                {"Kategori": "Bobokan", "h_awal": "28/12/2025", "h_akhir": "30/12/2025"}
            ]
        
        Returns:
        --------
        dict: Hasil operasi dengan status dan detail
        """
        # Konversi ke format data_list
        data_list = []
        for item in kategori_data:
            data_list.append({
                config.COLUMN_NAMES.LOKASI: nomor_ulok,
                config.COLUMN_NAMES.LINGKUP_PEKERJAAN: lingkup_pekerjaan,
                "Kategori": item.get("Kategori", ""),
                config.COLUMN_NAMES.HARI_AWAL: item.get("h_awal", item.get(config.COLUMN_NAMES.HARI_AWAL, "")),
                config.COLUMN_NAMES.HARI_AKHIR: item.get("h_akhir", item.get(config.COLUMN_NAMES.HARI_AKHIR, ""))
            })
        
        return self.insert_day_gantt_chart_data(data_list)

    
    # akhir gantt chart

    def get_user_info_by_cabang(self, cabang):
        pic_list, koordinator_info, manager_info = [], {}, {}
        try:
            cabang_sheet = self.sheet.worksheet(config.CABANG_SHEET_NAME)
            records = cabang_sheet.get_all_records()
            for record in records:
                if str(record.get('CABANG', '')).strip().lower() == cabang.strip().lower():
                    jabatan = str(record.get('JABATAN', '')).strip().upper()
                    email = str(record.get('EMAIL_SAT', '')).strip()
                    nama = str(record.get('NAMA LENGKAP', '')).strip()
                    
                    if "SUPPORT" in jabatan:
                        pic_list.append({'email': email, 'nama': nama})
                    elif "COORDINATOR" in jabatan:
                        koordinator_info = {'email': email, 'nama': nama}
                    elif "MANAGER" in jabatan and "BRANCH MANAGER" not in jabatan:
                         manager_info = {'email': email, 'nama': nama}
        except Exception as e:
            print(f"Error saat mengambil info user by cabang: {e}")
        return pic_list, koordinator_info, manager_info

    def get_kode_ulok_by_cabang(self, cabang):
        ulok_list = []
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            records = worksheet.get_all_records()
            for record in records:
                if str(record.get('Cabang', '')).strip().lower() == cabang.strip().lower():
                    ulok = str(record.get('Nomor Ulok', '')).strip()
                    if ulok:
                        ulok_list.append(ulok)
            return sorted(list(set(ulok_list)))
        except Exception as e:
            print(f"Error saat mengambil kode ulok by cabang: {e}")
            return []

    def get_kode_ulok_by_cabang_kedua(self, cabang):
        ulok_list = []
        try:
            worksheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME_RAB_2)
            records = worksheet.get_all_records()
            for record in records:
                if str(record.get('Cabang', '')).strip().lower() == cabang.strip().lower():
                    ulok = str(record.get('Nomor Ulok', '')).strip()
                    if ulok:
                        ulok_list.append(ulok)
            return sorted(list(set(ulok_list)))
        except Exception as e:
            print(f"Error saat mengambil kode ulok by cabang: {e}")
            return []


    def get_active_pengawasan_by_pic(self, pic_email):
        try:
            penugasan_sheet = self.gspread_client.open_by_key(config.PENGAWASAN_SPREADSHEET_ID).worksheet(config.PENUGASAN_SHEET_NAME)
            all_records = penugasan_sheet.get_all_records()
            
            projects = []
            for record in all_records:
                if str(record.get('Email_BBS', '')).strip().lower() == pic_email.strip().lower():
                    projects.append({
                        "kode_ulok": record.get("Kode_Ulok"),
                        "cabang": record.get("Cabang")
                    })
            return projects
        except Exception as e:
            print(f"Error getting active pengawasan projects: {e}")
            return []

    def get_pic_email_by_ulok(self, kode_ulok):
        try:
            penugasan_sheet = self.gspread_client.open_by_key(config.PENGAWASAN_SPREADSHEET_ID).worksheet(config.PENUGASAN_SHEET_NAME)
            all_records = penugasan_sheet.get_all_records()
            for record in reversed(all_records):
                if str(record.get('Kode_Ulok', '')).strip() == str(kode_ulok).strip():
                    return record.get('Email_BBS')
            return None
        except Exception as e:
            print(f"Error getting PIC email by Ulok: {e}")
            return None

    def upload_file_to_drive(self, file_bytes, filename, mimetype, folder_id):
        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
        file = self.drive_service.files().create(
            body=file_metadata, media_body=media, fields='id, webViewLink'
        ).execute()
        return file.get('webViewLink')

    def create_calendar_event(self, event_data):
        try:
            event = {
                'summary': event_data['title'],
                'description': event_data['description'],
                'start': {'date': event_data['date']},
                'end': {'date': event_data['date']},
                'attendees': [{'email': email} for email in event_data['guests']],
            }
            self.calendar_service.events().insert(
                calendarId='primary', body=event, sendUpdates='all'
            ).execute()
            print(f"✅ Event kalender berhasil dibuat untuk: {event_data['title']}")
        except Exception as e:
            print(f"❌ Gagal membuat event di Google Calendar: {e}")

    def send_email(self, to, subject, html_body, attachments=None, cc=None):
        try:
            message = MIMEMultipart()
            to_list = to if isinstance(to, list) else [to]
            cc_list = cc if isinstance(cc, list) else ([cc] if cc else [])
            message['to'] = ', '.join(filter(None, to_list))
            message['subject'] = subject
            if cc_list: message['cc'] = ', '.join(filter(None, cc_list))
            message.attach(MIMEText(html_body, 'html'))

            if attachments:
                for attachment_tuple in attachments:
                    filename, file_bytes, mimetype = attachment_tuple
                    main_type, sub_type = mimetype.split('/')
                    part = MIMEBase(main_type, sub_type)
                    part.set_payload(file_bytes)
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                    message.attach(part)

            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            self.gmail_service.users().messages().send(
                userId='me', body={'raw': raw_message}
            ).execute()
            print(f"Email sent successfully to {message['to']}")
        except Exception as e:
            print(f"An error occurred while sending email: {e}")
            raise

    def validate_user(self, email, cabang):
        try:
            cabang_sheet = self.sheet.worksheet(config.CABANG_SHEET_NAME)
            for record in cabang_sheet.get_all_records():
                if str(record.get('EMAIL_SAT', '')).strip().lower() == email.lower() and \
                   str(record.get('CABANG', '')).strip().lower() == cabang.lower():
                    return True, record.get('JABATAN', '')
        except gspread.exceptions.WorksheetNotFound:
            print(f"Error: Worksheet '{config.CABANG_SHEET_NAME}' not found.")
        return False, None

    def check_user_submissions(self, email, cabang):
        try:
            all_values = self.data_entry_sheet.get_all_values()
            if len(all_values) <= 1:
                return {"active_codes": {"pending": [], "approved": []}, "rejected_submissions": []}
            headers = all_values[0]
            records = [dict(zip(headers, row)) for row in all_values[1:]]
            
            pending_codes, approved_codes, rejected_submissions = [], [], []
            processed_keys = set() # Ganti nama set agar lebih jelas
            user_cabang = str(cabang).strip().lower()

            for record in reversed(records):
                lokasi = str(record.get(config.COLUMN_NAMES.LOKASI, "")).strip().upper()
                lingkup = str(record.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, "")).strip().upper() # Ambil lingkup
                
                # Buat kunci unik gabungan
                unique_key = f"{lokasi}_{lingkup}"

                if not lokasi or unique_key in processed_keys: continue
                
                status = str(record.get(config.COLUMN_NAMES.STATUS, "")).strip()
                record_cabang = str(record.get(config.COLUMN_NAMES.CABANG, "")).strip().lower()
                
                # Cek filter berdasarkan user cabang (opsional: tambah cek email jika perlu sangat spesifik)
                if record_cabang == user_cabang:
                    if status in [config.STATUS.WAITING_FOR_COORDINATOR, config.STATUS.WAITING_FOR_MANAGER]:
                        pending_codes.append(lokasi)
                    elif status == config.STATUS.APPROVED:
                        approved_codes.append(lokasi)
                    elif status in [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER]:
                        item_details_json = record.get('Item_Details_JSON', '{}')
                        if item_details_json:
                            try:
                                item_details = json.loads(item_details_json)
                                record.update(item_details)
                            except json.JSONDecodeError:
                                print(f"Warning: Could not decode Item_Details_JSON for {lokasi}")
                        rejected_submissions.append(record)
                
                processed_keys.add(unique_key)

            return {"active_codes": {"pending": pending_codes, "approved": approved_codes}, "rejected_submissions": rejected_submissions}
        except Exception as e:
            raise e

    def check_user_submissions_rab_2(self, email, cabang):
        """Mirrors check_user_submissions but reads from Spreadsheet RAB 2."""
        try:
            # Open RAB 2 spreadsheet and target worksheet
            spreadsheet = self.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
            worksheet = spreadsheet.worksheet(config.DATA_ENTRY_SHEET_NAME_RAB_2)

            all_values = worksheet.get_all_values()
            if len(all_values) <= 1:
                return {"active_codes": {"pending": [], "approved": []}, "rejected_submissions": []}

            headers = all_values[0]
            records = [dict(zip(headers, row)) for row in all_values[1:]]

            pending_codes, approved_codes, rejected_submissions = [], [], []
            processed_locations = set()
            user_cabang = str(cabang).strip().lower()

            for record in reversed(records):
                lokasi = str(record.get(config.COLUMN_NAMES.LOKASI, "")).strip().upper()
                if not lokasi or lokasi in processed_locations:
                    continue

                status = str(record.get(config.COLUMN_NAMES.STATUS, "")).strip()
                record_cabang = str(record.get(config.COLUMN_NAMES.CABANG, "")).strip().lower()

                if status in [config.STATUS.WAITING_FOR_COORDINATOR, config.STATUS.WAITING_FOR_MANAGER]:
                    pending_codes.append(lokasi)
                elif status == config.STATUS.APPROVED:
                    approved_codes.append(lokasi)
                elif status in [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER] and record_cabang == user_cabang:
                    item_details_json = record.get('Item_Details_JSON', '{}')
                    if item_details_json:
                        try:
                            item_details = json.loads(item_details_json)
                            record.update(item_details)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not decode Item_Details_JSON for {lokasi}")
                    rejected_submissions.append(record)

                processed_locations.add(lokasi)

            return {"active_codes": {"pending": pending_codes, "approved": approved_codes}, "rejected_submissions": rejected_submissions}
        except Exception as e:
            raise e

    def get_sheet_headers(self, worksheet_name):
        return self.sheet.worksheet(worksheet_name).row_values(1)

    def append_to_sheet(self, data, worksheet_name):
        worksheet = self.sheet.worksheet(worksheet_name)
        headers = self.get_sheet_headers(worksheet_name)
        row_data = [data.get(header, "") for header in headers]
        
        # Gunakan append_row dengan value_input_option dan dapatkan response
        # untuk mendapatkan row index yang akurat (menghindari race condition)
        # insert_data_option='INSERT_ROWS' memastikan data ditambahkan sebagai baris baru
        result = worksheet.append_row(
            row_data, 
            value_input_option='USER_ENTERED',
            insert_data_option='INSERT_ROWS'
        )
        
        # Parse response untuk mendapatkan row number yang sebenarnya
        # Response format: {'updates': {'updatedRange': 'SheetName!A100:XX100', ...}}
        try:
            updated_range = result.get('updates', {}).get('updatedRange', '')
            # Extract row number from range like "Form2!A100:XX100"
            if updated_range:
                # Ambil bagian setelah "!" dan sebelum ":"
                range_part = updated_range.split('!')[-1]  # "A100:XX100"
                # Ambil angka dari cell pertama
                import re
                match = re.search(r'[A-Z]+(\d+)', range_part)
                if match:
                    return int(match.group(1))
        except Exception as e:
            print(f"Warning: Could not parse row index from append response: {e}")
        
        # Fallback ke method lama jika parsing gagal
        return len(worksheet.get_all_values())

    def get_row_data(self, row_index):
        records = self.data_entry_sheet.get_all_records()
        if 1 < row_index <= len(records) + 1:
            return records[row_index - 2]
        return {}

    def update_cell(self, row_index, column_name, value):
        try:
            col_index = self.data_entry_sheet.row_values(1).index(column_name) + 1
            self.data_entry_sheet.update_cell(row_index, col_index, value)
            return True
        except Exception as e:
            print(f"Error updating cell [{row_index}, {column_name}]: {e}")
            return False

    def get_email_by_jabatan(self, branch_name, jabatan):
        try:
            cabang_sheet = self.sheet.worksheet(config.CABANG_SHEET_NAME)
            for record in cabang_sheet.get_all_records():
                if branch_name and str(record.get('CABANG', '')).strip().lower() == branch_name.strip().lower() and \
                   str(record.get('JABATAN', '')).strip().upper() == jabatan.strip().upper():
                    return record.get('EMAIL_SAT')
        except gspread.exceptions.WorksheetNotFound:
            print(f"Error: Worksheet '{config.CABANG_SHEET_NAME}' not found.")
        return None

    def get_emails_by_jabatan(self, branch_name, jabatan):
        emails = []
        try:
            cabang_sheet = self.sheet.worksheet(config.CABANG_SHEET_NAME)
            for record in cabang_sheet.get_all_records():
                if branch_name and str(record.get('CABANG', '')).strip().lower() == branch_name.strip().lower() and \
                   str(record.get('JABATAN', '')).strip().upper() == jabatan.strip().upper():
                    email = record.get('EMAIL_SAT')
                    if email:
                        emails.append(email)
        except gspread.exceptions.WorksheetNotFound:
            print(f"Error: Worksheet '{config.CABANG_SHEET_NAME}' not found.")
        return emails

    def copy_to_approved_sheet(self, row_data):
        try:
            approved_sheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            headers = self.get_sheet_headers(config.APPROVED_DATA_SHEET_NAME)
            data_to_append = [row_data.get(header, "") for header in headers]
            # insert_data_option='INSERT_ROWS' memastikan data ditambahkan sebagai baris baru
            approved_sheet.append_row(
                data_to_append,
                value_input_option='USER_ENTERED',
                insert_data_option='INSERT_ROWS'
            )
            return True
        except Exception as e:
            print(f"Failed to copy data to approved sheet: {e}")
            return False

    def copy_to_approved_sheet_kedua(self, row_data):
        try:
            # 1. Buka Spreadsheet RAB 2 (Bukan yang default)
            spreadsheet = self.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
            
            # 2. Buka Worksheet Tujuan (Form3 / Approved Data untuk RAB 2)
            approved_sheet = spreadsheet.worksheet(config.APPROVED_DATA_SHEET_NAME_RAB_2)
            
            # 3. Ambil Header untuk pemetaan data yang benar
            headers = approved_sheet.row_values(1)
            
            # 4. Susun data sesuai urutan header
            data_to_append = [row_data.get(header, "") for header in headers]
            
            # 5. Simpan - insert_data_option='INSERT_ROWS' memastikan data ditambahkan sebagai baris baru
            approved_sheet.append_row(
                data_to_append,
                value_input_option='USER_ENTERED',
                insert_data_option='INSERT_ROWS'
            )
            print("Berhasil menyalin data ke sheet Approved RAB 2")
            return True
        except Exception as e:
            print(f"Gagal menyalin data ke approved sheet RAB 2: {e}")
            return False


    def delete_row(self, worksheet_name, row_index):
        try:
            worksheet = self.sheet.worksheet(worksheet_name)
            worksheet.delete_rows(row_index)
            return True
        except Exception as e:
            print(f"Failed to delete row {row_index} from {worksheet_name}: {e}")
            return False
            
    def get_sheet_data_by_id(self, spreadsheet_id):
        try:
            spreadsheet = self.gspread_client.open_by_key(spreadsheet_id)
            worksheet = spreadsheet.get_worksheet(0)
            return worksheet.get_all_values()
        except Exception as e:
            raise e

    def _normalize_ulok(self, ulok_string):
        """
        Membersihkan format Nomor Ulok.
        Z001-2512-D4D4-R  -> Z0012512D4D4R
        Z001 - 2512 - 0001 -> Z00125120001
        """
        if not ulok_string:
            return ""
        return str(ulok_string).replace("-", "").replace(" ", "").strip().upper()

    def _normalize_lingkup(self, lingkup_string):
        """
        Menyamakan format Lingkup Pekerjaan.
        PEKERJAAN SIPIL -> SIPIL
        MEKANIKAL -> ME
        """
        if not lingkup_string:
            return ""
        clean = str(lingkup_string).strip().upper()
        if "SIPIL" in clean:
            return "SIPIL"
        if "ME" in clean or "M.E" in clean or "MEKANIKAL" in clean:
            return "ME"
        return clean
    
    def check_ulok_exists(self, nomor_ulok_to_check, lingkup_pekerjaan_to_check):
        """
        Mengecek apakah RAB sudah ada dan sedang AKTIF (Waiting atau Approved).
        Jika fungsi ini return True, berarti user TIDAK BOLEH submit (Duplicate).
        """
        try:
            target_ulok = self._normalize_ulok(nomor_ulok_to_check)
            target_lingkup = self._normalize_lingkup(lingkup_pekerjaan_to_check)
            
            all_records = self.data_entry_sheet.get_all_records()
            
            for record in all_records:
                status = str(record.get(config.COLUMN_NAMES.STATUS, "")).strip()
                
                # Cek hanya status AKTIF.
                # Jika status DITOLAK, fungsi ini akan mengabaikannya (return False untuk row itu),
                # sehingga kode di app.py bisa lanjut mengecek apakah ini Revisi.
                active_statuses = [
                    config.STATUS.WAITING_FOR_COORDINATOR, 
                    config.STATUS.WAITING_FOR_MANAGER, 
                    config.STATUS.APPROVED
                ]

                if status in active_statuses:
                    # Ambil data dari row & Normalisasi
                    rec_ulok_raw = str(record.get(config.COLUMN_NAMES.LOKASI, ""))
                    
                    # Robust: Cek key 'Lingkup_Pekerjaan' (config) DAN 'Lingkup Pekerjaan' (header spasi)
                    rec_lingkup_raw = record.get(
                        config.COLUMN_NAMES.LINGKUP_PEKERJAAN, 
                        record.get('Lingkup Pekerjaan', '')
                    )

                    rec_ulok = self._normalize_ulok(rec_ulok_raw)
                    rec_lingkup = self._normalize_lingkup(rec_lingkup_raw)

                    # Bandingkan data yang sudah bersih
                    if rec_ulok == target_ulok and rec_lingkup == target_lingkup:
                        print(f"Duplikasi RAB Ditemukan: {target_ulok} ({target_lingkup})")
                        return True
                        
            return False
        except Exception as e:
            print(f"Error checking for existing RAB ulok: {e}")
            return False

    def check_spk_exists(self, nomor_ulok_to_check, lingkup_pekerjaan_to_check):
        """
        Mengecek duplikasi di Sheet SPK (SPK_Data).
        Memastikan Manager tidak submit SPK ganda untuk Ulok + Lingkup yang sama.
        """
        try:
            target_ulok = self._normalize_ulok(nomor_ulok_to_check)
            target_lingkup = self._normalize_lingkup(lingkup_pekerjaan_to_check)

            spk_sheet = self.sheet.worksheet(config.SPK_DATA_SHEET_NAME)
            all_records = spk_sheet.get_all_records()

            for record in all_records:
                status = str(record.get("Status", "")).strip()

                # Cek jika statusnya BUKAN Ditolak (artinya sedang Waiting BM atau Sudah Approved)
                if status != config.STATUS.SPK_REJECTED:
                    
                    rec_ulok_raw = str(record.get("Nomor Ulok", ""))
                    rec_lingkup_raw = record.get("Lingkup Pekerjaan", record.get("Lingkup_Pekerjaan", ""))

                    rec_ulok = self._normalize_ulok(rec_ulok_raw)
                    rec_lingkup = self._normalize_lingkup(rec_lingkup_raw)

                    if rec_ulok == target_ulok and rec_lingkup == target_lingkup:
                        print(f"Duplikasi SPK Ditemukan: {target_ulok} ({target_lingkup}) status {status}")
                        return True
            
            return False
        except Exception as e:
            print(f"Error checking for existing SPK ulok: {e}")
            return False
    
    def is_revision(self, nomor_ulok, email_pembuat, lingkup_pekerjaan=None):
        try:
            normalized_ulok = str(nomor_ulok).replace("-", "")
            target_lingkup = str(lingkup_pekerjaan).strip().upper() if lingkup_pekerjaan else None
            
            all_records = self.data_entry_sheet.get_all_records()
            for record in reversed(all_records):
                rec_ulok = str(record.get(config.COLUMN_NAMES.LOKASI, "")).replace("-", "")
                rec_email = record.get(config.COLUMN_NAMES.EMAIL_PEMBUAT, "")
                rec_lingkup = str(record.get(config.COLUMN_NAMES.LINGKUP_PEKERJAAN, "")).strip().upper()

                if rec_ulok == normalized_ulok and rec_email == email_pembuat:
                    # Jika lingkup spesifik diberikan, cek kecocokannya
                    if target_lingkup and rec_lingkup != target_lingkup:
                        continue
                        
                    return record.get(config.COLUMN_NAMES.STATUS, "") in [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER]
            return False
        except Exception:
            return False

    def get_approved_rab_by_cabang(self, user_cabang):
        try:
            approved_sheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME)
            all_records = approved_sheet.get_all_records()
            
            branch_groups = {
                "BANDUNG 1": ["BANDUNG 1", "BANDUNG 2"], "BANDUNG 2": ["BANDUNG 1", "BANDUNG 2"],
                "LOMBOK": ["LOMBOK", "SUMBAWA"], "SUMBAWA": ["LOMBOK", "SUMBAWA"],
                "MEDAN": ["MEDAN", "ACEH"], "ACEH": ["MEDAN", "ACEH"],
                "PALEMBANG": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"], "BENGKULU": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"],
                "BANGKA": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"], "BELITUNG": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"],
                "SIDOARJO": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"], "SIDOARJO BPN_SMD": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"],
                "MANOKWARI": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"], "NTT": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"],
                "SORONG": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"]
            }
            
            allowed_branches = [b.lower() for b in branch_groups.get(user_cabang.upper(), [user_cabang.upper()])]
            filtered_rabs = [rec for rec in all_records if str(rec.get('Cabang', '')).strip().lower() in allowed_branches]

            for rab in filtered_rabs:
                try:
                    grand_total_non_sbo_from_sheet = float(str(rab.get(config.COLUMN_NAMES.GRAND_TOTAL_NONSBO, 0)).replace(",", ""))
                except (ValueError, TypeError):
                    grand_total_non_sbo_from_sheet = 0
                
                final_total_with_ppn = grand_total_non_sbo_from_sheet * 1.11
                rab['Grand Total Non-SBO'] = final_total_with_ppn
                
            return filtered_rabs
        except Exception as e:
            print(f"Error getting approved RABs: {e}")
            raise e

    def get_approved_rab_by_cabang_kedua(self, user_cabang):
        try:
            approved_sheet = self.sheet.worksheet(config.APPROVED_DATA_SHEET_NAME_RAB_2)
            all_records = approved_sheet.get_all_records()
            
            branch_groups = {
                "BANDUNG 1": ["BANDUNG 1", "BANDUNG 2"], "BANDUNG 2": ["BANDUNG 1", "BANDUNG 2"],
                "LOMBOK": ["LOMBOK", "SUMBAWA"], "SUMBAWA": ["LOMBOK", "SUMBAWA"],
                "MEDAN": ["MEDAN", "ACEH"], "ACEH": ["MEDAN", "ACEH"],
                "PALEMBANG": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"], "BENGKULU": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"],
                "BANGKA": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"], "BELITUNG": ["PALEMBANG", "BENGKULU", "BANGKA", "BELITUNG"],
                "SIDOARJO": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"], "SIDOARJO BPN_SMD": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"],
                "MANOKWARI": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"], "NTT": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"],
                "SORONG": ["SIDOARJO", "SIDOARJO BPN_SMD", "MANOKWARI", "NTT", "SORONG"]
            }
            
            allowed_branches = [b.lower() for b in branch_groups.get(user_cabang.upper(), [user_cabang.upper()])]
            filtered_rabs = [rec for rec in all_records if str(rec.get('Cabang', '')).strip().lower() in allowed_branches]

            for rab in filtered_rabs:
                try:
                    grand_total_non_sbo_from_sheet = float(str(rab.get(config.COLUMN_NAMES.GRAND_TOTAL_NONSBO, 0)).replace(",", ""))
                except (ValueError, TypeError):
                    grand_total_non_sbo_from_sheet = 0
                
                final_total_with_ppn = grand_total_non_sbo_from_sheet * 1.11
                rab['Grand Total Non-SBO'] = final_total_with_ppn
                
            return filtered_rabs
        except Exception as e:
            print(f"Error getting approved RABs: {e}")
            raise e

    def get_kontraktor_by_cabang(self, user_cabang):
        try:
            kontraktor_sheet_object = self.gspread_client.open_by_key(config.KONTRAKTOR_SHEET_ID)
            worksheet = kontraktor_sheet_object.worksheet(config.KONTRAKTOR_SHEET_NAME)
            
            all_values = worksheet.get_all_values()
            if len(all_values) < 2: return []

            headers = all_values[1]
            records = [dict(zip(headers, row)) for row in all_values[2:]]
            
            allowed_branches_lower = user_cabang.strip().lower()
            kontraktor_list = []
            for record in records:
                if str(record.get('NAMA CABANG', '')).strip().lower() == allowed_branches_lower and \
                   str(record.get('STATUS KONTRAKTOR', '')).strip().upper() == 'AKTIF':
                    nama_kontraktor = str(record.get('NAMA KONTRAKTOR', '')).strip()
                    if nama_kontraktor and nama_kontraktor not in kontraktor_list:
                        kontraktor_list.append(nama_kontraktor)
            return sorted(kontraktor_list)
        except Exception as e:
            print(f"Error getting contractors: {e}")
            raise e
    
    def update_row(self, sheet_name, row_index, data_dict):
        """
        Update seluruh nilai 1 baris berdasarkan dictionary key=columnName.
        """
        sheet = self.sheet.worksheet(sheet_name)
        headers = sheet.row_values(1)

        # Ubah dictionary menjadi list sesuai urutan kolom
        updated_row = []
        for col in headers:
            updated_row.append(data_dict.get(col, ""))
        
        # Hitung jumlah kolom
        num_cols = len(headers)
        
        # Gunakan fungsi helper untuk mendapatkan huruf kolom terakhir yang benar (misal: "ZZ")
        last_col_letter = self._get_col_letter(num_cols)

        # Update seluruh baris sekaligus dengan range yang valid
        sheet.update(
            f"A{row_index}:{last_col_letter}{row_index}",
            [updated_row]
        )

    def get_row_data_by_sheet(self, worksheet, row_index):
        try:
            records = worksheet.get_all_records()
            if 1 < row_index <= len(records) + 1:
                return records[row_index - 2] 
            return {}
        except Exception as e:
            print(f"Error getting row data from {worksheet.title}: {e}")
            return {}

    def update_cell_by_sheet(self, worksheet, row_index, column_name, value):
        try:
            headers = worksheet.row_values(1)
            col_index = headers.index(column_name) + 1
            worksheet.update_cell(row_index, col_index, value)
            return True
        except Exception as e:
            print(f"Error updating cell [{row_index}, {column_name}] in {worksheet.title}: {e}")
            return False
        
    def get_rab_creator_by_ulok(self, nomor_ulok, lingkup_pekerjaan=None):
        try:
            rab_sheet = self.sheet.worksheet(config.DATA_ENTRY_SHEET_NAME)
            records = rab_sheet.get_all_records()
            
            # Mengubah "Z001-2512-D4D4-R" menjadi "Z0012512D4D4R" agar pencarian akurat
            target_ulok = str(nomor_ulok).replace("-", "").strip().upper()
            target_lingkup = str(lingkup_pekerjaan).strip().lower() if lingkup_pekerjaan else None

            # Loop dari bawah (terbaru)
            for record in reversed(records):
                # Ambil Ulok dari sheet dan normalisasi juga
                raw_ulok_sheet = str(record.get('Nomor Ulok', '')).replace("-", "").strip().upper()
                
                # Cek apakah Ulok cocok (setelah dinormalisasi)
                if raw_ulok_sheet == target_ulok:
                    # Jika lingkup_pekerjaan disertakan, cek kecocokannya
                    if target_lingkup:
                        # Cek kolom 'Lingkup_Pekerjaan' atau 'Lingkup Pekerjaan' (antisipasi spasi header)
                        rec_lingkup_raw = record.get('Lingkup_Pekerjaan', record.get('Lingkup Pekerjaan', ''))
                        rec_lingkup = str(rec_lingkup_raw).strip().lower()
                        
                        if rec_lingkup == target_lingkup:
                            return record.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)
                    else:
                        # Fallback jika tidak ada parameter lingkup
                        return record.get(config.COLUMN_NAMES.EMAIL_PEMBUAT)
            return None
        except Exception as e:
            print(f"Error saat mencari email pembuat RAB untuk Ulok {nomor_ulok}: {e}")
            return None
        
    def find_rejected_row_index(self, nomor_ulok, lingkup_pekerjaan):
        """
        Mencari index baris data yang DITOLAK untuk keperluan Revisi.
        """
        try:
            target_ulok = self._normalize_ulok(nomor_ulok)
            target_lingkup = self._normalize_lingkup(lingkup_pekerjaan)
            
            all_records = self.data_entry_sheet.get_all_records()
            
            for i, record in enumerate(all_records, start=2):
                status = str(record.get(config.COLUMN_NAMES.STATUS, "")).strip()

                if status in [config.STATUS.REJECTED_BY_COORDINATOR, config.STATUS.REJECTED_BY_MANAGER]:
                    rec_ulok_raw = str(record.get(config.COLUMN_NAMES.LOKASI, ""))
                    rec_lingkup_raw = record.get(
                        config.COLUMN_NAMES.LINGKUP_PEKERJAAN, 
                        record.get('Lingkup Pekerjaan', '')
                    )

                    rec_ulok = self._normalize_ulok(rec_ulok_raw)
                    rec_lingkup = self._normalize_lingkup(rec_lingkup_raw)

                    if rec_ulok == target_ulok and rec_lingkup == target_lingkup:
                        return i
            return None
        except Exception as e:
            print(f"Error searching for rejected row: {e}")
            return None
            
    def _get_col_letter(self, n):
        """Mengubah angka indeks kolom menjadi huruf Excel (misal: 1->A, 27->AA)"""
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    
    def ensure_header_exists(self, sheet_name, header_name):
        """
        Mengecek apakah header tertentu ada di baris 1.
        Jika tidak ada, tambahkan kolom baru di akhir header.
        """
        try:
            worksheet = self.sheet.worksheet(sheet_name)
            headers = worksheet.row_values(1)
            
            # Cek case-insensitive
            headers_lower = [h.strip().lower() for h in headers]
            
            if header_name.strip().lower() not in headers_lower:
                print(f"Header '{header_name}' tidak ditemukan di {sheet_name}. Menambahkan kolom baru...")
                # Tambahkan di kolom setelah kolom terakhir
                new_col_index = len(headers) + 1
                worksheet.update_cell(1, new_col_index, header_name)
                return True
            return False
        except Exception as e:
            print(f"Error ensuring header exists: {e}")
            return False

    def download_file_from_link(self, file_link):
        """Mendownload file dari link Google Drive publik/private"""
        try:
            print(f"DEBUG: Mencoba download dari link: {file_link}")
            if not file_link or "drive.google.com" not in file_link:
                print("DEBUG: Link tidak valid atau bukan Google Drive")
                return None, None, None

            # 1. Ekstrak File ID dari Link
            file_id = None
            if "/d/" in file_link:
                file_id = file_link.split("/d/")[1].split("/")[0]
            elif "id=" in file_link:
                file_id = file_link.split("id=")[1].split("&")[0]

            if not file_id: 
                print("DEBUG: File ID tidak ditemukan di link")
                return None, None, None

            # 2. Ambil Metadata (Nama File & MimeType)
            try:
                file_meta = self.drive_service.files().get(fileId=file_id, fields='name, mimeType').execute()
                filename = file_meta.get('name')
                mime_type = file_meta.get('mimeType')
            except Exception as e:
                print(f"DEBUG: Gagal ambil metadata file: {e}")
                # Fallback nama file jika metadata gagal
                return "Lampiran_Tambahan.pdf", None, "application/pdf"

            # 3. Download Content (Bytes)
            request = self.drive_service.files().get_media(fileId=file_id)
            file_stream = io.BytesIO()
            downloader = MediaIoBaseDownload(file_stream, request)

            done = False
            while done is False:
                status, done = downloader.next_chunk()

            file_stream.seek(0) # Reset pointer ke awal
            file_bytes = file_stream.read()

            print(f"DEBUG: Berhasil download {filename} ({len(file_bytes)} bytes)")
            return filename, file_bytes, mime_type

        except Exception as e:
            print(f"ERROR CRITICAL saat download file dari Drive: {e}")
            return None, None, None

    def check_ulok_exists_rab_2(self, nomor_ulok):
        """Mengecek apakah Nomor Ulok sudah ada di Spreadsheet RAB 2"""
        try:
            # Buka Spreadsheet RAB 2
            spreadsheet = self.gspread_client.open_by_key(config.SPREADSHEET_ID_RAB_2)
            worksheet = spreadsheet.worksheet(config.DATA_ENTRY_SHEET_NAME_RAB_2)
            
            # Ambil semua data
            all_records = worksheet.get_all_records()
            
            # Normalisasi input (hapus spasi/dash untuk perbandingan yang akurat)
            target_ulok = str(nomor_ulok).replace("-", "").strip().upper()

            for record in all_records:
                # Ambil Nomor Ulok dari record
                record_ulok = str(record.get(config.COLUMN_NAMES.LOKASI, "")).replace("-", "").strip().upper()
                
                if record_ulok == target_ulok:
                    # Jika ketemu, kembalikan statusnya
                    return {
                        "exists": True,
                        "status": record.get(config.COLUMN_NAMES.STATUS, "Unknown"),
                        "pembuat": record.get(config.COLUMN_NAMES.EMAIL_PEMBUAT, "-")
                    }
            
            return {"exists": False}
            
        except Exception as e:
            print(f"Error checking ULOK RAB 2: {e}")
            return {"exists": False, "error": str(e)}

    def get_nama_lengkap_dan_cabang_by_email(self, email):
        try:
            cabang_sheet = self.sheet.worksheet(config.CABANG_SHEET_NAME)
            for record in cabang_sheet.get_all_records():
                if str(record.get('EMAIL_SAT', '')).strip().lower() == email.lower():
                    return record.get('NAMA LENGKAP'), record.get('CABANG')
        except gspread.exceptions.WorksheetNotFound:
            print(f"Error: Worksheet '{config.CABANG_SHEET_NAME}' not found.")
        return None