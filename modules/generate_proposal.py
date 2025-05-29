import pandas as pd
import openpyxl
from docx import Document
import re
import os
import datetime
from docx.oxml.shared import qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX

def clean_filename(filename):
    """
    Membersihkan nama file dari karakter yang tidak valid
    
    Parameters:
    - filename: string nama file
    
    Returns:
    - string: nama file yang sudah dibersihkan
    """
    # Karakter yang tidak diizinkan dalam nama file Windows
    invalid_chars = '<>:"/\\|?*'
    
    # Ganti karakter tidak valid dengan underscore
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Ganti multiple spaces dengan single space
    filename = re.sub(r'\s+', ' ', filename)
    
    # Ganti multiple underscores dengan single underscore
    filename = re.sub(r'_+', '_', filename)
    
    # Trim spaces di awal dan akhir
    filename = filename.strip()
    
    # Batasi panjang nama file (Windows limit ~255 karakter, kita buat 200 untuk safety)
    if len(filename) > 200:
        # Potong tapi pertahankan ekstensi
        name, ext = os.path.splitext(filename)
        filename = name[:200-len(ext)] + ext
    
    return filename

def get_quotation_number_by_company(workbook, company_name):
    """
    Mencari quotation number berdasarkan company name dari sheet 'No of Quotation'
    
    Parameters:
    - workbook: openpyxl workbook object
    - company_name: nama perusahaan untuk dicari
    
    Returns:
    - string: quotation number atau default jika tidak ditemukan
    """
    try:
        if 'No of Quotation' not in workbook.sheetnames:
            print("Sheet 'No of Quotation' tidak ditemukan")
            return "021/01/PTC/MBR/2025"  # Default fallback
        
        # Akses sheet menggunakan openpyxl
        sheet = workbook['No of Quotation']
        
        print(f"Mencari quotation number untuk company: '{company_name}'")
        
        # Baca header dari baris pertama untuk menentukan kolom
        headers = []
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                headers.append(str(cell_value).strip())
            else:
                headers.append(f"Col_{col}")
        
        print(f"Available columns: {headers}")
        
        # Cari indeks kolom yang diperlukan dengan exact match
        company_col_idx = None
        quotation_col_idx = None
        
        for idx, header in enumerate(headers):
            col_idx = idx + 1  # openpyxl uses 1-based indexing
            
            # Exact match untuk menghindari ambiguitas
            if header == 'Company Name':  # Harus exact match
                company_col_idx = col_idx
                print(f"Found 'Company Name' at column {col_idx}")
            elif header == 'Quotation No.' or header == 'Quotation No':  # Support both formats
                quotation_col_idx = col_idx
                print(f"Found 'Quotation No.' at column {col_idx}")
        
        if company_col_idx is None:
            print("Kolom 'Company Name' tidak ditemukan")
            print("Header yang tersedia:")
            for idx, header in enumerate(headers):
                print(f"  Column {idx+1}: '{header}'")
            return "021/01/PTC/MBR/2025"
            
        if quotation_col_idx is None:
            print("Kolom 'Quotation No.' tidak ditemukan")
            print("Header yang tersedia:")
            for idx, header in enumerate(headers):
                print(f"  Column {idx+1}: '{header}'")
            return "021/01/PTC/MBR/2025"
        
        print(f"Company column index: {company_col_idx}, Quotation column index: {quotation_col_idx}")
        
        # Bersihkan nama perusahaan untuk pencarian yang lebih akurat
        company_name_clean = str(company_name).strip().lower()
        
        # Cari data mulai dari baris 2 (baris 1 adalah header)
        available_companies = []
        
        for row in range(2, sheet.max_row + 1):
            company_cell = sheet.cell(row=row, column=company_col_idx).value
            quotation_cell = sheet.cell(row=row, column=quotation_col_idx).value
            
            if company_cell and quotation_cell:
                row_company = str(company_cell).strip()
                row_company_clean = row_company.lower()
                quotation_no = str(quotation_cell).strip()
                
                available_companies.append(row_company)
                
                # Coba exact match terlebih dahulu
                if row_company_clean == company_name_clean:
                    print(f"Exact match ditemukan: {quotation_no}")
                    return quotation_no
        
        # Jika exact match tidak ditemukan, coba partial match
        for row in range(2, sheet.max_row + 1):
            company_cell = sheet.cell(row=row, column=company_col_idx).value
            quotation_cell = sheet.cell(row=row, column=quotation_col_idx).value
            
            if company_cell and quotation_cell:
                row_company_clean = str(company_cell).strip().lower()
                quotation_no = str(quotation_cell).strip()
                
                # Partial match - cek apakah company name ada dalam row atau sebaliknya
                if (company_name_clean in row_company_clean) or (row_company_clean in company_name_clean):
                    print(f"Partial match ditemukan: {quotation_no}")
                    return quotation_no
        
        print(f"Company '{company_name}' tidak ditemukan dalam database quotation")
        print("Available companies:")
        for company in available_companies:
            print(f"  - {company}")
            
        return "021/01/PTC/MBR/2025"  # Default fallback
        
    except Exception as e:
        print(f"Error saat mengambil quotation number: {str(e)}")
        import traceback
        traceback.print_exc()
        return "021/01/PTC/MBR/2025"  # Default fallback

def get_proposal_data_for_filename(workbook):
    """
    Mengambil data yang diperlukan untuk membuat nama file proposal
    
    Parameters:
    - workbook: openpyxl workbook object
    
    Returns:
    - dict: data untuk nama file
    """
    data = {
        'project_type': '',
        'capacity': '',
        'company_name': '',
        'user_code': ''
    }
    
    try:
        # 1. Project Type dari DIP_Project_Information.B3
        if 'DIP_Project Information' in workbook.sheetnames:
            sheet = workbook['DIP_Project Information']
            try:
                project_type = sheet['B3'].value
                if project_type:
                    data['project_type'] = str(project_type).strip()
                    print(f"Project Type: {data['project_type']}")
            except Exception as e:
                print(f"Error membaca project type: {str(e)}")
        
        # 2. Capacity dari DIP_Project_Information.B60
        if 'DIP_Project Information' in workbook.sheetnames:
            sheet = workbook['DIP_Project Information']
            try:
                capacity = sheet['B60'].value
                if capacity:
                    data['capacity'] = str(capacity).strip()
                    print(f"Capacity: {data['capacity']}")
            except Exception as e:
                print(f"Error membaca capacity: {str(e)}")
        
        # 3. Company Name dari DIP_Customer_Information.B4
        if 'DIP_Customer Information' in workbook.sheetnames:
            sheet = workbook['DIP_Customer Information']
            try:
                company_name = sheet['B4'].value
                if company_name:
                    data['company_name'] = str(company_name).strip()
                    print(f"Company Name: {data['company_name']}")
            except Exception as e:
                print(f"Error membaca company name: {str(e)}")
        
        # 4. User Code dari DATA_TEMP.B1
        if 'DATA_TEMP' in workbook.sheetnames:
            sheet = workbook['DATA_TEMP']
            try:
                user_code = sheet['B1'].value
                if user_code:
                    data['user_code'] = str(user_code).strip()
                    print(f"User Code: {data['user_code']}")
            except Exception as e:
                print(f"Error membaca user code: {str(e)}")
        
        return data
        
    except Exception as e:
        print(f"Error mengambil data untuk filename: {str(e)}")
        return data

def generate_dynamic_filename(workbook, fallback_customer_name=None, version="01"):
    """
    Generate nama file proposal yang dinamis
    
    Format: Commercial and Technical_Project Type Capacity CMD_Company Name_Ver.01_BDE/PSE Code
    
    Parameters:
    - workbook: openpyxl workbook object
    - fallback_customer_name: nama customer fallback jika tidak ada di Excel
    - version: versi proposal (default "01")
    
    Returns:
    - string: nama file yang sudah dibersihkan
    """
    try:
        # Ambil data dari Excel
        data = get_proposal_data_for_filename(workbook)
        
        # Component 1: Fixed prefix
        prefix = "Commercial and Technical"
        
        # Component 2: Project Type
        project_type = data['project_type'] if data['project_type'] else "WWTP Project"
        
        # Component 3: Capacity dengan CMD
        capacity = data['capacity'] if data['capacity'] else "100"
        capacity_part = f"{capacity} CMD"
        
        # Component 4: Company Name
        company_name = data['company_name'] if data['company_name'] else fallback_customer_name if fallback_customer_name else "Unknown Company"
        
        # Component 5: Version
        version_part = f"Ver.{version}"
        
        # Component 6: User Code
        user_code = data['user_code'] if data['user_code'] else "000"
        
        # Gabungkan semua komponen
        filename_parts = [
            prefix,
            f"{project_type} {capacity_part}",
            company_name,
            version_part,
            user_code
        ]
        
        # Gabung dengan separator underscore
        filename = "_".join(filename_parts)
        
        # Tambahkan ekstensi
        filename += ".docx"
        
        # Bersihkan nama file
        clean_name = clean_filename(filename)
        
        print(f"Generated filename: {clean_name}")
        return clean_name
        
    except Exception as e:
        print(f"Error generating dynamic filename: {str(e)}")
        # Fallback ke nama sederhana
        fallback_name = f"WWTP_Quotation_{fallback_customer_name or 'Result'}_{version}.docx"
        return clean_filename(fallback_name)

def get_user_data_by_code(workbook, user_code):
    """
    Mengambil data user berdasarkan user code dari sheet 'User Code'
    
    Parameters:
    - workbook: openpyxl workbook object
    - user_code: string user code yang dipilih
    
    Returns:
    - dict: data user atau dict kosong jika tidak ditemukan
    """
    try:
        if 'User Code' not in workbook.sheetnames:
            print("Sheet 'User Code' tidak ditemukan")
            return {}
        
        # Akses sheet menggunakan openpyxl
        sheet = workbook['User Code']
        
        print(f"Mencari user data untuk code: '{user_code}'")
        
        # Baca header dari baris pertama untuk menentukan kolom
        headers = []
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=1, column=col).value
            if cell_value:
                headers.append(str(cell_value).strip())
            else:
                headers.append(f"Col_{col}")
        
        print(f"Available columns in User Code sheet: {headers}")
        
        # Cari indeks kolom yang diperlukan
        code_col_idx = None
        
        # Mapping kolom untuk data user
        column_mapping = {}
        
        for idx, header in enumerate(headers):
            col_idx = idx + 1  # openpyxl uses 1-based indexing
            
            if 'Code' in header:
                code_col_idx = col_idx
                column_mapping['Code'] = col_idx
            elif 'Name' in header:
                column_mapping['Name'] = col_idx
            elif 'Position' in header:
                column_mapping['Position'] = col_idx
            elif 'Email' in header:
                column_mapping['Email'] = col_idx
            elif 'Mobile' in header or 'Phone' in header:
                column_mapping['Mobile'] = col_idx
            else:
                # Tambahkan kolom lain yang mungkin ada
                column_mapping[header] = col_idx
        
        if code_col_idx is None:
            print("Kolom 'Code' tidak ditemukan dalam sheet User Code")
            return {}
        
        print(f"Column mapping: {column_mapping}")
        
        # Bersihkan user code untuk pencarian
        user_code_clean = str(user_code).strip()
        
        # Cari data mulai dari baris 2 (baris 1 adalah header)
        available_codes = []
        
        for row in range(2, sheet.max_row + 1):
            code_cell = sheet.cell(row=row, column=code_col_idx).value
            
            if code_cell:
                row_code = str(code_cell).strip()
                available_codes.append(row_code)
                
                # Coba exact match
                if row_code == user_code_clean:
                    print(f"User code match ditemukan untuk: {user_code_clean}")
                    
                    # Ambil semua data untuk user ini
                    user_data = {}
                    for column_name, col_idx in column_mapping.items():
                        cell_value = sheet.cell(row=row, column=col_idx).value
                        if cell_value is not None:
                            user_data[column_name] = str(cell_value).strip()
                        else:
                            user_data[column_name] = ""
                    
                    print(f"Data user ditemukan untuk code '{user_code}': {user_data.get('Name', 'Unknown')}")
                    return user_data
        
        print(f"User code '{user_code}' tidak ditemukan dalam database")
        print("Available user codes:")
        for code in available_codes:
            print(f"  - {code}")
            
        return {}
        
    except Exception as e:
        print(f"Error saat mengambil data user: {str(e)}")
        import traceback
        traceback.print_exc()
        return {}

def get_selected_user_code_from_excel(workbook):
    """
    Mengambil user code yang dipilih dari DATA_TEMP.B1
    """
    try:
        # Lokasi penyimpanan selected user code: DATA_TEMP.B1
        if 'DATA_TEMP' in workbook.sheetnames:
            sheet = workbook['DATA_TEMP']
            try:
                cell_value = sheet['B1'].value
                if cell_value and str(cell_value).strip():
                    print(f"Selected user code ditemukan di DATA_TEMP.B1: {cell_value}")
                    return str(cell_value).strip()
                else:
                    print("DATA_TEMP.B1 kosong atau tidak valid")
                    return None
            except Exception as e:
                print(f"Error membaca DATA_TEMP.B1: {str(e)}")
                return None
        else:
            print("Sheet DATA_TEMP tidak ditemukan")
            return None
        
    except Exception as e:
        print(f"Error saat mencari selected user code: {str(e)}")
        return None

def excel_to_word_by_cell(excel_path, template_path, output_path, selected_user_code=None):
    """
    Mengisi template Word dengan data dari file Excel berdasarkan referensi sel.
    Mempertahankan format font saat mengganti placeholder.
    Menghilangkan prefix td_ dan tdi_ pada output.
    Mengubah karakter ü ke font Wingdings.
    Mengubah semua teks footer menjadi kapital.
    Mendukung placeholder tanggal.
    Mendukung placeholder matematika $P1$ - $P4$ dengan nilai tetap.
    Mendukung placeholder USER_CODE untuk data contact person.
    Mendukung placeholder QUOTATION_NO untuk nomor quotation otomatis.
    
    Parameters:
    - excel_path: Path ke file Excel yang berisi data
    - template_path: Path ke template Word
    - output_path: Path untuk menyimpan hasil output Word
    - selected_user_code: User code yang dipilih (optional)
    
    Returns:
    - bool: True jika berhasil, False jika gagal
    """
    print(f"Membuka file Excel: {excel_path}")
    print(f"Membuka template Word: {template_path}")
    print(f"Output akan disimpan ke: {output_path}")
    print(f"Selected user code: {selected_user_code}")
    
    # Siapkan data tanggal untuk placeholder
    now = datetime.datetime.now()
    
    # Format tanggal dengan format DD MM YYYY
    hari = now.strftime("%d")      # Format 2 digit: 01, 02, ..., 31
    bulan = now.strftime("%m")     # Format 2 digit: 01, 02, ..., 12
    tahun = now.strftime("%Y")     # Format 4 digit: 2025
    tanggal_lengkap = f"{hari} {bulan} {tahun}"
    
    # Buat data tanggal untuk placeholder
    tanggal_data = {
        "NOW": tanggal_lengkap,
        "DAY": hari,
        "MONTH": bulan,
        "YEAR": tahun
    }
    
    # Siapkan data untuk placeholder matematika
    math_placeholders = {
        "$P1$": "30",
        "$P2$": "50",
        "$P3$": "15",
        "$P4$": "5"
    }
    
    print(f"Tanggal hari ini: {tanggal_lengkap}")
    print(f"Placeholder matematika yang akan diganti: {', '.join(math_placeholders.keys())}")
    
    # Buka workbook Excel menggunakan openpyxl untuk akses sel langsung
    try:
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        print(f"Berhasil membuka file Excel. Sheet yang tersedia: {workbook.sheetnames}")
    except Exception as e:
        print(f"Error saat membuka file Excel: {e}")
        return False
    
    # Ambil selected user code jika tidak diberikan sebagai parameter
    if not selected_user_code:
        selected_user_code = get_selected_user_code_from_excel(workbook)
    
    # Ambil data user berdasarkan selected user code
    user_data = {}
    if selected_user_code:
        user_data = get_user_data_by_code(workbook, selected_user_code)
    
    # Siapkan data USER_CODE untuk placeholder
    user_code_data = {
        "NAME": user_data.get('Name', 'Tia Amelia'),  # Default fallback
        "POSITION": user_data.get('Position', 'Product Strategist Engineer'),
        "EMAIL": user_data.get('Email', 'tiamalia@grinvirobiotekno.com'),
        "MOBILE": user_data.get('Mobile', '+62 856-5504-9457')
    }
    
    print(f"Data USER_CODE yang akan digunakan: {user_code_data}")
    
    # Ambil company name untuk mencari quotation number
    company_name = ""
    try:
        if 'DIP_Customer Information' in workbook.sheetnames:
            sheet = workbook['DIP_Customer Information']
            company_cell = sheet['B4'].value
            if company_cell:
                company_name = str(company_cell).strip()
                print(f"Company name ditemukan: {company_name}")
    except Exception as e:
        print(f"Error mengambil company name: {str(e)}")
    
    # Dapatkan quotation number berdasarkan company name
    quotation_number = get_quotation_number_by_company(workbook, company_name)
    
    # Siapkan data QUOTATION_NO untuk placeholder
    quotation_data = {
        "NO": quotation_number
    }
    
    print(f"Data QUOTATION_NO yang akan digunakan: {quotation_data}")
    
    # Baca template Word
    try:
        doc = Document(template_path)
        print(f"Berhasil membuka template Word. Memiliki {len(doc.paragraphs)} paragraf dan {len(doc.tables)} tabel.")
    except Exception as e:
        print(f"Error saat membuka template Word: {e}")
        return False
    
    # Kelas untuk menghandle penggantian placeholder
    class Replacer:
        def __init__(self):
            self.replacement_count = 0
            self.math_replacement_count = 0
            self.user_code_replacement_count = 0
            self.quotation_no_replacement_count = 0
            
        def get_cell_value(self, sheet_name, cell_ref):
            # Cek apakah ini placeholder USER_CODE
            if sheet_name == "USER_CODE" and cell_ref in user_code_data:
                value = user_code_data[cell_ref]
                print(f"Menggunakan placeholder USER_CODE.{cell_ref} = {value}")
                self.user_code_replacement_count += 1
                return value
            
            # Cek apakah ini placeholder QUOTATION_NO
            if sheet_name == "QUOTATION_NO" and cell_ref in quotation_data:
                value = quotation_data[cell_ref]
                print(f"Menggunakan placeholder QUOTATION_NO.{cell_ref} = {value}")
                self.quotation_no_replacement_count += 1
                return value
            
            # Cek apakah ini placeholder tanggal khusus
            if sheet_name == "DATE" and cell_ref in tanggal_data:
                value = tanggal_data[cell_ref]
                print(f"Menggunakan placeholder tanggal DATE.{cell_ref} = {value}")
                return value
                
            # Kode yang sudah ada untuk mengambil nilai dari Excel
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                try:
                    cell_value = sheet[cell_ref].value
                    
                    # Proses nilai string dengan awalan td_ atau tdi_
                    if isinstance(cell_value, str):
                        original_value = cell_value
                        
                        # Menghilangkan prefix td_ dan tdi_
                        if cell_value.startswith("td_"):
                            cell_value = cell_value[3:]  # Hapus 3 karakter pertama (td_)
                            print(f"Membaca {sheet_name}.{cell_ref} = {original_value} -> Menghapus prefix td_ -> {cell_value}")
                        elif cell_value.startswith("tdi_"):
                            cell_value = cell_value[4:]  # Hapus 4 karakter pertama (tdi_)
                            print(f"Membaca {sheet_name}.{cell_ref} = {original_value} -> Menghapus prefix tdi_ -> {cell_value}")
                        else:
                            print(f"Membaca {sheet_name}.{cell_ref} = {cell_value}")
                    else:
                        print(f"Membaca {sheet_name}.{cell_ref} = {cell_value}")
                        
                    return cell_value
                except Exception as e:
                    print(f"Error pada sel {sheet_name}.{cell_ref}: {e}")
                    return f"ERROR: Invalid cell reference {cell_ref}"
            else:
                print(f"Sheet {sheet_name} tidak ditemukan")
                return f"ERROR: Sheet {sheet_name} not found"
        
        def set_wingdings_font(self, run):
            run.font.name = "Wingdings"
            # Menambahkan properti langsung ke XML untuk memastikan perubahan font
            run_properties = run._element.get_or_add_rPr()
            fonts = run_properties.xpath('./w:rFonts')
            if fonts:
                for font in fonts:
                    font.set(qn('w:ascii'), "Wingdings")
                    font.set(qn('w:hAnsi'), "Wingdings")
                    font.set(qn('w:cs'), "Wingdings")
            else:
                font_element = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Wingdings"/>')
                run_properties.append(font_element)
        
        def replace_math_placeholders(self, text):
            """Mengganti placeholder matematika $P1$-$P4$ dengan nilai tetap"""
            original_text = text
            replaced = False
            
            # Cari semua placeholder matematika dalam teks
            for placeholder, value in math_placeholders.items():
                if placeholder in text:
                    text = text.replace(placeholder, value)
                    self.math_replacement_count += 1
                    replaced = True
                    print(f"Mengganti placeholder matematika {placeholder} dengan {value}")
            
            return text, replaced
        
        def replace_in_paragraph_runs(self, paragraph, is_footer=False):
            """Mengganti placeholder dalam paragraf dengan mempertahankan format"""
            # Pola untuk mendeteksi placeholder Excel, mis: {{Sheet1.A1}}, {{DATE.NOW}}, {{USER_CODE.NAME}}, atau {{QUOTATION_NO.NO}}
            pattern = r'\{\{([^}]+)\.([A-Z0-9_]+)\}\}'
            
            # Sebelum memproses placeholder Excel, cek apakah ada placeholder matematika
            for i, run in enumerate(paragraph.runs):
                # Cek dan ganti placeholder matematika
                if any(math_ph in run.text for math_ph in math_placeholders.keys()):
                    new_text, replaced = self.replace_math_placeholders(run.text)
                    if replaced:
                        run.text = new_text
            
            # Kita perlu melacak perubahan pada struktur run, karena ini bisa berubah saat kita memodifikasi
            orig_runs = list(paragraph.runs)
            placeholder_runs = {}  # Menyimpan info tentang placeholder di run mana

            # Identifikasi run mana yang berisi bagian dari placeholder
            for i, run in enumerate(orig_runs):
                if '{{' in run.text:
                    # Kemungkinan awal dari placeholder
                    full_placeholder = run.text
                    run_index = i
                    
                    # Jika placeholder terbagi di beberapa run, kita perlu menyatukannya
                    while '}}' not in full_placeholder and run_index < len(orig_runs) - 1:
                        run_index += 1
                        full_placeholder += orig_runs[run_index].text
                    
                    # Ekstrak bagian placeholder menggunakan regex
                    matches = list(re.finditer(pattern, full_placeholder))
                    for match in matches:
                        # Simpan informasi tentang placeholder ini
                        placeholder_text = match.group(0)  # {{Sheet1.A1}}, {{USER_CODE.NAME}}, {{QUOTATION_NO.NO}}, dll
                        sheet_name = match.group(1)        # Sheet1, USER_CODE, DATE, QUOTATION_NO
                        cell_ref = match.group(2)          # A1, NAME, NOW, NO
                        
                        # Debug output
                        print(f"Menemukan placeholder: {placeholder_text}, sheet: {sheet_name}, ref: {cell_ref}")
                        
                        # Dapatkan nilai dari Excel, USER_CODE, QUOTATION_NO, atau placeholder tanggal
                        value = self.get_cell_value(sheet_name, cell_ref)
                        if value is not None:
                            value = str(value)
                            # Jika di footer, ubah ke kapital
                            if is_footer:
                                value = value.upper()
                            
                            # Periksa apakah nilai mengandung placeholder matematika dan ganti jika diperlukan
                            for math_ph, math_val in math_placeholders.items():
                                if math_ph in value:
                                    original_value = value
                                    value = value.replace(math_ph, math_val)
                                    self.math_replacement_count += 1
                                    print(f"Mengganti placeholder matematika {math_ph} dalam nilai '{original_value}' dengan '{value}'")
                        else:
                            value = ""
                        
                        # Simpan informasi untuk pemrosesan nanti
                        placeholder_runs[i] = {
                            'start_run': i,
                            'end_run': run_index,
                            'placeholder': placeholder_text,
                            'value': value,
                            'contains_umlaut': 'ü' in value
                        }
                        self.replacement_count += 1
            
            # Proses penggantian dimulai dari run terakhir untuk mencegah pergeseran indeks
            for start_run_idx in sorted(placeholder_runs.keys(), reverse=True):
                info = placeholder_runs[start_run_idx]
                
                # Kasus sederhana: placeholder ada dalam satu run
                if info['start_run'] == info['end_run']:
                    run = orig_runs[info['start_run']]
                    
                    # Cek apakah nilai mengandung karakter ü
                    if 'ü' in info['value']:
                        # Split teks berdasarkan karakter ü
                        parts = info['value'].split('ü')
                        
                        # Hapus placeholder di run asli dan ganti dengan bagian awal
                        run.text = run.text.replace(info['placeholder'], parts[0])
                        
                        # Untuk setiap ü dan teks setelahnya
                        for i in range(len(parts) - 1):
                            # Buat run baru untuk karakter ü dengan font Wingdings
                            u_run = paragraph.add_run('ü')
                            self.set_wingdings_font(u_run)
                            
                            # Buat run baru untuk teks setelah ü (jika ada)
                            if parts[i+1]:
                                paragraph.add_run(parts[i+1])
                    else:
                        # Tidak ada ü, cukup ganti placeholder
                        run.text = run.text.replace(info['placeholder'], info['value'])
                
                # Kasus kompleks: placeholder terbagi di beberapa run
                else:
                    # Ambil run pertama dan terakhir
                    first_run = orig_runs[info['start_run']]
                    last_run = orig_runs[info['end_run']]
                    
                    # Dapatkan teks sebelum placeholder di run pertama
                    before_text = first_run.text.split('{{')[0]
                    
                    # Dapatkan teks setelah placeholder di run terakhir
                    after_text = last_run.text.split('}}')[1] if '}}' in last_run.text else ''
                    
                    # Cek apakah nilai mengandung karakter ü
                    if 'ü' in info['value']:
                        # Split teks berdasarkan karakter ü
                        parts = info['value'].split('ü')
                        
                        # Set bagian awal ke run pertama
                        first_run.text = before_text + parts[0]
                        
                        # Hapus atau kosongkan run lain yang merupakan bagian dari placeholder
                        for i in range(info['start_run'] + 1, info['end_run'] + 1):
                            orig_runs[i].text = ''
                        
                        # Untuk setiap ü dan teks setelahnya
                        for i in range(len(parts) - 1):
                            # Buat run baru untuk karakter ü dengan font Wingdings
                            u_run = paragraph.add_run('ü')
                            self.set_wingdings_font(u_run)
                            
                            # Jika ini adalah bagian terakhir dan ada after_text
                            if i == len(parts) - 2 and after_text:
                                paragraph.add_run(parts[i+1] + after_text)
                            else:
                                # Buat run baru untuk teks setelah ü (jika ada)
                                if parts[i+1]:
                                    paragraph.add_run(parts[i+1])
                    else:
                        # Tidak ada ü, proses normal
                        first_run.text = before_text + info['value']
                        
                        # Hapus atau kosongkan run lain yang merupakan bagian dari placeholder
                        for i in range(info['start_run'] + 1, info['end_run'] + 1):
                            orig_runs[i].text = ''
                        
                        # Tambahkan teks setelah ke run terakhir jika ada
                        if after_text:
                            last_run.text = after_text
        
        def replace_in_tables(self, tables, is_footer=False):
            """Memproses tabel"""
            for table in tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            self.replace_in_paragraph_runs(para, is_footer)
        
        def replace_in_section_headers_footers(self, doc):
            """Memproses header dan footer"""
            for section in doc.sections:
                # Header
                header = section.header
                for para in header.paragraphs:
                    self.replace_in_paragraph_runs(para)
                self.replace_in_tables(header.tables)
                
                # Footer - dengan parameter is_footer=True untuk konversi ke kapital
                footer = section.footer
                for para in footer.paragraphs:
                    # Periksa apakah paragraf berisi field nomor halaman
                    contains_page_field = False
                    for run in para.runs:
                        # Field nomor halaman biasanya direpresentasikan dengan karakter khusus atau pattern tertentu
                        if run._element.xpath('.//w:fldChar') or "PAGE" in run.text or run._element.xpath('.//w:instrText'):
                            contains_page_field = True
                            print(f"Menemukan field nomor halaman, akan membiarkan paragraf ini utuh")
                            break
                    
                    # Hanya proses paragraf yang tidak berisi field nomor halaman
                    if not contains_page_field:
                        self.replace_in_paragraph_runs(para, is_footer=True)
                    else:
                        # Untuk paragraf dengan field nomor halaman, hanya ubah teks biasa menjadi kapital
                        # tanpa mengganggu field nomor halaman
                        for run in para.runs:
                            if not (run._element.xpath('.//w:fldChar') or run._element.xpath('.//w:instrText')):
                                # Ubah ke kapital hanya jika bukan bagian dari field
                                run.text = run.text.upper()
                
                self.replace_in_tables(footer.tables, is_footer=True)
        
        def process_document(self, doc):
            """Memproses seluruh dokumen"""
            # Proses paragraf utama
            for para in doc.paragraphs:
                self.replace_in_paragraph_runs(para)
            
            # Proses tabel
            self.replace_in_tables(doc.tables)
            
            # Proses header dan footer
            self.replace_in_section_headers_footers(doc)
            
            # Laporan hasil
            print(f"Total {self.replacement_count} penggantian placeholder Excel dilakukan.")
            print(f"Total {self.math_replacement_count} penggantian placeholder matematika dilakukan.")
            print(f"Total {self.user_code_replacement_count} penggantian placeholder USER_CODE dilakukan.")
            print(f"Total {self.quotation_no_replacement_count} penggantian placeholder QUOTATION_NO dilakukan.")
    
    # Gunakan kelas Replacer untuk memproses dokumen
    replacer = Replacer()
    replacer.process_document(doc)
    
    # Simpan hasil ke file baru
    try:
        # Pastikan direktori output ada
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        doc.save(output_path)
        print(f"Dokumen berhasil dihasilkan: {output_path}")
        success = True
    except Exception as e:
        print(f"Error saat menyimpan dokumen Word: {e}")
        success = False
    
    # Tutup workbook Excel
    workbook.close()
    
    return success

# Fungsi utama yang bisa dipanggil dari aplikasi lain
def generate_proposal(excel_path, template_path, output_dir, selected_user_code=None, customer_name=None, version="01"):
    """
    Fungsi utama untuk dipanggil dari aplikasi lain dengan dynamic filename.
    
    Parameters:
    - excel_path: Path ke file Excel yang berisi data
    - template_path: Path ke template Word
    - output_dir: Directory untuk menyimpan hasil output Word
    - selected_user_code: User code yang dipilih dari dropdown
    - customer_name: Nama customer untuk fallback filename
    - version: Versi proposal (default "01")
    
    Returns:
    - tuple: (bool success, str output_path) - True jika berhasil dan path file hasil
    """
    # Validasi path input
    if not os.path.exists(excel_path):
        print(f"Error: File Excel tidak ditemukan: {excel_path}")
        return False, ""
        
    if not os.path.exists(template_path):
        print(f"Error: Template Word tidak ditemukan: {template_path}")
        return False, ""
    
    # Pastikan direktori output ada
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"Membuat direktori output: {output_dir}")
        except Exception as e:
            print(f"Error membuat direktori output: {e}")
            return False, ""
    
    try:
        # Buka workbook untuk generate filename
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        
        # Generate nama file dinamis
        dynamic_filename = generate_dynamic_filename(
            workbook, 
            fallback_customer_name=customer_name, 
            version=version
        )
        
        # Tutup workbook sementara
        workbook.close()
        
        # Path output lengkap
        output_path = os.path.join(output_dir, dynamic_filename)
        
        print(f"Menggunakan nama file: {dynamic_filename}")
        
    except Exception as e:
        print(f"Error generating dynamic filename: {str(e)}")
        # Fallback ke nama file sederhana
        fallback_filename = f"WWTP_Quotation_{customer_name or 'Result'}_{version}.docx"
        output_path = os.path.join(output_dir, clean_filename(fallback_filename))
        print(f"Menggunakan nama file fallback: {os.path.basename(output_path)}")
            
    # Jalankan fungsi untuk membuat dokumen
    success = excel_to_word_by_cell(excel_path, template_path, output_path, selected_user_code)
    
    # Jika berhasil generate proposal, update sel A1 pada sheet DATA_PROPOSAL
    if success:
        try:
            # Buka workbook Excel yang sama
            wb = openpyxl.load_workbook(excel_path)
            
            # Cek apakah sheet DATA_PROPOSAL ada
            if 'DATA_PROPOSAL' in wb.sheetnames:
                sheet = wb['DATA_PROPOSAL']
                
                # Dapatkan path relatif dari output_path
                # Jika output_path adalah di folder customer, buat path relatif ke customer folder
                
                # Path dasar adalah root project
                root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                
                # Jika output_path merupakan subpath dari root_dir, buat path relatif
                if os.path.commonpath([root_dir]) == os.path.commonpath([root_dir, output_path]):
                    relative_path = os.path.relpath(output_path, root_dir)
                else:
                    # Jika berbeda drive atau tidak ada path umum, gunakan path absolut
                    relative_path = output_path
                
                # Konversi ke format path Windows dengan backslash
                relative_path = relative_path.replace('/', '\\')
                
                # Update sel A1 dengan path relatif
                sheet['A1'] = relative_path
                print(f"Menyimpan path relatif '{relative_path}' ke sel A1 di sheet DATA_PROPOSAL")
                
                # Simpan workbook
                wb.save(excel_path)
                print(f"Berhasil memperbarui file Excel: {excel_path}")
                wb.close()
            else:
                print("Sheet DATA_PROPOSAL tidak ditemukan di file Excel")
                wb.close()
        except Exception as e:
            print(f"Error saat memperbarui sel A1 di sheet DATA_PROPOSAL: {e}")
    
    return success, output_path if success else ""

# Jika script ini dijalankan secara langsung (bukan diimport)
if __name__ == "__main__":
    """Jika file ini dijalankan langsung (bukan diimpor), 
    kode ini memungkinkan penggunaan customer path jika diberikan sebagai argumen.
    
    Penggunaan:
    python generate_proposal.py [customer_name] [user_code] [version]
    
    Jika customer_name diberikan, script akan memproses file SET_BDU.xlsx
    di folder customer tersebut dan menyimpan output di folder yang sama.
    
    Jika user_code diberikan, script akan menggunakan data user tersebut.
    Jika version diberikan, akan digunakan untuk penamaan file.
    """
    # Import argparse untuk menangani argumen command line
    import argparse
    
    # Buat parser argumen
    parser = argparse.ArgumentParser(description='Generate proposal from SET_BDU.xlsx with dynamic filename and quotation number lookup')
    parser.add_argument('customer_name', nargs='?', help='Optional customer name')
    parser.add_argument('user_code', nargs='?', help='Optional user code')
    parser.add_argument('version', nargs='?', default='01', help='Optional version number (default: 01)')
    
    # Parse argumen
    args = parser.parse_args()
    
    # Tentukan path relatif terhadap root project
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    if args.customer_name:
        # Jika customer_name diberikan, gunakan path khusus customer
        from modules.fix_customer_system import clean_folder_name
        
        # Buat path folder customer
        customer_folder = os.path.join(root_dir, "data", "customers", clean_folder_name(args.customer_name))
        
        # Path file customer
        excel_file = os.path.join(customer_folder, "SET_BDU.xlsx")
        output_dir = customer_folder
        
        # Path template tetap mengacu ke folder data utama
        template_file = os.path.join(root_dir, "data", "Trial WWTP ANP Quotation Template.docx")
        
        print(f"Menggunakan file Excel customer: {excel_file}")
        print(f"Output akan disimpan ke: {output_dir}")
    else:
        # Gunakan path default jika tidak ada customer_name
        data_folder = os.path.join(root_dir, "data")
        
        # Nama file
        excel_filename = "SET_BDU.xlsx"
        template_filename = "Trial WWTP ANP Quotation Template.docx"
        
        # Path lengkap
        excel_file = os.path.join(data_folder, excel_filename)
        template_file = os.path.join(data_folder, template_filename)
        output_dir = data_folder
    
    # Pastikan folder data ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        print(f"Folder {output_dir} dibuat.")
    
    # Periksa apakah file yang diperlukan ada
    if not os.path.exists(excel_file):
        print(f"KESALAHAN: File Excel '{excel_file}' tidak ditemukan!")
    elif not os.path.exists(template_file):
        print(f"KESALAHAN: File template Word '{template_file}' tidak ditemukan!")
    else:
        print(f"Mulai memproses file...")
        # Panggil generate_proposal dengan parameter lengkap
        success, output_file = generate_proposal(
            excel_file, 
            template_file, 
            output_dir, 
            selected_user_code=args.user_code,
            customer_name=args.customer_name,
            version=args.version
        )
        
        if success:
            print(f"✅ Proposal berhasil digenerate: {output_file}")
        else:
            print("❌ Gagal menggenerate proposal")