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

def excel_to_word_by_cell(excel_path, template_path, output_path):
    """
    Mengisi template Word dengan data dari file Excel berdasarkan referensi sel.
    Mempertahankan format font saat mengganti placeholder.
    Menghilangkan prefix td_ dan tdi_ pada output.
    Mengubah karakter ü ke font Wingdings.
    Mengubah semua teks footer menjadi kapital.
    Mendukung placeholder tanggal.
    
    Parameters:
    - excel_path: Path ke file Excel yang berisi data
    - template_path: Path ke template Word
    - output_path: Path untuk menyimpan hasil output Word
    
    Returns:
    - bool: True jika berhasil, False jika gagal
    """
    print(f"Membuka file Excel: {excel_path}")
    print(f"Membuka template Word: {template_path}")
    
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
    
    print(f"Tanggal hari ini: {tanggal_lengkap}")
    
    # Buka workbook Excel menggunakan openpyxl untuk akses sel langsung
    try:
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        print(f"Berhasil membuka file Excel. Sheet yang tersedia: {workbook.sheetnames}")
    except Exception as e:
        print(f"Error saat membuka file Excel: {e}")
        return False
    
    # Fungsi untuk mendapatkan nilai sel dan memproses prefix
    def get_cell_value(sheet_name, cell_ref):
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
    
    # Baca template Word
    try:
        doc = Document(template_path)
        print(f"Berhasil membuka template Word. Memiliki {len(doc.paragraphs)} paragraf dan {len(doc.tables)} tabel.")
    except Exception as e:
        print(f"Error saat membuka template Word: {e}")
        return False
    
    # Pola untuk mendeteksi placeholder, mis: {{Sheet1.A1}} atau {{DATE.NOW}}
    # Modified pattern to support both Excel references and date placeholders
    pattern = r'\{\{([^}]+)\.([A-Z0-9_]+)\}\}'
    replacement_count = 0
    
    # Fungsi untuk mengubah font run ke Wingdings
    def set_wingdings_font(run):
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
    
    # Fungsi untuk mengganti placeholder dalam paragraf dengan mempertahankan format
    def replace_in_paragraph_runs(paragraph, is_footer=False):
        nonlocal replacement_count
        
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
                    placeholder_text = match.group(0)  # {{Sheet1.A1}} atau {{DATE.NOW}}
                    sheet_name = match.group(1)        # Sheet1 atau DATE
                    cell_ref = match.group(2)          # A1 atau NOW
                    
                    # Debug output
                    print(f"Menemukan placeholder: {placeholder_text}, sheet: {sheet_name}, ref: {cell_ref}")
                    
                    # Dapatkan nilai dari Excel atau placeholder tanggal
                    value = get_cell_value(sheet_name, cell_ref)
                    if value is not None:
                        value = str(value)
                        # Jika di footer, ubah ke kapital
                        if is_footer:
                            value = value.upper()
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
                    replacement_count += 1
        
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
                        set_wingdings_font(u_run)
                        
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
                        set_wingdings_font(u_run)
                        
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
    
    # Fungsi untuk memproses tabel
    def replace_in_tables(tables, is_footer=False):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph_runs(para, is_footer)
    
    # Fungsi untuk memproses header dan footer
    def replace_in_section_headers_footers(doc):
        for section in doc.sections:
            # Header
            header = section.header
            for para in header.paragraphs:
                replace_in_paragraph_runs(para)
            replace_in_tables(header.tables)
            
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
                    replace_in_paragraph_runs(para, is_footer=True)
                else:
                    # Untuk paragraf dengan field nomor halaman, hanya ubah teks biasa menjadi kapital
                    # tanpa mengganggu field nomor halaman
                    for run in para.runs:
                        if not (run._element.xpath('.//w:fldChar') or run._element.xpath('.//w:instrText')):
                            # Ubah ke kapital hanya jika bukan bagian dari field
                            run.text = run.text.upper()
            
            replace_in_tables(footer.tables, is_footer=True)
    
    # Proses paragraf utama
    for para in doc.paragraphs:
        replace_in_paragraph_runs(para)
    
    # Proses tabel
    replace_in_tables(doc.tables)
    
    # Proses header dan footer
    replace_in_section_headers_footers(doc)
    
    print(f"Total {replacement_count} penggantian dilakukan.")
    
    # Simpan hasil ke file baru
    try:
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
def generate_proposal(excel_path, template_path, output_path):
    """
    Fungsi utama untuk dipanggil dari aplikasi lain.
    
    Parameters:
    - excel_path: Path ke file Excel yang berisi data
    - template_path: Path ke template Word
    - output_path: Path untuk menyimpan hasil output Word
    
    Returns:
    - bool: True jika berhasil, False jika gagal
    """
    success = excel_to_word_by_cell(excel_path, template_path, output_path)
    
    # Jika berhasil generate proposal, update sel A1 pada sheet DATA_PROPOSAL
    if success:
        try:
            # Buka workbook Excel yang sama
            wb = openpyxl.load_workbook(excel_path)
            
            # Cek apakah sheet DATA_PROPOSAL ada
            if 'DATA_PROPOSAL' in wb.sheetnames:
                sheet = wb['DATA_PROPOSAL']
                
                # Dapatkan path relatif dari output_path
                root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                relative_path = os.path.relpath(output_path, root_dir)
                
                # Konversi ke format path Windows dengan backslash
                relative_path = relative_path.replace('/', '\\')
                
                # Update sel A1 dengan path relatif
                sheet['A1'] = relative_path
                print(f"Menyimpan path relatif '{relative_path}' ke sel A1 di sheet DATA_PROPOSAL")
                
                # Simpan workbook
                wb.save(excel_path)
                print(f"Berhasil memperbarui file Excel: {excel_path}")
            else:
                print("Sheet DATA_PROPOSAL tidak ditemukan di file Excel")
        except Exception as e:
            print(f"Error saat memperbarui sel A1 di sheet DATA_PROPOSAL: {e}")
    
    return success

# Jika script ini dijalankan secara langsung (bukan diimport)
if __name__ == "__main__":
    # Tentukan path relatif terhadap root project
    # Folder data akan berada di root project, bukan di folder modules
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_folder = os.path.join(root_dir, "data")
    
    # Nama file
    excel_filename = "SET_BDU.xlsx"
    template_filename = "Trial WWTP ANP Quotation Template.docx"
    output_filename = "WWTP_Quotation_Result.docx"
    
    # Path lengkap
    excel_file = os.path.join(data_folder, excel_filename)
    template_file = os.path.join(data_folder, template_filename)
    output_file = os.path.join(data_folder, output_filename)
    
    # Pastikan folder data ada
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
        print(f"Folder {data_folder} dibuat.")
    
    # Periksa apakah file yang diperlukan ada
    if not os.path.exists(excel_file):
        print(f"KESALAHAN: File Excel '{excel_file}' tidak ditemukan!")
    elif not os.path.exists(template_file):
        print(f"KESALAHAN: File template Word '{template_file}' tidak ditemukan!")
    else:
        print(f"Mulai memproses file...")
        # Panggil generate_proposal alih-alih excel_to_word_by_cell langsung
        # agar sel A1 pada DATA_PROPOSAL diperbarui
        generate_proposal(excel_file, template_file, output_file)