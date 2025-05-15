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
    Mendukung placeholder matematika $P1$ - $P4$ dengan nilai tetap.
    
    Parameters:
    - excel_path: Path ke file Excel yang berisi data
    - template_path: Path ke template Word
    - output_path: Path untuk menyimpan hasil output Word
    
    Returns:
    - bool: True jika berhasil, False jika gagal
    """
    print(f"Membuka file Excel: {excel_path}")
    print(f"Membuka template Word: {template_path}")
    print(f"Output akan disimpan ke: {output_path}")
    
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
            
        def get_cell_value(self, sheet_name, cell_ref):
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
            # Pola untuk mendeteksi placeholder Excel, mis: {{Sheet1.A1}} atau {{DATE.NOW}}
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
                        placeholder_text = match.group(0)  # {{Sheet1.A1}} atau {{DATE.NOW}}
                        sheet_name = match.group(1)        # Sheet1 atau DATE
                        cell_ref = match.group(2)          # A1 atau NOW
                        
                        # Debug output
                        print(f"Menemukan placeholder: {placeholder_text}, sheet: {sheet_name}, ref: {cell_ref}")
                        
                        # Dapatkan nilai dari Excel atau placeholder tanggal
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
    # Validasi path input
    if not os.path.exists(excel_path):
        print(f"Error: File Excel tidak ditemukan: {excel_path}")
        return False
        
    if not os.path.exists(template_path):
        print(f"Error: Template Word tidak ditemukan: {template_path}")
        return False
    
    # Pastikan direktori output ada
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"Membuat direktori output: {output_dir}")
        except Exception as e:
            print(f"Error membuat direktori output: {e}")
            return False
            
    # Jalankan fungsi untuk membuat dokumen
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
            else:
                print("Sheet DATA_PROPOSAL tidak ditemukan di file Excel")
        except Exception as e:
            print(f"Error saat memperbarui sel A1 di sheet DATA_PROPOSAL: {e}")
    
    return success

# Jika script ini dijalankan secara langsung (bukan diimport)
if __name__ == "__main__":
    """Jika file ini dijalankan langsung (bukan diimpor), 
    kode ini memungkinkan penggunaan customer path jika diberikan sebagai argumen.
    
    Penggunaan:
    python generate_proposal.py [customer_name]
    
    Jika customer_name diberikan, script akan memproses file SET_BDU.xlsx
    di folder customer tersebut dan menyimpan output di folder yang sama.
    """
    # Import argparse untuk menangani argumen command line
    import argparse
    
    # Buat parser argumen
    parser = argparse.ArgumentParser(description='Generate proposal from SET_BDU.xlsx')
    parser.add_argument('customer_name', nargs='?', help='Optional customer name')
    
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
        output_file = os.path.join(customer_folder, f"WWTP_Quotation_{args.customer_name}.docx")
        
        # Path template tetap mengacu ke folder data utama
        template_file = os.path.join(root_dir, "data", "Trial WWTP ANP Quotation Template.docx")
        
        print(f"Menggunakan file Excel customer: {excel_file}")
        print(f"Output akan disimpan ke: {output_file}")
    else:
        # Gunakan path default jika tidak ada customer_name
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
    output_folder = os.path.dirname(output_file)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)
        print(f"Folder {output_folder} dibuat.")
    
    # Periksa apakah file yang diperlukan ada
    if not os.path.exists(excel_file):
        print(f"KESALAHAN: File Excel '{excel_file}' tidak ditemukan!")
    elif not os.path.exists(template_file):
        print(f"KESALAHAN: File template Word '{template_file}' tidak ditemukan!")
    else:
        print(f"Mulai memproses file...")
        # Panggil generate_proposal
        generate_proposal(excel_file, template_file, output_file)