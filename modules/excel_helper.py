# modules/excel_helper.py
import os
import pandas as pd
import numpy as np

class ExcelHelper:
    """Helper class untuk membaca dan menulis data Excel"""
    
    @staticmethod
    def read_excel_file(file_path):
        """Membaca file Excel dan mengembalikan ExcelFile object"""
        try:
            if not os.path.exists(file_path):
                return None, f"File tidak ditemukan: {file_path}"
            
            excel_file = pd.ExcelFile(file_path)
            return excel_file, None
        except Exception as e:
            return None, f"Error membaca file Excel: {str(e)}"
    
    @staticmethod
    def get_dip_sheets(excel_file):
        """Mengembalikan daftar sheet yang dimulai dengan 'DIP'"""
        try:
            sheets = excel_file.sheet_names
            dip_sheets = [sheet for sheet in sheets if sheet.startswith('DIP')]
            return dip_sheets
        except Exception as e:
            return []
    
    @staticmethod
    def read_sheet_data(excel_file, sheet_name):
        """Membaca data dari sheet tertentu"""
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            return df, None
        except Exception as e:
            return None, f"Error membaca sheet {sheet_name}: {str(e)}"
    
    @staticmethod
    def save_data_to_excel(file_path, sheet_name, data_dict):
        """Menyimpan data ke file Excel"""
        try:
            # Baca file Excel yang ada
            if os.path.exists(file_path):
                writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
                workbook = writer.book
                
                # Jika sheet sudah ada, hapus dulu
                if sheet_name in workbook.sheetnames:
                    idx = workbook.sheetnames.index(sheet_name)
                    workbook.remove(workbook.worksheets[idx])
                    
                # Buat DataFrame dari data_dict
                df = pd.DataFrame(data_dict)
                
                # Simpan ke Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                writer.close()
                
                return True, None
            else:
                return False, f"File tidak ditemukan: {file_path}"
        except Exception as e:
            return False, f"Error menyimpan data ke Excel: {str(e)}"
    
    @staticmethod
    def parse_sheet_structure(df):
        """Parse struktur sheet untuk mendapatkan struktur data dengan format baru"""
        structure = {
            'sections': [],
            'fields': []
        }
        
        current_section = None
        field_id = 0
        current_column_headers = []
        has_column_headers = False
        
        # Proses setiap baris
        for idx, row in df.iterrows():
            # Skip baris kosong
            if pd.isna(row).all():
                continue
            
            # Ambil kolom pertama untuk menentukan tipe
            first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
            
            # Periksa apakah ini section header (sub_)
            if isinstance(first_col, str) and first_col.startswith('sub_'):
                section_title = first_col[4:].strip()  # Hapus prefix 'sub_'
                current_section = {
                    'id': len(structure['sections']),
                    'title': section_title,
                    'fields': []
                }
                structure['sections'].append(current_section)
                # Reset column headers when entering a new section
                current_column_headers = []
                has_column_headers = False
                continue
            
            # Periksa apakah ini field header (fh_)
            if isinstance(first_col, str) and first_col.startswith('fh_'):
                if current_section is None:
                    continue
                
                field_header = first_col[3:].strip()  # Hapus prefix 'fh_'
                current_section['current_header'] = field_header
                # Reset column headers when adding a new field header
                current_column_headers = []
                has_column_headers = False
                continue
            
            # Periksa apakah ini column header (ch_)
            if isinstance(first_col, str) and first_col.startswith('ch_'):
                has_column_headers = True
                current_column_headers = []
                
                # Process header row and collect all ch_ columns
                for col_idx in range(df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip()
                        if col_value.startswith('ch_'):
                            header_text = col_value[3:].strip()  # Remove 'ch_' prefix
                            current_column_headers.append(header_text)
                continue
            
            # Periksa apakah ini field (f_)
            if isinstance(first_col, str) and first_col.startswith('f_'):
                if current_section is None:
                    continue
                
                field_name = first_col[2:].strip()  # Hapus prefix 'f_'
                field_id += 1
                
                # Tentukan tipe field dan opsi jika ada
                field_type = "text"  # Tipe default
                options = []
                
                # Periksa kolom kedua (tipe atau opsi)
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    second_col = str(row.iloc[1]).strip()
                    
                    # Periksa apakah ini dropdown
                    if "dropdown" in second_col.lower():
                        field_type = "dropdown"
                        
                        # Ekstrak opsi jika ada
                        if len(row) > 2 and not pd.isna(row.iloc[2]):
                            options_str = str(row.iloc[2]).strip()
                            options = [opt.strip() for opt in options_str.split(',')]
                
                # Buat field
                field = {
                    'id': field_id,
                    'name': field_name,
                    'type': field_type,
                    'options': options,
                    'section_id': current_section['id'],
                    'header': current_section.get('current_header', '')
                }
                
                current_section['fields'].append(field_id)
                structure['fields'].append(field)
                continue
            
            # Periksa apakah ini field dropdown (fd_)
            if isinstance(first_col, str) and first_col.startswith('fd_'):
                if current_section is None:
                    continue
                
                field_name = first_col[3:].strip()  # Hapus prefix 'fd_'
                field_id += 1
                
                options = []
                # Get dropdown options from next column
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    options_str = str(row.iloc[1]).strip()
                    options = [opt.strip() for opt in options_str.split(',')]
                
                # Buat field
                field = {
                    'id': field_id,
                    'name': field_name,
                    'type': 'dropdown',
                    'options': options,
                    'section_id': current_section['id'],
                    'header': current_section.get('current_header', '')
                }
                
                current_section['fields'].append(field_id)
                structure['fields'].append(field)
                continue
            
            # Periksa apakah ini field multiple (fm_)
            if isinstance(first_col, str) and first_col.startswith('fm_'):
                if current_section is None:
                    continue
                
                field_name = first_col[3:].strip()  # Hapus prefix 'fm_'
                
                # Jika ada column headers, buat satu field untuk setiap header
                if has_column_headers and len(current_column_headers) > 0:
                    for i, header in enumerate(current_column_headers):
                        field_id += 1
                        
                        # Buat field untuk kolom ini
                        field = {
                            'id': field_id,
                            'name': f"{field_name} - {header}",
                            'type': 'text',
                            'options': [],
                            'section_id': current_section['id'],
                            'header': current_section.get('current_header', ''),
                            'parent_field': field_name,
                            'column_index': i
                        }
                        
                        current_section['fields'].append(field_id)
                        structure['fields'].append(field)
                else:
                    # Default 2 kolom jika tidak ada column headers
                    field_names = ["Name", "Contact"]
                    for i in range(2):
                        field_id += 1
                        
                        # Buat field untuk kolom ini
                        field = {
                            'id': field_id,
                            'name': f"{field_name} - {field_names[i]}",
                            'type': 'text',
                            'options': [],
                            'section_id': current_section['id'],
                            'header': current_section.get('current_header', ''),
                            'parent_field': field_name,
                            'column_index': i
                        }
                        
                        current_section['fields'].append(field_id)
                        structure['fields'].append(field)
                
                continue
                
        return structure