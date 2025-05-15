import os
import sys
import io
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFrame, QGridLayout, QSpacerItem,
                             QSizePolicy, QScrollArea, QApplication, QMenu, QAction,
                             QTabWidget, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QFileDialog, QDateEdit, QCheckBox)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QCursor, QImage
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QPoint, QDate, QThread
import tempfile
import shutil
import subprocess

# Import local modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_NAME, SECONDARY_COLOR, PRIMARY_COLOR, BG_COLOR, DEPARTMENTS

INDUSTRY_SUBTYPE_MAPPING = {
    "Palm Oil": ["CPO&CPKO", "EFB", "Palm Oil Plantation"],
    "Mining": ["Coal Mining", "Gold Mining", "Nickel Mining", "Tin Mining", "Bauxite Mining"],
    "Oil & Gas": ["Upstream", "Midstream", "Downstream"],
    "Non Food Industry-A": ["Apparel & Footwear (Textile)", "Manufacturing/Heavy Industry", 
                           "Technology & Telecommunication", "Transportation & Logistics"],
    "F&B": ["Processed Food", "Beverages", "Dairy Products", "Confectionary", 
           "Meat Processing", "Seasoning"],
    "Agro Industry": ["Fishery & Aquaculture Products", "Food Crops (Cofee, Cocoa)", 
                     "Tobacco", "Sugar", "Livestock & Poultry"],
    "Non Food Industry-B": ["Tourism & Hospitality", "Construction & Real Estate", "Residential"]
}

SEISMIC_ZONE_DESCRIPTIONS = {
    "ZONE-1": "2.5 or less. Usually not felt, but can be recorded by seismographs",
    "ZONE-2": "2.5 - 5.4. Often felt, but causes only minor damage",
    "ZONE-3": "5.5 - 6.0. Can cause slight damage to buildings and other structures",
    "ZONE-4": "6.6 - 6.9. Can cause significant damage in populated areas",
    "Zone-5": "7.0 - 7.9. Major earthquake with serious damage.",
    "Zone-6": "8.0 or larger. Great earthquake. Can destroy communities near the epicenter."
}

WIND_SPEED_DESCRIPTIONS = {
    "LEVEL-1": "0 - 2,0 m/s",
    "LEVEL-2": "2,0 - 4,0 m/s",
    "LEVEL-3": "4,0 - 6,0 m/s",
    "LEVEL-4": "6,0 - 8,0 m/s",
    "LEVEL-5": "8,0 - 10,0 m/s",
    "LEVEL-6": "10,0 m/s"
}

# Pengecekan opsional untuk modul docx2pdf
HAS_DOCX2PDF = False
try:
    import docx2pdf
    HAS_DOCX2PDF = True
except ImportError:
    pass

# Pengecekan opsional untuk modul PyMuPDF (fitz)
HAS_PYMUPDF = False
try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    pass

# Kelas thread untuk mengkonversi dokumen Word ke gambar thumbnail
class ThumbnailGeneratorThread(QThread):
    # Signal saat thumbnail siap
    thumbnail_ready = pyqtSignal(int, QPixmap)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, file_path, max_pages=100):  # Ubah max_pages menjadi 100
        super().__init__()
        self.file_path = file_path
        self.max_pages = max_pages
        self.temp_dir = None
        
    def run(self):
        try:
             # Inisialisasi COM library untuk thread ini
            import pythoncom
            pythoncom.CoInitialize()
            
            # Buat direktori sementara
            self.temp_dir = tempfile.mkdtemp()
            
            # Konversi Word ke PDF jika itu docx dan modul docx2pdf tersedia
            _, ext = os.path.splitext(self.file_path)
            if ext.lower() == '.docx' and HAS_DOCX2PDF:
                pdf_path = os.path.join(self.temp_dir, "temp.pdf")
                docx2pdf.convert(self.file_path, pdf_path)
                
                # Gunakan PDF untuk menghasilkan thumbnail jika PyMuPDF tersedia
                if HAS_PYMUPDF:
                    self.generate_thumbnails_from_pdf(pdf_path)
                else:
                    self.generate_thumbnails_alternative()
            else:
                # Gunakan pendekatan alternatif untuk file non-PDF dan jika docx2pdf tidak tersedia
                self.generate_thumbnails_alternative()
                
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))
        finally:
            # Bersihkan direktori sementara saat selesai
            if self.temp_dir and os.path.exists(self.temp_dir):
                try:
                    shutil.rmtree(self.temp_dir)
                except:
                    pass
                
            # Uninisialisasi COM library
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass

    def generate_thumbnails_from_pdf(self, pdf_path):
        """Menghasilkan thumbnail dari file PDF menggunakan PyMuPDF jika tersedia"""
        try:
            # Pastikan PyMuPDF tersedia
            if not HAS_PYMUPDF:
                self.error.emit("PyMuPDF (modul fitz) tidak tersedia. Menggunakan metode alternatif.")
                self.generate_thumbnails_alternative()
                return
                
            doc = fitz.open(pdf_path)
            # Proses semua halaman tanpa batasan 10
            for page_num in range(min(self.max_pages, doc.page_count)):
                page = doc.load_page(page_num)
                
                # Render halaman sebagai gambar
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                img_data = pix.samples
                
                # Buat QImage dari data piksel
                img = QImage(img_data, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                
                # Kirim thumbnail ke UI
                self.thumbnail_ready.emit(page_num, pixmap)
            
            doc.close()
        except Exception as e:
            self.error.emit(f"Error PyMuPDF: {str(e)}. Menggunakan metode alternatif.")
            self.generate_thumbnails_alternative()
    
    def generate_thumbnails_alternative(self):
        """Metode alternatif untuk menghasilkan thumbnail sederhana"""
        try:
            import docx
            
            # Buat thumbnail sederhana yang menunjukkan judul dokumen dan beberapa baris teks
            doc = docx.Document(self.file_path)
            
            # Hitung jumlah halaman (perkiraan)
            total_paragraphs = len(doc.paragraphs)
            estimated_pages = max(1, total_paragraphs // 20)  # Asumsikan ~20 paragraf per halaman
            
            # Tampilkan semua halaman yang diperkirakan, tanpa dibatasi 10
            for page_num in range(min(self.max_pages, estimated_pages)):
                # Buat gambar kosong
                width, height = 800, 1000
                img = QImage(width, height, QImage.Format_RGB888)
                img.fill(Qt.white)
                
                # Buat objek painter
                painter = QtGui.QPainter(img)
                painter.setFont(QFont("Arial", 12))
                
                # Gambar judul dokumen
                painter.setFont(QFont("Arial", 16, QFont.Bold))
                painter.drawText(40, 40, "Page {}".format(page_num + 1))
                
                # Gambar beberapa teks dari dokumen
                painter.setFont(QFont("Arial", 11))
                
                y_pos = 80
                start_para = page_num * 20
                end_para = min(start_para + 20, total_paragraphs)
                
                for i in range(start_para, end_para):
                    text = doc.paragraphs[i].text
                    if text:
                        wrapped_text = self.wrap_text(painter, text, width - 80)
                        for line in wrapped_text:
                            painter.drawText(40, y_pos, line)
                            y_pos += 20
                            if y_pos > height - 40:
                                break
                    y_pos += 10
                    if y_pos > height - 40:
                        break
                
                painter.end()
                
                # Konversi ke QPixmap
                pixmap = QPixmap.fromImage(img)
                
                # Kirim thumbnail ke UI
                self.thumbnail_ready.emit(page_num, pixmap)
        except Exception as e:
            self.error.emit(f"Gagal membuat thumbnail alternatif: {str(e)}")
            
            # Buat thumbnail sangat sederhana sebagai fallback terakhir
            self.generate_simple_fallback(estimated_pages=min(100, estimated_pages))  # Tingkatkan jumlah halaman
    
    def generate_simple_fallback(self, estimated_pages=100):
        """Membuat thumbnail sederhana sebagai fallback terakhir"""
        try:
            for page_num in range(min(self.max_pages, estimated_pages)):
                # Buat gambar kosong
                width, height = 800, 1000
                img = QImage(width, height, QImage.Format_RGB888)
                img.fill(Qt.white)
                
                # Buat objek painter
                painter = QtGui.QPainter(img)
                
                # Gambar judul dokumen
                painter.setFont(QFont("Arial", 16, QFont.Bold))
                painter.drawText(40, 40, "Page {}".format(page_num + 1))
                
                # Gambar pesan kesalahan atau info
                painter.setFont(QFont("Arial", 12))
                painter.drawText(40, 80, "Pratinjau dokumen tidak tersedia.")
                painter.drawText(40, 110, "Silakan gunakan tombol 'Open in Word'")
                painter.drawText(40, 140, "untuk membuka dokumen langsung di Word.")
                
                # Gambar ikon dokumen besar di tengah
                painter.setFont(QFont("Arial", 100))
                painter.drawText(width/2 - 80, height/2, "📄")
                
                painter.end()
                
                # Konversi ke QPixmap
                pixmap = QPixmap.fromImage(img)
                
                # Kirim thumbnail ke UI
                self.thumbnail_ready.emit(page_num, pixmap)
        except Exception as e:
            self.error.emit(f"Gagal membuat thumbnail fallback: {str(e)}")
    
    def wrap_text(self, painter, text, max_width):
        """Membantu memotong teks untuk muat dalam lebar tertentu"""
        words = text.split()
        lines = []
        current_line = ""
        
        for word in words:
            test_line = current_line + " " + word if current_line else word
            width = painter.fontMetrics().horizontalAdvance(test_line)
            
            if width <= max_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        
        if current_line:
            lines.append(current_line)
            
        return lines
class BDUGroupView(QMainWindow):
    """View untuk BDU Group"""
    back_to_dashboard = pyqtSignal()
    
    def __init__(self, auth_manager):
        super().__init__()
        self.auth_manager = auth_manager
        self.current_user = auth_manager.get_current_user()
        self.excel_data = None
        self.sheet_tabs = {}
        self.data_fields = {}
        self.linked_dropdowns = {}
        
        # Excel path
        self.excel_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "SET_BDU.xlsx")
        
        self.initUI()
        self.load_excel_data()
    
    def initUI(self):
        """Initialize the UI"""
        # Set window properties
        self.setWindowTitle(f"BDU Group - {APP_NAME}")
        self.setMinimumSize(1000, 700)
        
        # Set central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Header
        self.setup_header()
        main_layout.addWidget(self.header_widget)
        
        # Content area
        content_widget = QScrollArea()
        content_widget.setWidgetResizable(True)
        content_widget.setStyleSheet(f"background-color: {BG_COLOR}; border: none;")
        
        # Main content
        self.main_content = QWidget()
        self.main_content.setStyleSheet(f"background-color: {BG_COLOR};")
        
        content_layout = QVBoxLayout(self.main_content)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)
        
        # Page title
        title_layout = QHBoxLayout()
        
        # Back button
        back_btn = QPushButton()
        back_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ArrowBack))
        back_btn.setFixedSize(36, 36)
        back_btn.setCursor(Qt.PointingHandCursor)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {PRIMARY_COLOR};
                border-radius: 18px;
                color: white;
            }}
            QPushButton:hover {{
                background-color: #1f2c39;
            }}
        """)
        back_btn.clicked.connect(self.go_back_to_dashboard)
        
        page_title = QLabel("BDU Group")
        page_title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        page_title.setStyleSheet(f"color: {PRIMARY_COLOR};")
        
        # Get BDU icon from departments
        bdu_dept = next((dept for dept in DEPARTMENTS if dept["id"] == "BDU"), None)
        if bdu_dept:
            dept_icon = QLabel(bdu_dept["emoji"])
            dept_icon.setFont(QFont("Segoe UI", 22))
            dept_icon.setStyleSheet(f"color: {bdu_dept['color']};")
            title_layout.addWidget(dept_icon)
        
        title_layout.addWidget(back_btn)
        title_layout.addSpacing(10)
        title_layout.addWidget(page_title)
        title_layout.addStretch()
        
        # Add refresh button
        refresh_btn = QPushButton("Refresh Data")
        refresh_btn.setFont(QFont("Segoe UI", 10))
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {SECONDARY_COLOR};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: #2980B9;
            }}
        """)
        refresh_btn.clicked.connect(self.load_excel_data)
        
        title_layout.addWidget(refresh_btn)
        
        content_layout.addLayout(title_layout)
        
        # Tab widget for different sheets
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid #cccccc;
                background-color: white;
                border-radius: 5px;
            }}
            QTabBar::tab {{
                background-color: #f0f0f0;
                color: #333333;
                padding: 8px 15px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                border: 1px solid #cccccc;
                border-bottom: none;
            }}
            QTabBar::tab:selected {{
                background-color: white;
                border-bottom-color: white;
            }}
            QTabBar::tab:hover {{
                background-color: #e0e0e0;
            }}
        """)
        
        content_layout.addWidget(self.tab_widget)
        
        # Loading message
        self.loading_label = QLabel("Loading data from SET_BDU.xlsx...")
        self.loading_label.setFont(QFont("Segoe UI", 12))
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setStyleSheet("color: #666; margin: 20px;")
        
        content_layout.addWidget(self.loading_label)
        
        # Set the main content to the scroll area
        content_widget.setWidget(self.main_content)
        main_layout.addWidget(content_widget)
        
        # Status bar
        self.statusBar().showMessage(f"BDU Group Module | User: {self.current_user['username']}")
        self.statusBar().setStyleSheet("background-color: #f0f0f0; color: #555;")
    
    def setup_header(self):
        """Setup header widget dengan logo, judul, dan menu"""
        self.header_widget = QWidget()
        self.header_widget.setFixedHeight(60)
        self.header_widget.setStyleSheet(f"background-color: {PRIMARY_COLOR};")
        
        header_layout = QHBoxLayout(self.header_widget)
        header_layout.setContentsMargins(15, 0, 15, 0)
        
        # BDU title with icon
        bdu_dept = next((dept for dept in DEPARTMENTS if dept["id"] == "BDU"), None)
        if bdu_dept:
            dept_icon = QLabel(bdu_dept["emoji"])
            dept_icon.setFont(QFont("Segoe UI", 18))
            dept_icon.setStyleSheet("color: white;")
            header_layout.addWidget(dept_icon)
        
        # App title
        title_label = QLabel(f"{APP_NAME} - BDU Group")
        title_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title_label.setStyleSheet("color: white;")
        
        # Spacer
        spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        # User button (menggunakan teks emoji)
        user_menu_btn = QPushButton("👤")
        user_menu_btn.setFont(QFont("Segoe UI", 14))
        user_menu_btn.setStyleSheet("""
            QPushButton {
                border: none;
                background-color: transparent;
                color: white;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.2);
                border-radius: 5px;
            }
        """)
        user_menu_btn.setFixedSize(36, 36)
        user_menu_btn.setCursor(Qt.PointingHandCursor)
        user_menu_btn.clicked.connect(self.show_user_menu)
        
        # Add all elements to header layout
        header_layout.addWidget(title_label)
        header_layout.addItem(spacer)
        header_layout.addWidget(user_menu_btn)
    
    def show_user_menu(self):
        """Tampilkan menu pengguna"""
        sender = self.sender()
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 5px;
            }
            QMenu::item {
                padding: 5px 15px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #f0f0f0;
            }
        """)
        
        # Add user info at the top (non-clickable)
        user_info = QAction(f"{self.current_user['name']} ({self.current_user['department']})", self)
        user_info.setEnabled(False)
        menu.addAction(user_info)
        
        menu.addSeparator()
        
        # Add menu actions
        dashboard_action = QAction("Back to Dashboard", self)
        settings_action = QAction("Settings", self)
        
        menu.addAction(dashboard_action)
        menu.addAction(settings_action)
        
        # Connect actions
        dashboard_action.triggered.connect(self.go_back_to_dashboard)
        
        # Show menu at button position
        menu.exec_(sender.mapToGlobal(QPoint(0, sender.height())))
    
    def go_back_to_dashboard(self):
        """Kembali ke dashboard"""
        self.back_to_dashboard.emit()
        self.close()
    
    def update_dependent_dropdown(self, parent_value, child_dropdown):
        """Update dropdown yang bergantung berdasarkan nilai dropdown utama"""
        # Bersihkan dropdown anak
        child_dropdown.clear()
        
        # Simpan nilai sub-industry saat ini agar bisa digunakan nanti
        current_sub_industry = getattr(child_dropdown, "_last_value", None)
        
        # Isi dengan nilai-nilai yang sesuai berdasarkan pilihan pada dropdown utama
        if parent_value in INDUSTRY_SUBTYPE_MAPPING:
            child_dropdown.addItems(INDUSTRY_SUBTYPE_MAPPING[parent_value])
        else:
            # Jika tidak ada mapping, tambahkan placeholder atau biarkan kosong
            child_dropdown.addItem("-- Select Sub Industry --")
            
    def get_absolute_path(self, relative_path):
        """Konversi path relatif menjadi absolut relatif terhadap root project"""
        if os.path.isabs(relative_path):
            return relative_path
        
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(project_root, relative_path)
    
    def load_excel_data(self):
        """Load data from SET_BDU.xlsx"""
        try:
            if not os.path.exists(self.excel_path):
                self.loading_label.setText(f"Error: File SET_BDU.xlsx not found in the data directory.")
                self.loading_label.setStyleSheet("color: #E74C3C; margin: 20px;")
                return

            # Clear existing tabs
            self.tab_widget.clear()
            self.sheet_tabs = {}
            self.data_fields = {}

            # Hide loading message when tabs exist
            self.loading_label.setVisible(True)

            # Read Excel file
            xl = pd.ExcelFile(self.excel_path)
            sheet_names = xl.sheet_names

            # Variabel untuk melacak posisi proposal sheet dalam daftar
            proposal_sheet_index = -1
            for i, sheet_name in enumerate(sheet_names):
                if sheet_name == "DATA_PROPOSAL":
                    proposal_sheet_index = i
                    break

            # Filter sheets to only include those starting with DATA_ or DIP_
            filtered_sheets = [sheet for sheet in sheet_names if sheet.startswith("DATA_") or sheet.startswith("DIP_")]

            if len(filtered_sheets) == 0:
                self.loading_label.setText("No DATA_ or DIP_ sheets found in SET_BDU.xlsx.")
                return

            # Hide loading label as we have data
            self.loading_label.setVisible(False)

            # Create a tab for each filtered sheet
            for sheet_name in filtered_sheets:
                # Get display name (remove DIP_ or DATA_ prefix)
                display_name = sheet_name
                if sheet_name.startswith("DIP_"):
                    display_name = sheet_name[4:]  # Remove "DIP_"
                elif sheet_name.startswith("DATA_"):
                    display_name = sheet_name[5:]  # Remove "DATA_"

                try:
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
                    
                    # Khusus untuk DATA_PROPOSAL
                    if sheet_name == "DATA_PROPOSAL":
                        # Check apakah cell A1 ada dan berisi file path
                        file_exists = False
                        relative_word_file_path = ""
                        
                        if not df.empty and not pd.isna(df.iloc[0, 0]):
                            relative_word_file_path = str(df.iloc[0, 0]).strip()
                            # Konversi ke path absolut untuk pengecekan file
                            absolute_word_file_path = self.get_absolute_path(relative_word_file_path)
                            file_exists = os.path.exists(absolute_word_file_path)
                        
                        # Buat tab PROPOSAL dengan tampilan yang sesuai
                        scroll_area = QScrollArea()
                        scroll_area.setWidgetResizable(True)
                        
                        # Buat widget untuk proposal
                        proposal_widget = QWidget()
                        proposal_layout = QVBoxLayout(proposal_widget)
                        proposal_layout.setContentsMargins(10, 10, 10, 0)
                        proposal_layout.setSpacing(0)
                        
                        # Jika file ada, langsung proses dokumen
                        if file_exists:
                            # Process the Word document - simpan path relatif untuk referensi, tapi gunakan path absolut untuk proses
                            self.proposal_relative_path = relative_word_file_path
                            if self.process_proposal_document(absolute_word_file_path):
                                # Add the proposal document widget
                                if hasattr(self, 'proposal_document_widget'):
                                    proposal_layout.addWidget(self.proposal_document_widget)
                        else:
                            # Jika file TIDAK ada, tampilkan header dengan icon dan tombol
                            # Container untuk header file yang tidak ada
                            top_container = QWidget()
                            top_layout = QHBoxLayout(top_container)
                            top_layout.setContentsMargins(0, 0, 0, 0)
                            
                            # Panel kiri untuk ikon dan label
                            left_panel = QWidget()
                            left_panel.setMaximumWidth(300)
                            left_layout = QHBoxLayout(left_panel)
                            left_layout.setContentsMargins(0, 0, 0, 0)
                            left_layout.setSpacing(5)
                            
                            # Ikon file
                            file_icon = QLabel("📄")
                            file_icon.setFont(QFont("Segoe UI", 14))
                            file_icon.setStyleSheet("color: #3498DB; background-color: transparent;")
                            left_layout.addWidget(file_icon)
                            
                            # Label status proposal
                            file_label = QLabel("Proposal Not Generated")
                            file_label.setFont(QFont("Segoe UI", 11))
                            file_label.setStyleSheet("color: #333; background-color: transparent;")
                            left_layout.addWidget(file_label)
                            left_layout.addStretch()
                            
                            # Panel kanan untuk tombol
                            right_panel = QWidget()
                            right_layout = QHBoxLayout(right_panel)
                            right_layout.setContentsMargins(0, 0, 0, 0)
                            right_layout.setAlignment(Qt.AlignRight | Qt.AlignTop)
                            
                            # Tambahkan tombol Run Projection
                            run_projection_btn = QPushButton("Run Projection")
                            run_projection_btn.setFont(QFont("Segoe UI", 10))
                            run_projection_btn.setCursor(Qt.PointingHandCursor)
                            run_projection_btn.setStyleSheet(f"""
                                QPushButton {{
                                    background-color: #3498db;
                                    color: white;
                                    border: none;
                                    border-radius: 4px;
                                    padding: 5px 10px;
                                    margin-right: 5px;
                                }}
                                QPushButton:hover {{
                                    background-color: #2980b9;
                                }}
                            """)
                            # Connect to function if needed
                            run_projection_btn.clicked.connect(self.run_projection)
                            right_layout.addWidget(run_projection_btn)
                            
                            # Tambahkan tombol Generate Proposal
                            generate_btn = QPushButton("Generate Proposal")
                            generate_btn.setFont(QFont("Segoe UI", 10))
                            generate_btn.setCursor(Qt.PointingHandCursor)
                            generate_btn.setStyleSheet(f"""
                                QPushButton {{
                                    background-color: #27ae60;
                                    color: white;
                                    border: none;
                                    border-radius: 4px;
                                    padding: 5px 10px;
                                }}
                                QPushButton:hover {{
                                    background-color: #2ecc71;
                                }}
                            """)
                            generate_btn.clicked.connect(self.run_generate_proposal)
                            right_layout.addWidget(generate_btn)
                            
                            # Menyusun panel kiri dan kanan
                            top_layout.addWidget(left_panel)
                            top_layout.addWidget(right_panel)
                            
                            # Tambahkan container utama ke layout proposal
                            proposal_layout.addWidget(top_container)
                            
                            # Tambahkan instruksi yang singkat
                            info_label = QLabel("Click 'Generate Proposal' button to create a new proposal document.")
                            info_label.setAlignment(Qt.AlignCenter)
                            info_label.setFont(QFont("Segoe UI", 10))
                            info_label.setStyleSheet("color: #666;")
                            info_label.setMaximumHeight(20)
                            proposal_layout.addWidget(info_label)
                            
                            # Tambahkan spacer
                            proposal_layout.addStretch(1)
                        
                        scroll_area.setWidget(proposal_widget)
                        self.tab_widget.addTab(scroll_area, display_name)
                        self.sheet_tabs[sheet_name] = proposal_widget
                    else:
                        # Regular processing for other sheets
                        scroll_area = QScrollArea()
                        scroll_area.setWidgetResizable(True)
                        
                        sheet_widget = QWidget()
                        sheet_layout = QVBoxLayout(sheet_widget)
                        sheet_layout.setContentsMargins(15, 15, 15, 15)
                        
                        # Process the sheet data
                        self.process_sheet_data(df, sheet_name, sheet_layout)
                        
                        scroll_area.setWidget(sheet_widget)
                        self.tab_widget.addTab(scroll_area, display_name)
                        self.sheet_tabs[sheet_name] = sheet_widget
                        
                except Exception as e:
                    print(f"Error processing sheet {sheet_name}: {str(e)}")
                    # Create an error tab for this sheet
                    error_widget = QWidget()
                    error_layout = QVBoxLayout(error_widget)
                    
                    error_label = QLabel(f"Error loading {sheet_name}: {str(e)}")
                    error_label.setStyleSheet("color: #E74C3C;")
                    error_layout.addWidget(error_label)
                    
                    self.tab_widget.addTab(error_widget, display_name)
                    
            # Setelah selesai memuat semua sheet dan tab, tambahkan kode untuk menyimpan nilai yang saat ini dipilih
            for sheet_name, sheet_widget in self.sheet_tabs.items():
                # Cari dropdown Sub Industry
                for key, widget in self.data_fields.items():
                    if isinstance(widget, QComboBox) and key.startswith(sheet_name):
                        # Simpan nilai saat ini ke dalam property widget
                        widget._last_value = widget.currentText()
            
            # Hapus pesan loading status bar  
            self.statusBar().clearMessage()
                
        except Exception as e:
            self.loading_label.setText(f"Error loading data: {str(e)}")
            self.loading_label.setStyleSheet("color: #E74C3C; margin: 20px;")
            self.loading_label.setVisible(True)
            print(f"Error loading Excel data: {str(e)}")
            
    def run_projection(self):
        """Fungsi untuk menjalankan projection dari data BDU ke ANAPAK ke PUMP dan kembali ke BDU"""
        try:
            import os
            import pandas as pd
            from openpyxl import load_workbook
            import time
            import sys
            import subprocess
            from PyQt5.QtWidgets import QApplication, QMessageBox
            
            # Use the customer-specific Excel file if available
            if hasattr(self, 'excel_path'):
                set_bdu_path = self.excel_path
                print(f"Using customer-specific SET_BDU.xlsx: {set_bdu_path}")
            else:
                # Fall back to default path if customer file isn't set
                data_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
                set_bdu_path = os.path.join(data_folder, "SET_BDU.xlsx")
                print(f"No customer file path found, using default SET_BDU.xlsx: {set_bdu_path}")
            
            # Base data folder path
            data_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
            
            # Path ke file-file Excel
            sbt_anapak_path = os.path.join(data_folder, "SBT_ANAPAK.xlsx")
            sbt_pump_path = os.path.join(data_folder, "SBT_PUMP.xlsm")
            all_udf_path = os.path.join(data_folder, "ALL_UDF.py")
            
            # Pastikan file yang dibutuhkan ada
            missing_files = []
            for file_path in [set_bdu_path, sbt_anapak_path, sbt_pump_path, all_udf_path]:
                if not os.path.exists(file_path):
                    missing_files.append(os.path.basename(file_path))
            
            if missing_files:
                QMessageBox.critical(
                    self, 
                    "File Tidak Ditemukan", 
                    f"Beberapa file berikut tidak ditemukan di folder data:\n{', '.join(missing_files)}"
                )
                return
            
            # Tampilkan pesan status
            self.statusBar().showMessage("Menjalankan projection - Transfer data dari SET_BDU ke SBT_ANAPAK...")
            QApplication.processEvents()
            
            # PROSES 1: Transfer data dari SET_BDU ke SBT_ANAPAK
            # Buka workbook SET_BDU untuk membaca data
            wb_bdu = load_workbook(set_bdu_path, data_only=True)
            
            # Ambil data yang diperlukan dari DIP_Technical Information
            sheet_dip = wb_bdu["DIP_Technical Information"]
            values_to_transfer = {
                'C37': sheet_dip['B9'].value,  # B9 -> C37
                'C38': sheet_dip['B10'].value, # B10 -> C38
                'C39': sheet_dip['B11'].value, # B11 -> C39
                'C40': sheet_dip['B12'].value, # B12 -> C40
                'C41': sheet_dip['B13'].value, # B13 -> C41
                'C42': sheet_dip['B16'].value, # B16 -> C42
                'C43': sheet_dip['B17'].value, # B17 -> C43
                'C46': sheet_dip['B21'].value  # B21 -> C46
            }
            
            # Tutup workbook SET_BDU
            wb_bdu.close()
            
            # Buka dan update workbook SBT_ANAPAK
            wb_anapak = load_workbook(sbt_anapak_path)
            sheet_anapak = wb_anapak["ANAPAK"]
            
            # Pindahkan data ke SBT_ANAPAK
            for cell_addr, value in values_to_transfer.items():
                sheet_anapak[cell_addr] = value
            
            # Simpan SBT_ANAPAK
            wb_anapak.save(sbt_anapak_path)
            wb_anapak.close()
            
            # LANGKAH BARU: Buka SBT_ANAPAK dengan Excel dan paksa kalkulasi
            self.statusBar().showMessage("Menjalankan projection - Mengkalkulasi SBT_ANAPAK...")
            QApplication.processEvents()
            
            try:
                # Buat VBS script untuk membuka dan mengkalkulasi SBT_ANAPAK
                import tempfile
                
                # Buat file VBS temporal
                vbs_calc_file = tempfile.NamedTemporaryFile(delete=False, suffix='.vbs')
                vbs_calc_path = vbs_calc_file.name
                
                # Tulis skrip VBS untuk membuka Excel, mengkalkulasi, dan menyimpan
                vbs_calc_script = f'''
                Set objExcel = CreateObject("Excel.Application")
                objExcel.DisplayAlerts = False
                objExcel.Visible = False
                
                ' Buka file SBT_ANAPAK.xlsx
                Set objWorkbook = objExcel.Workbooks.Open("{sbt_anapak_path}")
                
                ' Paksa kalkulasi untuk semua sheet
                objExcel.CalculateFullRebuild
                objWorkbook.Application.Calculate
                
                ' Simpan dan tutup
                objWorkbook.Save
                objWorkbook.Close
                objExcel.Quit
                
                Set objWorkbook = Nothing
                Set objExcel = Nothing
                '''
                
                vbs_calc_file.write(vbs_calc_script.encode('utf-8'))
                vbs_calc_file.close()
                
                # Jalankan script VBS
                if sys.platform == 'win32':
                    subprocess.call(['cscript.exe', '//nologo', vbs_calc_path])
                else:
                    QMessageBox.warning(
                        self, 
                        "Warning", 
                        "Menjalankan VBScript di platform non-Windows tidak didukung."
                    )
                
                # Hapus file VBS sementara
                try:
                    os.unlink(vbs_calc_path)
                except:
                    pass
                
                # Tunggu sebentar untuk memastikan file tersimpan
                time.sleep(2)
                
            except Exception as e:
                QMessageBox.warning(
                    self, 
                    "Warning", 
                    f"Gagal menjalankan kalkulasi otomatis SBT_ANAPAK: {str(e)}\nMelanjutkan proses..."
                )
            
            self.statusBar().showMessage("Menjalankan projection - Transfer data dari SBT_ANAPAK ke SBT_PUMP...")
            QApplication.processEvents()
            
            # PROSES 2: Ambil data dari SBT_ANAPAK untuk dimasukkan ke SBT_PUMP
            # Buka lagi SBT_ANAPAK untuk ambil data terbaru setelah perhitungan
            wb_anapak = load_workbook(sbt_anapak_path, data_only=True)
            sheet_anapak = wb_anapak["ANAPAK"]
            
            # Ambil nilai dari SBT_ANAPAK.ANAPAK.I66 dan K67
            value_i66 = sheet_anapak['I66'].value
            value_k67 = sheet_anapak['K67'].value
            
            # Tampilkan nilai untuk debugging
            print(f"DEBUG: Nilai dari SBT_ANAPAK.ANAPAK.I66: {value_i66}")
            print(f"DEBUG: Nilai dari SBT_ANAPAK.ANAPAK.K67: {value_k67}")
            
            # Periksa apakah nilai masih None
            if value_i66 is None or value_k67 is None:
                # Tampilkan pesan dan minta pengguna untuk membuka file secara manual
                QMessageBox.warning(
                    self,
                    "Nilai Kosong",
                    "Sel I66 dan/atau K67 di file SBT_ANAPAK.xlsx masih kosong setelah kalkulasi. "
                    "Hal ini mungkin karena:\n\n"
                    "1. File SBT_ANAPAK memerlukan kalkulasi manual\n"
                    "2. Rumus di I66/K67 perlu diperbarui\n\n"
                    "Klik OK untuk melanjutkan proses dengan nilai kosong, atau Cancel untuk berhenti. "
                    "Anda juga bisa membuka file SBT_ANAPAK.xlsx secara manual, memastikan nilainya "
                    "terkalkulasi, lalu menjalankan proyeksi ini lagi.",
                    QMessageBox.Ok | QMessageBox.Cancel
                ) 
                
                response = QMessageBox.question(
                    self,
                    "Buka SBT_ANAPAK secara manual?",
                    "Apakah Anda ingin membuka SBT_ANAPAK.xlsx secara manual untuk memeriksa kalkulasi?",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                )
                
                if response == QMessageBox.Cancel:
                    wb_anapak.close()
                    self.statusBar().clearMessage()
                    return
                elif response == QMessageBox.Yes:
                    # Buka file Excel secara manual
                    if sys.platform == 'win32':
                        os.startfile(sbt_anapak_path)
                    else:
                        if sys.platform == 'darwin':  # macOS
                            subprocess.call(['open', sbt_anapak_path])
                        else:  # Linux
                            subprocess.call(['xdg-open', sbt_anapak_path])
                    
                    # Tanyakan apakah pengguna ingin melanjutkan setelah memeriksa file
                    continue_response = QMessageBox.question(
                        self,
                        "Lanjutkan Proses?",
                        "Setelah Anda memeriksa dan menyimpan file SBT_ANAPAK.xlsx, "
                        "apakah Anda ingin melanjutkan proses?",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    
                    if continue_response == QMessageBox.No:
                        wb_anapak.close()
                        self.statusBar().clearMessage()
                        return
                    
                    # Buka kembali file yang mungkin telah diubah manual
                    wb_anapak.close()
                    wb_anapak = load_workbook(sbt_anapak_path, data_only=True)
                    sheet_anapak = wb_anapak["ANAPAK"]
                    
                    # Ambil nilai yang diperbarui
                    value_i66 = sheet_anapak['I66'].value
                    value_k67 = sheet_anapak['K67'].value
                    
                    print(f"DEBUG setelah pembaruan manual: Nilai dari SBT_ANAPAK.ANAPAK.I66: {value_i66}")
                    print(f"DEBUG setelah pembaruan manual: Nilai dari SBT_ANAPAK.ANAPAK.K67: {value_k67}")
            
            wb_anapak.close()
            
            # Gunakan nilai default jika masih None
            if value_i66 is None:
                value_i66 = 0
                print("DEBUG: Menggunakan nilai default 0 untuk I66 karena nilainya None")
            
            if value_k67 is None:
                value_k67 = 0
                print("DEBUG: Menggunakan nilai default 0 untuk K67 karena nilainya None")
            
            # PERBAIKAN: Transfer data dari SBT_ANAPAK ke SBT_PUMP menggunakan openpyxl saja
            # (tidak menggunakan VBScript yang menyebabkan error)
            try:
                self.statusBar().showMessage("Mentransfer data dari SBT_ANAPAK ke SBT_PUMP menggunakan openpyxl...")
                QApplication.processEvents()
                
                # Cek sheet names di SBT_PUMP sebelum membuka
                try:
                    # Load workbook untuk melihat sheet names
                    temp_wb = load_workbook(sbt_pump_path, read_only=True)
                    sheet_names = temp_wb.sheetnames
                    print(f"DEBUG: Sheet names di SBT_PUMP.xlsm: {sheet_names}")
                    temp_wb.close()
                except Exception as e:
                    print(f"DEBUG: Error saat memeriksa sheet names: {str(e)}")
                
                # Buka SBT_PUMP.xlsm
                wb_pump = load_workbook(sbt_pump_path, keep_vba=True)
                
                # Cari sheet 'DATA INPUT' dengan case insensitive match jika perlu
                target_sheet_name = "DATA INPUT"
                found_sheet = False
                
                for sheet_name in wb_pump.sheetnames:
                    if sheet_name.upper() == target_sheet_name.upper():
                        target_sheet_name = sheet_name  # Gunakan nama sheet yang benar
                        found_sheet = True
                        break
                
                if not found_sheet:
                    QMessageBox.critical(
                        self,
                        "Error",
                        f"Sheet 'DATA INPUT' tidak ditemukan di file SBT_PUMP.xlsm. "
                        f"Sheet yang tersedia: {', '.join(wb_pump.sheetnames)}"
                    )
                    wb_pump.close()
                    return
                
                # Gunakan nama sheet yang ditemukan
                sheet_pump = wb_pump[target_sheet_name]
                
                # Transfer data ke SBT_PUMP
                sheet_pump['B13'] = value_i66
                sheet_pump['B14'] = value_k67
                
                # Log untuk debugging
                print(f"DEBUG: Menulis nilai {value_i66} ke SBT_PUMP.{target_sheet_name}.B13")
                print(f"DEBUG: Menulis nilai {value_k67} ke SBT_PUMP.{target_sheet_name}.B14")
                
                # Simpan SBT_PUMP
                wb_pump.save(sbt_pump_path)
                wb_pump.close()
                
                # Tunggu sedikit untuk memastikan file tersimpan
                time.sleep(1)
                
                # Verifikasi nilai telah tersimpan
                verify_wb = load_workbook(sbt_pump_path, data_only=True)
                verify_sheet = verify_wb[target_sheet_name]
                
                verify_b13 = verify_sheet['B13'].value
                verify_b14 = verify_sheet['B14'].value
                
                print(f"DEBUG: Verifikasi - Nilai di SBT_PUMP.{target_sheet_name}.B13: {verify_b13}")
                print(f"DEBUG: Verifikasi - Nilai di SBT_PUMP.{target_sheet_name}.B14: {verify_b14}")
                
                verify_wb.close()
                
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Gagal mentransfer data ke SBT_PUMP: {str(e)}"
                )
                print(f"DEBUG: Error saat transfer data ke SBT_PUMP: {str(e)}")
                import traceback
                traceback.print_exc()
                return
            
            self.statusBar().showMessage("Menjalankan projection - Menjalankan macro GENERATE_REPORT di SBT_PUMP...")
            QApplication.processEvents()
            
            # PROSES 3: Menjalankan macro di SBT_PUMP
            self.statusBar().showMessage("Menjalankan projection - Menjalankan macro GENERATE_REPORT di SBT_PUMP...")
            QApplication.processEvents()

            try:
                # Import modul xlwings
                import xlwings as xw
                
                self.statusBar().showMessage("Menjalankan projection - Mengimpor UDF dan menjalankan macro...")
                QApplication.processEvents()
                
                # Nonaktifkan tampilan Excel (berjalan di background)
                app = xw.App(visible=False)
                app.display_alerts = False
                
                # Pastikan macro diaktifkan
                app.api.AutomationSecurity = 1  # msoAutomationSecurityLow
                
                # Buka workbook SBT_PUMP
                wb = xw.Book(sbt_pump_path)
                
                # Metode alternatif untuk mengimpor UDF - Menggunakan RunPython
                # Buat modul sementara untuk mengimpor UDF
                import tempfile
                import importlib.util
                
                # Buat file Python sementara untuk mengimpor ALL_UDF
                temp_py = tempfile.NamedTemporaryFile(delete=False, suffix='.py')
                temp_py_path = temp_py.name
                
                # Isi file dengan kode untuk mengimpor semua fungsi dari ALL_UDF.py
                with open(temp_py_path, 'w') as f:
                    f.write(f"""
            import sys
            import os

            # Tambahkan direktori ALL_UDF ke sys.path
            sys.path.append(os.path.dirname(r'{all_udf_path}'))

            # Import ALL_UDF
            from {os.path.splitext(os.path.basename(all_udf_path))[0]} import *

            def main():
                print("UDF berhasil diimpor")

            if __name__ == '__main__':
                main()
            """)
                
                # Jalankan file Python sementara menggunakan RunPython
                try:
                    print(f"DEBUG: Mencoba mengimpor UDF dari {all_udf_path}")
                    xw.Book(sbt_pump_path).api.Application.Run("ImportPythonUDFs")
                    print("DEBUG: Berhasil mengimpor UDF")
                except Exception as udf_err:
                    print(f"DEBUG: Error mengimpor UDF dengan RunPython: {str(udf_err)}")
                    pass  # Lanjut ke langkah berikutnya jika gagal
                
                # Tunggu sebentar
                time.sleep(2)
                
                # Jalankan langsung macro GENERATE_REPORT
                try:
                    print("DEBUG: Menjalankan macro GENERATE_REPORT")
                    wb.api.Application.Run("'SBT_PUMP.xlsm'!GENERATE_REPORT")
                except Exception as macro_err:
                    print(f"DEBUG: Error menjalankan macro dengan metode pertama: {str(macro_err)}")
                    # Coba metode alternatif jika metode pertama gagal
                    try:
                        wb.api.Application.Run("GENERATE_REPORT")
                    except Exception as alt_macro_err:
                        print(f"DEBUG: Error menjalankan macro dengan metode alternatif: {str(alt_macro_err)}")
                        raise  # Re-raise exception untuk ditangani oleh blok except di luar
                
                # Tunggu macro selesai dijalankan
                time.sleep(5)
                
                # Simpan workbook
                wb.save()
                
                # Tutup workbook dan application
                wb.close()
                app.quit()
                
                # Hapus file Python sementara
                try:
                    os.unlink(temp_py_path)
                except:
                    pass
                
                self.statusBar().showMessage("Macro GENERATE_REPORT berhasil dijalankan")
                print("DEBUG: Macro GENERATE_REPORT berhasil dijalankan")
                
            except ImportError:
                QMessageBox.critical(
                    self,
                    "Error",
                    "Module xlwings tidak ditemukan. Silakan install dengan menjalankan 'pip install xlwings'."
                )
                return
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Gagal menjalankan macro GENERATE_REPORT: {str(e)}"
                )
                print(f"DEBUG: Error saat menjalankan macro: {str(e)}")
                import traceback
                traceback.print_exc()
                
                # Alternatif menggunakan COM langsung dengan pywin32
                try:
                    self.statusBar().showMessage("Mencoba alternatif menggunakan pywin32...")
                    QApplication.processEvents()
                    
                    # Gunakan pywin32 untuk menjalankan macro secara langsung
                    import win32com.client
                    
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    
                    # Pastikan makro diizinkan
                    excel.AutomationSecurity = 1  # msoAutomationSecurityLow
                    
                    # Buka workbook
                    wb = excel.Workbooks.Open(sbt_pump_path)
                    
                    # Jalankan VBA untuk mengaktifkan makro jika diperlukan
                    excel.Run("Application.AutomationSecurity=1")
                    
                    # Coba jalankan macro
                    print("DEBUG: Menjalankan macro dengan pywin32")
                    try:
                        excel.Run("'SBT_PUMP.xlsm'!GENERATE_REPORT")
                    except:
                        # Jika gagal, coba tanpa nama file
                        try:
                            excel.Run("GENERATE_REPORT")
                        except Exception as win32_err:
                            print(f"DEBUG: Error menjalankan macro dengan pywin32: {str(win32_err)}")
                    
                    # Simpan dan tutup
                    wb.Save()
                    wb.Close()
                    excel.Quit()
                    
                    self.statusBar().showMessage("Macro GENERATE_REPORT berhasil dijalankan dengan metode alternatif")
                    print("DEBUG: Macro GENERATE_REPORT berhasil dijalankan dengan metode alternatif")
                    
                except Exception as win_err:
                    print(f"DEBUG: Error pada metode pywin32: {str(win_err)}")
                    
                    # Terakhir, coba dengan VBScript
                    try:
                        self.statusBar().showMessage("Mencoba alternatif dengan VBScript...")
                        QApplication.processEvents()
                        
                        import tempfile
                        vbs_file = tempfile.NamedTemporaryFile(delete=False, suffix='.vbs')
                        vbs_path = vbs_file.name
                        
                        vbs_script = f'''
                        Set objExcel = CreateObject("Excel.Application")
                        objExcel.DisplayAlerts = False
                        objExcel.Visible = False
                        objExcel.AutomationSecurity = 1  ' msoAutomationSecurityLow
                        
                        ' Buka file SBT_PUMP.xlsm
                        Set objWorkbook = objExcel.Workbooks.Open("{sbt_pump_path}")
                        
                        On Error Resume Next
                        ' Jalankan macro GENERATE_REPORT
                        objExcel.Run "GENERATE_REPORT"
                        
                        If Err.Number <> 0 Then
                            Err.Clear
                            ' Coba dengan nama workbook yang spesifik
                            objExcel.Run "'SBT_PUMP.xlsm'!GENERATE_REPORT"
                        End If
                        
                        ' Tunggu sebentar untuk memastikan macro selesai
                        WScript.Sleep 5000
                        
                        ' Simpan dan tutup
                        objWorkbook.Save
                        objWorkbook.Close
                        objExcel.Quit
                        
                        Set objWorkbook = Nothing
                        Set objExcel = Nothing
                        '''
                        
                        vbs_file.write(vbs_script.encode('utf-8'))
                        vbs_file.close()
                        
                        # Jalankan script VBS
                        if sys.platform == 'win32':
                            subprocess.call(['cscript.exe', '//nologo', vbs_path])
                            print("DEBUG: Menjalankan macro dengan VBScript sebagai alternatif terakhir")
                        else:
                            QMessageBox.warning(
                                self, 
                                "Warning", 
                                "Menjalankan VBScript di platform non-Windows tidak didukung."
                            )
                        
                        # Hapus file VBS sementara
                        try:
                            os.unlink(vbs_path)
                        except:
                            pass
                        
                        # Tunggu sebentar
                        time.sleep(2)
                        
                    except Exception as alt_e:
                        QMessageBox.critical(
                            self,
                            "Error",
                            f"Gagal menjalankan macro dengan semua metode alternatif: {str(alt_e)}"
                        )
                        print(f"DEBUG: Error pada semua metode: {str(alt_e)}")
                        import traceback
                        traceback.print_exc()
                        return
            
            self.statusBar().showMessage("Menjalankan projection - Transfer data dari SBT_ANAPAK dan SBT_PUMP ke SET_BDU...")
            QApplication.processEvents()
            
            # PROSES 4: Transfer data dari SBT_ANAPAK dan SBT_PUMP kembali ke SET_BDU
            # Buka SBT_ANAPAK untuk mendapatkan data output
            try:
                wb_anapak_output = load_workbook(sbt_anapak_path, data_only=True)
                
                # Pastikan sheet DATA_OUTPUT ada
                if "DATA_OUTPUT" not in wb_anapak_output.sheetnames:
                    QMessageBox.critical(
                        self, 
                        "Error", 
                        "Sheet DATA_OUTPUT tidak ditemukan di file SBT_ANAPAK.xlsx"
                    )
                    wb_anapak_output.close()
                    return
                    
                sheet_output = wb_anapak_output["DATA_OUTPUT"]
                
                # Ambil nilai dari sheet DATA_OUTPUT
                output_values = {}
                
                # Gunakan try-except untuk setiap sel untuk menangani kemungkinan kesalahan
                try:
                    output_values['B3'] = sheet_output['C20'].value  # C20 -> B3
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C20: {sheet_output['C20'].value}")
                except Exception as e:
                    print(f"DEBUG: Error saat mengambil C20: {str(e)}")
                    output_values['B3'] = None
                    
                try:
                    output_values['B4'] = sheet_output['C21'].value  # C21 -> B4
                    output_values['B5'] = sheet_output['C22'].value  # C22 -> B5
                    output_values['B6'] = sheet_output['C23'].value  # C23 -> B6
                    output_values['B7'] = sheet_output['C24'].value  # C24 -> B7
                    output_values['B8'] = sheet_output['C25'].value  # C25 -> B8
                    output_values['B10'] = sheet_output['C30'].value # C30 -> B10
                    output_values['B11'] = sheet_output['C31'].value # C31 -> B11
                    
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C21: {sheet_output['C21'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C22: {sheet_output['C22'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C23: {sheet_output['C23'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C24: {sheet_output['C24'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C25: {sheet_output['C25'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C30: {sheet_output['C30'].value}")
                    print(f"DEBUG: Nilai dari SBT_ANAPAK.DATA_OUTPUT.C31: {sheet_output['C31'].value}")
                except Exception as e:
                    print(f"DEBUG: Error saat mengambil beberapa nilai dari DATA_OUTPUT: {str(e)}")
                
                wb_anapak_output.close()
                
                # Buka lagi SBT_PUMP untuk mendapatkan data dari DATA INPUT dan DATA ENGINE
                wb_pump_input = load_workbook(sbt_pump_path, data_only=True)
                
                # Cari sheet 'DATA INPUT' dengan case insensitive match jika perlu
                data_input_sheet_name = "DATA INPUT"
                found_data_input = False
                
                for sheet_name in wb_pump_input.sheetnames:
                    if sheet_name.upper() == data_input_sheet_name.upper():
                        data_input_sheet_name = sheet_name  # Gunakan nama sheet yang benar
                        found_data_input = True
                        break
                
                if found_data_input:
                    sheet_pump_input = wb_pump_input[data_input_sheet_name]
                    # Tambahkan data dari SBT_PUMP.DATA INPUT
                    try:
                        output_values['B13'] = sheet_pump_input['B13'].value  # B13 -> B13
                        output_values['B14'] = sheet_pump_input['B14'].value  # B14 -> B14
                        
                        print(f"DEBUG: Nilai dari SBT_PUMP.{data_input_sheet_name}.B13: {sheet_pump_input['B13'].value}")
                        print(f"DEBUG: Nilai dari SBT_PUMP.{data_input_sheet_name}.B14: {sheet_pump_input['B14'].value}")
                    except Exception as e:
                        print(f"DEBUG: Error saat mengambil nilai dari DATA INPUT: {str(e)}")
                else:
                    QMessageBox.warning(
                        self, 
                        "Warning", 
                        f"Sheet 'DATA INPUT' tidak ditemukan di file SBT_PUMP.xlsm. "
                        f"Beberapa data mungkin tidak lengkap."
                    )
                
                # TAMBAHAN: Cari dan ambil data dari sheet 'DATA ENGINE'
                data_engine_sheet_name = "DATA ENGINE"
                found_data_engine = False
                
                for sheet_name in wb_pump_input.sheetnames:
                    if sheet_name.upper() == data_engine_sheet_name.upper():
                        data_engine_sheet_name = sheet_name  # Gunakan nama sheet yang benar
                        found_data_engine = True
                        break
                
                if found_data_engine:
                    sheet_pump_engine = wb_pump_input[data_engine_sheet_name]
                    # Tambahkan data dari SBT_PUMP.DATA ENGINE
                    try:
                        output_values['B15'] = sheet_pump_engine['B19'].value  # B19 -> B15
                        
                        print(f"DEBUG: Nilai dari SBT_PUMP.{data_engine_sheet_name}.B19: {sheet_pump_engine['B19'].value}")
                    except Exception as e:
                        print(f"DEBUG: Error saat mengambil nilai dari DATA ENGINE: {str(e)}")
                else:
                    QMessageBox.warning(
                        self, 
                        "Warning", 
                        f"Sheet 'DATA ENGINE' tidak ditemukan di file SBT_PUMP.xlsm. "
                        f"Beberapa data mungkin tidak lengkap."
                    )
                
                wb_pump_input.close()
                
                # Buka SET_BDU untuk update nilai akhir - PENTING: Gunakan file SET_BDU customer, bukan template
                wb_bdu_final = load_workbook(set_bdu_path)
                
                # Pastikan sheet DATA_TEMP ada
                if "DATA_TEMP" not in wb_bdu_final.sheetnames:
                    QMessageBox.critical(
                        self, 
                        "Error", 
                        "Sheet DATA_TEMP tidak ditemukan di file SET_BDU.xlsx"
                    )
                    wb_bdu_final.close()
                    return
                    
                sheet_temp = wb_bdu_final["DATA_TEMP"]
                
                # Transfer semua nilai ke SET_BDU.DATA_TEMP
                for cell_addr, value in output_values.items():
                    if value is not None:  # Hanya transfer nilai yang tidak None
                        sheet_temp[cell_addr] = value
                        print(f"DEBUG: Menyimpan nilai {value} ke SET_BDU.DATA_TEMP.{cell_addr}")
                    else:
                        print(f"DEBUG: Tidak menyimpan nilai None ke SET_BDU.DATA_TEMP.{cell_addr}")
                
                # Simpan SET_BDU
                wb_bdu_final.save(set_bdu_path)
                wb_bdu_final.close()
                
                self.statusBar().showMessage("Projection selesai - Semua data berhasil diproses")
                QMessageBox.information(
                    self, 
                    "Projection Selesai", 
                    "Proses projection telah berhasil dilakukan. Data telah dipindahkan dan diproses."
                )
                
                # Perbarui tampilan
                self.load_excel_data()
                
            except Exception as e:
                self.statusBar().clearMessage()
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"Terjadi kesalahan saat memproses data akhir: {str(e)}"
                )
                print(f"DEBUG: Error saat proses 4: {str(e)}")
                import traceback
                traceback.print_exc()
            
        except Exception as e:
            self.statusBar().clearMessage()
            QMessageBox.critical(
                self, 
                "Error", 
                f"Terjadi kesalahan saat menjalankan projection: {str(e)}"
            )
            print(f"Error running projection: {str(e)}")
            import traceback
            traceback.print_exc()
                    
    def run_generate_proposal(self):
        """Fungsi untuk menjalankan script generate_proposal.py dengan file customer"""
        try:
            # Use the customer-specific Excel file
            if not hasattr(self, 'customer_name') or not self.customer_name:
                # If no customer is set, use the default path
                excel_path = self.excel_path
                output_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            else:
                # Use customer-specific paths
                from modules.fix_customer_system import clean_folder_name
                
                # Get customer folder path
                customer_folder = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "data", "customers", clean_folder_name(self.customer_name)
                )
                
                # Set excel_path to the customer's SET_BDU.xlsx
                excel_path = os.path.join(customer_folder, "SET_BDU.xlsx")
                
                # Set output directory to customer folder
                output_dir = customer_folder
            
            # Path to the Word template
            template_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                                    "data", "Trial WWTP ANP Quotation Template.docx")
            
            # Path for the output file - save in customer folder
            output_filename = f"WWTP_Quotation_{self.customer_name or 'Result'}.docx"
            output_path = os.path.join(output_dir, output_filename)
            
            # Display status message
            self.statusBar().showMessage("Generating proposal...")
            
            # Import the generate_proposal module
            import importlib.util
            spec = importlib.util.spec_from_file_location(
                "generate_proposal", 
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                            "modules", "generate_proposal.py")
            )
            generate_proposal = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(generate_proposal)
            
            # Call the function with proper paths
            success = generate_proposal.generate_proposal(excel_path, template_path, output_path)
            
            if success:
                QMessageBox.information(
                    self,
                    "Success",
                    f"Proposal successfully generated and saved to: {output_path}"
                )
                # Refresh display to show the new file
                self.load_excel_data()
            else:
                QMessageBox.critical(
                    self,
                    "Error",
                    "Failed to generate proposal. Check the console for more details."
                )
            
            # Clear status message
            self.statusBar().clearMessage()
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error", 
                f"An error occurred while generating proposal: {str(e)}"
            )
            self.statusBar().clearMessage()
            print(f"Error in run_generate_proposal: {str(e)}")    
    
    def on_generate_proposal_finished(self, success, output):
        """Handler ketika proses generate proposal selesai"""
        if success:
            QMessageBox.information(
                self,
                "Sukses",
                "Proposal berhasil digenerate. Silakan refresh halaman untuk melihat hasilnya."
            )
            print("Output generate-proposal.py:", output)
        else:
            QMessageBox.critical(
                self,
                "Error",
                f"Gagal menghasilkan proposal: {output}"
            )
            print("Error generate-proposal.py:", output)
        
        # Hapus pesan status
        self.statusBar().clearMessage()

    def process_proposal_document(self, file_path):
        try:
            # Pastikan path adalah absolut
            abs_file_path = self.get_absolute_path(file_path)
            
            # Periksa apakah file ada
            if not os.path.exists(abs_file_path):
                print(f"Dokumen Word tidak ditemukan: {abs_file_path}")
                return False
                    
            # Dapatkan ekstensi file
            _, ext = os.path.splitext(abs_file_path)
            ext = ext.lower()
            
            # Periksa apakah itu dokumen Word
            if ext not in ['.docx', '.doc']:
                print(f"Bukan dokumen Word: {abs_file_path}")
                return False

            # Simpan path dokumen untuk referensi
            self.proposal_document_path = abs_file_path
                    
            # Buat widget untuk menampilkan dokumen
            document_widget = QWidget()
            document_layout = QVBoxLayout(document_widget)
            document_layout.setContentsMargins(0, 0, 0, 0)
            
            # Buat widget info file - layout lebih kompak 
            file_header = QWidget()
            file_header.setStyleSheet("background-color: white; border: 1px solid #ddd; border-radius: 4px;")
            file_header_layout = QHBoxLayout(file_header)
            file_header_layout.setContentsMargins(10, 5, 10, 5)
            
            # Ikon file
            file_icon = QLabel("📄")
            file_icon.setFont(QFont("Segoe UI", 14))
            file_icon.setStyleSheet("color: #3498DB;")
            
            # Nama file
            file_name = os.path.basename(file_path)
            file_label = QLabel(f"{file_name}")
            file_label.setFont(QFont("Segoe UI", 11))
            file_label.setStyleSheet("color: #333;")
            
            # Tombol-tombol menu
            run_projection_btn = QPushButton("Re-run Projection")
            run_projection_btn.setFont(QFont("Segoe UI", 10))
            run_projection_btn.setCursor(Qt.PointingHandCursor)
            run_projection_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: #3498db;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 10px;
                    margin-right: 5px;
                }}
                QPushButton:hover {{
                    background-color: #2980b9;
                }}
            """)
            run_projection_btn.clicked.connect(self.run_projection)
            
            generate_btn = QPushButton("Regenerate Proposal")
            generate_btn.setFont(QFont("Segoe UI", 10))
            generate_btn.setCursor(Qt.PointingHandCursor)
            generate_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: #27ae60;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 10px;
                    margin-right: 5px;
                }}
                QPushButton:hover {{
                    background-color: #2ecc71;
                }}
            """)
            generate_btn.clicked.connect(self.run_generate_proposal)
            
            open_btn = QPushButton("Open in Word")
            open_btn.setFont(QFont("Segoe UI", 10))
            open_btn.setCursor(Qt.PointingHandCursor)
            open_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {SECONDARY_COLOR};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 10px;
                }}
                QPushButton:hover {{
                    background-color: #2980B9;
                }}
            """)
            open_btn.clicked.connect(lambda: self.open_document(file_path))
            
            # Tambahkan widget ke layout header
            file_header_layout.addWidget(file_icon)
            file_header_layout.addWidget(file_label)
            file_header_layout.addStretch()
            file_header_layout.addWidget(run_projection_btn)
            file_header_layout.addWidget(generate_btn)
            file_header_layout.addWidget(open_btn)
            
            document_layout.addWidget(file_header)
            
            # Alternatif untuk pratinjau: Pesan sederhana dengan ikon dokumen
            preview_container = QWidget()
            preview_layout = QVBoxLayout(preview_container)
            preview_layout.setAlignment(Qt.AlignCenter)
            preview_container.setStyleSheet("""
                background-color: #f9f9f9;
                border: 1px solid #ddd;
                border-radius: 4px;
            """)
            
            # Info pesan
            info_label = QLabel("Preview tidak tersedia.")
            info_label.setFont(QFont("Segoe UI", 12))
            info_label.setAlignment(Qt.AlignCenter)
            info_label.setStyleSheet("color: #666; margin: 20px;")
            
            # Ikon dokumen besar
            doc_icon = QLabel("📄")
            doc_icon.setFont(QFont("Segoe UI", 48))
            doc_icon.setAlignment(Qt.AlignCenter)
            doc_icon.setStyleSheet("color: #3498DB; margin: 20px;")
            
            # Pesan untuk membuka di Word
            tip_label = QLabel("Gunakan tombol 'Open in Word' untuk melihat dokumen.")
            tip_label.setFont(QFont("Segoe UI", 10))
            tip_label.setAlignment(Qt.AlignCenter)
            tip_label.setStyleSheet("color: #666; margin: 10px;")
            
            # Tambahkan ke container
            preview_layout.addWidget(info_label)
            preview_layout.addWidget(doc_icon)
            preview_layout.addWidget(tip_label)
            
            document_layout.addWidget(preview_container, 1)  # Stretch factor 1
            
            # Coba gunakan metode thumbnail generator jika tersedia modul yang diperlukan
            try:
                # Cek apakah library yang diperlukan tersedia
                have_required_modules = False
                try:
                    import pythoncom
                    import docx2pdf
                    import fitz
                    have_required_modules = True
                except ImportError:
                    have_required_modules = False
                
                if have_required_modules:
                    # Membuat area scroll untuk thumbnail
                    scroll_area = QScrollArea()
                    scroll_area.setWidgetResizable(True)
                    scroll_area.setStyleSheet("""
                        QScrollArea {
                            border: 1px solid #ddd;
                            background-color: #f9f9f9;
                            border-radius: 4px;
                        }
                    """)
                    
                    # Container untuk thumbnail
                    thumbnail_container = QWidget()
                    thumbnail_layout = QVBoxLayout(thumbnail_container)
                    thumbnail_layout.setAlignment(Qt.AlignHCenter)
                    thumbnail_layout.setSpacing(20)
                    
                    # Label loading
                    loading_label = QLabel("Loading preview document...")
                    loading_label.setFont(QFont("Segoe UI", 12))
                    loading_label.setAlignment(Qt.AlignCenter)
                    loading_label.setStyleSheet("color: #666; margin: 20px;")
                    thumbnail_layout.addWidget(loading_label)
                    
                    # Simpan referensi layout thumbnail untuk digunakan nanti
                    self.thumbnail_layout = thumbnail_layout
                    self.loading_label = loading_label
                    
                    # Tambahkan container ke area gulir
                    scroll_area.setWidget(thumbnail_container)
                    document_layout.removeWidget(preview_container)
                    preview_container.deleteLater()
                    document_layout.addWidget(scroll_area, 1)  # Stretch factor 1
                    
                    # Inisialisasi COM library untuk main thread
                    pythoncom.CoInitialize()
                    
                    # Mulai thread generator thumbnail
                    self.thumbnail_thread = ThumbnailGeneratorThread(file_path, max_pages=100)
                    self.thumbnail_thread.thumbnail_ready.connect(self.add_thumbnail)
                    self.thumbnail_thread.finished.connect(self.generation_finished)
                    self.thumbnail_thread.error.connect(self.generation_error)
                    self.thumbnail_thread.start()
                else:
                    # Jika modul tidak tersedia, gunakan pratinjau sederhana
                    print("Library yang diperlukan tidak tersedia. Menggunakan pratinjau sederhana.")
            except Exception as e:
                print(f"Error saat membuat pratinjau: {e}")
            
            # Simpan widget untuk digunakan nanti
            self.proposal_document_widget = document_widget
            
            return True
                
        except Exception as e:
            print(f"Error memproses dokumen Word: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
            
    def add_thumbnail(self, page_num, pixmap):
        """Menambahkan thumbnail ke UI"""
        try:
            # Hapus label loading saat thumbnail pertama ditambahkan
            if page_num == 0 and self.loading_label:
                self.loading_label.setVisible(False)
            
            # Buat frame untuk thumbnail
            thumbnail_frame = QFrame()
            thumbnail_frame.setFrameShape(QFrame.StyledPanel)
            thumbnail_frame.setStyleSheet("""
                QFrame {
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    background-color: white;
                    padding: 10px;
                }
            """)
            frame_layout = QVBoxLayout(thumbnail_frame)
            
            # Tambahkan label halaman
            page_label = QLabel(f"Page {page_num + 1}")
            page_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
            page_label.setAlignment(Qt.AlignCenter)
            frame_layout.addWidget(page_label)
            
            # Skala thumbnail untuk fit dalam lebar 600px
            scaled_pixmap = pixmap.scaledToWidth(600, Qt.SmoothTransformation)
            
            # Buat label untuk thumbnail
            thumbnail_label = QLabel()
            thumbnail_label.setPixmap(scaled_pixmap)
            thumbnail_label.setAlignment(Qt.AlignCenter)
            thumbnail_label.setStyleSheet("background-color: white;")
            frame_layout.addWidget(thumbnail_label)
            
            # Tambahkan thumbnail ke layout
            self.thumbnail_layout.addWidget(thumbnail_frame)
        except Exception as e:
            print(f"Error menambahkan thumbnail: {str(e)}")

    def generation_finished(self):
        """Dipanggil saat generasi thumbnail selesai"""
        try:
            # Hapus label loading jika masih ada
            if self.loading_label:
                self.loading_label.setVisible(False)
                
        except Exception as e:
            print(f"Error dalam generation_finished: {str(e)}")

    def generation_error(self, error_message):
        """Dipanggil jika ada error saat generasi thumbnail"""
        try:
            if self.loading_label:
                self.loading_label.setText(f"Error saat memuat pratinjau: {error_message}")
                self.loading_label.setStyleSheet("color: #E74C3C; margin: 20px;")
        except Exception as e:
            print(f"Error dalam generation_error: {str(e)}")

    def open_document(self, file_path):
        """Membuka dokumen di aplikasi aslinya"""
        try:
            # Konversi ke path absolut
            abs_file_path = self.get_absolute_path(file_path)
                
            # Untuk Windows
            if sys.platform == 'win32':
                os.startfile(abs_file_path)
            # Untuk macOS
            elif sys.platform == 'darwin':
                subprocess.call(('open', abs_file_path))
            # Untuk Linux
            else:
                subprocess.call(('xdg-open', abs_file_path))
                    
            print(f"Membuka dokumen: {abs_file_path}")
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error Membuka File", 
                f"Tidak dapat membuka file: {str(e)}"
            )
            print(f"Error membuka dokumen: {str(e)}")
                    
    def get_validation_values(self, excel_path, sheet_name, cell_address):
        """Mengambil nilai dari data validation di sebuah sel Excel"""
        from openpyxl import load_workbook
        
        try:
            # Pastikan untuk memuat dengan data_only=False agar kita bisa mengakses validasi
            workbook = load_workbook(excel_path, data_only=False)
            
            if sheet_name not in workbook.sheetnames:
                return []
                
            sheet = workbook[sheet_name]
            
            # Periksa apakah cell address valid
            try:
                cell = sheet[cell_address]
            except:
                return []
            
            # Cek data validation secara eksplisit
            dv = sheet.data_validations.dataValidation
            for validation in dv:
                for coord in validation.sqref.ranges:
                    if cell.coordinate in str(coord):
                        # Ditemukan validasi untuk sel ini
                        if validation.type == "list":
                            formula = validation.formula1
                            
                            # Jika formula menggunakan referensi
                            if formula.startswith('='):
                                # Implementasi sama seperti sebelumnya...
                                pass
                            else:
                                # Untuk list langsung seperti "A,B,C"
                                if formula.startswith('"') and formula.endswith('"'):
                                    formula = formula[1:-1]
                                return [val.strip() for val in formula.split(',')]
            
            # Fallback: Coba cara lain untuk mendapatkan validation list
            try:
                # Untuk beberapa versi openpyxl, langsung coba akses data_validation
                if hasattr(cell, 'data_validation') and cell.data_validation and hasattr(cell.data_validation, 'type'):
                    if cell.data_validation.type == 'list':
                        formula = cell.data_validation.formula1
                        if formula.startswith('"') and formula.endswith('"'):
                            formula = formula[1:-1]
                        return [val.strip() for val in formula.split(',')]
            except:
                pass
                
            return []
        except Exception as e:
            print(f"Error saat membaca data validation: {str(e)}")
            return []
        
    def process_excel_images(self, sheet_name, layout):
        """Extract and display images from Excel sheet"""
        try:
            from openpyxl import load_workbook
            from openpyxl.drawing.image import Image
            from io import BytesIO
            from PIL import Image as PILImage
            
            # Load workbook
            wb = load_workbook(self.excel_path)
            if sheet_name not in wb.sheetnames:
                return False
                
            sheet = wb[sheet_name]
            
            # Create a frame for images
            images_frame = QWidget()
            images_layout = QVBoxLayout(images_frame)
            images_layout.setContentsMargins(10, 10, 10, 10)
            
            # Add a title for images section
            images_title = QLabel("Diagrams and Images")
            images_title.setFont(QFont("Segoe UI", 14, QFont.Bold))
            images_title.setStyleSheet(f"color: {PRIMARY_COLOR};")
            images_layout.addWidget(images_title)
            
            # Track if we found any images
            found_images = False
            
            # Process all images in the sheet
            for image in sheet._images:
                found_images = True
                
                # Create a label to display the image
                img_label = QLabel()
                img_label.setAlignment(Qt.AlignCenter)
                img_label.setStyleSheet("background-color: white; border: 1px solid #ddd; padding: 10px;")
                
                # Extract image data
                img_data = image._data()
                
                # Convert to QPixmap and set to label
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                
                # Scale image if too large
                if pixmap.width() > 800:
                    pixmap = pixmap.scaledToWidth(800, Qt.SmoothTransformation)
                    
                img_label.setPixmap(pixmap)
                images_layout.addWidget(img_label)
                
                # Add some spacing between images
                images_layout.addSpacing(20)
            
            # Add the images frame to the main layout if we found any
            if found_images:
                layout.addWidget(images_frame)
                return True
            
            return False
        except Exception as e:
            print(f"Error processing images from sheet {sheet_name}: {str(e)}")
            return False
    
    def process_sheet_data(self, df, sheet_name, layout):
        """Process the data from a sheet and create UI elements in a grid layout similar to Excel"""
        # Check if the sheet is a DATA sheet (just display as a table)
        if sheet_name.startswith("DATA_"):
            self.create_data_table(df, layout)
            self.process_excel_images(sheet_name, layout)
            return

        # For DIP sheets or other sheets, process as forms
        # Initialize variables
        current_section = None
        section_layout = None
        section_grid = None
        current_row = 0  # Track the current row in the grid
        current_header_labels = []  # For storing column headers from ch_
        has_column_headers = False
        
        # Track sections by column position
        left_section = None  # Main left section (columns 0-1)
        right_section = None  # Right section (columns 2-3)
        right_section_title = None

        # Field identification
        field_count = 0
        
        # Variables to track industry dropdown fields
        industry_dropdown = None
        sub_industry_dropdown = None
        
        # Untuk melacak pasangan field dropdown industry-subindustry
        industry_field_key = None
        sub_industry_field_key = None
        
        # Create a section for table if we find table formatting
        table_section = None
        table_grid = None
        in_table = False
        table_row = 0

        # Dictionary to track rowspans
        rowspans = {}
        tdm_counter = 0  # Counter for tdm_ cells in current row
        last_header = None  # Track last header for rowspans
        current_tdm_row = 0  # Track current row for tdm counters

        # Check if the dataframe is empty
        if df.empty:
            empty_label = QLabel("No data found in this sheet.")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("color: #666; margin: 20px;")
            layout.addWidget(empty_label)
            return

        # First pass: pre-scan for column headers
        # This will help us collect all ch_ headers before processing fields
        for index, row in df.iterrows():
            # Skip empty rows
            if pd.isna(row).all():
                continue
            
            first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
            if not isinstance(first_col, str):
                try:
                    first_col = str(first_col)
                except:
                    continue
                    
            # If this is a 'Contact' type row with headers
            if first_col.startswith('fh_') and any(isinstance(cell, str) and cell.startswith('ch_') for cell in row if not pd.isna(cell)):
                # Process all cells in this row to find ch_ headers
                header_row_labels = []
                for col_idx in range(df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip() if isinstance(row[col_idx], str) else ""
                        if col_value.startswith('ch_'):
                            header_text = col_value[3:].strip()  # Remove 'ch_' prefix
                            header_row_labels.append(header_text)
                
                if len(header_row_labels) > 0:
                    current_header_labels = header_row_labels
                    has_column_headers = True
                    
                # Also look for any right headers (fh_) in the same row
                for col_idx in range(1, df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip() if isinstance(row[col_idx], str) else ""
                        if col_value.startswith('fh_'):
                            right_header_text = col_value[3:].strip()  # Remove 'fh_' prefix
                            # If this is first header for right section, treat it as section title
                            if right_section is None:
                                right_section = right_header_text

        # Process each row
        for index, row in df.iterrows():
            # Skip empty rows
            if pd.isna(row).all():
                continue
            
            # Get the first column to determine the type
            first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""

            # Convert to string for startswith checks, but only if we're not in a DATA_ sheet
            if not isinstance(first_col, str):
                try:
                    first_col = str(first_col)
                except:
                    continue  # Skip if can't convert to string
            
            # Ambil nilai kolom A (indeks 0)
            col_a = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
            
            # Ubah ke string untuk pemeriksaan dengan startswith
            if not isinstance(col_a, str):
                try:
                    col_a = str(col_a)
                except:
                    continue  # Lewati jika tidak bisa dikonversi ke string
            
            # Pemrosesan header (th_ atau thr_) di kolom A
            if col_a.startswith('th_') or col_a.startswith('thr_'):
                # Jika kita sedang dalam tabel tapi ini header baru, reset penghitung
                if in_table and table_grid is None:
                    # Buat struktur tabel baru
                    table_frame = QWidget()
                    table_frame.setStyleSheet("background-color: white; border: 1px solid #ddd;")
                    
                    table_section = QVBoxLayout(table_frame)
                    table_section.setContentsMargins(0, 0, 0, 0)
                    table_section.setSpacing(0)
                    
                    table_grid = QGridLayout()
                    table_grid.setSpacing(0)
                    table_grid.setContentsMargins(0, 0, 0, 0)
                    
                    section_layout.addWidget(table_frame)
                    table_section.addLayout(table_grid)
                    
                    table_row = 0
                elif not in_table:
                    # Buat tabel baru jika belum ada
                    in_table = True
                    
                    # Buat section untuk tabel jika belum ada
                    if section_layout is None:
                        section_frame = QWidget()
                        section_frame.setStyleSheet("background-color: white;")
                        section_layout = QVBoxLayout(section_frame)
                        section_layout.setContentsMargins(15, 15, 15, 15)
                        section_layout.setSpacing(15)
                        layout.addWidget(section_frame)
                    
                    # Buat frame tabel
                    table_frame = QWidget()
                    table_frame.setStyleSheet("background-color: white; border: 1px solid #ddd;")
                    
                    table_section = QVBoxLayout(table_frame)
                    table_section.setContentsMargins(0, 0, 0, 0)
                    table_section.setSpacing(0)
                    
                    # Buat grid untuk tabel
                    table_grid = QGridLayout()
                    table_grid.setSpacing(0)
                    table_grid.setContentsMargins(0, 0, 0, 0)
                    
                    # Tambahkan ke section
                    section_layout.addWidget(table_frame)
                    table_section.addLayout(table_grid)
                    
                    # Reset penghitung tabel
                    table_row = 0
                
                # Header biasa (th_)
                if col_a.startswith('th_'):
                    header_text = col_a[3:].strip()  # Hapus awalan 'th_'
                    header_label = QLabel(header_text)
                    header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                    header_label.setStyleSheet("""
                        background-color: #f5f5f5; 
                        color: #333;
                        padding: 8px;
                        border: 1px solid #ddd;
                    """)
                    header_label.setMinimumWidth(250)
                    
                    # Tambahkan ke grid
                    table_grid.addWidget(header_label, table_row, 0)
                    
                    # Cek kolom B (indeks 1) untuk data
                    if len(row) > 1 and not pd.isna(row.iloc[1]):
                        col_b = row.iloc[1]
                        
                        # Ubah ke string untuk pemeriksaan
                        if not isinstance(col_b, str):
                            try:
                                col_b = str(col_b)
                            except:
                                col_b = ""
                        
                        # Proses data berdasarkan jenisnya
                        if col_b.startswith('td_'):
                            # Data biasa
                            data_text = col_b[3:].strip()  # Hapus awalan 'td_'
                            data_label = QLabel(data_text)
                            data_label.setFont(QFont("Segoe UI", 11))
                            data_label.setStyleSheet("""
                                padding: 8px;
                                border: 1px solid #ddd;
                            """)
                            data_label.setWordWrap(True)
                            
                            # Tambahkan ke grid, span 2 kolom
                            table_grid.addWidget(data_label, table_row, 1, 1, 2)
                        
                        elif col_b.startswith('tdi_'):
                            # Data dengan input field
                            data_text = col_b[4:].strip()  # Hapus awalan 'tdi_'
                            
                            # Cek placeholder seperti $P1$
                            import re
                            placeholders = re.findall(r'\$P\d+\$', data_text)
                            
                            if placeholders:
                                # Buat container untuk input field + teks
                                container = QWidget()
                                container_layout = QHBoxLayout(container)
                                container_layout.setContentsMargins(8, 8, 8, 8)
                                
                                # Pisah teks berdasarkan placeholder
                                parts = re.split(r'(\$P\d+\$)', data_text)
                                
                                for part in parts:
                                    if re.match(r'\$P\d+\$', part):
                                        # Ini placeholder, buat input field
                                        input_field = QLineEdit()
                                        input_field.setFixedWidth(50)
                                        input_field.setStyleSheet("""
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                        """)
                                        container_layout.addWidget(input_field)
                                        
                                        # Daftarkan field di data_fields
                                        field_key = f"tdi_{part}_{header_text}"
                                        self.data_fields[field_key] = input_field
                                    else:
                                        # Teks biasa
                                        text_label = QLabel(part)
                                        text_label.setFont(QFont("Segoe UI", 11))
                                        container_layout.addWidget(text_label)
                                
                                container_layout.addStretch()
                                
                                # Tambahkan ke grid, span 2 kolom
                                table_grid.addWidget(container, table_row, 1, 1, 2)
                            else:
                                # Tidak ada placeholder, teks biasa
                                data_label = QLabel(data_text)
                                data_label.setFont(QFont("Segoe UI", 11))
                                data_label.setStyleSheet("""
                                    padding: 8px;
                                    border: 1px solid #ddd;
                                """)
                                data_label.setWordWrap(True)
                                
                                # Tambahkan ke grid, span 2 kolom
                                table_grid.addWidget(data_label, table_row, 1, 1, 2)
                        
                        elif col_b.startswith('tdm_'):
                            # Multi-kolom data (2 kolom)
                            data_text_b = col_b[4:].strip()  # Hapus awalan 'tdm_'
                            
                            # Buat label untuk kolom pertama (B)
                            data_label_b = QLabel(data_text_b)
                            data_label_b.setFont(QFont("Segoe UI", 11))
                            data_label_b.setStyleSheet("""
                                padding: 8px;
                                border: 1px solid #ddd;
                            """)
                            data_label_b.setWordWrap(True)
                            
                            # Tambahkan ke grid di kolom 1
                            table_grid.addWidget(data_label_b, table_row, 1)
                            
                            # Cek kolom C (indeks 2) untuk data kedua
                            if len(row) > 2 and not pd.isna(row.iloc[2]):
                                col_c = row.iloc[2]
                                
                                # Ubah ke string untuk pemeriksaan
                                if not isinstance(col_c, str):
                                    try:
                                        col_c = str(col_c)
                                    except:
                                        col_c = ""
                                
                                if col_c.startswith('tdm_'):
                                    # Multi-kolom data kedua
                                    data_text_c = col_c[4:].strip()  # Hapus awalan 'tdm_'
                                    
                                    # Buat label untuk kolom kedua (C)
                                    data_label_c = QLabel(data_text_c)
                                    data_label_c.setFont(QFont("Segoe UI", 11))
                                    data_label_c.setStyleSheet("""
                                        padding: 8px;
                                        border: 1px solid #ddd;
                                    """)
                                    data_label_c.setWordWrap(True)
                                    
                                    # Tambahkan ke grid di kolom 2
                                    table_grid.addWidget(data_label_c, table_row, 2)
                        
                        else:
                            # Data tanpa awalan khusus, tampilkan sebagai teks biasa
                            data_text = str(col_b).strip()
                            data_label = QLabel(data_text)
                            data_label.setFont(QFont("Segoe UI", 11))
                            data_label.setStyleSheet("""
                                padding: 8px;
                                border: 1px solid #ddd;
                            """)
                            data_label.setWordWrap(True)
                            
                            # Tambahkan ke grid, span 2 kolom
                            table_grid.addWidget(data_label, table_row, 1, 1, 2)
                    
                    # Simpan header terakhir
                    last_header = header_text
                    table_row += 1
                
                # Untuk header dengan rowspan (thr_)
                elif col_a.startswith('thr_'):
                    header_text = col_a[4:].strip()  # Hapus awalan 'thr_'
                    header_label = QLabel(header_text)
                    header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                    header_label.setStyleSheet("""
                        background-color: #f5f5f5; 
                        color: #333;
                        padding: 8px;
                        border: 1px solid #ddd;
                    """)
                    header_label.setMinimumWidth(250)
                    header_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
                    
                    # Hitung rowspan dengan melihat baris-baris berikutnya
                    rowspan = 0
                    tdm_count = 0
                    current_row = index
                    
                    # Simpan baris-baris data yang akan diproses nanti
                    data_rows = []
                    
                    # Periksa baris-baris berikutnya sampai ketemu header baru
                    while current_row + 1 < len(df):
                        next_row = current_row + 1
                        next_row_data = df.iloc[next_row]
                        next_col_a = next_row_data.iloc[0] if not pd.isna(next_row_data.iloc[0]) else ""
                        
                        # Ubah ke string jika perlu
                        if not isinstance(next_col_a, str):
                            try:
                                next_col_a = str(next_col_a)
                            except:
                                next_col_a = ""
                        
                        # Jika ketemu header baru, berhenti
                        if next_col_a.startswith('th_') or next_col_a.startswith('thr_'):
                            break
                        
                        # Periksa kolom B untuk data, meskipun kolom A kosong
                        if len(next_row_data) > 1:
                            next_col_b = next_row_data.iloc[1] if not pd.isna(next_row_data.iloc[1]) else ""
                            
                            # Ubah ke string jika perlu
                            if not isinstance(next_col_b, str):
                                try:
                                    next_col_b = str(next_col_b)
                                except:
                                    next_col_b = ""
                            
                            # Jika ada data di kolom B, tambah rowspan dan simpan data
                            if (next_col_b.startswith('td_') or 
                                next_col_b.startswith('tdi_') or
                                next_col_b.startswith('tdm_')):
                                rowspan += 1
                                data_rows.append(next_row_data)
                        
                        current_row = next_row
                    
                    # Perlu ditambahkan 1 untuk baris pertama
                    rowspan = max(1, rowspan + 1)
                    
                    # Tambahkan baris saat ini ke daftar data
                    data_rows.insert(0, row)
                    
                    # Simpan informasi rowspan
                    rowspans[table_row] = {
                        'header': header_text,
                        'rowspan': rowspan,
                        'widget': header_label
                    }
                    
                    # Tambahkan ke grid dengan rowspan
                    table_grid.addWidget(header_label, table_row, 0, rowspan, 1)
                    
                    # Proses semua data dalam rowspan
                    current_data_row = table_row
                    
                    # Proses semua baris data yang telah dikumpulkan
                    for data_row in data_rows:
                        # Cek kolom B untuk data
                        if len(data_row) > 1 and not pd.isna(data_row.iloc[1]):
                            col_b = data_row.iloc[1]
                            
                            # Ubah ke string jika perlu
                            if not isinstance(col_b, str):
                                try:
                                    col_b = str(col_b)
                                except:
                                    col_b = ""
                            
                            # Proses berdasarkan jenis data
                            if col_b.startswith('td_'):
                                data_text = col_b[3:].strip()
                                data_label = QLabel(data_text)
                                data_label.setFont(QFont("Segoe UI", 11))
                                data_label.setStyleSheet("""
                                    padding: 8px;
                                    border: 1px solid #ddd;
                                """)
                                data_label.setWordWrap(True)
                                table_grid.addWidget(data_label, current_data_row, 1, 1, 2)
                                current_data_row += 1
                            
                            # Untuk memproses tdi_
                            elif col_b.startswith('tdi_'):
                                data_text = col_b[4:].strip()  # Hapus awalan 'tdi_'
                                
                                import re
                                placeholders = re.findall(r'\$P\d+\$', data_text)
                                
                                if placeholders:
                                    # Buat container untuk input field + teks
                                    container = QWidget()
                                    container_layout = QHBoxLayout(container)
                                    container_layout.setContentsMargins(8, 8, 8, 8)
                                    
                                    # Split teks berdasarkan placeholder
                                    parts = re.split(r'(\$P\d+\$)', data_text)
                                    
                                    for part in parts:
                                        if re.match(r'\$P\d+\$', part):
                                            # Ini placeholder, buat input field
                                            input_field = QLineEdit()
                                            input_field.setFixedWidth(50)
                                            input_field.setStyleSheet("""
                                                padding: 5px;
                                                border: 1px solid #ccc;
                                                border-radius: 4px;
                                            """)
                                            
                                            # Validator untuk membatasi input hanya angka 0-100
                                            from PyQt5.QtGui import QIntValidator
                                            validator = QIntValidator(0, 100)
                                            input_field.setValidator(validator)
                                            
                                            # Set nilai default berdasarkan placeholder
                                            if part == "$P1$":
                                                input_field.setText("30")
                                            elif part == "$P2$":
                                                input_field.setText("50")
                                            elif part == "$P3$":
                                                input_field.setText("15")
                                            elif part == "$P4$":
                                                input_field.setText("5")
                                            
                                            container_layout.addWidget(input_field)
                                            
                                            # Daftarkan field di data_fields
                                            field_key = f"tdi_{part}_{header_text}_{current_data_row}"
                                            self.data_fields[field_key] = input_field
                                        else:
                                            # Teks biasa
                                            text_label = QLabel(part)
                                            text_label.setFont(QFont("Segoe UI", 11))
                                            # Tidak ada border pada teks
                                            container_layout.addWidget(text_label)
                                    
                                    container_layout.addStretch()
                                    
                                    # Container tidak memiliki border
                                    container.setStyleSheet("background-color: transparent; border: none;")
                                    
                                    # Tambahkan ke grid, span seluruh kolom
                                    table_grid.addWidget(container, current_data_row, 1, 1, 2)
                                else:
                                    # Tidak ada placeholder, tampilkan sebagai teks biasa
                                    data_label = QLabel(data_text)
                                    data_label.setFont(QFont("Segoe UI", 11))
                                    # Tidak ada border pada teks
                                    data_label.setStyleSheet("padding: 8px; border: none;")
                                    data_label.setWordWrap(True)
                                    
                                    # Tambahkan ke grid, span seluruh kolom
                                    table_grid.addWidget(data_label, current_data_row, 1, 1, 2)
                                
                                current_data_row += 1
                            
                            elif col_b.startswith('tdm_'):
                                data_text_b = col_b[4:].strip()
                                
                                data_label_b = QLabel(data_text_b)
                                data_label_b.setFont(QFont("Segoe UI", 11))
                                data_label_b.setStyleSheet("""
                                    padding: 8px;
                                    border: 1px solid #ddd;
                                """)
                                data_label_b.setWordWrap(True)
                                table_grid.addWidget(data_label_b, current_data_row, 1)
                                
                                # Cek kolom C untuk tdm_ pasangan
                                if len(data_row) > 2 and not pd.isna(data_row.iloc[2]):
                                    col_c = data_row.iloc[2]
                                    
                                    if not isinstance(col_c, str):
                                        try:
                                            col_c = str(col_c)
                                        except:
                                            col_c = ""
                                    
                                    if col_c.startswith('tdm_'):
                                        data_text_c = col_c[4:].strip()
                                        
                                        data_label_c = QLabel(data_text_c)
                                        data_label_c.setFont(QFont("Segoe UI", 11))
                                        data_label_c.setStyleSheet("""
                                            padding: 8px;
                                            border: 1px solid #ddd;
                                        """)
                                        data_label_c.setWordWrap(True)
                                        table_grid.addWidget(data_label_c, current_data_row, 2)
                                
                                current_data_row += 1
                            
                            else:
                                # Data biasa tanpa prefix
                                data_text = str(col_b).strip()
                                data_label = QLabel(data_text)
                                data_label.setFont(QFont("Segoe UI", 11))
                                data_label.setStyleSheet("""
                                    padding: 8px;
                                    border: 1px solid #ddd;
                                """)
                                data_label.setWordWrap(True)
                                table_grid.addWidget(data_label, current_data_row, 1, 1, 2)
                                current_data_row += 1
                    
                    # Update table_row setelah memproses semua data
                    table_row = current_data_row
                    
                    # Simpan header terakhir
                    last_header = header_text
            else:
                # Ini bukan header di kolom A, mungkin baris lanjutan atau bukan bagian tabel
                # Jika sebelumnya kita dalam tabel tapi sekarang tidak ada header/data yang sesuai,
                # tandai bahwa kita keluar dari tabel
                if in_table and not col_a:
                    in_table = False
                    table_section = None
                    table_grid = None
                
            # Check if it's a section header (sub_)
            if isinstance(first_col, str) and first_col.startswith('sub_'):
                # Create a new section
                section_title = first_col[4:].strip()  # Remove 'sub_' prefix

                # Create section frame - removed border
                section_frame = QWidget()
                section_frame.setStyleSheet("""
                    background-color: white;
                """)

                section_layout = QVBoxLayout(section_frame)
                section_layout.setContentsMargins(15, 15, 15, 15)
                section_layout.setSpacing(15)

                # Add section title - ensure it's aligned with the left margin, no indent
                title_label = QLabel(section_title)
                title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
                title_label.setStyleSheet(f"color: {PRIMARY_COLOR}; background-color: transparent;")
                title_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                title_label.setIndent(0)  # Ensure no text indentation

                section_layout.addWidget(title_label)
                
                # Create a grid for this section
                section_grid = QGridLayout()
                section_grid.setHorizontalSpacing(30)  # Increase horizontal spacing
                section_grid.setVerticalSpacing(10)
                section_grid.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                section_layout.addLayout(section_grid)

                # Add section to main layout
                layout.addWidget(section_frame)
                layout.addSpacing(20)

                current_section = section_title
                left_section = section_title  # Track as left section
                current_row = 0  # Reset row counter for new section
                field_count = 0

                # Also check if there are additional sections in this row (columns to the right)
                for col_idx in range(1, df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        right_col_value = str(row[col_idx]).strip() if isinstance(row[col_idx], str) else ""
                        if right_col_value.startswith('sub_'):
                            right_section_title = right_col_value[4:].strip()  # Remove 'sub_' prefix
                            right_section = right_section_title  # Track as right section
                            break

                continue
                
            # Ensure we have a section grid to add fields to
            if section_layout is None:
                # If no section is defined yet, create a default one
                section_frame = QWidget()
                section_frame.setStyleSheet("""
                    background-color: white;
                """)

                section_layout = QVBoxLayout(section_frame)
                section_layout.setContentsMargins(15, 15, 15, 15)
                section_layout.setSpacing(15)
                
                # Create a grid for this section
                section_grid = QGridLayout()
                section_grid.setHorizontalSpacing(30)  # Increase horizontal spacing
                section_grid.setVerticalSpacing(10)
                section_layout.addLayout(section_grid)

                # Add to main layout
                layout.addWidget(section_frame)
                current_section = "Default"
                current_row = 0
            
            # Check if it's a field header (fh_)
            if first_col.startswith('fh_'):
                field_header = first_col[3:].strip()  # Remove 'fh_' prefix
                
                # Field header label
                header_label = QLabel(field_header)
                header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                header_label.setStyleSheet("color: #555; margin-top: 5px;")
                
                # Crucial settings to prevent indentation
                header_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                header_label.setIndent(0)  # Set text indent to 0
                
                # Add header to the grid - in column 0, same as regular field labels
                # Don't span columns to keep consistent alignment with field labels
                section_grid.addWidget(header_label, current_row, 0, 1, 1)
                
                # Check if there are column headers (ch_) in the same row
                has_ch_in_row = False
                ch_headers = []
                ch_columns = []  # Track the column positions of the headers
                
                for col_idx in range(1, df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            col_value = str(row[col_idx]).strip()
                        
                        if col_value.startswith('ch_'):
                            has_ch_in_row = True
                            ch_header_text = col_value[3:].strip()
                            ch_headers.append(ch_header_text)
                            ch_columns.append(col_idx)  # Save the actual column position
                
                # If we found ch_ headers in this row, update our current headers
                if has_ch_in_row:
                    current_header_labels = ch_headers
                    has_column_headers = True
                    
                    # Add the column headers to the grid in the same row as the field header
                    # This is the key change - align column headers horizontally with the field header
                    for i, header_text in enumerate(current_header_labels):
                        col_header_label = QLabel(header_text)
                        col_header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                        col_header_label.setStyleSheet("color: #555; margin-top: 5px; padding-left: 0px;")
                        col_header_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                        col_header_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                        col_header_label.setIndent(0)  # Prevent text indentation
                        
                        # Position headers at the same row as the field header
                        # For the first header (Name), put it at column 1
                        # For the second header (Phone No/Email), put it at column 2
                        section_grid.addWidget(col_header_label, current_row, i+1)
                
                # Check if there are fields or headers in columns C and beyond in the same row
                # Keep track of found right header to avoid duplicates
                right_header_found = False
                
                for col_idx in range(1, df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # Skip ch_ headers as we've already processed them
                        if col_value.startswith('ch_'):
                            continue
                            
                        # Check for any field type in the right columns - not just fh_
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fh_') or 
                            col_value.startswith('fm_')) and not right_header_found:
                            
                            # Handle right headers by prefix type
                            prefix = col_value[:2] if col_value.startswith('f_') else col_value[:3]
                            suffix = col_value[2:] if col_value.startswith('f_') else col_value[3:]
                            right_content = suffix.strip()
                            
                            right_header_found = True  # Mark that we found a right content
                            
                            # For header fields (fh_)
                            if col_value.startswith('fh_'):
                                # It's another header in the same row - this will be for the right section
                                right_header = right_content
                                
                                # If this is first header for right section, treat it as section title if we don't have one yet
                                if right_section is None:
                                    right_section = right_header
                                    
                                right_header_label = QLabel(right_header)
                                right_header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                                right_header_label.setStyleSheet("color: #555; margin-top: 5px;")
                                
                                # Position in the grid correctly - at the same row level as the left header
                                # But in columns 3-4 (index 3-4) to create proper separation
                                section_grid.addWidget(right_header_label, current_row, 3, 1, 2)
                            else:
                                # Handle non-header fields in the right section immediately after a header in the left
                                right_section_name = right_section if right_section else current_section
                                right_field_key = f"{sheet_name}_{right_section_name}_{field_count}"
                                field_count += 1
                                
                                # Create right field label
                                right_label = QLabel(right_content)
                                right_label.setFont(QFont("Segoe UI", 11))
                                right_label.setStyleSheet("color: #333; background-color: transparent;")
                                right_label.setMinimumWidth(250)  # Set minimum width for consistent layout
                                
                                # Add label to grid - position at the same row level as current header
                                section_grid.addWidget(right_label, current_row, 3)
                                
                                # Create input field based on type
                                if col_value.startswith('fd_'):
                                    # It's a dropdown
                                    right_input_field = QComboBox()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setMinimumWidth(200)
                                    right_input_field.setStyleSheet("""
                                        QComboBox {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QComboBox:hover {
                                            border: 1px solid #3498DB;
                                        }
                                        QComboBox::drop-down {
                                            subcontrol-origin: padding;
                                            subcontrol-position: top right;
                                            width: 20px;
                                            border-left-width: 1px;
                                            border-left-color: #ccc;
                                            border-left-style: solid;
                                            border-top-right-radius: 4px;
                                            border-bottom-right-radius: 4px;
                                        }
                                    """)
                                    
                                    # Get options for dropdown
                                    options = []
                                    right_cell_col = col_idx + 1
                                    right_cell_address = f"{chr(ord('A') + right_cell_col)}{index + 1}"

                                    # Get validation values first
                                    right_validation_options = self.get_validation_values(self.excel_path, sheet_name, right_cell_address)

                                    if right_validation_options:
                                        options = right_validation_options
                                    else:
                                        # Fallback if data validation not found
                                        if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                            options_str = str(row.iloc[col_idx + 1]).strip()
                                            options = [opt.strip() for opt in options_str.split(',')]
                                    
                                    # Add options and set default
                                    right_input_field.addItems(options)
                                    if len(options) > 0:
                                        right_input_field.setCurrentText(options[0])
                                        
                                    # Set value if available
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                        value = str(row.iloc[col_idx + 1]).strip()
                                        if value in options:
                                            right_input_field.setCurrentText(value)
                                    
                                    # Add to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                else:
                                    # It's a regular input field (f_ or fm_)
                                    right_input_field = QLineEdit()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setPlaceholderText(f"Enter {right_content}")
                                    right_input_field.setMinimumWidth(200)
                                    right_input_field.setStyleSheet("""
                                        QLineEdit {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QLineEdit:hover {
                                            border: 1px solid #3498DB;
                                        }
                                    """)
                                    
                                    # Set value if available
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                        right_input_field.setText(str(row.iloc[col_idx + 1]).strip())
                                    
                                    # Add to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                
                                # Register the field
                                self.data_fields[right_field_key] = right_input_field
                
                current_row += 1
                continue

            # Check if it's a column header row (first cell starts with ch_ or has ch_ cells in the row)
            elif first_col.startswith('ch_') or any(isinstance(cell, str) and cell.startswith('ch_') for cell in row if not pd.isna(cell)):
                has_column_headers = True
                current_header_labels = []
                
                # Process header row and collect all ch_ columns
                for col_idx in range(df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip() if isinstance(row[col_idx], str) else str(row[col_idx]).strip()
                        if col_value.startswith('ch_'):
                            header_text = col_value[3:].strip()  # Remove 'ch_' prefix
                            current_header_labels.append(header_text)
                
                # If we have header labels, create a header row
                if len(current_header_labels) > 0:
                    for col_idx, header_text in enumerate(current_header_labels):
                        col_header_label = QLabel(header_text)
                        col_header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                        col_header_label.setStyleSheet("color: #555; margin-top: 5px; padding-left: 0px;")
                        col_header_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                        col_header_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                        col_header_label.setIndent(0)  # Prevent text indentation
                        
                        # Grid column position depends on the header position
                        grid_col = col_idx  # Each field takes its own column in our grid
                        section_grid.addWidget(col_header_label, current_row, grid_col + 1)  # +1 to leave room for labels
                    
                    current_row += 1
                
                continue
            
            # Check if it's a field (f_)
            if first_col.startswith('f_'):
                field_name = first_col[2:].strip()  # Remove 'f_' prefix
        
                # Extract display name by removing numeric prefix if present
                display_name = field_name
                if field_name and field_name[0].isdigit():
                    # Remove digit prefix from display name
                    for i, char in enumerate(field_name):
                        if not char.isdigit():
                            display_name = field_name[i:]
                            break
                
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field label
                label = QLabel(display_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(250)  # Set minimum width for consistent layout
                label.setContentsMargins(0, 0, 0, 0)  # Ensure no margins
                label.setIndent(0)  # Set text indent to 0 for proper alignment
                
                # Add label to grid
                section_grid.addWidget(label, current_row, 0)

                # Create input field
                input_field = QLineEdit()
                input_field.setFont(QFont("Segoe UI", 11))
                input_field.setPlaceholderText(f"Enter {display_name}")
                input_field.setMinimumWidth(200)  # Set minimum width for consistent layout
                input_field.setStyleSheet("""
                    QLineEdit {
                        padding: 5px;
                        border: 1px solid #ccc;
                        border-radius: 4px;
                        background-color: white;
                        min-height: 28px;
                    }
                    QLineEdit:hover {
                        border: 1px solid #3498DB;
                    }
                """)
                
                # Set default value if available
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    input_field.setText(str(row.iloc[1]).strip())
                
                # Check if there's a unit value (fu_) in the next column
                has_unit = False
                unit_value = ""
                
                if len(row) > 2 and not pd.isna(row.iloc[2]):
                    unit_cell = str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) else ""
                    # Only process unit values with fu_ prefix
                    if isinstance(unit_cell, str) and unit_cell.startswith('fu_'):
                        has_unit = True
                        unit_value = unit_cell[3:].strip()  # Remove 'fu_' prefix
                
                # Check if there's any field in the right columns (columns C and beyond)
                has_right_field = False
                for col_idx in range(2, min(len(row), df.shape[1])):  # Start checking from column 2
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # If we find a unit (fu_), skip it for right field check
                        if isinstance(col_value, str) and col_value.startswith('fu_'):
                            continue
                            
                        # Check for any field prefix in the right columns
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fh_') or 
                            col_value.startswith('fm_')):
                            has_right_field = True
                            break
                
                # If there's a unit value, add it to the layout
                if has_unit:
                    # Create unit label
                    unit_label = QLabel(unit_value)
                    unit_label.setFont(QFont("Segoe UI", 11))
                    unit_label.setStyleSheet("color: #666; margin-left: 5px;")
                    unit_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    
                    # Add unit label without changing the original input field position
                    section_grid.addWidget(input_field, current_row, 1)
                    section_grid.addWidget(unit_label, current_row, 2)
                else:
                    # If there's no field in the right columns, make this field span wider
                    if not has_right_field:
                        section_grid.addWidget(input_field, current_row, 1, 1, 2)  # Span 2 columns
                    else:
                        # Add input field to grid normally
                        section_grid.addWidget(input_field, current_row, 1)
                
                # Register the field
                self.data_fields[field_key] = input_field
                
                # Check for fields in right section (columns to the right)
                # Process any field type (f_, fd_, fh_) in the right section
                right_field_found = False
                
                for col_idx in range(2, min(len(row), df.shape[1])):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            # Convert non-string values to string
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # Check for any field prefix in the right section (f_, fd_, fh_)
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fh_') or 
                            col_value.startswith('fm_')):
                            
                            # Extract the right field prefix and name accordingly
                            prefix = col_value[:2] if col_value.startswith('f_') else col_value[:3]
                            suffix = col_value[2:] if col_value.startswith('f_') else col_value[3:]
                            right_field_name = suffix.strip()
                            
                            # If it's a header field (fh_)
                            if col_value.startswith('fh_'):
                                if not right_field_found:  # Only process the first header in this row
                                    right_header_label = QLabel(right_field_name)
                                    right_header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                                    right_header_label.setStyleSheet("color: #555; margin-top: 5px;")
                                    right_header_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                                    right_header_label.setIndent(0)  # Prevent text indentation
                                    
                                    # Add header to the right section - position on column 3 only (don't span)
                                    section_grid.addWidget(right_header_label, current_row, 3, 1, 1)
                                    right_field_found = True
                                continue  # Skip further processing for headers
                                
                            # For regular fields or dropdowns
                            if not right_field_found:  # Only process the first field in this row
                                right_section_name = right_section if right_section else current_section
                                right_field_key = f"{sheet_name}_{right_section_name}_{field_count}"
                                field_count += 1
                                right_field_found = True
                                
                                # Create right field label
                                right_label = QLabel(right_field_name)
                                right_label.setFont(QFont("Segoe UI", 11))
                                right_label.setStyleSheet("color: #333; background-color: transparent;")
                                right_label.setMinimumWidth(250)  # Set minimum width for consistent layout
                                right_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                                right_label.setIndent(0)  # Prevent text indentation
                                
                                # Add label to grid - position at the same row level as current field
                                section_grid.addWidget(right_label, current_row, 3)
                                
                                # Handle different input field types based on prefix
                                if col_value.startswith('fd_'):
                                    # Create dropdown for right field
                                    right_input_field = QComboBox()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setMinimumWidth(200)  # Set minimum width for consistent layout
                                    right_input_field.setStyleSheet("""
                                        QComboBox {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QComboBox:hover {
                                            border: 1px solid #3498DB;
                                        }
                                        QComboBox::drop-down {
                                            subcontrol-origin: padding;
                                            subcontrol-position: top right;
                                            width: 20px;
                                            border-left-width: 1px;
                                            border-left-color: #ccc;
                                            border-left-style: solid;
                                            border-top-right-radius: 4px;
                                            border-bottom-right-radius: 4px;
                                        }
                                    """)
                                    
                                    # Get options for right dropdown
                                    right_options = []
                                    right_cell_col = col_idx + 1
                                    right_cell_address = f"{chr(ord('A') + right_cell_col)}{index + 1}"
                                    
                                    right_validation_options = self.get_validation_values(self.excel_path, sheet_name, right_cell_address)
                                    
                                    if right_validation_options:
                                        right_options = right_validation_options
                                    else:
                                        # Fallback if data validation not found
                                        if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                            right_options_str = str(row.iloc[col_idx + 1]).strip()
                                            right_options = [opt.strip() for opt in right_options_str.split(',')]
                                    
                                    # Add options to right dropdown
                                    right_input_field.addItems(right_options)
                                    
                                    # Set default value if available
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]) and str(row.iloc[col_idx + 1]).strip() in right_options:
                                        right_input_field.setCurrentText(str(row.iloc[col_idx + 1]).strip())
                                    elif len(right_options) > 0:
                                        right_input_field.setCurrentText(right_options[0])
                                    
                                    # Add dropdown to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                else:
                                    # Regular input field (f_ or fm_)
                                    right_input_field = QLineEdit()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setPlaceholderText(f"Enter {right_field_name}")
                                    right_input_field.setMinimumWidth(200)  # Set minimum width for consistent layout
                                    right_input_field.setStyleSheet("""
                                        QLineEdit {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QLineEdit:hover {
                                            border: 1px solid #3498DB;
                                        }
                                    """)
                                    
                                    # Set default value if available
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                        right_input_field.setText(str(row.iloc[col_idx + 1]).strip())
                                    
                                    # Add input field to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                
                                # Register the field
                                self.data_fields[right_field_key] = right_input_field
                
                current_row += 1
                continue
            
            # Check if it's a field dropdown (fd_)
            if first_col.startswith('fd_'):
                field_name = first_col[3:].strip()  # Remove 'fd_' prefix
        
                # Extract display name by removing numeric prefix
                display_name = field_name
                if field_name and field_name[0].isdigit():
                    for i, char in enumerate(field_name):
                        if not char.isdigit():
                            display_name = field_name[i:]
                            break
                        
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field label
                label = QLabel(display_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(250)  # Set minimum width for consistent layout
                
                # Add label to grid
                section_grid.addWidget(label, current_row, 0)

                # Create dropdown
                input_field = QComboBox()
                input_field.setFont(QFont("Segoe UI", 11))
                input_field.setMinimumWidth(200)  # Set minimum width for consistent layout
                input_field.setStyleSheet("""
                    QComboBox {
                        padding: 5px;
                        border: 1px solid #ccc;
                        border-radius: 4px;
                        background-color: white;
                        min-height: 28px;
                    }
                    QComboBox:hover {
                        border: 1px solid #3498DB;
                    }
                    QComboBox::drop-down {
                        subcontrol-origin: padding;
                        subcontrol-position: top right;
                        width: 20px;
                        border-left-width: 1px;
                        border-left-color: #ccc;
                        border-left-style: solid;
                        border-top-right-radius: 4px;
                        border-bottom-right-radius: 4px;
                    }
                """)
                
                is_industry_field = field_name == "Industry Classification"
                is_sub_industry_field = field_name == "Sub Industry Specification"
                
                
                # Try to get options from data validation or from second column
                options = []
                # If Excel stores cell position
                row_index = index + 1  # +1 because Excel starts from 1
                col_index = 2  # Assuming column B for dropdown value
                cell_address = f"{chr(ord('A') + col_index-1)}{row_index}"
                
                # Untuk field Industry Classification, gunakan data dari mapping
                if is_industry_field:
                    # Gunakan list industry dari mapping
                    options = list(INDUSTRY_SUBTYPE_MAPPING.keys())
                    industry_dropdown = input_field
                    industry_field_key = field_key
                else:
                    # Untuk field lain, gunakan metode biasa
                    # Try to get from data validation
                    validation_options = self.get_validation_values(self.excel_path, sheet_name, cell_address)
                    
                    if validation_options:
                        options = validation_options
                    else:
                        # Fallback to old method if data validation not found
                        if len(row) > 1 and not pd.isna(row.iloc[1]):
                            options_str = str(row.iloc[1]).strip()
                            options = [opt.strip() for opt in options_str.split(',')]
                
                # Jika ini adalah Sub Industry Specification field
                if is_sub_industry_field:
                    sub_industry_dropdown = input_field
                    sub_industry_field_key = field_key
                    # Tambahkan placeholder, nilai sebenarnya akan diisi nanti saat industry dipilih
                    options = ["-- Select Industry first --"]
                
                # Add options and set default if available
                input_field.addItems(options)
                
                # Tambahkan stylesheet untuk tooltip
                app = QApplication.instance()
                app.setStyleSheet("""
                    QToolTip {
                        background-color: #F5F5F5;
                        color: #333333;
                        border: 1px solid #CCCCCC;
                        padding: 5px;
                        font: 10pt "Segoe UI";
                        opacity: 255;
                    }
                """)
                
                # Periksa field untuk menambahkan tooltip yang sesuai
                if field_name == "Seismic Hazard Zone":
                    # Tambahkan tooltip untuk setiap zona
                    for i in range(input_field.count()):
                        zone_text = input_field.itemText(i)
                        if zone_text in SEISMIC_ZONE_DESCRIPTIONS:
                            input_field.setItemData(i, SEISMIC_ZONE_DESCRIPTIONS[zone_text], Qt.ToolTipRole)
                
                # Tambahkan kondisi untuk Wind Speed Zone
                elif field_name == "Wind Speed Zone":
                    # Tambahkan tooltip untuk setiap level
                    for i in range(input_field.count()):
                        level_text = input_field.itemText(i)
                        if level_text in WIND_SPEED_DESCRIPTIONS:
                            input_field.setItemData(i, WIND_SPEED_DESCRIPTIONS[level_text], Qt.ToolTipRole)
                
                # Set default value if available
                default_value = ""
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    default_value = str(row.iloc[1]).strip()
                    
                if default_value and default_value in options:
                    input_field.setCurrentText(default_value)
                elif len(options) > 0:
                    input_field.setCurrentText(options[0])
                    
                # Check if there's any field in the right columns (columns C and beyond)
                has_right_field = False
                for col_idx in range(2, min(len(row), df.shape[1])):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # Check for any field prefix in the right columns
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fh_') or 
                            col_value.startswith('fm_')):
                            has_right_field = True
                            break
                
                # If there's no field in the right columns, make this field span to match fm_ fields
                if not has_right_field:
                    # Set dropdown to span only columns 1 and 2 to align with fm_ fields
                    section_grid.addWidget(input_field, current_row, 1, 1, 2)  # Span 2 columns
                else:
                    # Add dropdown to grid normally
                    section_grid.addWidget(input_field, current_row, 1)
                    
                # Register the field
                self.data_fields[field_key] = input_field
                
                # Register the field
                self.data_fields[field_key] = input_field
                
                # Check for fields in right section (columns to the right)
                # Process any field type (f_, fd_, fh_, fm_) in the right section
                right_field_found = False
                right_header_found = False
                
                for col_idx in range(2, min(len(row), df.shape[1])):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            # Convert non-string values to string
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # Check for any field prefix in the right section
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fh_') or 
                            col_value.startswith('fm_')):
                            
                            # Extract the right field prefix and name accordingly
                            prefix = col_value[:2] if col_value.startswith('f_') else col_value[:3]
                            suffix = col_value[2:] if col_value.startswith('f_') else col_value[3:]
                            right_field_name = suffix.strip()
                            
                            # Special handling for header fields (fh_) in the right section
                            if col_value.startswith('fh_') and not right_header_found:
                                right_header_found = True  # Track that we found a header
                                right_field_found = True   # Consider this as a field being processed
                                
                                # Create right field header
                                right_header_label = QLabel(right_field_name)
                                right_header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                                right_header_label.setStyleSheet("color: #555; margin-top: 5px;")
                                right_header_label.setContentsMargins(0, 0, 0, 0)  # Remove any default margins
                                right_header_label.setIndent(0)  # Prevent text indentation
                                
                                # Add header to the grid at the same row level - don't span columns
                                section_grid.addWidget(right_header_label, current_row, 3, 1, 1)
                                
                                # If this is first header for right section, treat it as section title if we don't have one yet
                                if right_section is None:
                                    right_section = right_field_name
                            
                            # For regular fields or dropdowns (if no header was found yet)
                            elif not right_field_found:
                                right_section_name = right_section if right_section else current_section
                                right_field_key = f"{sheet_name}_{right_section_name}_{field_count}"
                                field_count += 1
                                right_field_found = True
                                
                                # Create right field label
                                right_label = QLabel(right_field_name)
                                right_label.setFont(QFont("Segoe UI", 11))
                                right_label.setStyleSheet("color: #333; background-color: transparent;")
                                right_label.setMinimumWidth(250)  # Set minimum width for consistent layout
                                
                                # Add label to grid - position at the same row level as current field
                                section_grid.addWidget(right_label, current_row, 3)
                                
                                # Handle different input field types based on prefix
                                if col_value.startswith('fd_'):
                                    # Create dropdown
                                    right_input_field = QComboBox()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setMinimumWidth(200)
                                    right_input_field.setStyleSheet("""
                                        QComboBox {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QComboBox:hover {
                                            border: 1px solid #3498DB;
                                        }
                                        QComboBox::drop-down {
                                            subcontrol-origin: padding;
                                            subcontrol-position: top right;
                                            width: 20px;
                                            border-left-width: 1px;
                                            border-left-color: #ccc;
                                            border-left-style: solid;
                                            border-top-right-radius: 4px;
                                            border-bottom-right-radius: 4px;
                                        }
                                    """)
                                    
                                    # Get options for dropdown
                                    right_options = []
                                    right_cell_col = col_idx + 1
                                    right_cell_address = f"{chr(ord('A') + right_cell_col)}{index + 1}"
                                    
                                    right_validation_options = self.get_validation_values(self.excel_path, sheet_name, right_cell_address)
                                    
                                    if right_validation_options:
                                        right_options = right_validation_options
                                    else:
                                        if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                            right_options_str = str(row.iloc[col_idx + 1]).strip()
                                            right_options = [opt.strip() for opt in right_options_str.split(',')]
                                    
                                    # Add options
                                    right_input_field.addItems(right_options)
                                    
                                    # Set default value
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]) and str(row.iloc[col_idx + 1]).strip() in right_options:
                                        right_input_field.setCurrentText(str(row.iloc[col_idx + 1]).strip())
                                    elif len(right_options) > 0:
                                        right_input_field.setCurrentText(right_options[0])
                                    
                                    # Add to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                else:
                                    # Regular input field
                                    right_input_field = QLineEdit()
                                    right_input_field.setFont(QFont("Segoe UI", 11))
                                    right_input_field.setPlaceholderText(f"Enter {right_field_name}")
                                    right_input_field.setMinimumWidth(200)
                                    right_input_field.setStyleSheet("""
                                        QLineEdit {
                                            padding: 5px;
                                            border: 1px solid #ccc;
                                            border-radius: 4px;
                                            background-color: white;
                                            min-height: 28px;
                                        }
                                        QLineEdit:hover {
                                            border: 1px solid #3498DB;
                                        }
                                    """)
                                    
                                    # Set value if available
                                    if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                        right_input_field.setText(str(row.iloc[col_idx + 1]).strip())
                                    
                                    # Add to grid
                                    section_grid.addWidget(right_input_field, current_row, 4)
                                
                                # Register field
                                self.data_fields[right_field_key] = right_input_field
                
                current_row += 1
                continue
                
            # Check if it's a field multiple (fm_)
            if first_col.startswith('fm_'):
                field_name = first_col[3:].strip()  # Remove 'fd_' prefix
        
                # Extract display name by removing numeric prefix
                display_name = field_name
                if field_name and field_name[0].isdigit():
                    for i, char in enumerate(field_name):
                        if not char.isdigit():
                            display_name = field_name[i:]
                            break
                        
                field_key_base = f"{sheet_name}_{current_section}_{field_name}"
                
                # Create field label
                label = QLabel(display_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(250)  # Set minimum width for consistent layout
                
                # Add label to grid
                section_grid.addWidget(label, current_row, 0)
                
                # Determine if we have column headers to use for field names
                header_names = []
                if has_column_headers and len(current_header_labels) > 0:
                    header_names = current_header_labels
                else:
                    # Default fallback column names based on the screenshot
                    header_names = ["Name", "Phone No/Email"]
                
                # KEY CHANGE: Instead of creating a container widget, we'll place fields directly in the grid
                # This aligns the fields with the column headers above
                
                # Create input fields based on the number of headers
                for i, header in enumerate(header_names):
                    if i >= 2:  # Limit to 2 columns in case there are more headers
                        break
                        
                    input_field = QLineEdit()
                    input_field.setFont(QFont("Segoe UI", 11))
                    input_field.setMinimumWidth(180)
                    
                    # Include the field name in the placeholder text
                    input_field.setPlaceholderText(f"Enter {display_name} {header}")
                    
                    input_field.setStyleSheet("""
                        QLineEdit {
                            padding: 5px;
                            border: 1px solid #ccc;
                            border-radius: 4px;
                            background-color: white;
                            min-height: 28px;
                        }
                        QLineEdit:hover {
                            border: 1px solid #3498DB;
                        }
                    """)
                    
                    # Set value if available
                    if i+1 < len(row) and not pd.isna(row.iloc[i+1]):
                        input_field.setText(str(row.iloc[i+1]).strip())
                    
                    # Place each field directly in its corresponding column position
                    # First field (Name) goes in column 1, second field (Phone No/Email) goes in column 2
                    section_grid.addWidget(input_field, current_row, i+1)
                    
                    # Register field
                    self.data_fields[f"{field_key_base}_{i}"] = input_field
                    field_count += 1
                
                # Check for fields in right section (columns to the right)
                right_field_found = False
                
                for col_idx in range(3, min(len(row), df.shape[1])):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = ""
                        if isinstance(row[col_idx], str):
                            col_value = row[col_idx].strip()
                        elif not pd.isna(row[col_idx]):
                            col_value = str(row[col_idx]).strip()
                        else:
                            continue
                            
                        # Check for any field prefix in right section
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fm_') or
                            col_value.startswith('fh_')):
                            
                            # Extract the prefix and name
                            if col_value.startswith('f_'):
                                prefix = col_value[:2]
                                suffix = col_value[2:]
                            else:
                                prefix = col_value[:3]
                                suffix = col_value[3:]
                            
                            right_field_name = suffix.strip()
                            
                            # Skip headers
                            if col_value.startswith('fh_'):
                                continue
                                
                            # Process the first field found in right section
                            if not right_field_found:
                                right_field_found = True
                                
                                right_section_name = right_section if right_section else current_section
                                right_field_key_base = f"{sheet_name}_{right_section_name}_{right_field_name}"
                                
                                # Create right field label
                                right_label = QLabel(right_field_name)
                                right_label.setFont(QFont("Segoe UI", 11))
                                right_label.setStyleSheet("color: #333; background-color: transparent;")
                                right_label.setMinimumWidth(250)
                                
                                section_grid.addWidget(right_label, current_row, 3)
                                
                                # For multiple fields (fm_) in right section
                                if col_value.startswith('fm_'):
                                    # Create container for right fields
                                    right_container = QWidget()
                                    right_layout = QHBoxLayout(right_container)
                                    right_layout.setContentsMargins(0, 0, 0, 0)
                                    right_layout.setSpacing(10)
                                    right_layout.setAlignment(Qt.AlignLeft)
                                    
                                    # Create fields based on number of headers
                                    for i, header in enumerate(header_names):
                                        if i >= 2:  # Limit to 2 columns
                                            break
                                            
                                        right_input = QLineEdit()
                                        right_input.setFont(QFont("Segoe UI", 11))
                                        right_input.setPlaceholderText(f"Enter {header}")
                                        right_input.setStyleSheet("""
                                            QLineEdit {
                                                padding: 5px;
                                                border: 1px solid #ccc;
                                                border-radius: 4px;
                                                background-color: white;
                                                min-height: 28px;
                                            }
                                            QLineEdit:hover {
                                                border: 1px solid #3498DB;
                                            }
                                        """)
                                        
                                        # Set value if available
                                        if col_idx + i + 1 < len(row) and not pd.isna(row.iloc[col_idx + i + 1]):
                                            right_input.setText(str(row.iloc[col_idx + i + 1]).strip())
                                        
                                        right_layout.addWidget(right_input)
                                        
                                        # Register field
                                        self.data_fields[f"{right_field_key_base}_{i}"] = right_input
                                        field_count += 1
                                    
                                    # Add container to grid
                                    section_grid.addWidget(right_container, current_row, 4, 1, 1)
                                else:
                                    # For other field types
                                    if col_value.startswith('fd_'):
                                        # Create dropdown
                                        right_input = QComboBox()
                                        right_input.setFont(QFont("Segoe UI", 11))
                                        right_input.setMinimumWidth(200)
                                        right_input.setStyleSheet("""
                                            QComboBox {
                                                padding: 5px;
                                                border: 1px solid #ccc;
                                                border-radius: 4px;
                                                background-color: white;
                                                min-height: 28px;
                                            }
                                            QComboBox:hover {
                                                border: 1px solid #3498DB;
                                            }
                                            QComboBox::drop-down {
                                                subcontrol-origin: padding;
                                                subcontrol-position: top right;
                                                width: 20px;
                                                border-left-width: 1px;
                                                border-left-color: #ccc;
                                                border-left-style: solid;
                                                border-top-right-radius: 4px;
                                                border-bottom-right-radius: 4px;
                                            }
                                        """)
                                        
                                        # Get options
                                        right_options = []
                                        if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                            right_options_str = str(row.iloc[col_idx + 1]).strip()
                                            right_options = [opt.strip() for opt in right_options_str.split(',')]
                                        
                                        right_input.addItems(right_options)
                                        
                                        if len(right_options) > 0:
                                            right_input.setCurrentText(right_options[0])
                                        
                                        section_grid.addWidget(right_input, current_row, 4)
                                    else:
                                        # Regular input field
                                        right_input = QLineEdit()
                                        right_input.setFont(QFont("Segoe UI", 11))
                                        right_input.setPlaceholderText(f"Enter {right_field_name}")
                                        right_input.setStyleSheet("""
                                            QLineEdit {
                                                padding: 5px;
                                                border: 1px solid #ccc;
                                                border-radius: 4px;
                                                background-color: white;
                                                min-height: 28px;
                                            }
                                            QLineEdit:hover {
                                                border: 1px solid #3498DB;
                                            }
                                        """)
                                        
                                        # Set value
                                        if col_idx + 1 < len(row) and not pd.isna(row.iloc[col_idx + 1]):
                                            right_input.setText(str(row.iloc[col_idx + 1]).strip())
                                        
                                        section_grid.addWidget(right_input, current_row, 4)
                                    
                                    # Register field
                                    self.data_fields[f"{right_field_key_base}_0"] = right_input
                                    field_count += 1
                            
                            break  # Only process the first field with a prefix
                
                current_row += 1
                continue
            
            # Check if it's a table group (ftg_)
            if first_col.startswith('ftg_'):
                group_name = first_col[4:].strip()  # Remove 'ftg_' prefix
                
                # Create group header label - span across all columns
                group_label = QLabel(group_name)
                group_label.setFont(QFont("Segoe UI", 11, QFont.Bold))  # Make it bold
                group_label.setStyleSheet("""
                    color: #333; 
                    background-color: #f0f0f0; 
                    padding: 5px;
                    border-bottom: 1px solid #ccc;
                """)
                group_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                
                # Add label to grid - span across all columns (4 columns: task, client, contractor, remarks)
                section_grid.addWidget(group_label, current_row, 0, 1, 4)
                
                current_row += 1
                continue
            
            # Check if it's a table note (ftn_)
            if first_col.startswith('ftn_'):
                note_text = first_col[4:].strip()  # Remove 'ftn_' prefix
                
                # Create note label with italic style
                note_label = QLabel(note_text)
                note_label.setFont(QFont("Segoe UI", 10))
                note_label.setStyleSheet("""
                    color: #666; 
                    background-color: #f9f9f9; 
                    padding: 3px 5px;
                    font-style: italic;
                """)
                note_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                
                # Add label to grid - span across all columns
                section_grid.addWidget(note_label, current_row, 0, 1, 4)
                
                current_row += 1
                continue

            # Check if it's a table item (ft_)
            elif first_col.startswith('ft_'):
                task_name = first_col[3:].strip()  # Remove 'ft_' prefix
                field_key_base = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1
                
                # Create task label
                task_label = QLabel(task_name)
                task_label.setFont(QFont("Segoe UI", 11))
                task_label.setStyleSheet("color: #333; background-color: transparent;")
                task_label.setMinimumWidth(1000)  # Wider for task text
                task_label.setWordWrap(True)  # Allow text wrapping for long task names
                
                # Add label to grid - span only column 0
                section_grid.addWidget(task_label, current_row, 0)
                
                # Create checkbox for Client
                client_checkbox = QCheckBox()
                client_checkbox.setStyleSheet("""
                    QCheckBox {
                        min-height: 20px;
                    }
                    QCheckBox::indicator {
                        width: 18px;
                        height: 18px;
                    }
                """)
                
                # Create container widget for client checkbox to center it
                client_container = QWidget()
                client_layout = QHBoxLayout(client_container)
                client_layout.setAlignment(Qt.AlignCenter)
                client_layout.setContentsMargins(0, 0, 0, 0)
                client_layout.addWidget(client_checkbox)
                
                section_grid.addWidget(client_container, current_row, 1)
                
                # Create checkbox for Contractor
                contractor_checkbox = QCheckBox()
                contractor_checkbox.setStyleSheet("""
                    QCheckBox {
                        min-height: 20px;
                    }
                    QCheckBox::indicator {
                        width: 18px;
                        height: 18px;
                    }
                """)
                
                # Create container widget for contractor checkbox to center it
                contractor_container = QWidget()
                contractor_layout = QHBoxLayout(contractor_container)
                contractor_layout.setAlignment(Qt.AlignCenter)
                contractor_layout.setContentsMargins(0, 0, 0, 0)
                contractor_layout.addWidget(contractor_checkbox)
                
                section_grid.addWidget(contractor_container, current_row, 2)
                
                # Create input field for Remarks
                remarks_input = QLineEdit()
                remarks_input.setFont(QFont("Segoe UI", 11))
                remarks_input.setPlaceholderText("Enter remarks")
                remarks_input.setStyleSheet("""
                    QLineEdit {
                        padding: 5px;
                        border: 1px solid #ccc;
                        border-radius: 4px;
                        background-color: white;
                        min-height: 28px;
                    }
                    QLineEdit:hover {
                        border: 1px solid #3498DB;
                    }
                """)
                section_grid.addWidget(remarks_input, current_row, 3)
                
                # Set default values if available
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    # For Client checkbox
                    client_value = str(row.iloc[1]).strip()
                    if client_value == 'ü' or client_value.lower() == 'true':
                        client_checkbox.setChecked(True)
                        
                if len(row) > 2 and not pd.isna(row.iloc[2]):
                    # For Contractor checkbox
                    contractor_value = str(row.iloc[2]).strip()
                    if contractor_value == 'ü' or contractor_value.lower() == 'true':
                        contractor_checkbox.setChecked(True)
                        
                if len(row) > 3 and not pd.isna(row.iloc[3]):
                    # For Remarks input
                    remarks_input.setText(str(row.iloc[3]).strip())
                
                def create_checkbox_handler(client_cb, contractor_cb, source):
                    def handler():
                        if source == 'client' and client_cb.isChecked():
                            contractor_cb.setChecked(False)
                        elif source == 'contractor' and contractor_cb.isChecked():
                            client_cb.setChecked(False)
                    return handler

                # Hubungkan dengan cara yang lebih eksplisit
                client_handler = create_checkbox_handler(client_checkbox, contractor_checkbox, 'client')
                contractor_handler = create_checkbox_handler(client_checkbox, contractor_checkbox, 'contractor')

                client_checkbox.clicked.connect(client_handler)
                contractor_checkbox.clicked.connect(contractor_handler)
                
                # Register the fields
                self.data_fields[f"{field_key_base}_client"] = client_checkbox
                self.data_fields[f"{field_key_base}_contractor"] = contractor_checkbox
                self.data_fields[f"{field_key_base}_remarks"] = remarks_input
                
                current_row += 1
                continue

        # Setelah semua field diproses, hubungkan dropdown yang saling terkait
        if industry_dropdown and sub_industry_dropdown:
            # Simpan pasangan dropdown untuk referensi nanti
            self.linked_dropdowns[industry_field_key] = sub_industry_field_key
            
            # Hubungkan event dropdown industry ke fungsi update untuk dropdown sub-industry
            industry_dropdown.currentTextChanged.connect(
                lambda text, child=sub_industry_dropdown: self.update_dependent_dropdown(text, child))
            
            # Trigger update awal untuk mengisi sub-industry berdasarkan nilai industry yang sudah dipilih
            industry_value = industry_dropdown.currentText()
            if industry_value:
                self.update_dependent_dropdown(industry_value, sub_industry_dropdown)
                
                # Cari nilai sub-industry yang tersimpan di Excel
                sub_industry_value = None
                sub_industry_row = None
                
                # Cari baris yang berisi 'fd_Sub Industry Specification'
                for row_idx, row in df.iterrows():
                    if pd.isna(row.iloc[0]):
                        continue
                        
                    first_col = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                    if first_col == 'fd_Sub Industry Specification':
                        sub_industry_row = row_idx
                        # Ambil nilai dari kolom B (indeks 1)
                        if len(row) > 1 and not pd.isna(row.iloc[1]):
                            sub_industry_value = str(row.iloc[1]).strip()
                        break
                
                # Jika nilai sub-industry ditemukan, set ke dropdown
                if sub_industry_value and sub_industry_value in [sub_industry_dropdown.itemText(i) for i in range(sub_industry_dropdown.count())]:
                    sub_industry_dropdown.setCurrentText(sub_industry_value)
                    print(f"Setting Sub Industry dropdown to saved value: {sub_industry_value}")

        # Process any excel images that may exist in this sheet
        if section_layout:
            self.process_excel_images(sheet_name, section_layout)
            
        # Add a save button for the sheet at the end (only for DIP sheets)
        if section_layout and sheet_name.startswith("DIP_"):
            # Add spacer
            section_layout.addSpacing(10)

            # Save button
            save_btn = QPushButton("Save Changes")
            save_btn.setFont(QFont("Segoe UI", 11, QFont.Bold))
            save_btn.setCursor(Qt.PointingHandCursor)
            save_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {SECONDARY_COLOR};
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 15px;
                }}
                QPushButton:hover {{
                    background-color: #2980B9;
                }}
            """)
            save_btn.clicked.connect(lambda: self.save_sheet_data(sheet_name))

            button_layout = QHBoxLayout()
            button_layout.addStretch()
            button_layout.addWidget(save_btn)

            section_layout.addLayout(button_layout)
            
        # If no content was added, add a default message
        if not layout.count():
            no_data_label = QLabel("No form fields found in this sheet.")
            no_data_label.setAlignment(Qt.AlignCenter)
            no_data_label.setStyleSheet("color: #666; margin: 20px;")
            layout.addWidget(no_data_label)
    
    # Implementasi method lainnya yang perlu tetap ada (dari kode asli)
    def get_validation_values(self, excel_path, sheet_name, cell_address):
        """Mengambil nilai dari data validation di sebuah sel Excel"""
        from openpyxl import load_workbook
        
        try:
            # Pastikan untuk memuat dengan data_only=False agar kita bisa mengakses validasi
            workbook = load_workbook(excel_path, data_only=False)
            
            if sheet_name not in workbook.sheetnames:
                return []
                
            sheet = workbook[sheet_name]
            
            # Periksa apakah cell address valid
            try:
                cell = sheet[cell_address]
            except:
                return []
            
            # Cek data validation secara eksplisit
            dv = sheet.data_validations.dataValidation
            for validation in dv:
                for coord in validation.sqref.ranges:
                    if cell.coordinate in str(coord):
                        # Ditemukan validasi untuk sel ini
                        if validation.type == "list":
                            formula = validation.formula1
                            
                            # Jika formula menggunakan referensi
                            if formula.startswith('='):
                                # Implementasi sama seperti sebelumnya...
                                pass
                            else:
                                # Untuk list langsung seperti "A,B,C"
                                if formula.startswith('"') and formula.endswith('"'):
                                    formula = formula[1:-1]
                                return [val.strip() for val in formula.split(',')]
            
            # Fallback: Coba cara lain untuk mendapatkan validation list
            try:
                # Untuk beberapa versi openpyxl, langsung coba akses data_validation
                if hasattr(cell, 'data_validation') and cell.data_validation and hasattr(cell.data_validation, 'type'):
                    if cell.data_validation.type == 'list':
                        formula = cell.data_validation.formula1
                        if formula.startswith('"') and formula.endswith('"'):
                            formula = formula[1:-1]
                        return [val.strip() for val in formula.split(',')]
            except:
                pass
                
            return []
        except Exception as e:
            print(f"Error saat membaca data validation: {str(e)}")
            return []
        
    # Kode Baru
    def save_sheet_data(self, sheet_name):
        """Save the form data back to the Excel file"""
        try:
            # Check if the sheet name exists
            if not sheet_name.startswith("DIP_"):
                QMessageBox.warning(self, "Warning", "Only DIP sheets can be saved.")
                return
            
            import pandas as pd
            from openpyxl import load_workbook
            import os
            import time
            
            # Check if file exists and is accessible
            if not os.path.exists(self.excel_path):
                QMessageBox.critical(self, "Error", f"Excel file not found: {self.excel_path}")
                return
                
            # Check if file is not opened by another process
            try:
                # Try to open file in append mode to check if it's locked
                with open(self.excel_path, 'a'):
                    pass
            except PermissionError:
                QMessageBox.critical(self, "Error", "Excel file is currently opened by another application. Please close it and try again.")
                return
            
            print(f"Starting to save data to sheet: {sheet_name}")
            
            # Create a backup of the Excel file
            backup_path = self.excel_path + ".bak"
            try:
                import shutil
                shutil.copy2(self.excel_path, backup_path)
                print(f"Backup created at: {backup_path}")
            except Exception as e:
                print(f"Warning: Could not create backup: {str(e)}")
            
            # Load the Excel workbook with openpyxl
            wb = load_workbook(self.excel_path)
            
            if sheet_name not in wb.sheetnames:
                QMessageBox.critical(self, "Error", f"Sheet '{sheet_name}' not found in the Excel file.")
                return
            
            # Get the sheet
            sheet = wb[sheet_name]
            
            # Also load with pandas to help us find the field positions
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
            
            # Create a mapping of field IDs to row indices and track dropdown options
            field_to_row = {}
            right_fields = {}  # Track right fields by row and column
            dropdown_options = {}  # Track dropdown options by row and column
            
            # First pass: scan the Excel to build a map of field identifiers to row numbers
            # and collect dropdown options
            for row_idx, row in df.iterrows():
                # Skip empty rows
                if pd.isna(row).all():
                    continue
                
                # Get first column value
                first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
                
                # Convert to string if not already
                if not isinstance(first_col, str):
                    try:
                        first_col = str(first_col)
                    except:
                        continue
                
                # Store the row index for this field
                if first_col.startswith('f_') or first_col.startswith('fd_') or first_col.startswith('fm_') or first_col.startswith('ft_'):
                    field_to_row[first_col] = row_idx
                    
                    # Extract the field name without prefix
                    if first_col.startswith('f_'):
                        field_name = first_col[2:].strip()
                    elif first_col.startswith('fd_') or first_col.startswith('fm_') or first_col.startswith('ft_'):
                        field_name = first_col[3:].strip()
                    else:
                        field_name = first_col
                        
                    # Store the original field identifier with row index for easy lookup
                    field_to_row[f"row_{row_idx}"] = first_col
                    
                    # For dropdown fields (fd_), store their options
                    if first_col.startswith('fd_'):
                        # Options should be in column B (index 1)
                        if len(row) > 1 and not pd.isna(row.iloc[1]):
                            options_str = str(row.iloc[1]).strip()
                            options = [opt.strip() for opt in options_str.split(',')]
                            dropdown_options[row_idx] = {
                                'col': 1,  # Column B
                                'options': options
                            }
                            print(f"Found dropdown options for row {row_idx}: {options}")
                
                # IMPORTANT: Process fields in ALL columns to track dropdown options
                for col_idx in range(len(row)):
                    if not pd.isna(row[col_idx]):
                        col_value = row.iloc[col_idx] if not pd.isna(row.iloc[col_idx]) else ""
                        
                        # Convert to string if needed
                        if not isinstance(col_value, str):
                            try:
                                col_value = str(col_value)
                            except:
                                continue
                        
                        # Store field identifiers from ALL columns
                        if (col_value.startswith('f_') or 
                            col_value.startswith('fd_') or 
                            col_value.startswith('fm_') or
                            col_value.startswith('fh_')):
                            
                            # Create a key that includes the row and column position
                            field_key = f"col_{col_idx}_row_{row_idx}"
                            right_fields[field_key] = col_value
                            
                            # Also store by the field identifier for reverse lookup
                            right_fields[col_value] = {"row": row_idx, "col": col_idx}
                            
                            # For dropdown fields (fd_), store their options
                            if col_value.startswith('fd_'):
                                # Options should be in the next column
                                next_col = col_idx + 1
                                if next_col < len(row) and not pd.isna(row.iloc[next_col]):
                                    options_str = str(row.iloc[next_col]).strip()
                                    options = [opt.strip() for opt in options_str.split(',')]
                                    dropdown_options[f"{row_idx}_{col_idx}"] = {
                                        'col': next_col,
                                        'options': options
                                    }
                                    print(f"Found dropdown options for row {row_idx}, col {col_idx}: {options}")
            
            print(f"Found {len(field_to_row)} field mappings in the Excel sheet")
            print(f"Found {len(dropdown_options)} dropdown option sets")
            
            # Create validation data maps to help match dropdowns with their correct options
            # This will store cell positions and their validation options
            validation_data = {}
            
            # Load validation info from workbook
            from openpyxl.worksheet.datavalidation import DataValidation
            
            # Process all data validations in the sheet
            if hasattr(sheet, 'data_validations'):
                for validation in sheet.data_validations.dataValidation:
                    if validation.type == 'list':
                        # This is a dropdown validation
                        formula = validation.formula1
                        options = []
                        
                        # Extract options from formula
                        if formula.startswith('"') and formula.endswith('"'):
                            formula = formula[1:-1]
                            options = [opt.strip() for opt in formula.split(',')]
                        
                        # Get all cells this validation applies to
                        for coord in validation.sqref.ranges:
                            for row in range(coord.min_row, coord.max_row + 1):
                                for col in range(coord.min_col, coord.max_col + 1):
                                    cell_key = f"row_{row-1}_col_{col-1}"  # Adjust for 0-based indexing
                                    validation_data[cell_key] = {
                                        'options': options,
                                        'cell': f"{chr(64 + col)}{row}"  # Convert to A1 notation
                                    }
                                    print(f"Found validation at {chr(64 + col)}{row}: {options}")
            
            # Find industry and sub-industry dropdown widgets
            industry_dropdown = None
            sub_industry_dropdown = None
            
            # Track the widget indices to make sure we match them properly
            industry_index = -1
            sub_industry_index = -1
            
            # Find dropdowns for Industry and Sub-Industry by going through all widgets
            for key, widget in self.data_fields.items():
                if not key.startswith(sheet_name):
                    continue  # Skip widgets from other sheets
                
                if isinstance(widget, QComboBox):
                    # Extract current section and field index
                    parts = key.split('_')
                    if len(parts) >= 3:
                        try:
                            field_index = int(parts[-1])
                            # Assuming industry comes before sub-industry
                            if field_index > industry_index:
                                # First dropdown after index is industry
                                if industry_dropdown is None:
                                    industry_dropdown = widget
                                    industry_index = field_index
                                    print(f"Found Industry dropdown with key {key}")
                                # Second dropdown after index is sub-industry    
                                elif sub_industry_dropdown is None:
                                    sub_industry_dropdown = widget
                                    sub_industry_index = field_index
                                    print(f"Found Sub-Industry dropdown with key {key}")
                        except ValueError:
                            # Not an integer index
                            pass
            
            # Track changes
            changes_made = 0
            changes_log = []
            
            # Improved approach: Create a direct mapping between Excel cells and widgets
            # This will help us match the correct widget to each cell
            cell_to_widget_map = {}
            
            # First, map all left column fields (Column A/B pairs)
            for row_idx in range(len(df)):
                # Skip empty rows
                if pd.isna(df.iloc[row_idx]).all():
                    continue
                    
                first_col = df.iloc[row_idx, 0] if not pd.isna(df.iloc[row_idx, 0]) else ""
                if not isinstance(first_col, str):
                    try:
                        first_col = str(first_col)
                    except:
                        continue
                
                # Only process rows that start with field identifiers
                if not (first_col.startswith('f_') or 
                    first_col.startswith('fd_') or 
                    first_col.startswith('fm_') or
                    first_col.startswith('ft_')):
                    continue
                
                # Get field properties
                field_type = first_col[:3] if not first_col.startswith('f_') else first_col[:2]
                field_name = first_col[3:] if not first_col.startswith('f_') else first_col[2:]
                
                # Clean up the display name
                display_name = field_name.strip()
                if display_name and display_name[0].isdigit():
                    for i, char in enumerate(display_name):
                        if not char.isdigit():
                            display_name = display_name[i:]
                            break
                
                # Get the section for this row
                section_name = self._get_section_for_row(df, row_idx)
                
                # Find the matching widget based on field type
                if field_type == 'f_':
                    # Regular input field
                    field_id = first_col  # Simpan field identifier asli (dengan angka)
                    original_field_name = field_name.strip()  # Simpan nama field asli (dengan angka)
                    
                    # Buat key khusus untuk pencocokan widget
                    widget_key_prefix = f"{sheet_name}_{section_name}_"
                    possible_keys = []
                    
                    # Coba berbagai kemungkinan kunci
                    for i in range(100):
                        possible_keys.append(f"{widget_key_prefix}{i}")
                        possible_keys.append(f"{widget_key_prefix}{original_field_name}_{i}")
                        possible_keys.append(f"{widget_key_prefix}{display_name}_{i}")
                    
                    # Tambahkan kunci berdasarkan nomor baris
                    possible_keys.append(f"{widget_key_prefix}{row_idx}")
                    
                    # Coba semua kunci tersebut secara berurutan
                    found_widget = None
                    found_key = None
                    
                    for test_key in possible_keys:
                        if test_key in self.data_fields:
                            widget = self.data_fields[test_key]
                            if (isinstance(widget, QLineEdit) and 
                                test_key not in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info]):
                                
                                placeholder = widget.placeholderText()
                                # Prioritaskan kecocokan yang lebih spesifik
                                if placeholder:
                                    # Kecocokan paling spesifik: memuat field ID lengkap (dengan prefix angka)
                                    if field_id in placeholder:
                                        found_widget = widget
                                        found_key = test_key
                                        break
                                    # Kecocokan level kedua: memuat nama field asli (dengan angka)
                                    elif original_field_name in placeholder:
                                        found_widget = widget
                                        found_key = test_key
                                        # Jangan break, karena masih mencari kecocokan terbaik
                                    # Kecocokan level ketiga: memuat display name (tanpa angka)
                                    elif display_name in placeholder and not found_widget:
                                        found_widget = widget
                                        found_key = test_key
                                        # Jangan break, karena masih mencari kecocokan yang lebih baik
                    
                    # Jika widget ditemukan, petakan
                    if found_widget:
                        cell_key = f"row_{row_idx}_col_1"  # Column B
                        cell_to_widget_map[cell_key] = {
                            'widget': found_widget,
                            'key': found_key,
                            'type': 'text',
                            'display_name': display_name,
                            'original_field': original_field_name,
                            'field_id': field_id
                        }
                    else:
                        print(f"Warning: Could not find widget for field: {field_id}")
                                        
                elif field_type == 'fd_':
                    # Dropdown field - requires special handling
                    # For Industry Classification
                    if "Industry Classification" in display_name and industry_dropdown:
                        cell_key = f"row_{row_idx}_col_1"  # Column B
                        cell_to_widget_map[cell_key] = {
                            'widget': industry_dropdown,
                            'key': 'industry_dropdown',
                            'type': 'dropdown'
                        }
                    
                    # For Sub Industry Specification
                    elif "Sub Industry Specification" in display_name and sub_industry_dropdown:
                        cell_key = f"row_{row_idx}_col_1"  # Column B
                        cell_to_widget_map[cell_key] = {
                            'widget': sub_industry_dropdown,
                            'key': 'sub_industry_dropdown',
                            'type': 'dropdown'
                        }
                    
                    # For other dropdowns
                    else:
                        # Find the options for this dropdown
                        dropdown_key = row_idx
                        options = []
                        
                        if dropdown_key in dropdown_options:
                            options = dropdown_options[dropdown_key]['options']
                        
                        # Find a QComboBox that isn't already mapped and has matching options
                        for key, widget in self.data_fields.items():
                            if not key.startswith(sheet_name):
                                continue
                                
                            if isinstance(widget, QComboBox) and key not in cell_to_widget_map.values():
                                # Skip industry and sub-industry dropdowns
                                if widget is industry_dropdown or widget is sub_industry_dropdown:
                                    continue
                                    
                                # Check if widget options match
                                widget_options = [widget.itemText(i) for i in range(widget.count())]
                                option_match = False
                                
                                # If there are options, check for matches
                                if options:
                                    option_match = any(opt in widget_options for opt in options)
                                
                                # If no specific options or matching options found, use this widget
                                if not options or option_match or len(options) == 0:
                                    cell_key = f"row_{row_idx}_col_1"  # Column B
                                    cell_to_widget_map[cell_key] = {
                                        'widget': widget,
                                        'key': key,
                                        'type': 'dropdown'
                                    }
                                    break
                
                elif field_type == 'fm_':
                    # Multiple field (e.g., Name and Phone/Email)
                    found_base_key = None
                    
                    # Cari berdasarkan nama bidang, dan juga coba variasi lain
                    field_base = f"{sheet_name}_{section_name}_{field_name}"
                    field_key_0 = f"{field_base}_0"
                    field_key_1 = f"{field_base}_1"
                    
                    # Coba cari widget berdasarkan nama bidang lengkap terlebih dahulu
                    if field_key_0 in self.data_fields and field_key_1 in self.data_fields:
                        found_key_base = field_base
                        widget_0 = self.data_fields[field_key_0]
                        widget_1 = self.data_fields[field_key_1]
                        
                        # Map both widgets
                        cell_key_0 = f"row_{row_idx}_col_1"  # Column B
                        cell_key_1 = f"row_{row_idx}_col_2"  # Column C
                        
                        cell_to_widget_map[cell_key_0] = {
                            'widget': widget_0,
                            'key': field_key_0,
                            'type': 'text'
                        }
                        
                        cell_to_widget_map[cell_key_1] = {
                            'widget': widget_1,
                            'key': field_key_1,
                            'type': 'text'
                        }
                    else:
                        # Jika tidak ditemukan, cari berdasarkan indeks dan placeholder
                        for i in range(100):
                            test_base = f"{sheet_name}_{section_name}_{i}"
                            test_key_0 = f"{test_base}_0"
                            test_key_1 = f"{test_base}_1"
                            
                            if test_key_0 in self.data_fields and test_key_1 in self.data_fields:
                                widget_0 = self.data_fields[test_key_0]
                                widget_1 = self.data_fields[test_key_1]
                                
                                # Skip jika sudah dipetakan
                                if (test_key_0 in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info] or
                                    test_key_1 in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info]):
                                    continue
                                
                                if isinstance(widget_0, QLineEdit) and isinstance(widget_1, QLineEdit):
                                    # Periksa placeholder dengan kriteria yang lebih longgar
                                    placeholder_0 = widget_0.placeholderText()
                                    placeholder_1 = widget_1.placeholderText()
                                    
                                    # Gunakan pendekatan lebih fleksibel untuk mencocokkan placeholder
                                    # Periksa jika nama tampilan ada di salah satu placeholder
                                    # atau jika kita belum memetakan cukup banyak widget fm_
                                    if ((placeholder_0 and display_name in placeholder_0) or 
                                        (placeholder_1 and display_name in placeholder_1) or
                                        len([k for k in cell_to_widget_map.keys() if '_col_2' in k]) < 10):  # Pastikan kita memetakan cukup banyak widget fm_
                                        
                                        # Map kedua widget
                                        cell_key_0 = f"row_{row_idx}_col_1"  # Column B
                                        cell_key_1 = f"row_{row_idx}_col_2"  # Column C
                                        
                                        cell_to_widget_map[cell_key_0] = {
                                            'widget': widget_0,
                                            'key': test_key_0,
                                            'type': 'text'
                                        }
                                        
                                        cell_to_widget_map[cell_key_1] = {
                                            'widget': widget_1,
                                            'key': test_key_1,
                                            'type': 'text'
                                        }
                                        
                                        found_base_key = test_base
                                        break
                    
                    if not found_base_key:
                        print(f"Warning: Could not find widgets for multiple field: {display_name}")
                
                elif field_type == 'ft_':
                    # Table item (checkboxes and remarks)
                    # Keep track of which table items we've processed
                    processed_table_items = getattr(self, '_processed_table_items', set())
                    
                    # For each row, find the FIRST unprocessed matching widget set
                    found_base_key = None
                    
                    for i in range(100):  # Try different indexes
                        base_key = f"{sheet_name}_{section_name}_{i}"
                        client_key = f"{base_key}_client"
                        contractor_key = f"{base_key}_contractor"
                        remarks_key = f"{base_key}_remarks"
                        
                        # Check if all three widgets exist and none of them are already processed
                        if (client_key in self.data_fields and 
                            contractor_key in self.data_fields and 
                            remarks_key in self.data_fields and
                            base_key not in processed_table_items):
                            
                            client_widget = self.data_fields[client_key]
                            contractor_widget = self.data_fields[contractor_key]
                            remarks_widget = self.data_fields[remarks_key]
                            
                            if (isinstance(client_widget, QCheckBox) and 
                                isinstance(contractor_widget, QCheckBox) and 
                                isinstance(remarks_widget, QLineEdit)):
                                
                                # Map all three widgets
                                cell_key_client = f"row_{row_idx}_col_1"  # Column B
                                cell_key_contractor = f"row_{row_idx}_col_2"  # Column C
                                cell_key_remarks = f"row_{row_idx}_col_3"  # Column D
                                
                                cell_to_widget_map[cell_key_client] = {
                                    'widget': client_widget,
                                    'key': client_key,
                                    'type': 'checkbox'
                                }
                                
                                cell_to_widget_map[cell_key_contractor] = {
                                    'widget': contractor_widget,
                                    'key': contractor_key,
                                    'type': 'checkbox'
                                }
                                
                                cell_to_widget_map[cell_key_remarks] = {
                                    'widget': remarks_widget,
                                    'key': remarks_key,
                                    'type': 'text'
                                }
                                
                                found_base_key = base_key
                                processed_table_items.add(base_key)  # Mark this widget set as processed
                                break
                    
                    # Store the processed items for later use
                    self._processed_table_items = processed_table_items
            
            # Now, handle right column fields (Columns C/D, E/F, etc.)
            for row_idx in range(len(df)):
                # Skip empty rows
                if pd.isna(df.iloc[row_idx]).all():
                    continue
                
                # Check columns starting from the third column (index 2 = column C)
                # This is where right fields start
                for col_idx in range(2, len(df.columns)):
                    if col_idx >= len(df.columns):
                        break
                        
                    # Get the value in this cell
                    cell_value = ""
                    if col_idx < len(df.iloc[row_idx]) and not pd.isna(df.iloc[row_idx, col_idx]):
                        cell_value = df.iloc[row_idx, col_idx]
                        
                        # Convert to string if needed
                        if not isinstance(cell_value, str):
                            try:
                                cell_value = str(cell_value)
                            except:
                                continue
                    
                    # Only process cells with field identifiers
                    if not (cell_value.startswith('f_') or 
                        cell_value.startswith('fd_') or 
                        cell_value.startswith('fm_')):
                        continue
                    
                    # Skip header cells
                    if cell_value.startswith('fh_'):
                        continue
                    
                    # Get field properties
                    field_type = cell_value[:3] if not cell_value.startswith('f_') else cell_value[:2]
                    field_name = cell_value[3:] if not cell_value.startswith('f_') else cell_value[2:]
                    
                    # Clean up the display name
                    display_name = field_name.strip()
                    if display_name and display_name[0].isdigit():
                        for i, char in enumerate(display_name):
                            if not char.isdigit():
                                display_name = display_name[i:]
                                break
                    
                    # Get the section for this row (either right section or default)
                    section_name = self._get_right_section_for_row(df, row_idx, col_idx)
                    if not section_name:
                        section_name = self._get_section_for_row(df, row_idx)
                    
                    # Find the matching widget based on field type
                    if field_type == 'f_':
                        # Regular input field
                        # Look for a QLineEdit with this name in the placeholder
                        for key, widget in self.data_fields.items():
                            if not key.startswith(sheet_name):
                                continue
                                
                            if isinstance(widget, QLineEdit) and key not in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info]:
                                placeholder = widget.placeholderText()
                                if placeholder and display_name in placeholder:
                                    target_col = col_idx + 1  # Next column
                                    cell_key = f"row_{row_idx}_col_{target_col}"
                                    cell_to_widget_map[cell_key] = {
                                        'widget': widget,
                                        'key': key,
                                        'type': 'text'
                                    }
                                    break
                    
                    elif field_type == 'fd_':
                        # Dropdown field
                        # Find the options for this dropdown
                        dropdown_key = f"{row_idx}_{col_idx}"
                        options = []
                        
                        if dropdown_key in dropdown_options:
                            options = dropdown_options[dropdown_key]['options']
                        else:
                            # Try to get options from validation data
                            validation_key = f"row_{row_idx}_col_{col_idx+1}"  # Next column
                            if validation_key in validation_data:
                                options = validation_data[validation_key]['options']
                        
                        # Find a QComboBox that isn't already mapped
                        for key, widget in self.data_fields.items():
                            if not key.startswith(sheet_name):
                                continue
                                
                            if isinstance(widget, QComboBox) and key not in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info]:
                                # Skip industry and sub-industry dropdowns
                                if widget is industry_dropdown or widget is sub_industry_dropdown:
                                    continue
                                    
                                # Check if widget options match
                                widget_options = [widget.itemText(i) for i in range(widget.count())]
                                option_match = False
                                
                                # If there are options, check for matches
                                if options:
                                    option_match = any(opt in widget_options for opt in options)
                                
                                # If no specific options or matching options found, use this widget
                                if not options or option_match or len(options) == 0:
                                    target_col = col_idx + 1  # Next column
                                    cell_key = f"row_{row_idx}_col_{target_col}"
                                    cell_to_widget_map[cell_key] = {
                                        'widget': widget,
                                        'key': key,
                                        'type': 'dropdown',
                                        'options': options
                                    }
                                    break
                    
                    elif field_type == 'fm_':
                        # Multiple field (e.g., Name and Phone/Email)
                        found_base_key = None
                        
                        # Look for a pair of QLineEdit widgets for this field
                        for i in range(100):
                            base_key = f"{sheet_name}_{section_name}_{i}"
                            key_0 = f"{base_key}_0"
                            key_1 = f"{base_key}_1"
                            
                            if key_0 in self.data_fields and key_1 in self.data_fields:
                                widget_0 = self.data_fields[key_0]
                                widget_1 = self.data_fields[key_1]
                                
                                # Skip if already mapped
                                if (key_0 in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info] or
                                    key_1 in [info.get('key') for info in cell_to_widget_map.values() if 'key' in info]):
                                    continue
                                
                                if isinstance(widget_0, QLineEdit) and isinstance(widget_1, QLineEdit):
                                    # Check if placeholders contain the display name
                                    placeholder_0 = widget_0.placeholderText()
                                    placeholder_1 = widget_1.placeholderText()
                                    
                                    if ((placeholder_0 and display_name in placeholder_0) or 
                                        (placeholder_1 and display_name in placeholder_1)):
                                            
                                        # Map both widgets
                                        target_col_0 = col_idx + 1  # Next column
                                        target_col_1 = col_idx + 2  # Next column + 1
                                        
                                        cell_key_0 = f"row_{row_idx}_col_{target_col_0}"
                                        cell_key_1 = f"row_{row_idx}_col_{target_col_1}"
                                        
                                        cell_to_widget_map[cell_key_0] = {
                                            'widget': widget_0,
                                            'key': key_0,
                                            'type': 'text'
                                        }
                                        
                                        cell_to_widget_map[cell_key_1] = {
                                            'widget': widget_1,
                                            'key': key_1,
                                            'type': 'text'
                                        }
                                        
                                        found_base_key = base_key
                                        break
            
            # Create a reverse mapping for debugging
            widget_to_cell_map = {}
            for cell_key, info in cell_to_widget_map.items():
                if 'key' in info:
                    widget_key = info['key']
                    widget_to_cell_map[widget_key] = cell_key
            
            print(f"Mapped {len(cell_to_widget_map)} Excel cells to widgets")
            
            # Now use the cell-to-widget mapping to update cells
            for cell_key, info in cell_to_widget_map.items():
                if 'widget' not in info:
                    continue
                    
                widget = info['widget']
                widget_type = info.get('type', 'text')
                
                # Parse row and column from cell key
                parts = cell_key.split('_')
                row_idx = int(parts[1])
                col_idx = int(parts[3])
                
                # Convert to Excel coordinates (1-based)
                excel_row = row_idx + 1
                excel_col = col_idx + 1
                col_letter = chr(64 + excel_col)  # A=1, B=2, etc.
                
                # Get cell value based on widget type
                if widget_type == 'text' and isinstance(widget, QLineEdit):
                    value = widget.text()
                elif widget_type == 'dropdown' and isinstance(widget, QComboBox):
                    value = widget.currentText()
                elif widget_type == 'checkbox' and isinstance(widget, QCheckBox):
                    # Use 'ü' for checked boxes, empty string for unchecked
                    value = 'ü' if widget.isChecked() else ''
                    
                    # Get the cell to update
                    target_cell = sheet.cell(row=excel_row, column=excel_col)
                    old_value = target_cell.value
                    target_cell.value = value
                    
                    # Set the font to Wingdings for checkbox cells
                    from openpyxl.styles import Font
                    target_cell.font = Font(name='Wingdings', size=11)
                else:
                    continue  # Skip invalid combinations
                
                # Get display name for logging
                display_name = ""
                
                # For left columns (A), get from column A
                if col_idx == 1:  # Column B
                    cell_a = df.iloc[row_idx, 0] if not pd.isna(df.iloc[row_idx, 0]) else ""
                    if isinstance(cell_a, str):
                        if cell_a.startswith('f_'):
                            display_name = cell_a[2:].strip()
                        elif cell_a.startswith('fd_') or cell_a.startswith('fm_') or cell_a.startswith('ft_'):
                            display_name = cell_a[3:].strip()
                # For right columns (C and beyond), get from the actual column
                elif col_idx >= 2:  # Column C or beyond
                    identifier_col = col_idx - 1
                    if identifier_col < len(df.columns) and not pd.isna(df.iloc[row_idx, identifier_col]):
                        cell_value = df.iloc[row_idx, identifier_col]
                        if isinstance(cell_value, str):
                            if cell_value.startswith('f_'):
                                display_name = cell_value[2:].strip()
                            elif cell_value.startswith('fd_') or cell_value.startswith('fm_'):
                                display_name = cell_value[3:].strip()
                
                # Clean up display name
                if display_name and display_name[0].isdigit():
                    for i, char in enumerate(display_name):
                        if not char.isdigit():
                            display_name = display_name[i:]
                            break
                
                # Add "Right" prefix for right columns
                if col_idx >= 3:  # Column D and beyond
                    display_name = f"Right {display_name}"
                
                # Update the cell
                target_cell = sheet.cell(row=excel_row, column=excel_col)
                old_value = target_cell.value
                target_cell.value = value
                
                changes_made += 1
                changes_log.append(f"Updated cell {col_letter}{excel_row} ({display_name}): {old_value} -> {value}")
                print(f"Updated cell {col_letter}{excel_row} ({display_name}): {old_value} -> {value}")
            
            # Create special mapping for tdi_ fields with placeholders
            placeholder_widgets = {}

            # First, scan the Excel file to find cells with tdi_ and $P placeholders
            for row_idx, row in df.iterrows():
                if pd.isna(row).all():
                    continue
                    
                # Check each column in the row
                for col_idx in range(len(row)):
                    cell_value = row.iloc[col_idx] if col_idx < len(row) and not pd.isna(row.iloc[col_idx]) else ""
                    
                    # Convert to string if needed
                    if not isinstance(cell_value, str):
                        try:
                            cell_value = str(cell_value)
                        except:
                            continue
                            
                    # Check if it's a tdi_ cell with placeholder
                    if cell_value.startswith('tdi_') and ('$P1$' in cell_value or '$P2$' in cell_value or 
                                                        '$P3$' in cell_value or '$P4$' in cell_value):
                        # Store this cell location for later processing
                        cell_key = f"row_{row_idx}_col_{col_idx}"
                        placeholder_widgets[cell_key] = {
                            'cell_value': cell_value,
                            'row': row_idx,
                            'col': col_idx,
                            'placeholders': []
                        }
                        
                        # Find all placeholders in this cell
                        import re
                        placeholders = re.findall(r'\$P\d+\$', cell_value)
                        placeholder_widgets[cell_key]['placeholders'] = placeholders

            # Now match up the input fields with the placeholders
            for cell_key, info in placeholder_widgets.items():
                row_idx = info['row']
                col_idx = info['col']
                cell_text = info['cell_value']
                header_text = ""
                
                # Try to find the header for this row
                # First, look for a header in the same row (for th_ fields)
                if col_idx > 0:  # If not in first column
                    header_cell = df.iloc[row_idx, 0] if not pd.isna(df.iloc[row_idx, 0]) else ""
                    if isinstance(header_cell, str) and (header_cell.startswith('th_') or header_cell.startswith('thr_')):
                        if header_cell.startswith('th_'):
                            header_text = header_cell[3:].strip()  # Remove 'th_' prefix
                        else:
                            header_text = header_cell[4:].strip()  # Remove 'thr_' prefix
                
                # Now find matching input fields for each placeholder
                placeholders = info['placeholders']
                
                # Look through all input fields to find matches
                for key, widget in self.data_fields.items():
                    # For each placeholder in this cell
                    for placeholder in placeholders:
                        # The key for input fields related to these placeholders follows this pattern
                        placeholder_key = f"tdi_{placeholder}_{header_text}"
                        placeholder_key_with_row = f"tdi_{placeholder}_{header_text}_{row_idx}"
                        
                        # Check both possible key patterns
                        if (key == placeholder_key or key == placeholder_key_with_row) and isinstance(widget, QLineEdit):
                            # Found a match!
                            value = widget.text()
                            
                            # Add to our cell-to-widget mapping for the placeholder
                            # We'll replace the placeholder in the cell value later
                            if 'placeholder_inputs' not in placeholder_widgets[cell_key]:
                                placeholder_widgets[cell_key]['placeholder_inputs'] = {}
                                
                            placeholder_widgets[cell_key]['placeholder_inputs'][placeholder] = {
                                'widget': widget,
                                'value': value,
                                'key': key
                            }
                            
                            print(f"Matched placeholder {placeholder} in row {row_idx}, col {col_idx} to widget {key} with value: {value}")
                            break  # Found a match for this placeholder

            # Now process all cells with placeholders
            # We'll add this after the regular cell processing

            # Store replacements for later processing
            placeholder_replacements = []

            # After regular cell processing is complete, handle placeholder cells
            for cell_key, info in placeholder_widgets.items():
                row_idx = info['row']
                col_idx = info['col']
                cell_text = info['cell_value']
                
                # Skip if no input widgets were found
                if 'placeholder_inputs' not in info:
                    continue
                    
                # Get Excel coordinates
                excel_row = row_idx + 1
                excel_col = col_idx + 1
                col_letter = chr(64 + excel_col)  # Convert to A1 notation
                
                # Make a copy of the original text to modify
                new_text = cell_text
                
                # Replace each placeholder with its input value
                for placeholder, input_info in info['placeholder_inputs'].items():
                    value = input_info['value']
                    new_text = new_text.replace(placeholder, value)
                
                # Remove the 'tdi_' prefix from the beginning for Excel
                if new_text.startswith('tdi_'):
                    new_text = new_text[4:]
                
                # Store for later processing to avoid modifying the sheet while iterating
                placeholder_replacements.append({
                    'row': excel_row,
                    'col': excel_col, 
                    'value': new_text,
                    'original': sheet.cell(row=excel_row, column=excel_col).value
                })

            # Now apply all placeholder replacements
            for replacement in placeholder_replacements:
                target_cell = sheet.cell(row=replacement['row'], column=replacement['col'])
                old_value = replacement['original']
                new_value = replacement['value']
                
                # Update the cell
                target_cell.value = new_value
                
                # Log the change
                col_letter = chr(64 + replacement['col'])
                cell_address = f"{col_letter}{replacement['row']}"
                changes_made += 1
                changes_log.append(f"Updated placeholder cell {cell_address}: {old_value} -> {new_value}")
                print(f"Updated placeholder cell {cell_address}: {old_value} -> {new_value}")
                        
            # Save the workbook
            print("Attempting to save workbook...")
            try:
                wb.save(self.excel_path)
                print(f"Workbook saved successfully to: {self.excel_path}")
                
                # Show success message
                QMessageBox.information(
                    self, 
                    "Save Successful", 
                    f"Data successfully saved to {sheet_name} with {changes_made} changes."
                )
                
                # Write detailed log to file for debugging
                log_path = os.path.join(os.path.dirname(self.excel_path), "save_changes_log.txt")
                with open(log_path, 'w') as log_file:
                    log_file.write(f"Save operation at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                    log_file.write(f"Excel file: {self.excel_path}\n")
                    log_file.write(f"Sheet: {sheet_name}\n")
                    log_file.write(f"Total changes: {changes_made}\n\n")
                    log_file.write("Detailed changes:\n")
                    for change in changes_log:
                        log_file.write(f"- {change}\n")
                
                # Important: Force application to process events before reloading
                QApplication.processEvents()
                
                # Reload the data to reflect changes - just reload the current tab
                current_tab_index = self.tab_widget.currentIndex()
                current_tab_text = self.tab_widget.tabText(current_tab_index)
                
                print(f"Reloading data for tab: {current_tab_text}")
                
                # Only reload the data for the current sheet to avoid freezing
                if sheet_name in self.sheet_tabs:
                    try:
                        # Get the tab's scroll area and its widget
                        scroll_area = self.tab_widget.widget(current_tab_index)
                        if isinstance(scroll_area, QScrollArea):
                            # Save scroll position
                            scroll_pos = scroll_area.verticalScrollBar().value()
                            
                            # Clear and reload just this sheet
                            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
                            
                            sheet_widget = scroll_area.widget()
                            if sheet_widget:
                                sheet_layout = sheet_widget.layout()
                                if sheet_layout:
                                    # Clear existing layout
                                    while sheet_layout.count():
                                        item = sheet_layout.takeAt(0)
                                        widget = item.widget()
                                        if widget:
                                            widget.deleteLater()
                                    
                                    # Recreate sheet content
                                    self.process_sheet_data(df, sheet_name, sheet_layout)
                                    
                                    # Restore scroll position
                                    QApplication.processEvents()
                                    scroll_area.verticalScrollBar().setValue(scroll_pos)
                                    
                                    print(f"Successfully reloaded tab: {current_tab_text}")
                    except Exception as e:
                        print(f"Error reloading tab {current_tab_text}: {str(e)}")
                        import traceback
                        traceback.print_exc()
                
            except PermissionError:
                QMessageBox.critical(
                    self, 
                    "Permission Error", 
                    "Could not save the file. Make sure it is not open in Excel or another program."
                )
            except Exception as e:
                QMessageBox.critical(
                    self, 
                    "Error Saving File", 
                    f"Could not save the file: {str(e)}"
                )
                print(f"Error saving workbook: {str(e)}")
            
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error Saving Data", 
                f"An error occurred while processing data: {str(e)}"
            )
            print(f"Error in save_sheet_data: {str(e)}")
            import traceback
            traceback.print_exc()
    
    
    def _get_right_section_for_row(self, df, row_idx, col_idx):
        """Helper untuk mendapatkan section dari kolom kanan untuk baris tertentu"""
        # Start from the top and look for 'sub_' in the specified column
        current_right_section = None
        
        for i in range(row_idx):
            if i >= len(df):
                break
                
            row = df.iloc[i]
            
            # Skip if column doesn't exist
            if col_idx >= len(row):
                continue
                
            col_value = row.iloc[col_idx] if not pd.isna(row.iloc[col_idx]) else ""
            
            if isinstance(col_value, str) and col_value.startswith('sub_'):
                current_right_section = col_value[4:].strip()
        
        return current_right_section
                        
            
        
    def _get_section_for_row(self, df, row_idx):
        """Helper untuk mendapatkan section dari baris tertentu"""
        current_section = "Default"
        
        # Cari section terdekat sebelum row_idx
        for i in range(row_idx):
            if i >= len(df):
                break
                
            row = df.iloc[i]
            first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
            
            if isinstance(first_col, str) and first_col.startswith('sub_'):
                current_section = first_col[4:].strip()
        
        return current_section
        
    def create_data_table(self, df, layout):
        """Create a table view for DATA sheets"""
        # Create a table widget
        table = QTableWidget()
        
        # Set row and column count
        table.setRowCount(len(df))
        table.setColumnCount(len(df.columns))
        
        # Set headers - convert all column headers to strings
        headers = [str(col) for col in df.columns.tolist()]
        table.setHorizontalHeaderLabels(headers)
        
        # Fill data - convert all values to strings to avoid type issues
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                if pd.isna(value):
                    item = QTableWidgetItem("")
                else:
                    item = QTableWidgetItem(str(value))
                table.setItem(row_idx, col_idx, item)
        
        # Style the table
        table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background-color: white;
                gridline-color: #E0E0E0;
            }
            QHeaderView::section {
                background-color: #F5F5F5;
                padding: 5px;
                border: 1px solid #E0E0E0;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #D6EAF8;
            }
        """)
        
        # Resize columns to content
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        
        # Make it read-only for DATA sheets
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Add to layout
        layout.addWidget(table)
        
        # Add export button
        export_btn = QPushButton("Export to CSV")
        export_btn.setFont(QFont("Segoe UI", 11))
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {SECONDARY_COLOR};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                margin-top: 10px;
            }}
            QPushButton:hover {{
                background-color: #2980B9;
            }}
        """)
        export_btn.clicked.connect(lambda: self.export_data_table(df))
        
        export_layout = QHBoxLayout()
        export_layout.addStretch()
        export_layout.addWidget(export_btn)
        
        layout.addLayout(export_layout)
        
    def export_data_table(self, df):
        """Export DATA table to CSV"""
        try:
            # Ask for save location
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save CSV File", "", "CSV Files (*.csv);;All Files (*)"
            )
            
            if file_path:
                # Ensure it has .csv extension
                if not file_path.endswith('.csv'):
                    file_path += '.csv'
                
                # Save the dataframe
                df.to_csv(file_path, index=False)
                
                # Show success message
                QMessageBox.information(self, "Success", f"Data exported successfully to {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")
            print(f"Error exporting data: {str(e)}")