import os
import sys
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFrame, QGridLayout, QSpacerItem,
                             QSizePolicy, QScrollArea, QApplication, QMenu, QAction,
                             QTabWidget, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QFileDialog, QDateEdit)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QCursor
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QPoint, QDate

# Import local modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_NAME, SECONDARY_COLOR, PRIMARY_COLOR, BG_COLOR, DEPARTMENTS

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

            # Use sheet_names directly to preserve the order from Excel file
            # Instead of: sorted_sheets = sorted(sheet_names)

            if len(sheet_names) == 0:
                self.loading_label.setText("No sheets found in SET_BDU.xlsx.")
                return

            # Hide loading label as we have data
            self.loading_label.setVisible(False)

            # Create a tab for each sheet
            for sheet_name in sheet_names:
                # Get display name (remove DIP_ or DATA_ prefix)
                display_name = sheet_name
                if sheet_name.startswith("DIP_"):
                    display_name = sheet_name[4:]  # Remove "DIP_"
                elif sheet_name.startswith("DATA_"):
                    display_name = sheet_name[5:]  # Remove "DATA_"

                try:
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)

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

        except Exception as e:
            self.loading_label.setText(f"Error loading data: {str(e)}")
            self.loading_label.setStyleSheet("color: #E74C3C; margin: 20px;")
            self.loading_label.setVisible(True)
            print(f"Error loading Excel data: {str(e)}")
    
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
    
    def process_sheet_data(self, df, sheet_name, layout):
        """Process the data from a sheet and create UI elements"""
        # Check if the sheet is a DATA sheet (just display as a table)
        if sheet_name.startswith("DATA_"):
            self.create_data_table(df, layout)
            return

        # For DIP sheets or other sheets, process as forms
        # Initialize variables
        current_section = None
        section_layout = None
        current_header_labels = []  # For storing column headers from ch_
        has_column_headers = False

        # Field identification
        field_count = 0

        # Check if the dataframe is empty
        if df.empty:
            empty_label = QLabel("No data found in this sheet.")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("color: #666; margin: 20px;")
            layout.addWidget(empty_label)
            return

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
                
            # Check if it's a section header (sub_)
            if isinstance(first_col, str) and first_col.startswith('sub_'):
                # Create a new section
                section_title = first_col[4:].strip()  # Remove 'sub_' prefix

                # Create section frame - removed border
                section_frame = QWidget()  # Changed from QFrame to QWidget
                section_frame.setStyleSheet("""
                    background-color: white;
                """)

                section_layout = QVBoxLayout(section_frame)
                section_layout.setContentsMargins(15, 15, 15, 15)
                section_layout.setSpacing(15)

                # Add section title
                title_label = QLabel(section_title)
                title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
                title_label.setStyleSheet(f"color: {PRIMARY_COLOR}; background-color: transparent;")

                section_layout.addWidget(title_label)

                # Add section to main layout
                layout.addWidget(section_frame)
                layout.addSpacing(20)

                current_section = section_title
                field_count = 0
                # Reset column headers when entering a new section
                current_header_labels = []
                has_column_headers = False

                continue
            
            # Di dalam method process_sheet_data, perbarui bagian yang menangani field header (fh_) dan column header (ch_)

            # Check if it's a field header (fh_)
            if first_col.startswith('fh_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()
                    section_frame.setStyleSheet("""
                        background-color: white;
                    """)

                    section_layout = QVBoxLayout(section_frame)
                    section_layout.setContentsMargins(15, 15, 15, 15)
                    section_layout.setSpacing(15)

                    # Add to main layout
                    layout.addWidget(section_frame)
                    current_section = "Default"

                # Create field header
                field_header = first_col[3:].strip()  # Remove 'fh_' prefix
                
                # Periksa apakah ada column header di samping field header
                has_column_headers_in_row = False
                column_headers_in_row = []
                
                for col_idx in range(1, df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip()
                        if col_value.startswith('ch_'):
                            has_column_headers_in_row = True
                            header_text = col_value[3:].strip()  # Remove 'ch_' prefix
                            column_headers_in_row.append(header_text)
                
                if has_column_headers_in_row:
                    # Buat header row dengan multiple column
                    header_widget = QWidget()
                    header_widget.setStyleSheet("background-color: transparent;")
                    header_layout = QHBoxLayout(header_widget)
                    header_layout.setContentsMargins(5, 5, 5, 5)
                    
                    # Field header label
                    header_label = QLabel(field_header)
                    header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                    header_label.setStyleSheet("color: #555; margin-top: 5px;")
                    header_label.setMinimumWidth(350)
                    
                    header_layout.addWidget(header_label)
                    
                    # Add column header labels
                    for col_header in column_headers_in_row:
                        col_header_label = QLabel(col_header)
                        col_header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                        col_header_label.setStyleSheet("color: #555; margin-top: 5px;")
                        
                        # Set size policy untuk memastikan lebar yang konsisten
                        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                        col_header_label.setSizePolicy(size_policy)
                        
                        header_layout.addWidget(col_header_label)
                    
                    section_layout.addWidget(header_widget)
                    
                    # Simpan column headers untuk field multiple berikutnya
                    current_header_labels = column_headers_in_row
                    has_column_headers = True
                else:
                    # Tampilan header biasa jika tidak ada column header
                    header_label = QLabel(field_header)
                    header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                    header_label.setStyleSheet("color: #555; margin-top: 5px;")
                    
                    section_layout.addWidget(header_label)
                    
                    # Reset column headers jika tidak ada di row ini
                    current_header_labels = []
                    has_column_headers = False
                
                continue

            # Check if it's a column header row (first cell starts with ch_)
            elif first_col.startswith('ch_'):
                has_column_headers = True
                current_header_labels = []
                
                # Process header row and collect all ch_ columns
                for col_idx in range(df.shape[1]):
                    if col_idx < len(row) and not pd.isna(row[col_idx]):
                        col_value = str(row[col_idx]).strip()
                        if col_value.startswith('ch_'):
                            header_text = col_value[3:].strip()  # Remove 'ch_' prefix
                            current_header_labels.append(header_text)
                
                # Jika ini adalah row column header yang berdiri sendiri (tidak di samping fh_)
                # dan kita memiliki section layout, buat header row
                if section_layout is not None and len(current_header_labels) > 0:
                    header_widget = QWidget()
                    header_widget.setStyleSheet("background-color: transparent;")
                    header_layout = QHBoxLayout(header_widget)
                    header_layout.setContentsMargins(5, 5, 5, 5)
                    
                    # Tambahkan spacer untuk menyelaraskan dengan field label
                    spacer_label = QLabel("")
                    spacer_label.setMinimumWidth(350)
                    header_layout.addWidget(spacer_label)
                    
                    # Add column header labels
                    for col_header in current_header_labels:
                        col_header_label = QLabel(col_header)
                        col_header_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
                        col_header_label.setStyleSheet("color: #555; margin-top: 5px;")
                        
                        # Set size policy untuk memastikan lebar yang konsisten
                        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                        col_header_label.setSizePolicy(size_policy)
                        
                        header_layout.addWidget(col_header_label)
                    
                    section_layout.addWidget(header_widget)
                
                continue
            
            # Check if it's a field (f_)
            if first_col.startswith('f_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()
                    section_frame.setStyleSheet("""
                        background-color: white;
                    """)

                    section_layout = QVBoxLayout(section_frame)
                    section_layout.setContentsMargins(15, 15, 15, 15)
                    section_layout.setSpacing(15)

                    # Add to main layout
                    layout.addWidget(section_frame)
                    current_section = "Default"

                field_name = first_col[2:].strip()  # Remove 'f_' prefix
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field row with transparent background
                field_widget = QWidget()
                field_widget.setStyleSheet("background-color: transparent;")
                field_layout = QHBoxLayout(field_widget)
                field_layout.setContentsMargins(5, 5, 5, 5)

                # Field label - completely transparent with no background
                label = QLabel(field_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(350)

                field_layout.addWidget(label)

                # Process field type based on the second column
                field_type = "text"  # Default type
                options = []
                default_value = ""

                # Check for second column (type or options)
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    second_col = str(row.iloc[1]).strip()

                    # Check if it's a dropdown type
                    if "dropdown" in second_col.lower():
                        field_type = "dropdown"

                        # Extract options if they exist
                        if len(row) > 2 and not pd.isna(row.iloc[2]):
                            options_str = str(row.iloc[2]).strip()
                            options = [opt.strip() for opt in options_str.split(',')]

                    # Check if it's a date type
                    elif "date" in second_col.lower():
                        field_type = "date"

                    # Check if it's a number type
                    elif "number" in second_col.lower() or "numeric" in second_col.lower():
                        field_type = "number"

                # Create input field based on type
                input_field = None

                if field_type == "dropdown":
                    input_field = QComboBox()
                    input_field.addItems(options)
                    input_field.setFont(QFont("Segoe UI", 11))
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
                elif field_type == "date":
                    input_field = QDateEdit()
                    input_field.setFont(QFont("Segoe UI", 11))
                    input_field.setCalendarPopup(True)
                    input_field.setDate(QDate.currentDate())
                    input_field.setStyleSheet("""
                        QDateEdit {
                            padding: 5px;
                            border: 1px solid #ccc;
                            border-radius: 4px;
                            background-color: white;
                            min-height: 28px;
                        }
                        QDateEdit:hover {
                            border: 1px solid #3498DB;
                        }
                    """)
                elif field_type == "number":
                    input_field = QLineEdit()
                    input_field.setFont(QFont("Segoe UI", 11))
                    input_field.setPlaceholderText(f"Enter {field_name}")
                    # Only allow numbers and decimal point
                    input_field.setValidator(QtGui.QDoubleValidator())
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
                else:
                    input_field = QLineEdit()
                    input_field.setFont(QFont("Segoe UI", 11))
                    input_field.setPlaceholderText(f"Enter {field_name}")
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
                    default_value = str(row.iloc[1]).strip()
                    if field_type == "dropdown" and default_value in options:
                        input_field.setCurrentText(default_value)
                    elif field_type == "date":
                        try:
                            date_parts = default_value.split("-")
                            if len(date_parts) == 3:
                                input_field.setDate(QDate(int(date_parts[0]), int(date_parts[1]), int(date_parts[2])))
                        except:
                            pass  # If date parsing fails, use current date
                    elif field_type != "dropdown":
                        input_field.setText(default_value)

                # Save the field reference for later use
                self.data_fields[field_key] = input_field

                field_layout.addWidget(input_field)
                section_layout.addWidget(field_widget)

                continue
            
            # Check if it's a field dropdown (fd_)
            if first_col.startswith('fd_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()
                    section_frame.setStyleSheet("""
                        background-color: white;
                    """)

                    section_layout = QVBoxLayout(section_frame)
                    section_layout.setContentsMargins(15, 15, 15, 15)
                    section_layout.setSpacing(15)

                    # Add to main layout
                    layout.addWidget(section_frame)
                    current_section = "Default"

                field_name = first_col[3:].strip()  # Remove 'fd_' prefix
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field row with transparent background
                field_widget = QWidget()
                field_widget.setStyleSheet("background-color: transparent;")
                field_layout = QHBoxLayout(field_widget)
                field_layout.setContentsMargins(5, 5, 5, 5)

                # Field label - completely transparent with no background
                label = QLabel(field_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(350)

                field_layout.addWidget(label)

                # Create dropdown combo box
                input_field = QComboBox()
                input_field.setFont(QFont("Segoe UI", 11))
                
                # Atur properti size policy agar dropdown dapat diperluas sesuai layout
                size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                input_field.setSizePolicy(size_policy)
                
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
                
                # Coba ambil options dari data validation jika row dan column index diketahui
                options = []
                # Jika Excel menyimpan posisi sel
                row_index = index + 1  # +1 karena Excel mulai dari 1
                col_index = 2  # Asumsi kolom B untuk nilai dropdown
                cell_address = f"{chr(ord('A') + col_index-1)}{row_index}"
                
                # Coba ambil dari data validation
                validation_options = self.get_validation_values(self.excel_path, sheet_name, cell_address)
                
                if validation_options:
                    options = validation_options
                else:
                    # Fallback ke metode lama jika data validation tidak ditemukan
                    if len(row) > 1 and not pd.isna(row.iloc[1]):
                        options_str = str(row.iloc[1]).strip()
                        options = [opt.strip() for opt in options_str.split(',')]
                
                # Add options and set default if available
                input_field.addItems(options)
                if len(options) > 0:
                    input_field.setCurrentText(options[0])
                        
                # Save the field reference for later use
                self.data_fields[field_key] = input_field

                field_layout.addWidget(input_field)
                section_layout.addWidget(field_widget)

                continue
                
            # Check if it's a field multiple (fm_)
            if first_col.startswith('fm_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()
                    section_frame.setStyleSheet("""
                        background-color: white;
                    """)

                    section_layout = QVBoxLayout(section_frame)
                    section_layout.setContentsMargins(15, 15, 15, 15)
                    section_layout.setSpacing(15)

                    # Add to main layout
                    layout.addWidget(section_frame)
                    current_section = "Default"

                field_name = first_col[3:].strip()  # Remove 'fm_' prefix
                
                # Create field row with transparent background
                field_widget = QWidget()
                field_widget.setStyleSheet("background-color: transparent;")
                field_layout = QHBoxLayout(field_widget)
                field_layout.setContentsMargins(5, 5, 5, 5)

                # Field label - completely transparent with no background
                label = QLabel(field_name)
                label.setFont(QFont("Segoe UI", 11))
                label.setStyleSheet("color: #333; background-color: transparent;")
                label.setMinimumWidth(350)

                field_layout.addWidget(label)
                
                # Add input fields based on column headers or default to 2 columns
                if has_column_headers and len(current_header_labels) > 0:
                    # Create input fields for each column header
                    for i, header in enumerate(current_header_labels):
                        field_key = f"{sheet_name}_{current_section}_{field_name}_{i}"
                        field_count += 1
                        
                        # Create input field 
                        input_field = QLineEdit()
                        input_field.setFont(QFont("Segoe UI", 11))
                        
                        # Tambahkan nama field ke placeholder
                        input_field.setPlaceholderText(f"Enter {header} {field_name}")
                        
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
                        
                        # Set size policy untuk memastikan lebar yang konsisten
                        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                        input_field.setSizePolicy(size_policy)
                        
                        # Set default value if available
                        if i+1 < len(row) and not pd.isna(row.iloc[i+1]):
                            input_field.setText(str(row.iloc[i+1]).strip())
                        
                        # Save the field reference for later use
                        self.data_fields[field_key] = input_field
                        
                        field_layout.addWidget(input_field)
                else:
                    # Default to 2 columns if no column headers
                    column_names = ["Name", "Phone No/Email"]
                    for i in range(2):
                        field_key = f"{sheet_name}_{current_section}_{field_name}_{i}"
                        field_count += 1
                        
                        # Create input field 
                        input_field = QLineEdit()
                        input_field.setFont(QFont("Segoe UI", 11))
                        
                        # Tambahkan nama field ke placeholder
                        input_field.setPlaceholderText(f"Enter {column_names[i]} {field_name}")
                        
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
                        
                        # Set size policy untuk memastikan lebar yang konsisten
                        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                        input_field.setSizePolicy(size_policy)
                        
                        # Set default value if available
                        if i+1 < len(row) and not pd.isna(row.iloc[i+1]):
                            input_field.setText(str(row.iloc[i+1]).strip())
                        
                        # Save the field reference for later use
                        self.data_fields[field_key] = input_field
                        
                        field_layout.addWidget(input_field)

                section_layout.addWidget(field_widget)
                continue

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
        
    def save_sheet_data(self, sheet_name):
        """Simpan data dari suatu sheet ke file Excel"""
        try:
            # Validasi file Excel masih ada
            if not os.path.exists(self.excel_path):
                QMessageBox.critical(self, "Error", f"File Excel tidak ditemukan: {self.excel_path}")
                return
            
            # Baca file Excel asli untuk mendapatkan struktur
            excel_data = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
            
            # Buat dictionary untuk menyimpan data yang akan disimpan
            data_to_save = {}
            
            # Kumpulkan semua data dari fields
            for field_key, field_input in self.data_fields.items():
                # Periksa apakah field ini termasuk dalam sheet yang disimpan
                if not field_key.startswith(f"{sheet_name}_"):
                    continue
                
                # Parse field key untuk mendapatkan informasi
                parts = field_key.split('_')
                
                # Ambil nilai dari input field
                if isinstance(field_input, QLineEdit):
                    value = field_input.text()
                elif isinstance(field_input, QComboBox):
                    value = field_input.currentText()
                elif isinstance(field_input, QDateEdit):
                    value = field_input.date().toString("yyyy-MM-dd")
                else:
                    value = ""
                
                # Untuk field multiple (fm_), kita perlu menyimpan data khusus
                # Format: sheet_name_section_field_name_column_index
                if len(parts) > 4 and "_".join(parts[3:-1]) in field_key:
                    section = parts[1]
                    field_name = "_".join(parts[3:-1])  # Gabungkan semua bagian nama field
                    column_index = int(parts[-1])
                    
                    # Buat key untuk field multiple
                    fm_key = f"{section}_{field_name}"
                    
                    # Initialize array jika belum ada
                    if fm_key not in data_to_save:
                        data_to_save[fm_key] = [None, None]
                    
                    # Simpan value ke array sesuai column_index
                    data_to_save[fm_key][column_index] = value
                else:
                    # Untuk field biasa
                    section = parts[1]
                    field_name = parts[2]
                    data_to_save[f"{section}_{field_name}"] = value
            
            # Buat workbook baru dengan openpyxl
            from openpyxl import load_workbook
            
            # Load workbook yang ada
            workbook = load_workbook(self.excel_path)
            
            # Ambil sheet yang akan diupdate
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Proses setiap baris di sheet
                row_idx = 0
                for _, row in excel_data.iterrows():
                    row_idx += 1
                    
                    # Skip baris kosong
                    if pd.isna(row).all():
                        continue
                    
                    # Ambil kolom pertama untuk menentukan tipe
                    first_col = row.iloc[0] if not pd.isna(row.iloc[0]) else ""
                    if not isinstance(first_col, str):
                        try:
                            first_col = str(first_col)
                        except:
                            continue
                    
                    # Proses berdasarkan tipe prefix
                    if first_col.startswith('sub_') or first_col.startswith('fh_') or first_col.startswith('ch_'):
                        # Jangan ubah baris header
                        continue
                    
                    # Proses field biasa (f_)
                    if first_col.startswith('f_'):
                        field_name = first_col[2:].strip()
                        current_section = self._get_section_for_row(excel_data, row_idx)
                        
                        # Cari nilai di data_to_save
                        key = f"{current_section}_{field_name}"
                        if key in data_to_save:
                            # Update nilai di cell kedua
                            sheet.cell(row=row_idx, column=2).value = data_to_save[key]
                    
                    # Proses field dropdown (fd_)
                    elif first_col.startswith('fd_'):
                        field_name = first_col[3:].strip()
                        current_section = self._get_section_for_row(excel_data, row_idx)
                        
                        # Cari nilai di data_to_save
                        key = f"{current_section}_{field_name}"
                        if key in data_to_save:
                            # Update nilai di cell kedua
                            sheet.cell(row=row_idx, column=2).value = data_to_save[key]
                    
                    # Proses field multiple (fm_)
                    elif first_col.startswith('fm_'):
                        field_name = first_col[3:].strip()
                        current_section = self._get_section_for_row(excel_data, row_idx)
                        
                        # Cari nilai di data_to_save
                        key = f"{current_section}_{field_name}"
                        if key in data_to_save:
                            # Update nilai di cell kedua dan ketiga
                            values = data_to_save[key]
                            if values[0] is not None:
                                sheet.cell(row=row_idx, column=2).value = values[0]
                            if values[1] is not None and len(values) > 1:
                                sheet.cell(row=row_idx, column=3).value = values[1]
                
                # Simpan workbook
                workbook.save(self.excel_path)
                
                # Tampilkan pesan sukses
                QMessageBox.information(self, "Sukses", f"Data dalam sheet {sheet_name} berhasil disimpan!")
                
                # Reload data
                self.load_excel_data()
            else:
                QMessageBox.warning(self, "Peringatan", f"Sheet {sheet_name} tidak ditemukan dalam file Excel.")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal menyimpan data: {str(e)}")
            print(f"Error saving sheet data: {str(e)}")
        
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