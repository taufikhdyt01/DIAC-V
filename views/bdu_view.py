import os
import sys
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFrame, QGridLayout, QSpacerItem,
                             QSizePolicy, QScrollArea, QApplication, QMenu, QAction,
                             QTabWidget, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QFileDialog, QDateEdit, QCheckBox)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QCursor
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QPoint, QDate

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
        
        # Isi dengan nilai-nilai yang sesuai berdasarkan pilihan pada dropdown utama
        if parent_value in INDUSTRY_SUBTYPE_MAPPING:
            child_dropdown.addItems(INDUSTRY_SUBTYPE_MAPPING[parent_value])
        else:
            # Jika tidak ada mapping, tambahkan placeholder atau biarkan kosong
            child_dropdown.addItem("-- Select Sub Industry --")
            
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
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field label
                label = QLabel(field_name)
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
                input_field.setPlaceholderText(f"Enter {field_name}")
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
                field_key = f"{sheet_name}_{current_section}_{field_count}"
                field_count += 1

                # Create field label
                label = QLabel(field_name)
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
                field_name = first_col[3:].strip()  # Remove 'fm_' prefix
                field_key_base = f"{sheet_name}_{current_section}_{field_name}"
                
                # Create field label
                label = QLabel(field_name)
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
                    input_field.setPlaceholderText(f"Enter {field_name} {header}")
                    
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
                task_label.setMinimumWidth(300)  # Wider for task text
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
                if len(parts) > 4 and parts[-1].isdigit():
                    section = parts[1]
                    
                    # Handle more complex field names that might contain underscores
                    if len(parts) > 5:
                        # Join all parts between section and column index
                        field_name = "_".join(parts[2:-1])
                    else:
                        field_name = parts[2]
                        
                    column_index = int(parts[-1])
                    
                    # Buat key untuk field multiple
                    fm_key = f"{section}_{field_name}"
                    
                    # Initialize array jika belum ada - with enough slots for all columns
                    max_columns = 10  # Allocate enough space for multiple columns
                    if fm_key not in data_to_save:
                        data_to_save[fm_key] = [None] * max_columns
                    
                    # Simpan value ke array sesuai column_index
                    # Ensure we have enough space in the array
                    if column_index >= len(data_to_save[fm_key]):
                        # Extend the array if needed
                        data_to_save[fm_key].extend([None] * (column_index - len(data_to_save[fm_key]) + 1))
                    
                    data_to_save[fm_key][column_index] = value
                else:
                    # Untuk field biasa
                    if len(parts) >= 3:
                        section = parts[1]
                        
                        # Handle field names that might contain underscores
                        if len(parts) > 3:
                            # Join all parts after section
                            field_name = "_".join(parts[2:])
                        else:
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
                            continue  # Skip if can't convert to string
                    
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
                        
                        # Also check if there are fields in columns C and beyond
                        for col_idx in range(2, len(row)):
                            col_value = row.iloc[col_idx] if not pd.isna(row.iloc[col_idx]) else ""
                            
                            # Check for any field type (f_, fd_, fh_) in right columns
                            if isinstance(col_value, str) and (col_value.startswith('f_') or 
                                col_value.startswith('fd_') or 
                                col_value.startswith('fh_')):
                                
                                # This is a field in a right column
                                # Extract the prefix and field name
                                if col_value.startswith('f_'):
                                    prefix = col_value[:2]
                                    right_field_name = col_value[2:].strip()
                                else:
                                    prefix = col_value[:3]
                                    right_field_name = col_value[3:].strip()
                                
                                # Skip headers as we don't save them
                                if prefix == 'fh_':
                                    continue
                                    
                                right_section = self._get_right_section_for_row(excel_data, row_idx, col_idx)
                                
                                if not right_section:
                                    right_section = current_section
                                    
                                right_key = f"{right_section}_{right_field_name}"
                                
                                if right_key in data_to_save:
                                    # Update value in the right column's value cell
                                    sheet.cell(row=row_idx, column=col_idx+1).value = data_to_save[right_key]
                    
                    # Proses field dropdown (fd_)
                    elif first_col.startswith('fd_'):
                        field_name = first_col[3:].strip()
                        current_section = self._get_section_for_row(excel_data, row_idx)
                        
                        # Cari nilai di data_to_save
                        key = f"{current_section}_{field_name}"
                        if key in data_to_save:
                            # Update nilai di cell kedua
                            sheet.cell(row=row_idx, column=2).value = data_to_save[key]
                        
                        # Also check if there are fields in columns C and beyond
                        for col_idx in range(2, len(row)):
                            col_value = row.iloc[col_idx] if not pd.isna(row.iloc[col_idx]) else ""
                            
                            # Check for any field type (f_, fd_, fh_) in right columns
                            if isinstance(col_value, str) and (col_value.startswith('f_') or 
                                col_value.startswith('fd_') or 
                                col_value.startswith('fh_')):
                                
                                # This is a field in a right column
                                # Extract the prefix and field name
                                if col_value.startswith('f_'):
                                    prefix = col_value[:2]
                                    right_field_name = col_value[2:].strip()
                                else:
                                    prefix = col_value[:3]
                                    right_field_name = col_value[3:].strip()
                                
                                # Skip headers as we don't save them
                                if prefix == 'fh_':
                                    continue
                                    
                                right_section = self._get_right_section_for_row(excel_data, row_idx, col_idx)
                                
                                if not right_section:
                                    right_section = current_section
                                    
                                right_key = f"{right_section}_{right_field_name}"
                                
                                if right_key in data_to_save:
                                    # Update value in the right column's value cell
                                    sheet.cell(row=row_idx, column=col_idx+1).value = data_to_save[right_key]
                    
                    # Proses field multiple (fm_)
                    elif first_col.startswith('fm_'):
                        field_name = first_col[3:].strip()
                        current_section = self._get_section_for_row(excel_data, row_idx)
                        
                        # Cari nilai di data_to_save
                        key = f"{current_section}_{field_name}"
                        if key in data_to_save:
                            # Update nilai di semua cells untuk field multiple
                            values = data_to_save[key]
                            
                            # Loop through all possible values (up to 10 or length of values array)
                            for i in range(min(len(values), 10)):  # limit to prevent index errors
                                if values[i] is not None:
                                    # col_idx is i+2 because Excel columns start at 1, and the first column is for the field name
                                    sheet.cell(row=row_idx, column=i+2).value = values[i]
                        
                        # Also check if there are fields in columns C and beyond
                        for col_idx in range(2, len(row)):
                            col_value = row.iloc[col_idx] if not pd.isna(row.iloc[col_idx]) else ""
                            
                            # Check for any field type in right columns
                            if isinstance(col_value, str) and (col_value.startswith('f_') or 
                                col_value.startswith('fd_') or 
                                col_value.startswith('fm_') or
                                col_value.startswith('fh_')):
                                
                                # This is a field in a right column
                                # Extract the prefix and field name
                                if col_value.startswith('f_'):
                                    prefix = col_value[:2]
                                    right_field_name = col_value[2:].strip()
                                else:
                                    prefix = col_value[:3]
                                    right_field_name = col_value[3:].strip()
                                
                                # Skip headers as we don't save them
                                if prefix == 'fh_':
                                    continue
                                    
                                right_section = self._get_right_section_for_row(excel_data, row_idx, col_idx)
                                
                                if not right_section:
                                    right_section = current_section
                                    
                                right_key = f"{right_section}_{right_field_name}"
                                
                                if right_key in data_to_save:
                                    # For regular fields or dropdowns
                                    if not col_value.startswith('fm_'):
                                        sheet.cell(row=row_idx, column=col_idx+1).value = data_to_save[right_key]
                                    else:
                                        # For multiple fields
                                        right_values = data_to_save[right_key]
                                        
                                        # Loop through all possible values
                                        for i in range(min(len(right_values), 5)):  # limit to prevent errors
                                            if right_values[i] is not None:
                                                # col_idx+i+1 because we start at the column after the field name
                                                sheet.cell(row=row_idx, column=col_idx+i+1).value = right_values[i]
            
                if '_client' in field_key or '_contractor' in field_key or '_remarks' in field_key:
                    # Parse parts dari field key untuk mendapatkan informasi lengkap
                    parts = field_key.split('_')
                    suffix = parts[-1]  # 'client', 'contractor', atau 'remarks'
                    
                    # Dapatkan sheet_name, section, dan base_key
                    sheet_part = parts[0]
                    section_part = parts[1]
                    
                    # Jika ada banyak underscore dalam task name, kita perlu merekonstruksi dengan benar
                    if len(parts) > 3:
                        # Gabungkan semua bagian di tengah untuk mendapatkan task name dengan underscore
                        task_id_parts = parts[2:-1]  # Ambil semua kecuali sheet, section, dan suffix
                        task_id = '_'.join(task_id_parts)
                    else:
                        # Format sederhana: sheet_section_task_suffix
                        task_id = parts[2]
                    
                    # Cari task name yang sesuai dengan ID ini dalam excel_data
                    task_name = None
                    row_idx_to_update = None
                    
                    # Scan Excel untuk menemukan baris yang sesuai
                    for idx, excel_row in excel_data.iterrows():
                        # Skip baris kosong
                        if pd.isna(excel_row).all():
                            continue
                            
                        first_col = excel_row.iloc[0] if not pd.isna(excel_row.iloc[0]) else ""
                        if not isinstance(first_col, str):
                            try:
                                first_col = str(first_col)
                            except:
                                continue
                                
                        # Periksa apakah ini adalah item table (ft_)
                        if first_col.startswith('ft_'):
                            task_name_in_excel = first_col[3:].strip()
                            current_task_id = f"{sheet_part}_{section_part}_{idx}"
                            
                            # Jika ini adalah baris yang kita cari berdasarkan indeks baris
                            if str(idx) == task_id or current_task_id == f"{sheet_part}_{section_part}_{task_id}":
                                row_idx_to_update = idx + 1  # +1 karena Excel row mulai dari 1
                                task_name = task_name_in_excel
                                break
                    
                    # Jika menemukan baris yang cocok
                    if row_idx_to_update is not None:
                        # Tentukan kolom berdasarkan jenis field
                        col_idx = None
                        value = None
                        
                        if suffix == 'client':
                            col_idx = 1  # Kolom B
                            value = 'ü' if field_input.isChecked() else ''
                        elif suffix == 'contractor':
                            col_idx = 2  # Kolom C
                            value = 'ü' if field_input.isChecked() else ''
                        elif suffix == 'remarks':
                            col_idx = 3  # Kolom D
                            value = field_input.text()
                        
                        # Update nilai di cell Excel
                        if col_idx is not None:
                            sheet.cell(row=row_idx_to_update, column=col_idx+1).value = value
                            
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