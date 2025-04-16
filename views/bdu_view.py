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
        header_layout = None  # For bh_ headers

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

                continue
            
            # Check if it's a big header (bh_)
            if first_col.startswith('bh_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()  # Changed from QFrame to QWidget
                    section_frame.setStyleSheet("""
                        background-color: white;
                    """)

                    section_layout = QVBoxLayout(section_frame)
                    section_layout.setContentsMargins(15, 15, 15, 15)
                    section_layout.setSpacing(15)

                    # Add to main layout
                    layout.addWidget(section_frame)
                    current_section = "Default"

                # Create big header
                header_text = first_col[3:].strip()  # Remove 'bh_' prefix

                big_header = QLabel(header_text)
                big_header.setFont(QFont("Segoe UI", 14, QFont.Bold))
                big_header.setStyleSheet(f"color: {PRIMARY_COLOR}; margin-top: 10px;")

                section_layout.addWidget(big_header)

                # Create a horizontal line below the header
                line = QFrame()
                line.setFrameShape(QFrame.HLine)
                line.setFrameShadow(QFrame.Sunken)
                line.setStyleSheet(f"background-color: {PRIMARY_COLOR}; margin-bottom: 10px;")
                line.setFixedHeight(2)

                section_layout.addWidget(line)
                continue
            
            # Check if it's a field header (fh_)
            if first_col.startswith('fh_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()  # Changed from QFrame to QWidget
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

                header_label = QLabel(field_header)
                header_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
                header_label.setStyleSheet("color: #555; margin-top: 5px;")

                section_layout.addWidget(header_label)
                continue
            
            # Check if it's a field (f_)
            if first_col.startswith('f_'):
                if section_layout is None:
                    # If no section is defined yet, create a default one
                    section_frame = QWidget()  # Changed from QFrame to QWidget
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
                label.setMinimumWidth(250)

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
                if len(row) > 3 and not pd.isna(row.iloc[3]):
                    default_value = str(row.iloc[3]).strip()
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