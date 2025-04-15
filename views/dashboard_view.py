import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFrame, QGridLayout, QSpacerItem,
                             QSizePolicy, QScrollArea, QApplication, QMenu, QAction)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QCursor
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QPoint

# Import local modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_NAME, APP_LOGO, DEPARTMENTS, PRIMARY_COLOR, SECONDARY_COLOR, BG_COLOR

class DepartmentCard(QFrame):
    """Widget kartu untuk menampilkan departemen"""
    clicked = pyqtSignal(str)
    
    def __init__(self, dept_id, name, color, emoji):
        super().__init__()
        self.dept_id = dept_id
        self.color = color
        self.initUI(name, color, emoji)
    
    def initUI(self, name, color, emoji):
        # Set frame properties
        self.setFixedSize(220, 180)  # Larger card size
        self.setStyleSheet("""
            background-color: transparent;
            border-radius: 8px;
            border: none;
        """)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        
        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 15, 12, 15)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignCenter)
        
        # Emoji as icon
        icon_label = QLabel(emoji)
        icon_label.setFont(QFont("Segoe UI", 32))
        icon_label.setStyleSheet(f"color: {color}; background-color: transparent;")
        icon_label.setAlignment(Qt.AlignCenter)
        
        # Colored horizontal line
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet(f"background-color: {color};")
        line.setFixedHeight(2)
        
        # Department name
        name_label = QLabel(name)
        name_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        name_label.setAlignment(Qt.AlignCenter)
        name_label.setStyleSheet(f"color: {PRIMARY_COLOR}; background-color: transparent;")
        
        # Add all elements to layout
        layout.addWidget(icon_label)
        layout.addWidget(line)
        layout.addWidget(name_label)
        layout.addStretch()
    
    def mousePressEvent(self, event):
        """Event yang terjadi saat kartu diklik"""
        self.clicked.emit(self.dept_id)
        super().mousePressEvent(event)
    
    # Add hover event manually instead of in stylesheet
    def enterEvent(self, event):
        self.setStyleSheet(f"""
            background-color: rgba(249, 249, 249, 0.5);
            border-radius: 8px;
            border: none;
        """)
        super().enterEvent(event)
    
    def leaveEvent(self, event):
        self.setStyleSheet("""
            background-color: transparent;
            border-radius: 8px;
            border: none;
        """)
        super().leaveEvent(event)

class DashboardView(QMainWindow):
    """UI utama untuk Dashboard"""
    # Signal untuk logout
    logout_signal = pyqtSignal()
    
    # Signal untuk membuka department
    open_department_signal = pyqtSignal(str)  # New signal for opening departments
    
    def __init__(self, auth_manager):
        super().__init__()
        self.auth_manager = auth_manager
        self.current_user = auth_manager.get_current_user()
        self.accessible_depts = auth_manager.get_accessible_departments()
        self.initUI()
    
    def initUI(self):
        # Set window properties
        self.setWindowTitle(f"Dashboard - {APP_NAME}")
        self.setMinimumSize(1000, 700)
        self.setWindowIcon(QIcon(APP_LOGO))
        
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
        content_widget = QWidget()
        content_widget.setStyleSheet(f"background-color: {BG_COLOR};")
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        
        # Welcome section
        welcome_layout = QHBoxLayout()
        
        welcome_text = QLabel(f"Welcome, {self.current_user['name']}")
        welcome_text.setFont(QFont("Segoe UI", 22, QFont.Bold))
        welcome_text.setStyleSheet(f"color: {PRIMARY_COLOR};")
        
        date_label = QLabel(QtCore.QDate.currentDate().toString("dddd, MMMM d, yyyy"))
        date_label.setFont(QFont("Segoe UI", 12))
        date_label.setStyleSheet("color: #666;")
        date_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        welcome_layout.addWidget(welcome_text)
        welcome_layout.addStretch()
        welcome_layout.addWidget(date_label)
        
        content_layout.addLayout(welcome_layout)
        
        # Department heading
        dept_heading_layout = QHBoxLayout()
        
        dept_title = QLabel("Departments")
        dept_title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        dept_title.setStyleSheet(f"color: {PRIMARY_COLOR};")
        
        dept_subtitle = QLabel("Access your department modules")
        dept_subtitle.setFont(QFont("Segoe UI", 11))
        dept_subtitle.setStyleSheet("color: #666;")
        
        dept_heading_layout.addWidget(dept_title)
        dept_heading_layout.addSpacing(10)
        dept_heading_layout.addWidget(dept_subtitle)
        dept_heading_layout.addStretch()
        
        content_layout.addSpacing(15)
        content_layout.addLayout(dept_heading_layout)
        content_layout.addSpacing(15)
        
        # Top row - 3 departments
        top_row_layout = QHBoxLayout()
        top_row_layout.setSpacing(30)  # Increased spacing between cards
        
        # Bottom row - 4 departments
        bottom_row_layout = QHBoxLayout()
        bottom_row_layout.setSpacing(30)  # Increased spacing between cards
        
        # Add department cards to their respective rows
        dept_count = 0
        row_size = 3  # First row has 3 items
        
        for dept in DEPARTMENTS:
            dept_id = dept["id"]
            dept_card = DepartmentCard(
                dept_id=dept_id,
                name=dept["name"],
                color=dept["color"],
                emoji=dept["emoji"]
            )
            
            # If user doesn't have access, make it look disabled
            if dept_id not in self.accessible_depts:
                dept_card.setEnabled(False)
                dept_card.setStyleSheet("""
                    background-color: rgba(240, 240, 240, 0.5);
                    border-radius: 8px;
                    border: none;
                    opacity: 0.5;
                """)
                dept_card.setCursor(QCursor(Qt.ForbiddenCursor))
            
            # Connect click signal
            dept_card.clicked.connect(self.open_department)
            
            # Add to appropriate row
            if dept_count < row_size:
                top_row_layout.addWidget(dept_card)
            else:
                bottom_row_layout.addWidget(dept_card)
            
            dept_count += 1
        
        # Add stretches to center cards
        top_row_layout.insertStretch(0)
        top_row_layout.addStretch()
        
        bottom_row_layout.insertStretch(0)
        bottom_row_layout.addStretch()
        
        # Add rows to content layout
        content_layout.addLayout(top_row_layout)
        content_layout.addSpacing(30)  # Space between rows
        content_layout.addLayout(bottom_row_layout)
        content_layout.addStretch(1)  # Push content to the top
        
        # Add content to main layout
        main_layout.addWidget(content_widget)
        
        # Status bar
        self.statusBar().showMessage(f"Logged in as {self.current_user['username']} | Department: {self.current_user['department']}")
        self.statusBar().setStyleSheet("background-color: #f0f0f0; color: #555;")
    
    def setup_header(self):
        """Setup header widget dengan logo, judul, dan menu"""
        self.header_widget = QWidget()
        self.header_widget.setFixedHeight(60)
        self.header_widget.setStyleSheet(f"background-color: {PRIMARY_COLOR};")
        
        header_layout = QHBoxLayout(self.header_widget)
        header_layout.setContentsMargins(15, 0, 15, 0)
        
        # Logo
        logo_label = QLabel()
        logo_pixmap = QPixmap(APP_LOGO)
        logo_pixmap = logo_pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(logo_pixmap)
        
        # App title
        title_label = QLabel(APP_NAME)
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
        header_layout.addWidget(logo_label)
        header_layout.addSpacing(8)
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
        profile_action = QAction("My Profile", self)
        settings_action = QAction("Settings", self)
        logout_action = QAction("Logout", self)
        
        menu.addAction(profile_action)
        menu.addAction(settings_action)
        menu.addSeparator()
        menu.addAction(logout_action)
        
        # Connect actions
        logout_action.triggered.connect(self.logout)
        
        # Show menu at button position
        menu.exec_(sender.mapToGlobal(QPoint(0, sender.height())))
    
    def open_department(self, dept_id):
        """Buka modul departemen yang diklik"""
        # Emit signal with department ID
        self.open_department_signal.emit(dept_id)
        print(f"Opening department: {dept_id}")
    
    def logout(self):
        """Logout dari aplikasi"""
        self.auth_manager.logout()
        # Emit signal untuk menampilkan login view
        self.logout_signal.emit()
        self.close()
    
    # Override untuk memastikan tampilan maximized
    def showEvent(self, event):
        """Event handler saat window ditampilkan"""
        super().showEvent(event)
        self.showMaximized()  # Pastikan window benar-benar maximized