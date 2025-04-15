import os
import sys
import re
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QFrame, QGridLayout, QMessageBox, QApplication,
                             QGraphicsDropShadowEffect, QSpacerItem, QSizePolicy, QStyle)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QMovie, QCursor
from PyQt5.QtCore import Qt, QSize, QTimer
from PyQt5.QtSvg import QSvgWidget

# Import local modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_NAME, APP_LOGO, PRIMARY_COLOR, SECONDARY_COLOR, BG_COLOR

# SVG Resources untuk ikon
SVG_RESOURCES = {
    "user": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2023 Fonticons, Inc.--><path fill="#555555" d="M224 256A128 128 0 1 0 224 0a128 128 0 1 0 0 256zm-45.7 48C79.8 304 0 383.8 0 482.3C0 498.7 13.3 512 29.7 512H418.3c16.4 0 29.7-13.3 29.7-29.7C448 383.8 368.2 304 269.7 304H178.3z"/></svg>''',
    
    "lock": '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512"><!--!Font Awesome Free 6.5.1 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2023 Fonticons, Inc.--><path fill="#555555" d="M144 144v48H304V144c0-44.2-35.8-80-80-80s-80 35.8-80 80zM80 192V144C80 64.5 144.5 0 224 0s144 64.5 144 144v48h16c35.3 0 64 28.7 64 64V448c0 35.3-28.7 64-64 64H64c-35.3 0-64-28.7-64-64V256c0-35.3 28.7-64 64-64H80z"/></svg>'''
}

def get_svg_with_color(svg_key, color="#3498DB"):
    svg_content = SVG_RESOURCES[svg_key]
    
    svg_content = re.sub(r'fill="#[0-9A-Fa-f]{6}"', f'fill="{color}"', svg_content)
    
    return svg_content

class StyleHelper:
    
    @staticmethod
    def add_shadow(widget, radius=10, x_offset=0, y_offset=3, color="#32000000"):
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(radius)
        shadow.setOffset(x_offset, y_offset)
        shadow.setColor(QColor(color))
        widget.setGraphicsEffect(shadow)

class LoginView(QWidget):
    def __init__(self, auth_manager, on_login_success):
        super().__init__()
        self.auth_manager = auth_manager
        self.on_login_success = on_login_success
        
        # Status variabel
        self.login_in_progress = False
        
        # Setup UI
        self.initUI()
    
    def setup_line_edit_with_svg(self, line_edit, svg_key, placeholder_text, color=SECONDARY_COLOR, size=20):
        # Buat layout horizontal untuk line edit
        container = QFrame()
        container.setObjectName("inputContainer")
        container.setStyleSheet("""
            #inputContainer {
                background-color: #f9f9f9;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            #inputContainer:focus-within {
                border: 1px solid #3498DB;
                background-color: white;
            }
        """)
        
        layout = QHBoxLayout(container)
        layout.setContentsMargins(10, 0, 10, 0)
        layout.setSpacing(8)
        
        # Buat widget SVG
        svg_widget = QSvgWidget()
        svg_widget.setFixedSize(size, size)
        svg_content = get_svg_with_color(svg_key, color)
        svg_bytes = QtCore.QByteArray(svg_content.encode('utf-8'))
        svg_widget.load(svg_bytes)
        
        # Reset line edit style dan tambahkan ke layout
        line_edit.setPlaceholderText(placeholder_text)
        line_edit.setMinimumHeight(42)
        line_edit.setFont(QFont("Segoe UI", 11))
        line_edit.setStyleSheet("""
            QLineEdit {
                border: none;
                background-color: transparent;
                padding: 0;
            }
        """)
        
        # Tambahkan widget ke layout
        layout.addWidget(svg_widget)
        layout.addWidget(line_edit)
        
        return container
    
    def initUI(self):
        # Set window properties
        self.setWindowTitle(f"Login - {APP_NAME}")
        self.setMinimumSize(1000, 650)
        self.setWindowIcon(QIcon(APP_LOGO))
        
        # Main layout
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        self.setLayout(main_layout)
        
        # Left side (brand panel)
        left_panel = QFrame()
        left_panel.setStyleSheet(f"""
            background-color: {PRIMARY_COLOR};
        """)
        left_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        left_panel.setMinimumWidth(500)
        
        # Create a VBox layout for left panel
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignCenter)
        left_layout.setContentsMargins(40, 40, 40, 40)
        
        # Add logo
        logo_container = QFrame()
        logo_container.setFixedSize(180, 180)
        logo_container.setStyleSheet("background-color: transparent; border: none;")
        
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setAlignment(Qt.AlignCenter)
        
        logo_label = QLabel()
        logo_pixmap = QPixmap(APP_LOGO)
        logo_pixmap = logo_pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(logo_pixmap)
        logo_label.setAlignment(Qt.AlignCenter)
        
        logo_layout.addWidget(logo_label)
        StyleHelper.add_shadow(logo_container, radius=20, color="#80000000")
        
        # Add app name
        app_name_label = QLabel(APP_NAME)
        app_name_label.setFont(QFont("Segoe UI", 28, QFont.Bold))
        app_name_label.setStyleSheet("color: white;")
        app_name_label.setAlignment(Qt.AlignCenter)
        
        # Add tagline
        tagline_label = QLabel("Enterprise Platform")
        tagline_label.setFont(QFont("Segoe UI", 14))
        tagline_label.setStyleSheet("color: rgba(255, 255, 255, 0.85);")
        tagline_label.setAlignment(Qt.AlignCenter)
        
        # Add version
        version_label = QLabel("Version 1.0.0")
        version_label.setFont(QFont("Segoe UI", 10))
        version_label.setStyleSheet("color: rgba(255, 255, 255, 0.5);")
        version_label.setAlignment(Qt.AlignCenter)
        
        # Add elements to left layout
        left_layout.addStretch()
        left_layout.addWidget(logo_container, 0, Qt.AlignCenter)
        left_layout.addSpacing(20)
        left_layout.addWidget(app_name_label)
        left_layout.addWidget(tagline_label)
        left_layout.addSpacing(10)
        left_layout.addWidget(version_label)
        left_layout.addStretch()
        
        # Right side (login form)
        right_panel = QFrame()
        right_panel.setStyleSheet(f"""
            background-color: {BG_COLOR};
        """)
        right_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Create a VBox layout for right panel
        right_layout = QVBoxLayout(right_panel)
        right_layout.setAlignment(Qt.AlignCenter)
        right_layout.setContentsMargins(60, 40, 60, 40)
        
        # Welcome text
        welcome_label = QLabel("Welcome Back")
        welcome_label.setFont(QFont("Segoe UI", 26, QFont.Bold))
        welcome_label.setStyleSheet(f"color: {PRIMARY_COLOR};")
        
        welcome_subtitle = QLabel("Please sign in to your account")
        welcome_subtitle.setFont(QFont("Segoe UI", 12))
        welcome_subtitle.setStyleSheet("color: #666;")
        
        right_layout.addStretch()
        right_layout.addWidget(welcome_label)
        right_layout.addWidget(welcome_subtitle)
        right_layout.addSpacing(40)
        
        # Login form container
        form_container = QFrame()
        form_container.setMaximumWidth(380)
        form_container.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 15px;
                border: 1px solid #e0e0e0;
            }
        """)
        StyleHelper.add_shadow(form_container, radius=15, y_offset=4, color="#20000000")
        
        form_layout = QVBoxLayout(form_container)
        form_layout.setContentsMargins(30, 30, 30, 30)
        form_layout.setSpacing(20)
        
        # Username field - Tanpa border pada label
        username_label = QLabel("Username")
        username_label.setFont(QFont("Segoe UI", 11))
        username_label.setStyleSheet("color: #555; background: transparent; border: none;")
        
        self.username_input = QLineEdit()
        
        # Setup username input dengan SVG icon dan warna SECONDARY_COLOR
        username_container = self.setup_line_edit_with_svg(
            self.username_input, 
            "user",
            "Enter your username",
            PRIMARY_COLOR  
        )
        
        # Password field
        password_label = QLabel("Password")
        password_label.setFont(QFont("Segoe UI", 11))
        password_label.setStyleSheet("color: #555; background: transparent; border: none;")
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        
        # Setup password input 
        password_container = self.setup_line_edit_with_svg(
            self.password_input, 
            "lock",
            "Enter your password",
            PRIMARY_COLOR  
        )
        
        # Remember me & Forgot password row
        options_layout = QHBoxLayout()
        
        self.remember_checkbox = QtWidgets.QCheckBox("Remember me")
        self.remember_checkbox.setFont(QFont("Segoe UI", 10))
        self.remember_checkbox.setStyleSheet("""
            QCheckBox {
                color: #555;
                background: transparent;
                border: none;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        
        forgot_password = QLabel("Forgot password?")
        forgot_password.setFont(QFont("Segoe UI", 10))
        forgot_password.setStyleSheet(f"color: {SECONDARY_COLOR}; text-decoration: underline; background: transparent; border: none;")
        forgot_password.setCursor(Qt.PointingHandCursor)
        
        options_layout.addWidget(self.remember_checkbox)
        options_layout.addStretch()
        options_layout.addWidget(forgot_password)
        
        # Login button with loading state
        self.login_button = QPushButton("Login")
        self.login_button.setMinimumHeight(48)
        self.login_button.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.login_button.setCursor(Qt.PointingHandCursor)
        self.login_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {SECONDARY_COLOR};
                color: white;
                border: none;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: #2980B9;
            }}
            QPushButton:pressed {{
                background-color: #1F618D;
            }}
            QPushButton:disabled {{
                background-color: #7fb6dd;
            }}
        """)
        StyleHelper.add_shadow(self.login_button, radius=10, y_offset=3, color="#40000000")
        
        # Connect button to login function
        self.login_button.clicked.connect(self.attempt_login)
        self.password_input.returnPressed.connect(self.attempt_login)
        
        # Error message (initially hidden)
        self.error_label = QLabel("")
        self.error_label.setFont(QFont("Segoe UI", 10))
        self.error_label.setStyleSheet("color: #E74C3C;")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setWordWrap(True)
        self.error_label.setVisible(False)
        
        # Add all elements to form layout
        form_layout.addWidget(username_label)
        form_layout.addWidget(username_container)
        form_layout.addWidget(password_label)
        form_layout.addWidget(password_container)
        form_layout.addLayout(options_layout)
        form_layout.addWidget(self.error_label)
        form_layout.addWidget(self.login_button)
        
        # Add form to right layout
        right_layout.addWidget(form_container, 0, Qt.AlignCenter)
        right_layout.addStretch()
        
        # Add copyright notice
        copyright_label = QLabel("© 2025 DIAC-V. All rights reserved.")
        copyright_label.setFont(QFont("Segoe UI", 9))
        copyright_label.setStyleSheet("color: #999;")
        copyright_label.setAlignment(Qt.AlignCenter)
        
        right_layout.addWidget(copyright_label)
        
        # Add both panels to main layout
        main_layout.addWidget(left_panel, 1)  # Proportion 1
        main_layout.addWidget(right_panel, 1)  # Proportion 1
        
        # Set focus to username
        self.username_input.setFocus()
    
    def show_error(self, message):
        """Tampilkan pesan error"""
        self.error_label.setText(message)
        self.error_label.setVisible(True)
        
        # Shake effect untuk form
        self.shake_effect()
        
        # Sembunyikan error setelah beberapa detik
        QTimer.singleShot(5000, lambda: self.error_label.setVisible(False))
    
    def shake_effect(self):
        """Buat efek getar pada form saat login gagal"""
        self.original_pos = self.pos()
        
        shake_distance = 10
        shake_count = 5
        shake_duration = 50  # milliseconds
        
        for i in range(shake_count):
            QTimer.singleShot(i * shake_duration, lambda d=shake_distance if i % 2 == 0 else -shake_distance: 
                self.move(self.x() + d, self.y()))
            
        # Kembalikan ke posisi awal
        QTimer.singleShot(shake_count * shake_duration, lambda: self.move(self.original_pos))
    
    def set_loading_state(self, is_loading=True):
        """Atur status loading pada form login"""
        self.login_in_progress = is_loading
        self.login_button.setEnabled(not is_loading)
        self.username_input.setEnabled(not is_loading)
        self.password_input.setEnabled(not is_loading)
        self.remember_checkbox.setEnabled(not is_loading)
        
        if is_loading:
            self.login_button.setText("Signing in...")
        else:
            self.login_button.setText("Login")
    
    def attempt_login(self):
        """Mencoba melakukan login dengan kredensi yang dimasukkan"""
        # Jika sudah dalam proses login, abaikan
        if self.login_in_progress:
            return
            
        username = self.username_input.text().strip()
        password = self.password_input.text()
        
        if not username or not password:
            self.show_error("Please enter both username and password.")
            return
        
        # Set status loading
        self.set_loading_state(True)
        
        # Simulasikan delay jaringan (untuk UX yang lebih baik)
        QTimer.singleShot(800, lambda: self.process_login(username, password))
    
    def process_login(self, username, password):
        """Proses login setelah delay"""
        success, message = self.auth_manager.login(username, password)
        
        if success:
            # Login berhasil
            self.on_login_success()
        else:
            # Login gagal
            self.set_loading_state(False)
            self.show_error(message)
            self.password_input.clear()
            self.password_input.setFocus()
    
    def showEvent(self, event):
        """Reset field-field saat form ditampilkan"""
        self.username_input.clear()
        self.password_input.clear()
        self.username_input.setFocus()
        self.error_label.setVisible(False)
        self.set_loading_state(False)
        super().showEvent(event)