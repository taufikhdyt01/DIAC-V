#!/usr/bin/env python3

import os
import sys
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QSplashScreen, QMessageBox
from PyQt5.QtGui import QPixmap, QIcon

app = QApplication(sys.argv)

# Import modules lokal
from modules.auth import AuthManager
from views.login_view import LoginView
from views.dashboard_view import DashboardView
from views.bdu_view import BDUGroupView  # Import the BDU view
from config import APP_NAME, APP_LOGO

class DIACApplication:
    
    def __init__(self):
        # Gunakan aplikasi Qt yang sudah dibuat
        self.app = app
        self.app.setApplicationName(APP_NAME)
        self.app.setWindowIcon(QIcon(APP_LOGO))
        
        # Set style aplikasi
        self.setup_styles()
        
        # Inisialisasi auth manager
        self.auth_manager = AuthManager()
        
        # Inisialisasi splash screen
        self.show_splash_screen()
        
        # Inisialisasi views
        self.login_view = None
        self.dashboard_view = None
        self.bdu_view = None  # Add BDU view
        
        # Timer untuk menampilkan login setelah splash
        QtCore.QTimer.singleShot(2000, self.init_views)
    
    def setup_styles(self):
        # Set style aplikasi dengan styleheet global
        stylesheet = """
        QWidget {
            font-family: "Segoe UI", Arial, sans-serif;
        }
        QToolTip {
            border: 1px solid #2C3E50;
            background-color: #34495E;
            color: white;
            padding: 5px;
            border-radius: 3px;
        }
        QScrollBar:vertical {
            border: none;
            background: #F5F5F5;
            width: 12px;
            margin: 0px;
        }
        QScrollBar::handle:vertical {
            background: #CCCCCC;
            min-height: 20px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical:hover {
            background: #BBBBBB;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            border: none;
            background: none;
            height: 0px;
        }
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
            background: none;
        }
        QMessageBox {
            font-size: 13px;
        }
        QMessageBox QPushButton {
            width: 80px;
            height: 30px;
        }
        """
        self.app.setStyleSheet(stylesheet)
    
    def show_splash_screen(self):
        """Tampilkan splash screen saat aplikasi dimulai"""
        # Cek apakah file logo ada
        if not os.path.exists(APP_LOGO):
            # Buat direktori aset jika belum ada
            os.makedirs(os.path.dirname(APP_LOGO), exist_ok=True)
            
            # Jika logo tidak ada, coba buat logo default
            self.create_default_logo()
        
        # Buat splash screen
        splash_pixmap = QPixmap(APP_LOGO).scaled(200, 200, QtCore.Qt.KeepAspectRatio, 
                                                QtCore.Qt.SmoothTransformation)
        self.splash = QSplashScreen(splash_pixmap, QtCore.Qt.WindowStaysOnTopHint)
        
        # Tambahkan teks versi
        self.splash.showMessage(f"Version 1.0.0", QtCore.Qt.AlignBottom | QtCore.Qt.AlignRight, QtCore.Qt.white)
        
        # Tampilkan splash screen
        self.splash.show()
        self.app.processEvents()
    
    def create_default_logo(self):
        """Buat logo default jika logo tidak ada"""
        try:
            # Buat image 200x200 dengan teks
            from PIL import Image, ImageDraw, ImageFont
            
            img = Image.new('RGB', (200, 200), color=(44, 62, 80))
            d = ImageDraw.Draw(img)
            
            # Coba gunakan font sistem
            try:
                font = ImageFont.truetype("arial.ttf", 40)
            except:
                font = ImageFont.load_default()
                
            # Tambahkan teks
            d.text((50, 80), "DIAC-V", fill=(255, 255, 255), font=font)
            
            # Simpan image
            img.save(APP_LOGO)
        except Exception as e:
            print(f"Error creating default logo: {str(e)}")
    
    def init_views(self):
        """Inisialisasi views aplikasi"""
        # Tutup splash screen
        self.splash.finish(None)
        
        # Inisialisasi login view
        self.login_view = LoginView(self.auth_manager, self.on_login_success)
        
        # Tampilkan login view sebagai maximized
        self.login_view.showMaximized()
    
    def on_login_success(self):
        """Callback saat login berhasil"""
        # Sembunyikan login view
        self.login_view.hide()
        
        # Inisialisasi dan tampilkan dashboard
        self.dashboard_view = DashboardView(self.auth_manager)
        
        # Connect signal logout dari dashboard ke handler
        self.dashboard_view.logout_signal.connect(self.on_logout)
        
        # Connect department signals (for BDU opening)
        self.connect_department_signals()
        
        # Tampilkan dashboard sebagai maximized
        self.dashboard_view.showMaximized()
    
    def connect_department_signals(self):
        """Connect signals for opening department modules"""
        # Get the clicked signal from dashboard and connect it to open_department method
        if hasattr(self.dashboard_view, 'open_department_signal'):
            self.dashboard_view.open_department_signal.connect(self.open_department)
    
    def check_bdu_excel(self):
        """Check if SET_BDU.xlsx exists"""
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "SET_BDU.xlsx")
        return os.path.exists(excel_path)
    
    def open_department(self, dept_id):
        """Open department module based on department ID"""
        if dept_id == "BDU":
            # Check if SET_BDU.xlsx exists
            if not self.check_bdu_excel():
                QMessageBox.warning(self.dashboard_view, "File Not Found", 
                                   "SET_BDU.xlsx is missing in the data directory.\nPlease add the file and try again.")
                return
            
            # Hide dashboard
            self.dashboard_view.hide()
            
            # Initialize BDU view if not already initialized
            if not self.bdu_view:
                self.bdu_view = BDUGroupView(self.auth_manager)
                # Connect back signal
                self.bdu_view.back_to_dashboard.connect(self.on_back_to_dashboard)
            
            # Show BDU view
            self.bdu_view.showMaximized()
        else:
            # Notify user that other departments are not implemented yet
            QMessageBox.information(self.dashboard_view, "Department Access", 
                                   f"The {dept_id} module is not implemented yet.")
    
    def on_back_to_dashboard(self):
        """Callback when going back to dashboard from a department view"""
        # Hide department views
        if self.bdu_view:
            self.bdu_view.hide()
        
        # Show dashboard
        self.dashboard_view.showMaximized()
    
    def on_logout(self):
        """Callback saat logout"""
        # Tutup semua views
        if self.dashboard_view:
            self.dashboard_view.close()
            self.dashboard_view = None
        
        if self.bdu_view:
            self.bdu_view.close()
            self.bdu_view = None
        
        # Tampilkan login view kembali sebagai maximized
        self.login_view.showMaximized()
    
    def run(self):
        """Jalankan aplikasi"""
        return self.app.exec_()

# Entry point aplikasi
if __name__ == "__main__":
    # Buat direktori yang diperlukan
    os.makedirs(os.path.join("assets", "icons"), exist_ok=True)
    os.makedirs("data", exist_ok=True)
    
    # Jalankan aplikasi
    app_instance = DIACApplication()
    sys.exit(app_instance.run())