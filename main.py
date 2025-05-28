#!/usr/bin/env python3

import os
import sys
import warnings

# Suppress SIP deprecation warnings
warnings.filterwarnings("ignore", category=DeprecationWarning, module=".*sip.*")
warnings.filterwarnings("ignore", message=".*sipPyTypeDict.*")

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QSplashScreen, QMessageBox
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import QTimer, pyqtSignal, QObject

app = QApplication(sys.argv)

# Import modules lokal
from modules.auth import AuthManager
from views.login_view import LoginView
from views.dashboard_view import DashboardView
from views.customer_search_view import CustomerSearchView 
from views.bdu_view_extended import BDUGroupView
from config import APP_NAME, APP_LOGO

class ThreadSafeSignals(QObject):
    """Thread-safe signals untuk komunikasi antar komponen"""
    show_dashboard = pyqtSignal()
    show_customer_search = pyqtSignal()
    show_bdu_view = pyqtSignal()
    show_login = pyqtSignal()
    initialization_complete = pyqtSignal(bool, str)

class SimpleLoadingDialog(QtWidgets.QDialog):
    """Loading dialog sederhana tanpa threading"""
    
    def __init__(self, parent=None, title="Loading...", message="Please wait..."):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedSize(400, 150)
        self.setWindowFlags(QtCore.Qt.Dialog | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowTitleHint)
        self.setModal(False)  # Non-modal agar tidak blocking
        
        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Title
        title_label = QtWidgets.QLabel(title)
        title_label.setFont(QtGui.QFont("Segoe UI", 14, QtGui.QFont.Bold))
        title_label.setAlignment(QtCore.Qt.AlignCenter)
        title_label.setStyleSheet("color: #2C3E50;")
        
        # Message
        message_label = QtWidgets.QLabel(message)
        message_label.setFont(QtGui.QFont("Segoe UI", 11))
        message_label.setAlignment(QtCore.Qt.AlignCenter)
        message_label.setStyleSheet("color: #666;")
        message_label.setWordWrap(True)
        
        # Progress bar (indeterminate)
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setFixedHeight(20)
        
        layout.addWidget(title_label)
        layout.addWidget(message_label)
        layout.addWidget(self.progress_bar)
        
        # Center on screen
        self.center_on_screen()
    
    def center_on_screen(self):
        screen = QApplication.desktop().screenGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

class DIACApplication:
    
    def __init__(self):
        # Gunakan aplikasi Qt yang sudah dibuat
        self.app = app
        self.app.setApplicationName(APP_NAME)
        self.app.setWindowIcon(QIcon(APP_LOGO))
        
        # PENTING: Set agar aplikasi tidak keluar saat window disembunyikan
        self.app.setQuitOnLastWindowClosed(False)
        
        # Set style aplikasi
        self.setup_styles()
        
        # Inisialisasi signals thread-safe
        self.signals = ThreadSafeSignals()
        self.setup_signal_connections()
        
        # Inisialisasi auth manager
        self.auth_manager = AuthManager()
        
        # Inisialisasi splash screen
        self.show_splash_screen()
        
        # Inisialisasi views
        self.login_view = None
        self.dashboard_view = None
        self.customer_search_view = None
        self.bdu_view = None
        self.current_loading = None
        
        # Timer untuk menampilkan login setelah splash
        QtCore.QTimer.singleShot(2000, self.init_views)
    
    def setup_styles(self):
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
    
    def setup_signal_connections(self):
        """Setup thread-safe signal connections"""
        self.signals.show_dashboard.connect(self._do_show_dashboard)
        self.signals.show_customer_search.connect(self._do_show_customer_search)
        self.signals.show_bdu_view.connect(self._do_show_bdu_view)
        self.signals.show_login.connect(self._do_show_login)
    
    def show_simple_loading(self, title="Loading...", message="Please wait..."):
        """Show simple loading dialog tanpa threading"""
        if self.current_loading:
            self.current_loading.close()
            
        self.current_loading = SimpleLoadingDialog(None, title, message)
        self.current_loading.show()
        QApplication.processEvents()
        return self.current_loading
    
    def hide_current_loading(self):
        """Hide current loading dialog"""
        if self.current_loading:
            self.current_loading.close()
            self.current_loading = None
            QApplication.processEvents()
    
    def show_splash_screen(self):
        """Tampilkan splash screen saat aplikasi dimulai"""
        if not os.path.exists(APP_LOGO):
            os.makedirs(os.path.dirname(APP_LOGO), exist_ok=True)
            self.create_default_logo()
        
        splash_pixmap = QPixmap(APP_LOGO).scaled(200, 200, QtCore.Qt.KeepAspectRatio, 
                                                QtCore.Qt.SmoothTransformation)
        self.splash = QSplashScreen(splash_pixmap, QtCore.Qt.WindowStaysOnTopHint)
        self.splash.showMessage(f"Version 1.0.0", QtCore.Qt.AlignBottom | QtCore.Qt.AlignRight, QtCore.Qt.white)
        self.splash.show()
        self.app.processEvents()
    
    def create_default_logo(self):
        """Buat logo default jika logo tidak ada"""
        try:
            from PIL import Image, ImageDraw, ImageFont
            
            img = Image.new('RGB', (200, 200), color=(44, 62, 80))
            d = ImageDraw.Draw(img)
            
            try:
                font = ImageFont.truetype("arial.ttf", 40)
            except:
                font = ImageFont.load_default()
                
            d.text((50, 80), "DIAC-V", fill=(255, 255, 255), font=font)
            img.save(APP_LOGO)
        except Exception as e:
            print(f"Error creating default logo: {str(e)}")
    
    def init_views(self):
        """Inisialisasi views aplikasi"""
        self.splash.finish(None)
        
        # Inisialisasi login view
        self.login_view = LoginView(self.auth_manager, self.on_login_success)
        self.login_view.showMaximized()
    
    def on_login_success(self):
        """Callback saat login berhasil - THREAD SAFE"""
        print("Login success - starting dashboard initialization...")
        
        # Show loading
        loading = self.show_simple_loading("Initializing Dashboard", "Setting up your workspace...")
        
        # Use QTimer untuk memastikan eksekusi di main thread
        QTimer.singleShot(500, self.initialize_dashboard_delayed)
    
    def initialize_dashboard_delayed(self):
        """Initialize dashboard dengan delay untuk thread safety"""
        try:
            print("Creating dashboard in main thread...")
            
            # Hide login view
            if self.login_view:
                self.login_view.hide()
            
            # Create dashboard di main thread
            if not self.dashboard_view:
                self.dashboard_view = DashboardView(self.auth_manager)
                self.dashboard_view.logout_signal.connect(self.on_logout)
                self.connect_department_signals()
            
            # Hide loading
            self.hide_current_loading()
            
            # Show dashboard menggunakan signal
            QTimer.singleShot(100, self.signals.show_dashboard.emit)
            
        except Exception as e:
            print(f"Error in initialize_dashboard_delayed: {str(e)}")
            import traceback
            traceback.print_exc()
            
            self.hide_current_loading()
            QMessageBox.critical(None, "Initialization Error", f"Failed to initialize dashboard: {str(e)}")
            
            if self.login_view:
                self.login_view.showMaximized()
    
    def _do_show_dashboard(self):
        """Actually show dashboard - guaranteed main thread"""
        print("Showing dashboard in main thread...")
        if self.dashboard_view:
            self.dashboard_view.showMaximized()
            self.dashboard_view.raise_()
            self.dashboard_view.activateWindow()
            self.dashboard_view.setFocus()
            QApplication.processEvents()
    
    def _do_show_customer_search(self):
        """Actually show customer search - guaranteed main thread"""
        if self.customer_search_view:
            self.customer_search_view.showMaximized()
            self.customer_search_view.raise_()
            self.customer_search_view.activateWindow()
            QApplication.processEvents()
    
    def _do_show_bdu_view(self):
        """Actually show BDU view - guaranteed main thread"""
        if self.bdu_view:
            self.bdu_view.showMaximized()
            self.bdu_view.raise_()
            self.bdu_view.activateWindow()
            QApplication.processEvents()
    
    def _do_show_login(self):
        """Actually show login - guaranteed main thread"""
        if self.login_view:
            self.login_view.showMaximized()
            self.login_view.raise_()
            self.login_view.activateWindow()
            QApplication.processEvents()
    
    def connect_department_signals(self):
        """Connect signals for opening department modules"""
        if hasattr(self.dashboard_view, 'open_department_signal'):
            self.dashboard_view.open_department_signal.connect(self.open_department)
    
    def check_customer_database(self):
        """Check if database_customer.xlsx exists"""
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "database_customer.xlsx")
        return os.path.exists(db_path)
    
    def check_bdu_excel(self):
        """Check if SET_BDU.xlsx exists"""
        excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "SET_BDU.xlsx")
        return os.path.exists(excel_path)
    
    def open_department(self, dept_id):
        """Open department module based on department ID"""
        if dept_id == "BDU":
            if not self.check_bdu_excel():
                QMessageBox.warning(self.dashboard_view, "File Not Found", 
                                   "SET_BDU.xlsx is missing in the data directory.\nPlease add the file and try again.")
                return
            
            if not self.check_customer_database():
                QMessageBox.warning(self.dashboard_view, "File Not Found", 
                                   "database_customer.xlsx is missing in the data directory.\nPlease add the file and try again.")
                return
            
            # Show loading
            loading = self.show_simple_loading("Opening BDU Module", "Loading customer search interface...")
            
            # Initialize customer search dengan delay
            QTimer.singleShot(500, self.initialize_customer_search_delayed)
        else:
            QMessageBox.information(self.dashboard_view, "Department Access", 
                                   f"The {dept_id} module is not implemented yet.")
    
    def initialize_customer_search_delayed(self):
        """Initialize customer search dengan delay"""
        try:
            # Hide dashboard
            if self.dashboard_view:
                self.dashboard_view.hide()
            
            # Create customer search view
            if not self.customer_search_view:
                self.customer_search_view = CustomerSearchView(self.auth_manager)
                self.customer_search_view.back_to_dashboard.connect(self.on_back_to_dashboard)
                self.customer_search_view.open_bdu_view.connect(self.on_open_bdu_view)
            
            # Hide loading
            self.hide_current_loading()
            
            # Show customer search
            QTimer.singleShot(100, self.signals.show_customer_search.emit)
            
        except Exception as e:
            print(f"Error in initialize_customer_search_delayed: {str(e)}")
            self.hide_current_loading()
            QMessageBox.critical(self.dashboard_view, "Initialization Error", f"Failed to open customer search: {str(e)}")
            QTimer.singleShot(100, self.signals.show_dashboard.emit)
    
    def on_open_bdu_view(self, customer_name):
        """Callback when continuing to BDU view from customer search"""
        loading = self.show_simple_loading("Loading BDU Workspace", f"Setting up workspace for {customer_name}...")
        
        # Store customer name for delayed initialization
        self.current_customer = customer_name
        QTimer.singleShot(500, self.initialize_bdu_view_delayed)
    
    def initialize_bdu_view_delayed(self):
        """Initialize BDU view dengan delay"""
        try:
            customer_name = getattr(self, 'current_customer', 'Unknown')
            
            # Hide customer search view
            if self.customer_search_view:
                self.customer_search_view.hide()
            
            # Create or update BDU view
            if not self.bdu_view:
                self.bdu_view = BDUGroupView(self.auth_manager, customer_name)
                self.bdu_view.back_to_dashboard.connect(self.on_back_to_dashboard)
            else:
                if hasattr(self.bdu_view, 'set_current_customer'):
                    self.bdu_view.set_current_customer(customer_name)
            
            # Hide loading
            self.hide_current_loading()
            
            # Show BDU view
            QTimer.singleShot(100, self.signals.show_bdu_view.emit)
            
        except Exception as e:
            print(f"Error in initialize_bdu_view_delayed: {str(e)}")
            self.hide_current_loading()
            QMessageBox.critical(None, "Initialization Error", f"Failed to load BDU workspace: {str(e)}")
            
            if self.customer_search_view:
                QTimer.singleShot(100, self.signals.show_customer_search.emit)
    
    def on_back_to_dashboard(self):
        """Callback when going back to dashboard from any view"""
        loading = self.show_simple_loading("Returning to Dashboard", "Saving your session...")
        
        QTimer.singleShot(500, self.return_to_dashboard_delayed)
    
    def return_to_dashboard_delayed(self):
        """Return to dashboard dengan delay"""
        try:
            # Hide all department views
            if self.customer_search_view:
                self.customer_search_view.hide()
            
            if self.bdu_view:
                self.bdu_view.hide()
            
            # Hide loading
            self.hide_current_loading()
            
            # Show dashboard
            QTimer.singleShot(100, self.signals.show_dashboard.emit)
            
        except Exception as e:
            print(f"Error in return_to_dashboard_delayed: {str(e)}")
            self.hide_current_loading()
            QTimer.singleShot(100, self.signals.show_dashboard.emit)
    
    def on_logout(self):
        """Callback saat logout"""
        loading = self.show_simple_loading("Logging Out", "Saving your session...")
        
        QTimer.singleShot(500, self.logout_delayed)
    
    def logout_delayed(self):
        """Logout dengan delay"""
        try:
            # Close all views
            if self.dashboard_view:
                self.dashboard_view.close()
                self.dashboard_view = None
            
            if self.customer_search_view:
                self.customer_search_view.close()
                self.customer_search_view = None
                
            if self.bdu_view:
                self.bdu_view.close()
                self.bdu_view = None
            
            # Hide loading
            self.hide_current_loading()
            
            # Show login
            QTimer.singleShot(100, self.signals.show_login.emit)
            
        except Exception as e:
            print(f"Error in logout_delayed: {str(e)}")
            self.hide_current_loading()
            QTimer.singleShot(100, self.signals.show_login.emit)
    
    def quit_application(self):
        """Properly quit the application"""
        try:
            print("Quitting application...")
            
            # Hide loading
            self.hide_current_loading()
            
            # Close all windows
            if self.dashboard_view:
                self.dashboard_view.close()
            if self.customer_search_view:
                self.customer_search_view.close()
            if self.bdu_view:
                self.bdu_view.close()
            if self.login_view:
                self.login_view.close()
            
            # Quit application
            self.app.quit()
            
        except Exception as e:
            print(f"Error quitting application: {str(e)}")
            self.app.quit()
    
    def run(self):
        """Jalankan aplikasi"""
        try:
            print("Starting application...")
            result = self.app.exec_()
            print(f"Application finished with code: {result}")
            return result
            
        except Exception as e:
            print(f"Error in application run: {str(e)}")
            import traceback
            traceback.print_exc()
            return 1

# Entry point aplikasi
if __name__ == "__main__":
    try:
        print("=== DIAC Application Starting ===")
        
        # Buat direktori yang diperlukan
        os.makedirs(os.path.join("assets", "icons"), exist_ok=True)
        os.makedirs("data", exist_ok=True)
        
        # Jalankan aplikasi
        app_instance = DIACApplication()
        exit_code = app_instance.run()
        
        print(f"=== Application Finished with code: {exit_code} ===")
        sys.exit(exit_code)
        
    except Exception as e:
        print(f"=== FATAL ERROR ===")
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        try:
            error_app = QApplication.instance()
            if not error_app:
                error_app = QApplication(sys.argv)
                
            QMessageBox.critical(None, "Fatal Error", 
                               f"Application failed to start:\n\n{str(e)}")
        except:
            pass
            
        sys.exit(1)