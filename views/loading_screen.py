import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QProgressBar, QFrame, QApplication)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QMovie
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QPropertyAnimation, QEasingCurve

# Import local modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_NAME, PRIMARY_COLOR, SECONDARY_COLOR

class LoadingWorker(QThread):
    """Worker thread untuk menjalankan task di background"""
    progress_updated = pyqtSignal(int, str)  # progress percentage, status text
    task_completed = pyqtSignal(bool, str)  # success, result/error message
    
    def __init__(self, task_function, *args, **kwargs):
        super().__init__()
        self.task_function = task_function
        self.args = args
        self.kwargs = kwargs
        self.progress_callback = None
        self._stop_requested = False
        
    def set_progress_callback(self, callback):
        """Set callback function untuk update progress"""
        self.progress_callback = callback
        
    def update_progress(self, percentage, message=""):
        """Update progress dari task function"""
        if not self._stop_requested:
            self.progress_updated.emit(percentage, message)
            QApplication.processEvents()
        
    def stop(self):
        """Request worker thread to stop"""
        self._stop_requested = True
        
    def run(self):
        """Run the task function"""
        try:
            if self._stop_requested:
                return
                
            # Pass update_progress sebagai callback jika task function membutuhkannya
            if 'progress_callback' in self.kwargs:
                self.kwargs['progress_callback'] = self.update_progress
            elif len(self.args) > 0 and hasattr(self.args[0], '__call__'):
                # Jika args pertama adalah function, tambahkan progress_callback
                result = self.task_function(*self.args, progress_callback=self.update_progress, **self.kwargs)
            else:
                result = self.task_function(*self.args, **self.kwargs)
            
            if not self._stop_requested:
                self.task_completed.emit(True, str(result) if result else "Task completed successfully")
        except Exception as e:
            if not self._stop_requested:
                self.task_completed.emit(False, str(e))

class LoadingScreen(QWidget):
    """Modern loading screen dengan progress bar tanpa logo"""
    
    def __init__(self, parent=None, title="Loading...", message="Please wait while we process your request"):
        super().__init__(parent)
        self.title = title
        self.message = message
        self.worker = None
        self._is_closing = False
        
        self.setupUI()
        self.setupAnimations()
        
    def setupUI(self):
        """Setup UI loading screen tanpa logo"""
        # Set window properties
        self.setWindowTitle("Loading")
        self.setFixedSize(500, 300)  # Ukuran lebih kecil tanpa logo
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Main container dengan border radius
        self.container = QFrame()
        self.container.setObjectName("loadingContainer")
        self.container.setStyleSheet(f"""
            #loadingContainer {{
                background-color: white;
                border-radius: 15px;
                border: 2px solid {PRIMARY_COLOR};
            }}
        """)
        
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(self.container)
        
        # Container layout
        container_layout = QVBoxLayout(self.container)
        container_layout.setContentsMargins(40, 40, 40, 40)
        container_layout.setSpacing(25)
        container_layout.setAlignment(Qt.AlignCenter)
         
        # Title
        self.title_label = QLabel(self.title)
        self.title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        self.title_label.setStyleSheet(f"color: {PRIMARY_COLOR};")
        self.title_label.setAlignment(Qt.AlignCenter)
        container_layout.addWidget(self.title_label)
        
        # Message
        self.message_label = QLabel(self.message)
        self.message_label.setFont(QFont("Segoe UI", 12))
        self.message_label.setStyleSheet("color: #666;")
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setWordWrap(True)
        container_layout.addWidget(self.message_label)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(25)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid #E0E0E0;
                border-radius: 12px;
                background-color: #F5F5F5;
                text-align: center;
                font-weight: bold;
                color: {PRIMARY_COLOR};
            }}
            QProgressBar::chunk {{
                background-color: {SECONDARY_COLOR};
                border-radius: 10px;
                margin: 2px;
            }}
        """)
        container_layout.addWidget(self.progress_bar)
        
        # Status text
        self.status_label = QLabel("Initializing...")
        self.status_label.setFont(QFont("Segoe UI", 10))
        self.status_label.setStyleSheet("color: #888;")
        self.status_label.setAlignment(Qt.AlignCenter)
        container_layout.addWidget(self.status_label)
        
        # Center the window
        self.center_on_screen()
        
    def setupAnimations(self):
        """Setup animasi untuk loading screen tanpa logo"""
        # Timer untuk simulasi progress jika tidak ada worker
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.simulate_progress)
        self.current_progress = 0
        
    def center_on_screen(self):
        """Center window di tengah layar"""
        screen = QApplication.desktop().screenGeometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)
        
    def start_loading(self, task_function=None, *args, **kwargs):
        """Start loading process"""
        if self._is_closing:
            return
                 
        if task_function:
            # Jalankan task di background thread
            self.worker = LoadingWorker(task_function, *args, **kwargs)
            self.worker.progress_updated.connect(self.update_progress)
            self.worker.task_completed.connect(self.on_task_completed)
            self.worker.start()
        else:
            # Simulasi progress jika tidak ada task
            self.simulate_progress_start()
            
    def simulate_progress_start(self):
        """Start simulasi progress"""
        if self._is_closing:
            return
            
        self.current_progress = 0
        self.progress_timer.start(100)  # Update setiap 100ms
        
    def simulate_progress(self):
        """Simulasi progress untuk demo"""
        if self._is_closing:
            self.progress_timer.stop()
            return
            
        if self.current_progress < 90:
            # Progress lambat di awal, cepat di tengah, lambat di akhir
            if self.current_progress < 20:
                increment = 1
            elif self.current_progress < 70:
                increment = 3
            else:
                increment = 1
                
            self.current_progress += increment
            self.update_progress(self.current_progress, f"Processing... {self.current_progress}%")
        else:
            self.progress_timer.stop()
            
    def update_progress(self, percentage, message=""):
        """Update progress bar dan status"""
        if self._is_closing:
            return
            
        self.progress_bar.setValue(min(percentage, 100))
        if message:
            self.status_label.setText(message)
        
        # Update progress text di progress bar
        self.progress_bar.setFormat(f"{percentage}%")
        
        QApplication.processEvents()
        
    def on_task_completed(self, success, message):
        """Callback ketika task selesai"""
        if self._is_closing:
            return
            
        if success:
            self.update_progress(100, "Completed successfully!")
            QTimer.singleShot(500, self.close_with_success)
        else:
            self.update_progress(100, f"Error: {message}")
            self.progress_bar.setStyleSheet(f"""
                QProgressBar {{
                    border: 2px solid #E74C3C;
                    border-radius: 12px;
                    background-color: #F5F5F5;
                    text-align: center;
                    font-weight: bold;
                    color: #E74C3C;
                }}
                QProgressBar::chunk {{
                    background-color: #E74C3C;
                    border-radius: 10px;
                    margin: 2px;
                }}
            """)
            QTimer.singleShot(2000, self.close_with_error)
            
    def close_with_success(self):
        """Close dengan status success"""
        if self._is_closing:
            return
            
        self._is_closing = True
        self.cleanup_and_close()
        
    def close_with_error(self):
        """Close dengan status error"""
        if self._is_closing:
            return
            
        self._is_closing = True
        self.cleanup_and_close()
        
    def cleanup_and_close(self):
        """Cleanup resources dan tutup loading screen"""
        try:
            # Stop timers
            if self.progress_timer:
                self.progress_timer.stop()
                     
            # Stop worker thread dengan aman
            if self.worker and self.worker.isRunning():
                self.worker.stop()
                self.worker.quit()
                if not self.worker.wait(1000):  # Wait max 1 second
                    self.worker.terminate()
                    self.worker.wait()
                    
            # Close window
            self.close()
            
        except Exception as e:
            print(f"Error during cleanup: {str(e)}")
            self.close()
        
    def set_title(self, title):
        """Update title"""
        if not self._is_closing:
            self.title = title
            self.title_label.setText(title)
        
    def set_message(self, message):
        """Update message"""
        if not self._is_closing:
            self.message = message
            self.message_label.setText(message)
        
    def closeEvent(self, event):
        """Handle close event"""
        if not self._is_closing:
            self._is_closing = True
            
        try:
            # Stop timers
            if self.progress_timer:
                self.progress_timer.stop()
                     
            # Stop worker thread dengan aman
            if self.worker and self.worker.isRunning():
                self.worker.stop()
                self.worker.quit()
                if not self.worker.wait(1000):  # Wait max 1 second
                    self.worker.terminate()
                    self.worker.wait()
                    
        except Exception as e:
            print(f"Error in closeEvent: {str(e)}")
            
        event.accept()

class QuickLoadingDialog(QtWidgets.QDialog):
    """Dialog loading sederhana untuk operasi cepat tanpa logo"""
    
    def __init__(self, parent=None, message="Loading..."):
        super().__init__(parent)
        self.setWindowTitle("Loading")
        self.setFixedSize(300, 120)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 20, 30, 20)
         
        # Message
        message_label = QLabel(message)
        message_label.setFont(QFont("Segoe UI", 11))
        message_label.setAlignment(Qt.AlignCenter)
        message_label.setStyleSheet("color: #555;")
        layout.addWidget(message_label)
        
        # Center di parent
        if parent:
            parent_rect = parent.geometry()
            x = parent_rect.x() + (parent_rect.width() - self.width()) // 2
            y = parent_rect.y() + (parent_rect.height() - self.height()) // 2
            self.move(x, y)
        else:
            self.center_on_screen()
            
    def center_on_screen(self):
        """Center di layar"""
        screen = QApplication.desktop().screenGeometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

# Context manager untuk menggunakan loading screen dengan mudah
class LoadingContext:
    """Context manager untuk loading screen"""
    
    def __init__(self, parent=None, title="Loading...", message="Please wait...", task_function=None, *args, **kwargs):
        self.parent = parent
        self.title = title
        self.message = message
        self.task_function = task_function
        self.args = args
        self.kwargs = kwargs
        self.loading_screen = None
        self.result = None
        
    def __enter__(self):
        self.loading_screen = LoadingScreen(self.parent, self.title, self.message)
        self.loading_screen.show()
        
        if self.task_function:
            self.loading_screen.start_loading(self.task_function, *self.args, **self.kwargs)
        else:
            self.loading_screen.simulate_progress_start()
            
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.loading_screen:
            self.loading_screen.cleanup_and_close()
            
    def update_progress(self, percentage, message=""):
        """Update progress dari context"""
        if self.loading_screen and not self.loading_screen._is_closing:
            self.loading_screen.update_progress(percentage, message)
            
    def update_message(self, message):
        """Update message"""
        if self.loading_screen and not self.loading_screen._is_closing:
            self.loading_screen.set_message(message)

# Helper functions untuk kemudahan penggunaan
def show_loading_dialog(parent, message="Loading...", duration=2000):
    """Show quick loading dialog untuk operasi singkat"""
    dialog = QuickLoadingDialog(parent, message)
    dialog.show()
    QTimer.singleShot(duration, dialog.close)
    return dialog

def run_with_loading(parent, task_function, title="Loading...", message="Please wait...", *args, **kwargs):
    """Run function dengan loading screen"""
    loading = LoadingScreen(parent, title, message)
    loading.show()
    loading.start_loading(task_function, *args, **kwargs)
    
    # Return loading screen agar bisa dimonitor dari pemanggil
    return loading