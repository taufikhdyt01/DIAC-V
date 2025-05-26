# views/loading_utils.py - Utility functions untuk loading screen

import os
import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import QTimer
from views.loading_screen import LoadingScreen, QuickLoadingDialog

class LoadingManager:
    """Manager class untuk mengelola loading screens"""
    
    def __init__(self):
        self.current_loading = None
        self.loading_history = []
    
    def show_loading(self, parent=None, title="Loading...", message="Please wait...", 
                    task_function=None, *args, **kwargs):
        """Show loading screen dengan task function"""
        
        # Close any existing loading screen
        self.close_current_loading()
        
        # Create new loading screen
        self.current_loading = LoadingScreen(parent, title, message)
        self.current_loading.show()
        
        if task_function:
            self.current_loading.start_loading(task_function, *args, **kwargs)
            
            # Track completion
            self.current_loading.worker.task_completed.connect(self._on_task_complete)
        
        return self.current_loading
    
    def show_quick_loading(self, parent=None, message="Loading...", duration=2000):
        """Show quick loading dialog"""
        dialog = QuickLoadingDialog(parent, message)
        dialog.show()
        QTimer.singleShot(duration, dialog.close)
        return dialog
    
    def close_current_loading(self):
        """Close current loading screen"""
        if self.current_loading:
            self.current_loading.close()
            self.current_loading = None
    
    def _on_task_complete(self, success, message):
        """Internal handler for task completion"""
        self.loading_history.append({
            'success': success,
            'message': message,
            'timestamp': QTimer().remainingTime()
        })

# Global loading manager instance
loading_manager = LoadingManager()

# Decorator untuk function yang membutuhkan loading screen
def with_loading_screen(title="Processing...", message="Please wait...", 
                       show_success=True, show_error=True):
    """Decorator untuk menambahkan loading screen ke function"""
    def decorator(func):
        def wrapper(self, *args, **kwargs):
            def task_function(progress_callback=None):
                try:
                    # Call original function
                    if 'progress_callback' in func.__code__.co_varnames:
                        result = func(self, *args, progress_callback=progress_callback, **kwargs)
                    else:
                        result = func(self, *args, **kwargs)
                    
                    if progress_callback:
                        progress_callback(100, "Completed!")
                    
                    return result
                except Exception as e:
                    if progress_callback:
                        progress_callback(100, f"Error: {str(e)}")
                    raise e
            
            # Show loading screen
            loading_screen = LoadingScreen(
                parent=getattr(self, 'window', None) or getattr(self, 'parent', None),
                title=title,
                message=message
            )
            loading_screen.show()
            loading_screen.start_loading(task_function)
            
            # Handle completion
            def on_complete(success, result_message):
                if success and show_success:
                    QMessageBox.information(
                        getattr(self, 'window', None) or getattr(self, 'parent', None),
                        "Success",
                        "Operation completed successfully!"
                    )
                elif not success and show_error:
                    QMessageBox.critical(
                        getattr(self, 'window', None) or getattr(self, 'parent', None),
                        "Error",
                        f"Operation failed: {result_message}"
                    )
            
            loading_screen.worker.task_completed.connect(on_complete)
            return loading_screen
        
        return wrapper
    return decorator

# Utility functions
def run_with_progress(parent, task_func, title="Loading...", message="Please wait...", 
                     success_callback=None, error_callback=None, *args, **kwargs):
    """Run a function with progress loading screen"""
    
    def wrapped_task(progress_callback=None):
        try:
            # Check if task function accepts progress_callback
            import inspect
            sig = inspect.signature(task_func)
            if 'progress_callback' in sig.parameters:
                return task_func(*args, progress_callback=progress_callback, **kwargs)
            else:
                return task_func(*args, **kwargs)
        except Exception as e:
            if progress_callback:
                progress_callback(100, f"Error: {str(e)}")
            raise e
    
    loading_screen = LoadingScreen(parent, title, message)
    loading_screen.show()
    loading_screen.start_loading(wrapped_task)
    
    def on_complete(success, message):
        if success and success_callback:
            success_callback(message)
        elif not success and error_callback:
            error_callback(message)
        elif success:
            QMessageBox.information(parent, "Success", "Operation completed successfully!")
        else:
            QMessageBox.critical(parent, "Error", f"Operation failed: {message}")
    
    loading_screen.worker.task_completed.connect(on_complete)
    return loading_screen

def show_file_operation_loading(parent, operation_type, file_path=""):
    """Show loading for common file operations"""
    messages = {
        'save': f"Saving file{f': {os.path.basename(file_path)}' if file_path else ''}...",
        'load': f"Loading file{f': {os.path.basename(file_path)}' if file_path else ''}...",
        'export': f"Exporting data{f' to {os.path.basename(file_path)}' if file_path else ''}...",
        'import': f"Importing data{f' from {os.path.basename(file_path)}' if file_path else ''}...",
        'backup': f"Creating backup{f' of {os.path.basename(file_path)}' if file_path else ''}...",
        'calculate': "Calculating formulas and updating data..."
    }
    
    titles = {
        'save': "Saving Data",
        'load': "Loading Data", 
        'export': "Exporting Data",
        'import': "Importing Data",
        'backup': "Creating Backup",
        'calculate': "Updating Calculations"
    }
    
    title = titles.get(operation_type, "Processing...")
    message = messages.get(operation_type, "Please wait while operation completes...")
    
    return loading_manager.show_loading(parent, title, message)

def show_network_loading(parent, operation="Connecting"):
    """Show loading for network operations"""
    return loading_manager.show_loading(
        parent,
        f"{operation}...",
        "Please wait while we connect to the server..."
    )

def show_calculation_loading(parent, calculation_type="general"):
    """Show loading for calculation operations"""
    messages = {
        'general': "Performing calculations...",
        'excel': "Calculating Excel formulas...",
        'projection': "Running projection calculations...",
        'analysis': "Analyzing data...",
        'validation': "Validating data integrity..."
    }
    
    message = messages.get(calculation_type, "Performing calculations...")
    
    return loading_manager.show_loading(
        parent,
        "Calculating",
        message
    )

# Context managers for easier usage
class LoadingContext:
    """Context manager for loading screen"""
    
    def __init__(self, parent=None, title="Loading...", message="Please wait..."):
        self.parent = parent
        self.title = title
        self.message = message
        self.loading_screen = None
    
    def __enter__(self):
        self.loading_screen = LoadingScreen(self.parent, self.title, self.message)
        self.loading_screen.show()
        self.loading_screen.simulate_progress_start()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.loading_screen:
            self.loading_screen.close()
    
    def update_progress(self, percentage, message=""):
        """Update progress from context"""
        if self.loading_screen:
            self.loading_screen.update_progress(percentage, message)
    
    def update_message(self, message):
        """Update message"""
        if self.loading_screen:
            self.loading_screen.set_message(message)

class QuickLoadingContext:
    """Context manager for quick loading operations"""
    
    def __init__(self, parent=None, message="Loading...", min_duration=1000):
        self.parent = parent
        self.message = message
        self.min_duration = min_duration
        self.dialog = None
        self.start_time = None
    
    def __enter__(self):
        import time
        self.start_time = time.time()
        self.dialog = QuickLoadingDialog(self.parent, self.message)
        self.dialog.show()
        QApplication.processEvents()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.dialog:
            import time
            elapsed = (time.time() - self.start_time) * 1000
            
            if elapsed < self.min_duration:
                # Ensure minimum duration
                remaining = self.min_duration - elapsed
                QTimer.singleShot(int(remaining), self.dialog.close)
            else:
                self.dialog.close()

# Batch operation helpers
class BatchOperationLoader:
    """Helper for showing progress on batch operations"""
    
    def __init__(self, parent, title="Processing", message="Processing items..."):
        self.parent = parent
        self.title = title
        self.message = message
        self.loading_screen = None
        self.total_items = 0
        self.completed_items = 0
    
    def start(self, total_items):
        """Start batch operation"""
        self.total_items = total_items
        self.completed_items = 0
        
        self.loading_screen = LoadingScreen(
            self.parent, 
            self.title, 
            f"{self.message} (0/{total_items})"
        )
        self.loading_screen.show()
        self.loading_screen.update_progress(0, "Starting batch operation...")
    
    def update(self, completed_items, current_item_name=""):
        """Update progress"""
        self.completed_items = completed_items
        
        if self.loading_screen:
            percentage = int((completed_items / self.total_items) * 100) if self.total_items > 0 else 0
            status_message = f"Processing item {completed_items}/{self.total_items}"
            
            if current_item_name:
                status_message += f": {current_item_name}"
            
            self.loading_screen.update_progress(percentage, status_message)
            self.loading_screen.set_message(f"{self.message} ({completed_items}/{self.total_items})")
    
    def finish(self, success=True, message=""):
        """Finish batch operation"""
        if self.loading_screen:
            if success:
                self.loading_screen.update_progress(100, f"Completed! Processed {self.completed_items} items.")
                QTimer.singleShot(1000, self.loading_screen.close)
            else:
                self.loading_screen.update_progress(100, f"Error: {message}")
                QTimer.singleShot(3000, self.loading_screen.close)

# Example usage functions
def example_file_save_with_loading(parent, file_path, data):
    """Example: Save file with loading screen"""
    
    def save_file_task(progress_callback=None):
        import time
        
        if progress_callback:
            progress_callback(10, "Preparing data...")
        time.sleep(0.5)  # Simulate work
        
        if progress_callback:
            progress_callback(40, f"Writing to {os.path.basename(file_path)}...")
        time.sleep(1.0)  # Simulate file write
        
        if progress_callback:
            progress_callback(80, "Finalizing...")
        time.sleep(0.3)  # Simulate finalization
        
        if progress_callback:
            progress_callback(100, "File saved successfully!")
        
        return f"File saved to {file_path}"
    
    return run_with_progress(
        parent, 
        save_file_task,
        title="Saving File",
        message=f"Saving data to {os.path.basename(file_path)}..."
    )

def example_batch_processing_with_loading(parent, items):
    """Example: Process multiple items with batch loader"""
    
    batch_loader = BatchOperationLoader(
        parent,
        "Processing Items",
        "Processing data items..."
    )
    
    def process_batch():
        batch_loader.start(len(items))
        
        for i, item in enumerate(items):
            # Simulate processing
            import time
            time.sleep(0.1)
            
            batch_loader.update(i + 1, str(item))
            QApplication.processEvents()
        
        batch_loader.finish(success=True)
    
    # Run batch processing in a timer to avoid blocking UI
    QTimer.singleShot(100, process_batch)

# Error handling helpers
def safe_run_with_loading(parent, task_func, title="Processing", message="Please wait...", 
                         *args, **kwargs):
    """Safely run task with loading screen and proper error handling"""
    
    def safe_task(progress_callback=None):
        try:
            if progress_callback:
                progress_callback(5, "Initializing...")
            
            result = task_func(*args, **kwargs)
            
            if progress_callback:
                progress_callback(100, "Completed successfully!")
            
            return result
        except FileNotFoundError as e:
            error_msg = f"File not found: {str(e)}"
            if progress_callback:
                progress_callback(100, f"Error: {error_msg}")
            return error_msg
        except PermissionError as e:
            error_msg = f"Permission denied: {str(e)}"
            if progress_callback:
                progress_callback(100, f"Error: {error_msg}")
            return error_msg
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            if progress_callback:
                progress_callback(100, f"Error: {error_msg}")
            return error_msg
    
    return run_with_progress(parent, safe_task, title, message)