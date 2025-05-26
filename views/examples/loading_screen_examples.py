# examples/loading_screen_examples.py - Contoh praktis penggunaan loading screen

"""
PANDUAN IMPLEMENTASI LOADING SCREEN UNTUK DIAC-V

File ini berisi contoh-contoh praktis penggunaan loading screen
dalam berbagai skenario yang umum terjadi di aplikasi ERP.
"""

import os
import sys
import time
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QMessageBox
from PyQt5.QtCore import QTimer

# Import loading screen components
from views.loading_screen import LoadingScreen, LoadingContext, QuickLoadingDialog
from views.loading_utils import (
    run_with_progress, 
    show_file_operation_loading,
    BatchOperationLoader,
    with_loading_screen,
    LoadingContext,
    safe_run_with_loading
)

class ExampleMainWindow(QMainWindow):
    """Contoh window dengan berbagai implementasi loading screen"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loading Screen Examples - DIAC-V")
        self.setGeometry(100, 100, 800, 600)
        
        # Setup UI
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Example buttons
        examples = [
            ("1. File Loading Example", self.example_file_loading),
            ("2. Excel Processing Example", self.example_excel_processing),
            ("3. Batch Operation Example", self.example_batch_operation),
            ("4. Network Operation Example", self.example_network_operation),
            ("5. Using Context Manager", self.example_context_manager),
            ("6. Using Decorator", self.example_decorator_usage),
            ("7. Quick Loading Dialog", self.example_quick_loading),
            ("8. Error Handling Example", self.example_error_handling),
            ("9. Customer File Creation", self.example_customer_file_creation),
            ("10. Excel Calculation", self.example_excel_calculation)
        ]
        
        for title, handler in examples:
            btn = QPushButton(title)
            btn.clicked.connect(handler)
            layout.addWidget(btn)

    # EXAMPLE 1: File Loading dengan Progress Detail
    def example_file_loading(self):
        """Contoh loading file besar dengan progress detail"""
        
        def load_large_file(progress_callback=None):
            """Simulasi loading file besar"""
            file_path = "large_database.xlsx"
            
            if progress_callback:
                progress_callback(10, f"Opening {file_path}...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(30, "Reading headers...")
            time.sleep(0.8)
            
            if progress_callback:
                progress_callback(60, "Processing data rows...")
            time.sleep(1.2)
            
            if progress_callback:
                progress_callback(80, "Validating data...")
            time.sleep(0.6)
            
            if progress_callback:
                progress_callback(95, "Finalizing...")
            time.sleep(0.3)
            
            if progress_callback:
                progress_callback(100, "File loaded successfully!")
            
            return f"Successfully loaded {file_path} with 10,000 records"
        
        # Show loading screen
        loading_screen = LoadingScreen(
            parent=self,
            title="Loading Database",
            message="Loading large Excel database file..."
        )
        loading_screen.show()
        loading_screen.start_loading(load_large_file)
        
        # Handle completion
        def on_complete(success, message):
            if success:
                QMessageBox.information(self, "Success", message)
            else:
                QMessageBox.critical(self, "Error", f"Failed to load file: {message}")
        
        loading_screen.worker.task_completed.connect(on_complete)

    # EXAMPLE 2: Excel Processing seperti di BDU
    def example_excel_processing(self):
        """Contoh processing Excel seperti di modul BDU"""
        
        def process_excel_data(progress_callback=None):
            """Simulasi processing Excel yang kompleks"""
            steps = [
                (5, "Opening Excel workbook..."),
                (15, "Reading customer data sheet..."),
                (25, "Processing DIP_Technical Information..."),
                (35, "Validating input data..."),
                (45, "Calculating formulas..."),
                (55, "Transferring data to ANAPAK..."),
                (65, "Running ANAPAK calculations..."),
                (75, "Processing PUMP module..."),
                (85, "Generating output data..."),
                (95, "Saving results..."),
                (100, "Excel processing completed!")
            ]
            
            for percentage, message in steps:
                if progress_callback:
                    progress_callback(percentage, message)
                time.sleep(0.3)  # Simulate work
            
            return "Excel data processed successfully with 15 sheets updated"
        
        loading_screen = LoadingScreen(
            parent=self,
            title="Processing Excel Data",
            message="Running calculations and updating formulas..."
        )
        loading_screen.show()
        loading_screen.start_loading(process_excel_data)
        
        loading_screen.worker.task_completed.connect(
            lambda success, msg: QMessageBox.information(self, "Complete", msg) if success 
            else QMessageBox.critical(self, "Error", msg)
        )

    # EXAMPLE 3: Batch Operation
    def example_batch_operation(self):
        """Contoh pemrosesan batch seperti multiple customers"""
        
        # Simulate customer list
        customers = [f"Customer {i+1}" for i in range(20)]
        
        batch_loader = BatchOperationLoader(
            self,
            "Processing Customers",
            "Updating customer data files..."
        )
        
        def process_batch():
            batch_loader.start(len(customers))
            
            for i, customer in enumerate(customers):
                # Simulate processing each customer
                time.sleep(0.2)
                
                batch_loader.update(i + 1, customer)
                QApplication.processEvents()
                
                # Simulate occasional error
                if i == 10:  # Error on customer 11
                    batch_loader.finish(success=False, message="Failed to process Customer 11")
                    return
            
            batch_loader.finish(success=True)
            QMessageBox.information(self, "Batch Complete", f"Successfully processed {len(customers)} customers!")
        
        # Start batch processing
        QTimer.singleShot(100, process_batch)

    # EXAMPLE 4: Network Operation
    def example_network_operation(self):
        """Contoh operasi network/API"""
        
        def simulate_api_call(progress_callback=None):
            """Simulasi API call ke server"""
            if progress_callback:
                progress_callback(10, "Connecting to server...")
            time.sleep(0.8)
            
            if progress_callback:
                progress_callback(30, "Authenticating...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(50, "Uploading data...")
            time.sleep(1.0)
            
            if progress_callback:
                progress_callback(80, "Waiting for server response...")
            time.sleep(0.7)
            
            if progress_callback:
                progress_callback(100, "Data synchronized!")
            
            return "Successfully synchronized with server"
        
        run_with_progress(
            self,
            simulate_api_call,
            title="Synchronizing Data",
            message="Connecting to server and syncing data..."
        )

    # EXAMPLE 5: Context Manager Usage
    def example_context_manager(self):
        """Contoh penggunaan context manager"""
        
        def process_with_context():
            with LoadingContext(self, "Processing with Context", "Using context manager...") as loading:
                loading.update_progress(10, "Starting process...")
                time.sleep(0.5)
                
                loading.update_progress(30, "Processing data...")
                time.sleep(0.8)
                
                loading.update_progress(60, "Validating results...")
                time.sleep(0.6)
                
                loading.update_progress(90, "Finalizing...")
                time.sleep(0.4)
                
                loading.update_progress(100, "Process completed!")
                time.sleep(0.3)
            
            QMessageBox.information(self, "Context Manager", "Process completed using context manager!")
        
        # Run in timer to avoid blocking
        QTimer.singleShot(100, process_with_context)

    # EXAMPLE 6: Decorator Usage
    @with_loading_screen(title="Decorator Example", message="Processing with decorator...")
    def example_decorator_usage(self, progress_callback=None):
        """Contoh penggunaan decorator"""
        steps = [
            "Initializing process...",
            "Loading configuration...",
            "Processing data...",
            "Saving results...",
            "Cleaning up..."
        ]
        
        for i, step in enumerate(steps):
            if progress_callback:
                progress_callback((i + 1) * 20, step)
            time.sleep(0.4)
        
        return "Decorator example completed successfully!"

    # EXAMPLE 7: Quick Loading Dialog
    def example_quick_loading(self):
        """Contoh quick loading dialog untuk operasi cepat"""
        
        def quick_operation():
            dialog = QuickLoadingDialog(self, "Saving preferences...")
            dialog.show()
            
            # Simulate quick operation
            QTimer.singleShot(1500, dialog.close)
            QTimer.singleShot(1600, lambda: QMessageBox.information(self, "Quick Load", "Preferences saved!"))
        
        quick_operation()

    # EXAMPLE 8: Error Handling
    def example_error_handling(self):
        """Contoh error handling dalam loading screen"""
        
        def operation_with_error(progress_callback=None):
            """Operasi yang akan mengalami error"""
            if progress_callback:
                progress_callback(20, "Starting operation...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(50, "Processing data...")
            time.sleep(0.8)
            
            # Simulate error
            raise FileNotFoundError("Required file 'config.xml' not found")
        
        loading_screen = safe_run_with_loading(
            self,
            operation_with_error,
            title="Error Handling Example",
            message="This operation will encounter an error..."
        )

    # EXAMPLE 9: Customer File Creation (Real DIAC-V scenario)
    def example_customer_file_creation(self):
        """Contoh pembuatan file customer seperti di DIAC-V"""
        
        def create_customer_file(progress_callback=None):
            """Simulasi pembuatan file customer baru"""
            customer_name = "PT. Example Company"
            
            if progress_callback:
                progress_callback(10, f"Creating folder for {customer_name}...")
            time.sleep(0.4)
            
            if progress_callback:
                progress_callback(25, "Copying template file...")
            time.sleep(0.6)
            
            if progress_callback:
                progress_callback(45, "Reading customer data from database...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(65, "Updating customer information in Excel...")
            time.sleep(0.7)
            
            if progress_callback:
                progress_callback(80, "Applying company-specific settings...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(95, "Finalizing customer workspace...")
            time.sleep(0.3)
            
            if progress_callback:
                progress_callback(100, "Customer file created successfully!")
            
            return f"Successfully created workspace for {customer_name}"
        
        loading_screen = LoadingScreen(
            parent=self,
            title="Creating Customer Workspace",
            message="Setting up new customer environment..."
        )
        loading_screen.show()
        loading_screen.start_loading(create_customer_file)
        
        loading_screen.worker.task_completed.connect(
            lambda success, msg: QMessageBox.information(self, "Customer Created", msg) if success
            else QMessageBox.critical(self, "Creation Failed", msg)
        )

    # EXAMPLE 10: Excel Calculation (Real DIAC-V scenario)
    def example_excel_calculation(self):
        """Contoh kalkulasi Excel seperti di modul BDU"""
        
        def excel_calculation(progress_callback=None):
            """Simulasi kalkulasi Excel yang kompleks"""
            if progress_callback:
                progress_callback(5, "Opening Excel application...")
            time.sleep(0.6)
            
            if progress_callback:
                progress_callback(15, "Loading workbook...")
            time.sleep(0.5)
            
            if progress_callback:
                progress_callback(30, "Forcing formula recalculation...")
            time.sleep(1.2)
            
            if progress_callback:
                progress_callback(50, "Calculating dependent cells...")
            time.sleep(0.8)
            
            if progress_callback:
                progress_callback(70, "Updating linked worksheets...")
            time.sleep(0.9)
            
            if progress_callback:
                progress_callback(85, "Saving calculated values...")
            time.sleep(0.6)
            
            if progress_callback:
                progress_callback(95, "Closing Excel application...")
            time.sleep(0.4)
            
            if progress_callback:
                progress_callback(100, "Excel calculation completed!")
            
            return "All formulas recalculated and saved successfully"
        
        loading_screen = LoadingScreen(
            parent=self,
            title="Updating Excel Calculations",
            message="Recalculating formulas and updating data..."
        )
        loading_screen.show()
        loading_screen.start_loading(excel_calculation)
        
        loading_screen.worker.task_completed.connect(
            lambda success, msg: QMessageBox.information(self, "Calculation Complete", msg) if success
            else QMessageBox.critical(self, "Calculation Failed", msg)
        )

# QUICK INTEGRATION GUIDE untuk aplikasi yang sudah ada
"""
CARA CEPAT MENAMBAHKAN LOADING SCREEN KE FUNCTION YANG SUDAH ADA:

1. UNTUK FUNCTION SEDERHANA:
   
   Original code:
   def save_data(self):
       # ... lakukan saving ...
       QMessageBox.information(self, "Success", "Data saved!")
   
   Dengan loading screen:
   def save_data(self):
       def save_process(progress_callback=None):
           if progress_callback:
               progress_callback(50, "Saving data...")
           # ... lakukan saving ...
           if progress_callback:
               progress_callback(100, "Data saved!")
           return "Data saved successfully"
       
       loading_screen = LoadingScreen(self, "Saving", "Saving data...")
       loading_screen.show()
       loading_screen.start_loading(save_process)
       loading_screen.worker.task_completed.connect(
           lambda success, msg: QMessageBox.information(self, "Success", msg) if success 
           else QMessageBox.critical(self, "Error", msg)
       )

2. UNTUK FUNCTION DENGAN BANYAK STEPS:

   def complex_operation(self):
       def operation_process(progress_callback=None):
           steps = [
               (20, "Step 1: Preparing..."),
               (40, "Step 2: Processing..."),
               (60, "Step 3: Validating..."),
               (80, "Step 4: Saving..."),
               (100, "Completed!")
           ]
           
           for percentage, message in steps:
               if progress_callback:
                   progress_callback(percentage, message)
               # ... do actual work ...
               time.sleep(0.5)  # Replace with real work
           
           return "Operation completed successfully"
       
       run_with_progress(self, operation_process, "Processing", "Please wait...")

3. UNTUK BATCH OPERATIONS:

   def process_multiple_items(self, items):
       batch_loader = BatchOperationLoader(self, "Processing", "Processing items...")
       
       def process_batch():
           batch_loader.start(len(items))
           
           for i, item in enumerate(items):
               # Process each item
               process_single_item(item)
               batch_loader.update(i + 1, str(item))
               QApplication.processEvents()
           
           batch_loader.finish(success=True)
       
       QTimer.singleShot(100, process_batch)

4. MENGGUNAKAN CONTEXT MANAGER:

   def operation_with_context(self):
       with LoadingContext(self, "Processing", "Working...") as loading:
           loading.update_progress(25, "Starting...")
           # ... do work ...
           loading.update_progress(75, "Almost done...")
           # ... finish work ...
           loading.update_progress(100, "Complete!")

5. MENGGUNAKAN DECORATOR:

   @with_loading_screen("Processing Data", "Please wait...")
   def decorated_function(self, progress_callback=None):
       if progress_callback:
           progress_callback(50, "Working...")
       # ... do work ...
       return "Success!"
"""

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExampleMainWindow()
    window.show()
    sys.exit(app.exec_())