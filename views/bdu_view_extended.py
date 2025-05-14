import os
import sys
from .bdu_view import BDUGroupView as OriginalBDUGroupView

class BDUGroupView(OriginalBDUGroupView):
    """Extended BDU Group View with customer-specific file handling"""
    
    def __init__(self, auth_manager, customer_name=None):
        # Store customer name
        self.customer_name = customer_name
        
        # Call parent's __init__
        super().__init__(auth_manager)
        
        # If a customer name is provided, set up the customer-specific file
        if customer_name:
            self.setup_customer_file()
    
    def setup_customer_file(self):
        """Set up the customer-specific file path and load it"""
        if not self.customer_name:
            return
        
        # Get base path for customer folders
        base_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "data", "customers")
        
        # Import the clean_folder_name function to handle special characters
        from modules.fix_customer_system import clean_folder_name
        
        # Create a valid folder name from the customer name
        valid_folder_name = clean_folder_name(self.customer_name)
        customer_folder = os.path.join(base_path, valid_folder_name)
        
        # Check if the folder exists
        if not os.path.exists(customer_folder):
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(
                self,
                "Error",
                f"Customer folder not found: {customer_folder}"
            )
            return
        
        # Set the Excel path to the customer's file
        customer_file = os.path.join(customer_folder, "SET_BDU.xlsx")
        
        # Check if the file exists
        if not os.path.exists(customer_file):
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.critical(
                self,
                "Error",
                f"Customer BDU file not found: {customer_file}"
            )
            return
        
        # Update the Excel path
        self.excel_path = customer_file
        
        # Set window title to include customer name
        from config import APP_NAME
        self.setWindowTitle(f"BDU Group - {self.customer_name} - {APP_NAME}")
        
        # Update status bar
        self.statusBar().showMessage(f"BDU Group Module | Customer: {self.customer_name} | User: {self.current_user['username']}")
        
        # Load the customer-specific Excel data
        self.load_excel_data()
    
    def set_current_customer(self, customer_name):
        """Set the current customer and reload data"""
        self.customer_name = customer_name
        self.setup_customer_file()
        
        # Update UI elements to show customer name
        title_text = f"BDU Group - {customer_name}"
        
        # Find the page title label and update it
        from PyQt5.QtWidgets import QLabel
        for widget in self.findChildren(QLabel):
            if hasattr(widget, 'text') and widget.text() == "BDU Group":
                widget.setText(title_text)
                break