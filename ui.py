import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLineEdit, QHBoxLayout,
    QLabel, QMessageBox, QListWidget, QDialog, QScrollArea, QFormLayout, QDialogButtonBox,
    QFileDialog, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt
from datetime import datetime
from grant_management import GrantManagement
from excel_handler import ExcelHandler  # Import ExcelHandler

class GrantManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Grant Management System")
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("background-color: #cce7ff;")  # Light blue background

        # Initialize GrantManagement
        self.grant_management = GrantManagement()

        # Initialize ExcelHandler
        self.excel_handler = ExcelHandler(self, self.grant_management)

        # Main layout
        layout = QVBoxLayout()

        # Title labels
        title_label = QLabel("Grant Management System")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #333;")
        layout.addWidget(title_label)

        subtitle_label = QLabel("Faltas Lab")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #555; margin-top: -10px;")
        layout.addWidget(subtitle_label)

        # Create the timestamp label before calling update_timestamp
        self.timestamp_label = QLabel()
        self.timestamp_label.setAlignment(Qt.AlignCenter)
        self.timestamp_label.setStyleSheet("font-size: 14px; color: #777;")
        layout.addWidget(self.timestamp_label)

        # Display last accessed/updated date and time
        self.update_timestamp()

        # Buttons
        button_style = """
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """

        add_grant_btn = QPushButton("Add New Grant")
        add_grant_btn.setStyleSheet(button_style)
        add_grant_btn.clicked.connect(self.add_initial_grants)
        layout.addWidget(add_grant_btn)

        show_btn = QPushButton("Show Grants")
        show_btn.setStyleSheet(button_style)
        show_btn.clicked.connect(self.show_grants)
        layout.addWidget(show_btn)

        upload_btn = QPushButton("Upload Inventory Excel File")
        upload_btn.setStyleSheet(button_style)
        upload_btn.clicked.connect(self.excel_handler.upload_excel)  # Linked to ExcelHandler's upload_excel method
        layout.addWidget(upload_btn)

        display_saved_files_btn = QPushButton("Display Saved Excel Files")
        display_saved_files_btn.setStyleSheet(button_style)
        display_saved_files_btn.clicked.connect(self.excel_handler.display_saved_files)  # Linked to ExcelHandler's display_saved_files method
        layout.addWidget(display_saved_files_btn)

        add_rule_btn = QPushButton("Add Spending Rule")
        add_rule_btn.setStyleSheet(button_style)
        add_rule_btn.clicked.connect(self.choose_grant_for_rule)
        layout.addWidget(add_rule_btn)

        # Central widget
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def update_timestamp(self):
        """Update the last accessed/updated timestamp to 12-hour format."""
        self.last_accessed = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        self.timestamp_label.setText(f"Last accessed/updated: {self.last_accessed}")

    def show_grants(self):
        if self.grant_management.grant_data.empty:
            QMessageBox.information(self, "No Grants", "There are no grants in the database.")
        else:
            self.display_grants_popup()

    def display_grants_popup(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Existing Grants")
        dialog.setStyleSheet("background-color: #cce7ff;")
        dialog.resize(800, 600)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        widget = QWidget()
        vbox = QVBoxLayout(widget)

        for index, row in self.grant_management.grant_data.iterrows():
            grant_id_label = QLabel(f"Grant ID: {row['Grant ID']}")
            grant_id_label.setStyleSheet("font-size: 16px; color: #333;")
            vbox.addWidget(grant_id_label)

            grant_name_label = QLabel(f"Grant Name: {row['Grant Name']}")
            grant_name_label.setStyleSheet("font-size: 16px; color: #333;")
            vbox.addWidget(grant_name_label)

            total_balance_label = QLabel(f"Total Balance: ${row['Total Balance']:.2f}")
            total_balance_label.setStyleSheet("font-size: 16px; color: #333;")
            vbox.addWidget(total_balance_label)

            spending_rules_label = QLabel("Grant Spending Rules:")
            spending_rules_label.setStyleSheet("font-size: 16px; color: #333; font-weight: bold; margin-bottom: 5px;")
            vbox.addWidget(spending_rules_label)

            for item in row['Allowed Items']:
                item_label = QLabel(item)
                item_label.setStyleSheet(
                    "font-size: 16px; color: #333; background-color: #f9f9f9; "
                    "border: 2px solid #007BFF; padding: 5px; margin-bottom: 3px; border-radius: 5px;")
                vbox.addWidget(item_label)

            vbox.addWidget(QLabel("\n"))  # Add spacing between entries

        scroll.setWidget(widget)

        layout = QVBoxLayout(dialog)
        layout.addWidget(scroll)

        dialog.setLayout(layout)
        dialog.exec_()

    def add_initial_grants(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Grant")
        dialog.setStyleSheet("background-color: #cce7ff;")

        layout = QVBoxLayout()

        form_layout = QFormLayout()

        grant_id_input = QLineEdit()
        grant_id_input.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        form_layout.addRow("Grant ID:", grant_id_input)

        grant_name_input = QLineEdit()
        grant_name_input.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        form_layout.addRow("Grant Name:", grant_name_input)

        total_balance_input = QLineEdit()
        total_balance_input.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        form_layout.addRow("Total Balance:", total_balance_input)

        # Section for adding multiple allowed items
        allowed_items_layout = QVBoxLayout()
        allowed_items_label = QLabel("Allowed Items:")
        allowed_items_label.setStyleSheet("font-size: 16px; color: black;")
        allowed_items_layout.addWidget(allowed_items_label)

        add_item_input = QLineEdit()
        add_item_input.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        allowed_items_layout.addWidget(add_item_input)

        add_item_btn = QPushButton("Add Item")
        add_item_btn.setStyleSheet("""
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 5px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        allowed_items_layout.addWidget(add_item_btn)

        items_list_widget = QListWidget()
        items_list_widget.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        allowed_items_layout.addWidget(items_list_widget)

        add_item_btn.clicked.connect(lambda: self.add_item_to_list(add_item_input, items_list_widget))

        layout.addLayout(form_layout)
        layout.addLayout(allowed_items_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.setStyleSheet("""
            QDialogButtonBox {
                padding: 10px;
            }
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(button_box)

        button_box.accepted.connect(lambda: self.save_grant(dialog, grant_id_input.text(), grant_name_input.text(),
                                                           total_balance_input.text(), items_list_widget))
        button_box.rejected.connect(dialog.reject)
        dialog.exec_()

        self.update_timestamp()  # Update timestamp after saving

    def add_item_to_list(self, add_item_input, items_list_widget):
        item = add_item_input.text().strip()
        if item:
            items_list_widget.addItem(item)
            add_item_input.clear()
        else:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid item.")

    def save_grant(self, dialog, grant_id, grant_name, total_balance, items_list_widget):
        try:
            total_balance = float(total_balance)
            allowed_items_list = [items_list_widget.item(i).text() for i in range(items_list_widget.count())]

            self.grant_management.add_grant(grant_id, grant_name, total_balance, allowed_items_list)

            dialog.accept()
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number for the total balance.")

    def choose_grant_for_rule(self):
        if self.grant_management.grant_data.empty:
            QMessageBox.information(self, "No Grants", "There are no grants in the database.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Choose Grant for Spending Rule")
        dialog.setStyleSheet("background-color: #cce7ff;")

        layout = QVBoxLayout()

        list_widget = QListWidget()
        list_widget.setStyleSheet("font-size: 16px;")

        grant_id_map = {}
        for index, row in self.grant_management.grant_data.iterrows():
            list_text = f"{row['Grant ID']} - {row['Grant Name']}"
            list_widget.addItem(list_text)
            grant_id_map[list_text] = row['Grant ID']

        layout.addWidget(list_widget)

        select_button = QPushButton("Select")
        select_button.setStyleSheet("""
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        select_button.clicked.connect(lambda: self.add_rule(grant_id_map[list_widget.currentItem().text()]) if list_widget.currentItem() else None)
        layout.addWidget(select_button)

        dialog.setLayout(layout)
        dialog.exec_()

    def add_rule(self, grant_id=None):
        """Add spending rules to a selected grant."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Spending Rule")
        dialog.setStyleSheet("background-color: #cce7ff;")

        layout = QVBoxLayout()

        try:
            idx = self.grant_management.grant_data[self.grant_management.grant_data['Grant ID'] == grant_id].index[0]
        except IndexError:
            QMessageBox.critical(self, "Error", "Grant ID not found.")
            dialog.close()
            return

        label = QLabel(f"Grant ID: {grant_id}")
        label.setStyleSheet("font-size: 16px; color: black;")
        layout.addWidget(label)

        label = QLabel("Existing Spending Rules:")
        label.setStyleSheet("font-size: 16px; color: black;")
        layout.addWidget(label)

        rules_list_widget = QListWidget()
        rules_list_widget.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        for item in self.grant_management.grant_data.at[idx, 'Allowed Items']:
            rules_list_widget.addItem(item)
        layout.addWidget(rules_list_widget)

        label = QLabel("Add New Spending Rule:")
        label.setStyleSheet("font-size: 16px; color: black;")
        layout.addWidget(label)

        item_input = QLineEdit()
        item_input.setStyleSheet("font-size: 16px; color: black; background-color: white;")
        layout.addWidget(item_input)

        button_layout = QHBoxLayout()

        add_button = QPushButton("Add")
        add_button.setStyleSheet("""
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        add_button.clicked.connect(lambda: self.add_item_to_rules(rules_list_widget, item_input, idx))
        button_layout.addWidget(add_button)

        remove_button = QPushButton("Remove Selected")
        remove_button.setStyleSheet("""
            QPushButton {
                font-size: 16px;
                background-color: #f44336;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #e53935;
            }
        """)
        remove_button.clicked.connect(lambda: self.remove_selected_rule(rules_list_widget, idx))
        button_layout.addWidget(remove_button)

        layout.addLayout(button_layout)

        save_button = QPushButton("Save Changes")
        save_button.setStyleSheet("""
            QPushButton {
                font-size: 16px;
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        save_button.clicked.connect(lambda: self.save_rules_and_close(dialog))
        layout.addWidget(save_button)

        dialog.setLayout(layout)
        dialog.exec_()

    def add_item_to_rules(self, rules_list_widget, item_input, idx):
        item = item_input.text().strip()
        if item:
            self.grant_management.grant_data.at[idx, 'Allowed Items'].append(item)
            rules_list_widget.addItem(item)
            item_input.clear()
        else:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid item.")

    def remove_selected_rule(self, rules_list_widget, idx):
        selected_items = rules_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select an item to remove.")
            return
        for item in selected_items:
            self.grant_management.grant_data.at[idx, 'Allowed Items'].remove(item.text())
            rules_list_widget.takeItem(rules_list_widget.row(item))

    def save_rules_and_close(self, dialog):
        try:
            self.grant_management.save_grants()
            dialog.accept()
            self.update_timestamp()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            dialog.reject()

    def closeEvent(self, event):
        super().closeEvent(event)
        QApplication.quit()
        sys.exit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GrantManagementApp()
    window.show()
    sys.exit(app.exec_())
