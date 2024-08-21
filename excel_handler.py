import os
import pandas as pd
import re
from PyQt5.QtWidgets import (
    QDialog, QFileDialog, QVBoxLayout, QTabWidget, QTableWidget, QTableWidgetItem, 
    QMessageBox, QLabel, QHBoxLayout, QHeaderView, QDateEdit, QPushButton, QComboBox
)
from PyQt5.QtCore import Qt, QDate

class ExcelHandler:
    def __init__(self, parent, grant_management):
        self.parent = parent
        self.grant_management = grant_management
        self.total_cost = 0
        self.selected_sum_label = None
        self.sheet_data = None
        self.saved_excel_sheets = {}  # Dictionary to store saved Excel sheets

    def upload_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self.parent, "Upload Inventory Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        
        if file_path:
            try:
                # Read all sheets from the Excel file
                excel_data = pd.read_excel(file_path, sheet_name=None)
                
                # Remove empty tabs/sheets
                excel_data = {name: data for name, data in excel_data.items() if not data.empty}

                # Save the uploaded sheets to the dictionary
                self.saved_excel_sheets.update(excel_data)

                # Display the contents
                self.display_excel_contents(self.saved_excel_sheets)
            except Exception as e:
                QMessageBox.critical(self.parent, "Error", f"An error occurred while uploading the Excel file: {str(e)}")

    def display_excel_contents(self, excel_data):
        dialog = QDialog(self.parent)
        dialog.setWindowTitle("Excel File Contents")
        dialog.setStyleSheet("background-color: #cce7ff;")
        dialog.resize(1200, 1000)  # Extended the size of the dialog

        layout = QVBoxLayout()

        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("font-size: 14px;")

        self.total_cost = 0
        self.has_cost_column = False  # Track if any sheet has a cost column

        for sheet_name, sheet_data in excel_data.items():
            # Normalize column names to lowercase
            sheet_data.columns = sheet_data.columns.str.lower()

            sheet_data = sheet_data.fillna("")

            # Store the sheet data for date range filtering and cost allocation
            self.sheet_data = sheet_data

            # Sort by expiration date if present
            if 'expiration date' in sheet_data.columns:
                sheet_data['expiration date'] = pd.to_datetime(sheet_data['expiration date'], errors='coerce')
                sheet_data = sheet_data.sort_values(by='expiration date')

            table_widget = QTableWidget()
            table_widget.setRowCount(sheet_data.shape[0])
            table_widget.setColumnCount(sheet_data.shape[1])
            table_widget.setHorizontalHeaderLabels(sheet_data.columns)

            for i in range(sheet_data.shape[0]):
                for j in range(sheet_data.shape[1]):
                    item = QTableWidgetItem(str(sheet_data.iat[i, j]))
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    table_widget.setItem(i, j, item)

            # Enable editing of column names
            table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            table_widget.horizontalHeader().setSectionsMovable(True)
            table_widget.horizontalHeader().setSectionsClickable(True)

            table_widget.itemClicked.connect(self.update_selected_sum)
            tab_widget.addTab(table_widget, sheet_name)

            # Check if the 'cost' column exists and sum it if it does
            if 'cost' in sheet_data.columns:
                # Remove $ sign and other non-numeric characters before summing
                cleaned_costs = sheet_data['cost'].apply(lambda x: float(re.sub(r'[^\d.]', '', str(x))) if x != '' else 0)
                self.total_cost += cleaned_costs.sum()
                self.has_cost_column = True

        layout.addWidget(tab_widget)

        # Section for displaying total costs and selected sum
        cost_layout = QVBoxLayout()

        if self.has_cost_column:
            self.total_cost_label = QLabel(f"Total Cost of All Items: ${self.total_cost:.2f}")
        else:
            self.total_cost_label = QLabel("Total Cost: 'Cost' column not found.")
        
        self.total_cost_label.setStyleSheet("font-size: 16px; color: black;")
        cost_layout.addWidget(self.total_cost_label)

        self.selected_sum_label = QLabel("Selected Sum: $0.00")
        self.selected_sum_label.setStyleSheet("font-size: 16px; color: black;")
        cost_layout.addWidget(self.selected_sum_label)

        # Date range filtering section
        date_range_layout = QHBoxLayout()

        start_date_label = QLabel("Start Date:")
        start_date_label.setStyleSheet("font-size: 16px; color: black;")
        date_range_layout.addWidget(start_date_label)

        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate().addYears(-1))  # Default to one year ago
        date_range_layout.addWidget(self.start_date_edit)

        end_date_label = QLabel("End Date:")
        end_date_label.setStyleSheet("font-size: 16px; color: black;")
        date_range_layout.addWidget(end_date_label)

        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate())  # Default to today
        date_range_layout.addWidget(self.end_date_edit)

        filter_button = QPushButton("Filter Costs by Date Range")
        filter_button.setStyleSheet("font-size: 16px; color: white; background-color: #4CAF50;")
        filter_button.clicked.connect(self.filter_costs_by_date)
        date_range_layout.addWidget(filter_button)

        cost_layout.addLayout(date_range_layout)

        # Grant allocation section
        grant_layout = QHBoxLayout()

        grant_label = QLabel("Select Grant:")
        grant_label.setStyleSheet("font-size: 16px; color: black;")
        grant_layout.addWidget(grant_label)

        self.grant_combo = QComboBox()
        self.grant_combo.addItems(self.grant_management.get_grant_names())  # Assuming get_grant_names() returns a list of grant names
        grant_layout.addWidget(self.grant_combo)

        allocate_button = QPushButton("Allocate Selected Costs to Grant")
        allocate_button.setStyleSheet("font-size: 16px; color: white; background-color: #4CAF50;")
        allocate_button.clicked.connect(self.allocate_costs_to_grant)
        grant_layout.addWidget(allocate_button)

        self.net_amount_label = QLabel("Net Amount in Grant: $0.00")
        self.net_amount_label.setStyleSheet("font-size: 16px; color: black;")
        grant_layout.addWidget(self.net_amount_label)

        cost_layout.addLayout(grant_layout)

        layout.addLayout(cost_layout)

        dialog.setLayout(layout)
        dialog.exec_()

    def update_selected_sum(self, item):
        """Update the sum of selected costs."""
        selected_sum = 0.0
        for selected_item in item.tableWidget().selectedItems():
            try:
                selected_sum += float(re.sub(r'[^\d.]', '', selected_item.text()))
            except ValueError:
                continue  # Ignore non-numeric cells

        self.selected_sum_label.setText(f"Selected Sum: ${selected_sum:.2f}")

    def filter_costs_by_date(self):
        """Filter costs based on the selected date range."""
        if self.sheet_data is None or 'expiration date' not in self.sheet_data.columns:
            QMessageBox.warning(self.parent, "No Date Data", "No 'Expiration Date' column found for filtering.")
            return

        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()

        filtered_data = self.sheet_data[
            (self.sheet_data['expiration date'] >= pd.to_datetime(start_date)) &
            (self.sheet_data['expiration date'] <= pd.to_datetime(end_date))
        ]

        if 'cost' in filtered_data.columns:
            cleaned_costs = filtered_data['cost'].apply(lambda x: float(re.sub(r'[^\d.]', '', str(x))) if x != '' else 0)
            total_filtered_cost = cleaned_costs.sum()
            QMessageBox.information(self.parent, "Filtered Costs", f"Total Costs in Date Range: ${total_filtered_cost:.2f}")
        else:
            QMessageBox.warning(self.parent, "Cost Column Missing", "'Cost' column not found in the filtered data.")

    def allocate_costs_to_grant(self):
        """Allocate the selected costs to the selected grant."""
        selected_grant = self.grant_combo.currentText()
        selected_sum = float(re.sub(r'[^\d.]', '', self.selected_sum_label.text().split('$')[1]))

        grant_data = self.grant_management.get_grant_data(selected_grant)

        if grant_data:
            total_grant_amount = grant_data['Total Balance']
            updated_cost = grant_data.get('Allocated Costs', 0).iloc[0] + selected_sum
            net_amount = total_grant_amount - updated_cost

            # Update grant data with the allocated costs
            self.grant_management.update_grant_data(selected_grant, 'Allocated Costs', updated_cost)
            self.grant_management.update_grant_data(selected_grant, 'Net Amount', net_amount)

            # Update the UI with the net amount
            self.net_amount_label.setText(f"Net Amount in Grant: ${net_amount:.2f}")

            QMessageBox.information(self.parent, "Costs Allocated", f"Successfully allocated ${selected_sum:.2f} to the {selected_grant} grant.")
        else:
            QMessageBox.warning(self.parent, "Grant Not Found", f"The selected grant {selected_grant} could not be found.")
