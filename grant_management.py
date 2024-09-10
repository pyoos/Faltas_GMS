import os
import re
import pandas as pd
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QComboBox,
    QFileDialog, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout, QDateEdit
)
from PyQt5.QtCore import Qt, QDate

class GrantManagement:
    def __init__(self, directory_path='/Users/paul/Desktop/Faltas_GMS'):
        self.directory_path = directory_path
        self.file_path = os.path.join(self.directory_path, 'grants.csv')
        self.costs_file_path = os.path.join(self.directory_path, 'allocated_costs.csv')  # New file for costs
        self.required_columns = ['Grant ID', 'Grant Name', 'Total Balance', 'Allowed Items']
        self.grant_data = self.load_grants()
        self.allocated_costs = self.load_allocated_costs()

    def select_csv_file(self):
        """Search for all CSV files in the directory and prompt the user to select one."""
        csv_files = [f for f in os.listdir(self.directory_path) if f.endswith('.csv')]
        
        if not csv_files:
            QMessageBox.warning(None, "No CSV Files Found", "No CSV files were found in the directory.")
            return None
        elif len(csv_files) == 1:
            return os.path.join(self.directory_path, csv_files[0])
        else:
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getOpenFileName(None, "Select a CSV File", self.directory_path, "CSV Files (*.csv)", options=options)
            if file_path:
                return file_path
            else:
                return None
    def load_grants(self):
        if os.path.exists(self.file_path):
            try:
                data = pd.read_csv(self.file_path)
                if all(column in data.columns for column in self.required_columns):
                    if 'Allowed Items' in data.columns:
                        data['Allowed Items'] = data['Allowed Items'].apply(eval)
                    return data
                else:
                    QMessageBox.warning(None, "CSV Error", f"The file {self.file_path} does not contain the required columns.")
                    return self.initialize_csv()
            except Exception as e:
                QMessageBox.warning(None, "CSV Error", f"There was an error loading the file {self.file_path}: {str(e)}")
                return self.initialize_csv()
        else:
            return self.initialize_csv()
        
    def load_allocated_costs(self):
        """Load allocated costs from a CSV file."""
        if os.path.exists(self.costs_file_path):
            try:
                data = pd.read_csv(self.costs_file_path)
                if 'Grant ID' in data.columns and 'Cost' in data.columns:
                    return data
                else:
                    return self.initialize_costs_csv()
            except Exception as e:
                print(f"Error loading allocated costs: {str(e)}")
                return self.initialize_costs_csv()
        else:
            return self.initialize_costs_csv()

    def initialize_csv(self):
        data = pd.DataFrame(columns=self.required_columns)
        if self.file_path:
            data.to_csv(self.file_path, index=False)
        return data

    def initialize_costs_csv(self):
        """Initialize the costs CSV with default columns."""
        data = pd.DataFrame(columns=['Grant ID', 'Cost'])
        data.to_csv(self.costs_file_path, index=False)
        return data

    def delete_grant(self, grant_id):
        """Delete a grant from the system by Grant ID."""
        if grant_id in self.grant_data['Grant ID'].values:
            self.grant_data = self.grant_data[self.grant_data['Grant ID'] != grant_id]
            self.save_grants()
            return True
        else:
            return False

    def save_grants(self):
        if self.file_path:
            self.grant_data.to_csv(self.file_path, index=False)


    def save_allocated_costs(self):
        """Save the allocated costs to the CSV file."""
        if self.costs_file_path:
            self.allocated_costs.to_csv(self.costs_file_path, index=False)

    def add_allocated_cost(self, grant_id, cost):
        """Add a new allocated cost to the list for a specific grant."""
        new_cost = pd.DataFrame({'Grant ID': [grant_id], 'Cost': [cost]})
        self.allocated_costs = pd.concat([self.allocated_costs, new_cost], ignore_index=True)
        self.save_allocated_costs()

    def remove_allocated_cost(self, grant_id, cost):
        """Remove an allocated cost from the list for a specific grant."""
        self.allocated_costs = self.allocated_costs[
            ~((self.allocated_costs['Grant ID'] == grant_id) & (self.allocated_costs['Cost'] == cost))
        ]
        self.save_allocated_costs()

    def get_allocated_costs(self, grant_id):
        """Retrieve allocated costs for a specific grant."""
        return self.allocated_costs[self.allocated_costs['Grant ID'] == grant_id]

    def show_grants(self):
        """Show all the grants in the system."""
        try:
            if self.grant_management.grant_data.empty:
                QMessageBox.information(self, "No Grants", "There are no grants in the database.")
            else:
                dialog = QDialog(self)
                dialog.setWindowTitle("Existing Grants")
                dialog.setStyleSheet("background-color: #cce7ff;")
                dialog.resize(800, 600)

                scroll = QScrollArea()
                scroll.setWidgetResizable(True)

                widget = QWidget()
                vbox = QVBoxLayout(widget)

                for _, row in self.grant_management.grant_data.iterrows():
                    grant_id_label = QLabel(f"Grant ID: {row['Grant ID']}")
                    grant_id_label.setStyleSheet("font-size: 16px; color: #333; font-weight: bold; margin-bottom: 5px;")
                    vbox.addWidget(grant_id_label)

                    grant_name_label = QLabel(f"Grant Name: {row['Grant Name']}")
                    grant_name_label.setStyleSheet("font-size: 16px; color: #333; margin-bottom: 5px;")
                    vbox.addWidget(grant_name_label)

                    total_balance_label = QLabel(f"Total Balance: ${row['Total Balance']:.2f}")
                    total_balance_label.setStyleSheet("font-size: 16px; color: #333; margin-bottom: 5px;")
                    vbox.addWidget(total_balance_label)

                    spending_rules_label = QLabel("Grant Spending Rules:")
                    spending_rules_label.setStyleSheet("font-size: 16px; color: #333; font-weight: bold; margin-bottom: 5px;")
                    vbox.addWidget(spending_rules_label)

                    for item in row['Allowed Items']:
                        item_label = QLabel(item)
                        item_label.setStyleSheet("""
                            font-size: 16px;
                            color: #333;
                            background-color: #f0f8ff;
                            border: 1px solid #99c2ff;
                            border-radius: 5px;
                            padding: 5px;
                            margin-bottom: 3px;
                        """)
                        vbox.addWidget(item_label)

                    vbox.addWidget(QLabel("\n"))  # Add spacing between entries

                scroll.setWidget(widget)

                layout = QVBoxLayout(dialog)
                layout.addWidget(scroll)

                dialog.setLayout(layout)
                dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while displaying grants: {str(e)}")


    def get_grant_data(self, grant_name):
        """Retrieve data for a specific grant."""
        return self.grant_data[self.grant_data['Grant Name'] == grant_name]

    def update_grant_data(self, grant_name, key, value):
        """Update specific data in a grant."""
        self.grant_data.loc[self.grant_data['Grant Name'] == grant_name, key] = value
        self.save_grants()

    def get_grant_names(self):
        """Retrieve a list of all grant names."""
        return self.grant_data['Grant Name'].tolist()

    def add_grant(self, grant_id, grant_name, total_balance, allowed_items):
        """Add a new grant to the system."""
        if grant_name not in self.grant_data['Grant Name'].values:
            new_grant = pd.DataFrame({
                'Grant ID': [grant_id],
                'Grant Name': [grant_name],
                'Total Balance': [total_balance],
                'Allowed Items': [allowed_items]
            })

            # Handle empty DataFrame case explicitly
            if self.grant_data.empty:
                self.grant_data = new_grant
            else:
                self.grant_data = pd.concat([self.grant_data, new_grant], ignore_index=True)

            self.save_grants()
        else:
            QMessageBox.warning(None, "Duplicate Grant", f"The grant '{grant_name}' already exists.")


    def delete_grant(self, grant_id):
        """Delete a grant from the system by Grant ID."""
        if grant_id in self.grant_data['Grant ID'].values:
            # Remove the grant from the DataFrame
            self.grant_data = self.grant_data[self.grant_data['Grant ID'] != grant_id]
            self.save_grants()  # Save changes to the CSV
            return True
        else:
            return False

    def add_initial_grants(self):
        """Prompt the user to input a new grant."""
        dialog = QDialog()
        dialog.setWindowTitle("Add New Grant")
        dialog.setStyleSheet("background-color: #cce7ff;")  # Optional: Set background color
        dialog.resize(400, 300)  # Optional: Set initial size

        layout = QVBoxLayout()

        # Grant ID Input
        label_id = QLabel("Enter Grant ID:")
        layout.addWidget(label_id)
        grant_id_input = QLineEdit()
        layout.addWidget(grant_id_input)

        # Grant Name Input
        label_name = QLabel("Enter Grant Name:")
        layout.addWidget(label_name)
        grant_name_input = QLineEdit()
        layout.addWidget(grant_name_input)

        # Total Balance Input
        label_balance = QLabel("Enter Total Balance:")
        layout.addWidget(label_balance)
        grant_balance_input = QLineEdit()
        layout.addWidget(grant_balance_input)

        # Allowed Items Input
        label_items = QLabel("Enter Allowed Items (comma-separated):")
        layout.addWidget(label_items)
        allowed_items_input = QLineEdit()
        layout.addWidget(allowed_items_input)

        add_button = QPushButton("Add Grant")
        layout.addWidget(add_button)

        def add_grant_action():
            grant_id = grant_id_input.text().strip()
            grant_name = grant_name_input.text().strip()
            allowed_items = [item.strip() for item in allowed_items_input.text().split(',')]
            try:
                total_balance = float(grant_balance_input.text().strip())
                if grant_id and grant_name:
                    self.add_grant(grant_id, grant_name, total_balance, allowed_items)
                    QMessageBox.information(dialog, "Success", f"Grant '{grant_name}' added successfully.")
                    dialog.accept()
                else:
                    QMessageBox.warning(dialog, "Input Error", "Grant ID and Grant Name cannot be empty.")
            except ValueError:
                QMessageBox.warning(dialog, "Input Error", "Please enter a valid number for total balance.")

        add_button.clicked.connect(add_grant_action)

        dialog.setLayout(layout)
        dialog.exec_()


    def show_grants(self):
        """Show all the grants in the system."""
        try:
            if self.grant_data.empty:
                QMessageBox.information(None, "No Grants", "No grants available to display.")
            else:
                grants_info = "\n".join(
                    [
                        f"Grant Name: {row['Grant Name']}\n  Total Balance: ${row['Total Balance']:.2f}\n"
                        f"  Allowed Items: {', '.join(row['Allowed Items'])}\n"
                        for _, row in self.grant_data.iterrows()
                    ]
                )
                QMessageBox.information(None, "Current Grants", grants_info)
        except Exception as e:
            QMessageBox.critical(None, "Error", f"An error occurred while displaying grants: {str(e)}")

    def choose_grant_for_rule(self):
        """Allow the user to select a grant and add a spending rule."""
        if self.grant_data.empty:
            QMessageBox.information(None, "No Grants", "No grants available to add a spending rule.")
            return

        dialog = QDialog()
        dialog.setWindowTitle("Choose Grant for Spending Rule")
        layout = QVBoxLayout()

        label = QLabel("Select a grant to add a spending rule:")
        layout.addWidget(label)

        grant_combo = QComboBox()
        grant_combo.addItems(self.get_grant_names())
        layout.addWidget(grant_combo)

        add_rule_button = QPushButton("Add Spending Rule")
        add_rule_button.clicked.connect(lambda: self.add_spending_rule(grant_combo.currentText(), dialog))
        layout.addWidget(add_rule_button)

        dialog.setLayout(layout)
        dialog.exec_()

    def add_spending_rule(self, grant_name, dialog):
        """Add a spending rule to the selected grant."""
        QMessageBox.information(None, "Spending Rule", f"Spending rule added to grant: {grant_name}")
        dialog.accept()


# Integration with ExcelHandler
class ExcelHandler:
    def __init__(self, parent, grant_management):
        self.parent = parent
        self.grant_management = grant_management
        self.total_cost = 0
        self.selected_sum_label = None
        self.sheet_data = None

    def upload_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self.parent, "Upload Inventory Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        
        if file_path:
            try:
                excel_data = pd.read_excel(file_path, sheet_name=None)
                self.display_excel_contents(excel_data)
            except Exception as e:
                QMessageBox.critical(self.parent, "Error", f"An error occurred while uploading the Excel file: {str(e)}")

    def display_excel_contents(self, excel_data):
        dialog = QDialog(self.parent)
        dialog.setWindowTitle("Excel File Contents")
        dialog.setStyleSheet("background-color: #cce7ff;")
        dialog.resize(1200, 1000)

        layout = QVBoxLayout()

        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("font-size: 14px;")

        self.total_cost = 0
        self.has_cost_column = False

        for sheet_name, sheet_data in excel_data.items():
            if sheet_data.empty:
                continue

            sheet_data.columns = sheet_data.columns.str.lower()
            sheet_data = sheet_data.fillna("")

            self.sheet_data = sheet_data

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
                # Clean selected values by removing $ and other non-numeric characters
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

        if grant_data is not None:
            total_grant_amount = grant_data['Total Balance'].iloc[0]
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
