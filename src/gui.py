# gui.py
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QAction, QFileDialog, QListWidget, QStatusBar, QWidget, QVBoxLayout, QPlainTextEdit, QHBoxLayout, QLabel
import os
import subprocess
import traceback
from apj_handler import process_apj_file

class ExcelImporterGUI(QMainWindow):
    
    def __init__(self):
        super().__init__()

        # Declare apj_file_path as an instance variable
        self.apj_file_path = None

        # Create widgets
        self.list_widget = QListWidget(self)
        self.status_bar = QStatusBar()
        self.file_content_display = QPlainTextEdit(self)

        # Connect the item selection signal to a custom function
        self.list_widget.itemClicked.connect(self.item_selected)

        # Create layout
        layout = QHBoxLayout()  # Using QHBoxLayout to organize widgets horizontally
        left_widget = QWidget()
        right_widget = QWidget()

        # Create a QVBoxLayout for the left side (Status: toolbar and QListWidget)
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("Status"))  # Label for the left half plane
        left_layout.addWidget(self.list_widget)

        # Create a QVBoxLayout for the right side (Configuration: file_content_display)
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("Configuration"))  # Label for the right half plane
        right_layout.addWidget(self.file_content_display)

        # Add both left and right widgets to the main layout
        layout.addWidget(left_widget)
        layout.addWidget(right_widget)

        # Set the central widget
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Set the status bar
        self.setStatusBar(self.status_bar)

        # Set the main window properties
        self.setGeometry(300, 300, 1600, 1200)  # Set the window size to "1600 x 1200"
        self.setWindowTitle('Excel Importer')

        # Create menubar
        menubar = self.menuBar()

        # Create File menu
        file_menu = menubar.addMenu('File')

        # Create "Import Project" action and add it to the File menu
        import_project_action = QAction('Import Project', self)
        import_project_action.triggered.connect(self.import_apj)
        file_menu.addAction(import_project_action)

        # Create "Import Excel" action and add it to the File menu
        import_excel_action = QAction('Import Excel', self)
        import_excel_action.triggered.connect(self.import_excel)
        file_menu.addAction(import_excel_action)

    def import_apj(self):
        # Open a file dialog to select the APJ file
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle('Select APJ File')
        file_dialog.setNameFilter('APJ Files (*.apj)')

        if file_dialog.exec_() == QFileDialog.Accepted:
            # Assign the selected APJ file path to the instance variable
            self.apj_file_path = file_dialog.selectedFiles()[0]

            # Process the selected APJ file using apj_handler
            try:
                version, working_version = process_apj_file(self.apj_file_path)
                if version is not None and working_version is not None:
                    # Display version information in the status bar
                    self.status_bar.showMessage(f"AutomationStudio Version: {version}, WorkingVersion: {working_version}")

                    # Check for the /Physical subdirectory
                    physical_folder = os.path.join(os.path.dirname(self.apj_file_path), 'Physical')

                    if os.path.exists(physical_folder):
                        # Process the contents of /Physical subdirectory
                        contents = self.process_physical_folder(physical_folder)
                        self.display_data(contents)
                    else:
                        self.status_bar.showMessage("'/Physical' folder not found in the directory of the APJ file.")
                else:
                    self.status_bar.showMessage("AutomationStudio element not found in the APJ file.")
            except ValueError as e:
                self.status_bar.showMessage(str(e))

    def import_excel(self):
        # Open a file dialog to select the Excel file
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle('Select Excel File')
        file_dialog.setNameFilter('Excel Files (*.xls *.xlsx)')

        if file_dialog.exec_() == QFileDialog.Accepted:
            file_path = file_dialog.selectedFiles()[0]

            # Process the selected Excel file (replace with your logic)
            self.list_widget.clear()
            self.list_widget.addItem(f"Selected Excel File: {file_path}")

    def process_physical_folder(self, physical_folder):
        # Process the contents of the folder excluding *.pkg files
        contents = [item for item in os.listdir(physical_folder) if not item.endswith('.pkg')]
        return contents

    def display_data(self, data):
        # Display the data in the QListWidget
        self.list_widget.clear()
        self.list_widget.addItems(data)

    def item_selected(self, item):
        # Handle item selection
        selected_directory = item.text()

        if self.apj_file_path is not None:
            # Set the root directory based on the selected APJ file
            root_directory = os.path.dirname(self.apj_file_path)

            # Append '\Physical' between the root directory and the selected directory
            selected_directory_path = os.path.normpath(os.path.join(root_directory, 'Physical', selected_directory))

            try:
                # Check if 'Hardware.hw' exists in the selected directory
                hardware_hw_path = os.path.join(selected_directory_path, 'Hardware.hw')

                if os.path.exists(hardware_hw_path):
                    # Read and display the contents of 'Hardware.hw' in the QPlainTextEdit
                    with open(hardware_hw_path, 'r', encoding='utf-8') as file:
                        file_content = file.read()
                        self.file_content_display.setPlainText(file_content)

                    # Display a success popup message
                    success_message = f"'Hardware.hw' displayed in Configuration from: {selected_directory_path}"
                    QMessageBox.information(self, 'Success', success_message)

                else:
                    # Display a error popup message
                    error_message = f"'Hardware.hw' not found in Configuration: {selected_directory_path}"
                    QMessageBox.warning(self, 'Error', error_message)

            except Exception as e:
                print(f"Error finding and displaying 'Hardware.hw': {str(e)}")
                traceback.print_exc()  # Print the traceback to see detailed error information

                # Display an error popup message
                error_message = f"Error: {str(e)}"
                QMessageBox.critical(self, 'Error', error_message)

        else:
            # Display a message in a popup if no APJ file is imported
            QMessageBox.warning(self, 'Warning', "Please import an APJ file first.")