# gui.py
from PyQt5.QtWidgets import QMainWindow, QToolBar, QAction, QFileDialog, QListWidget, QStatusBar, QWidget, QVBoxLayout
import os
from apj_handler import process_apj_file

class ExcelImporterGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        # Declare apj_file_path as an instance variable
        self.apj_file_path = None

        self.init_ui()

    def init_ui(self):
        # Create widgets
        self.list_widget = QListWidget(self)
        self.status_bar = QStatusBar()

        # Create a toolbar
        toolbar = QToolBar("Toolbar")
        self.addToolBar(toolbar)

        # Create "Import Project" action and add it to the toolbar
        import_project_action = QAction('Import Project', self)
        import_project_action.triggered.connect(self.import_apj)
        toolbar.addAction(import_project_action)

        # Create "Import Excel" action and add it to the toolbar
        import_excel_action = QAction('Import Excel', self)
        import_excel_action.triggered.connect(self.import_excel)
        toolbar.addAction(import_excel_action)

        # Connect the item selection signal to a custom function
        self.list_widget.itemClicked.connect(self.item_selected)

        # Create layout
        layout = QWidget()
        layout_layout = QVBoxLayout(layout)
        layout_layout.addWidget(toolbar)
        layout_layout.addWidget(self.list_widget)

        # Set the central widget
        self.setCentralWidget(layout)

        # Set the status bar
        self.setStatusBar(self.status_bar)

        # Set the main window properties
        self.setGeometry(300, 300, 1600, 1200)  # Set the window size to "1600 x 1200"
        self.setWindowTitle('Excel Importer')

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

            print(f"Attempting to find 'Hardware.hw' in: {selected_directory_path}")

            try:
                # Check if 'Hardware.hw' exists in the selected directory
                hardware_hw_path = os.path.join(selected_directory_path, 'Hardware.hw')

                if os.path.exists(hardware_hw_path):
                    # Display a message or take action if 'Hardware.hw' is found
                    self.list_widget.clear()
                    self.list_widget.addItem(f"'Hardware.hw' found in: {selected_directory_path}")
                else:
                    self.list_widget.clear()
                    self.list_widget.addItem(f"'Hardware.hw' not found in: {selected_directory_path}")

            except Exception as e:
                print(f"Error finding 'Hardware.hw': {str(e)}")
                self.status_bar.showMessage(f"Error: {str(e)}")
        else:
            self.status_bar.showMessage("Please import an APJ file first.")

