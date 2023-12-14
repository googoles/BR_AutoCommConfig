# main.py
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QToolBar, QAction, QFileDialog, QTextBrowser, QStatusBar, QWidget, QVBoxLayout
import os
from apj_handler import process_apj_file

class ExcelImporterApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # Create widgets
        self.text_browser = QTextBrowser(self)
        self.status_bar = QStatusBar()

        # Create a toolbar
        toolbar = QToolBar("Toolbar")
        self.addToolBar(toolbar)

        # Create "Import Directory" action and add it to the toolbar
        import_dir_action = QAction('Import Directory', self)
        import_dir_action.triggered.connect(self.import_directory)
        toolbar.addAction(import_dir_action)

        # Create "Import Excel" action and add it to the toolbar
        import_excel_action = QAction('Import Excel', self)
        import_excel_action.triggered.connect(self.import_excel)
        toolbar.addAction(import_excel_action)

        # Create "Import APJ File" action and add it to the toolbar
        import_apj_action = QAction('Import APJ File', self)
        import_apj_action.triggered.connect(self.import_apj_file)
        toolbar.addAction(import_apj_action)

        # Create layout
        layout = QWidget()
        layout_layout = QVBoxLayout(layout)
        layout_layout.addWidget(toolbar)
        layout_layout.addWidget(self.text_browser)

        # Set the central widget
        self.setCentralWidget(layout)

        # Set the status bar
        self.setStatusBar(self.status_bar)

        # Set the main window properties
        self.setGeometry(300, 300, 500, 400)
        self.setWindowTitle('Excel Importer')

    def import_excel(self):
        # Open a file dialog to select the Excel file
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle('Select Excel File')
        file_dialog.setNameFilter('Excel Files (*.xls *.xlsx)')

        if file_dialog.exec_() == QFileDialog.Accepted:
            file_path = file_dialog.selectedFiles()[0]

            # Process the selected Excel file (replace with your logic)
            self.text_browser.clear()
            self.text_browser.setPlainText(f"Selected Excel File: {file_path}")

    def import_directory(self):
        # Open a directory dialog to select a directory
        dir_dialog = QFileDialog.getExistingDirectory(self, 'Select Directory')

        if dir_dialog:
            physical_folder = os.path.join(dir_dialog, 'Physical')

            # Check if "/Physical" folder exists
            if os.path.exists(physical_folder):
                # Process the contents of "/Physical" folder
                contents = self.process_physical_folder(physical_folder)
                self.display_data(contents)
            else:
                self.display_data("/Physical folder not found in the selected directory.")

    def import_apj_file(self):
        # Open a file dialog to select the APJ file
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle('Select APJ File')
        file_dialog.setNameFilter('APJ Files (*.apj)')

        if file_dialog.exec_() == QFileDialog.Accepted:
            apj_file_path = file_dialog.selectedFiles()[0]

            # Process the selected APJ file using apj_handler
            try:
                version, working_version = process_apj_file(apj_file_path)
                if version is not None and working_version is not None:
                    # Display version information in the status bar
                    self.status_bar.showMessage(f"AutomationStudio Version: {version}, WorkingVersion: {working_version}")
                else:
                    self.status_bar.showMessage("AutomationStudio element not found in the APJ file.")
            except ValueError as e:
                self.status_bar.showMessage(str(e))

    def process_physical_folder(self, physical_folder):
        # Process the contents of "/Physical" folder
        contents = os.listdir(physical_folder)
        return contents

    def display_data(self, data):
        # Display the data in the QTextBrowser
        self.text_browser.clear()
        self.text_browser.setPlainText(str(data))

def main():
    app = QApplication(sys.argv)
    excel_app = ExcelImporterApp()
    excel_app.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
