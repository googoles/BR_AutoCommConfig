# main.py
import sys
from PyQt5.QtWidgets import QApplication
from gui import ExcelImporterGUI

def main():
    app = QApplication(sys.argv)
    excel_gui = ExcelImporterGUI()
    excel_gui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
