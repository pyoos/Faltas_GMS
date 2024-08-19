import sys
from PyQt5.QtWidgets import QApplication
from ui import GrantManagementApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GrantManagementApp()
    window.show()
    sys.exit(app.exec_())
