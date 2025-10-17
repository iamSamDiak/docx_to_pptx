from src.gui import MainWindow
from PyQt5.QtWidgets import QApplication
import sys

__version__ = "1.0.0"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle(f"DOCX to PPTX Converter v{__version__}")
    window.show()
    sys.exit(app.exec_())