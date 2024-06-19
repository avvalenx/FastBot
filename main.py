import sys
from PyQt5.QtWidgets import QApplication
import fast_frontend

app = QApplication(sys.argv)

window = fast_frontend.MainWindow()
window.check_for_login_creds()
window.show()

app.exec()