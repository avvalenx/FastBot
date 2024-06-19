import sys
import os

from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import QSize, Qt, QThread
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QHBoxLayout, QLabel, QLineEdit

import fast_backend

# TODO
# fix threading so that finished signal works correctly
#     - finished signal not being put for each function run is causing issue
#     - moving to just 1 function does not matter if you execute multiple functions in the class
# format loading screen that counts up per fan done
# rename things to be more clear
# add error handling for bad input
# error handling if excel spreadsheet is open and instructions on that
# for any fans that are invalid keep track of business name
# don't overwrite excel spreadsheet that comes out
# let them set output
# message once finished
# opening spreadsheet can break it if reading or writing
# change login to attuid
# bold business name and fan or just make output look much better
# change title and descriptions for everything
# make it pretty dont forget error message
# optimization to make it launch faster

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # set base path for resources based on if application is raw code or packaged
        try:
            self.base_path = sys._MEIPASS
        except:
            self.base_path = os.getcwd()

        self.WINDOW_SIZE = QSize(1000, 600)
        self.WINDOW_TITLE = 'Fast Bot'
        self.setWindowIcon(QIcon(self.resource_path('fast_icon.ico')))

        self.error_message = None

    def resource_path(self, relative_path):
        return os.path.join(self.base_path, relative_path)
    
    def check_for_login_creds(self):
        if os.path.isfile(os.path.join(self.base_path, 'creds.txt')):
            self.show_home_page()
        else:
            self.show_welcome_page()

    def show_welcome_page(self):
        self.setWindowTitle(self.WINDOW_TITLE)
        self.setFixedSize(self.WINDOW_SIZE)
        self.setContentsMargins(0,0,10,10)

        att_logo = QLabel(self)
        pixmap = QPixmap(self.resource_path('att.png'))
        att_logo.setPixmap(pixmap)
        att_logo.setScaledContents(True)
        att_logo.setFixedWidth(500)

        self.title_label = QLabel('AT&T Fast Bot')
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setAlignment(Qt.AlignTop)
        self.title_label.setContentsMargins(100,0,0,0)
        title_font = self.title_label.font()
        title_font.setPointSize(19)
        self.title_label.setFont(title_font)
        
        self.intro_label = QLabel('Welcome to the AT&T Fast Bot. You can upload an Excel spreadsheet with a list of business names and FANs. The bot will create an Excel spreadsheet with all the contacts from those FANs.')
        self.intro_label.setAlignment(Qt.AlignCenter)
        self.intro_label.setWordWrap(True)
        intro_font = self.intro_label.font()
        intro_font.setPointSize(10)
        self.intro_label.setFont(intro_font)

        self.instruction_label = QLabel('First log in with your attuid global login credentials.')
        self.instruction_label.setAlignment(Qt.AlignCenter)
        self.instruction_label.setWordWrap(True)
        instruction_font = self.instruction_label.font()
        instruction_font.setPointSize(10)
        self.instruction_label.setFont(instruction_font)

        self.username_label = QLabel('Username:')

        self.username_input = QLineEdit()

        self.password_label = QLabel('Password:')

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.returnPressed.connect(self.write_credentials)

        self.submit_button = QPushButton('Continue')
        self.submit_button.clicked.connect(self.write_credentials)

        # create layouts
        vlayout = QVBoxLayout()
        hlayout = QHBoxLayout()
        hlayout.setContentsMargins(0,0,0,0)
        usernamelayout = QHBoxLayout()
        passwordlayout = QHBoxLayout()

        # add widgets to layouts
        usernamelayout.addWidget(self.username_label)
        usernamelayout.addWidget(self.username_input)
        passwordlayout.addWidget(self.password_label)
        passwordlayout.addWidget(self.password_input)
        vlayout.addWidget(self.title_label)
        if self.error_message != None:
            vlayout.addWidget(self.error_message)
        vlayout.addWidget(self.intro_label)
        vlayout.addWidget(self.instruction_label)
        vlayout.addLayout(usernamelayout)
        vlayout.addLayout(passwordlayout)
        vlayout.addWidget(self.submit_button)
        hlayout.addWidget(att_logo)
        hlayout.addLayout(vlayout)
        
        # Set the central widget of the Window.
        widget = QWidget()
        widget.setLayout(hlayout)
        self.setCentralWidget(widget)

    def show_home_page(self):
        self.setWindowTitle(self.WINDOW_TITLE)
        self.setFixedSize(self.WINDOW_SIZE)
        self.setContentsMargins(0,0,10,10)

        self.input_path = ''
        self.input_name = 'No file selected'

        att_logo = QLabel(self)
        pixmap = QPixmap(self.resource_path('att.png'))
        att_logo.setPixmap(pixmap)
        att_logo.setScaledContents(True)
        att_logo.setFixedWidth(500)

        self.excel_file_selected = QLabel(self.input_name)

        self.file_select_button = QPushButton("Select Excel File")
        self.file_select_button.clicked.connect(self.input_select)

        self.progress_label = QLabel()

        self.scrape_fast_contacts_button = QPushButton('Get Contact List')
        self.scrape_fast_contacts_button.setEnabled(False)
        self.scrape_fast_contacts_button.clicked.connect(self.scrape_fast_contacts)

        # create layout
        hlayout = QHBoxLayout()
        hlayout.setContentsMargins(0,0,0,0)
        hlayout_file_select = QHBoxLayout()
        hlayout_file_select.setContentsMargins(0,0,0,0)
        vlayout = QVBoxLayout()
        vlayout.setContentsMargins(0,0,0,0)

        hlayout.addWidget(att_logo)
        hlayout_file_select.addWidget(self.excel_file_selected)
        hlayout_file_select.addWidget(self.file_select_button)
        vlayout.addLayout(hlayout_file_select)
        vlayout.addWidget(self.progress_label)
        vlayout.addWidget(self.scrape_fast_contacts_button)
        hlayout.addLayout(vlayout)

        # Set the central widget of the Window.
        widget = QWidget()
        widget.setLayout(hlayout)
        self.setCentralWidget(widget)
    
    def write_credentials(self):
        with open(self.resource_path('creds.txt'), 'w') as creds:
            creds.write(self.username_input.text())
            creds.write('\n')
            creds.write(self.password_input.text())
        self.show_home_page()
    
    def input_select(self):
        dlg = QFileDialog().getOpenFileName(self, 'Open Excel Document', None, "Excel files (*.xlsx)")
        self.input_path = dlg[0]
        self.input_name = dlg[0].split('/')[-1]
        self.excel_file_selected.setText(self.input_name)
        if os.path.exists(self.input_path) and self.input_name.endswith('.xlsx'):
            self.scrape_fast_contacts_button.setEnabled(True)
        else:
            self.scrape_fast_contacts_button.setEnabled(False)
        
    def scrape_fast_contacts(self):
        self.thread = QThread()
        self.fast = fast_backend.Fast()
        self.fast.input_path = self.input_path
        self.fast.output_file_name = 'Fast Contacts.xlsx'
        self.fast.base_path = self.base_path

        self.fast.moveToThread(self.thread)
        self.thread.started.connect(self.fast.startup)
        self.thread.started.connect(self.fast.import_fans)
        self.thread.started.connect(self.fast.create_driver)
        self.thread.started.connect(self.fast.login)
        self.thread.started.connect(self.fast.scrape_contacts)
        self.thread.started.connect(self.fast.export_contacts_to_excel)
        self.fast.finished.connect(self.thread.quit)
        self.fast.finished.connect(self.fast.deleteLater)
        self.fast.finished.connect(self.thread.deleteLater)
        self.fast.progress.connect(self.report_progress)
        self.thread.start()

        self.scrape_fast_contacts_button.setEnabled(False)
        self.thread.finished.connect(lambda: self.scrape_fast_contacts_button.setEnabled(True))
    
    def report_progress(self, fan_counter):
        self.progress_label.setText(str(fan_counter))

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MainWindow()
    #window.check_for_login_creds()
    window.show_home_page()
    #window.show_welcome_page()
    window.show()

    app.exec()