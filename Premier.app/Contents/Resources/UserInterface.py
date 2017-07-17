import Premier
from PyQt5 import QtWidgets, QtGui, QtCore


class UI(QtWidgets.QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Premier")
        self.setGeometry(0, 0, 500, 500)
        self.move(QtWidgets.QApplication.desktop().screen().rect().center() - self.rect().center())
        self.setStyleSheet("background-color: #0074BB")

        self.vertical_layout = QtWidgets.QVBoxLayout()

        self.logo = QtWidgets.QLabel()
        self.logo.setPixmap(QtGui.QPixmap("Logo.png").scaled(150, 150, QtCore.Qt.KeepAspectRatioByExpanding))
        self.logo_layout = QtWidgets.QHBoxLayout()
        self.logo_layout.addStretch()
        self.logo_layout.addWidget(self.logo)
        self.logo_layout.addStretch()
        self.vertical_layout.addLayout(self.logo_layout)

        self.username = TextField("Username")
        self.vertical_layout.addLayout(self.username.layout)

        self.password = TextField("Password")
        self.vertical_layout.addLayout(self.password.layout)

        self.spg = TextField("SPG IDs")
        self.vertical_layout.addLayout(self.spg.layout)

        self.lidn = TextField("LIDN IDs")
        self.vertical_layout.addLayout(self.lidn.layout)

        self.run_button = QtWidgets.QPushButton()
        self.run_button.setText("Run")
        self.run_button.setFixedWidth(200)
        self.run_button.setFixedHeight(60)
        self.run_button.setStyleSheet("background-color: #60C578; color: white;")
        self.run_button.clicked.connect(self.run)
        self.run_button_layout = QtWidgets.QHBoxLayout()
        self.run_button_layout.addStretch()
        self.run_button_layout.addWidget(self.run_button)
        self.run_button_layout.addStretch()
        self.vertical_layout.addLayout(self.run_button_layout)

        self.help_button = QtWidgets.QPushButton()
        self.help_button.setText("Help")
        self.help_button.setFixedWidth(100)
        self.help_button.setFixedHeight(30)
        self.help_button.setStyleSheet("background-color: black; color: white;")
        self.help_button.clicked.connect(self.show_welcome_message)
        self.help_button_layout = QtWidgets.QHBoxLayout()
        self.help_button_layout.addStretch()
        self.help_button_layout.addWidget(self.help_button)
        self.help_button_layout.addStretch()
        self.vertical_layout.addLayout(self.help_button_layout)

        self.setLayout(self.vertical_layout)

    def run(self):
        self.username.setStyleSheet("background-color: white")
        self.password.setStyleSheet("background-color: white")
        Premier.run_script(self.username.text(),
                           self.password.text(),
                           self.spg.text().split(" "),
                           self.lidn.text().split(" "),
                           self)

    def login_failed(self):
        self.username.setStyleSheet("background-color: #e74c3c")
        self.password.setStyleSheet("background-color: #e74c3c")
        alert = QtWidgets.QMessageBox()
        alert.setIcon(QtWidgets.QMessageBox.Information)
        alert.setText("Invalid login.")
        alert.setInformativeText("Please try again.")
        alert.setStandardButtons(QtWidgets.QMessageBox.Ok)
        alert.exec_()
        self.username.clear()
        self.password.clear()

    @staticmethod
    def show_welcome_message():
        welcome_alert = QtWidgets.QMessageBox()
        welcome_alert.setIcon(QtWidgets.QMessageBox.Information)
        welcome_alert.setWindowTitle("Welcome")
        welcome_alert.setText("Enter your Premier Connect login information and the desired SPG and LIDN Ids.")
        welcome_alert.setInformativeText("Each ID number should be separated by a single space. Once you press run, it may take a few minutes depending on your network speed. The excel spreadsheet will open automatically, and then you can save it to your computer using File->Save As.")
        welcome_alert.setStandardButtons(QtWidgets.QMessageBox.Ok)
        welcome_alert.exec_()


class TextField(QtWidgets.QLineEdit):

    def __init__(self, placeholder):
        super().__init__()

        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setPlaceholderText(placeholder)
        self.setFrame(False)
        self.setFixedWidth(200)
        self.setFixedHeight(28)
        self.setStyleSheet("background-color: white")
        self.layout = QtWidgets.QHBoxLayout()
        self.layout.addStretch()
        self.layout.addWidget(self)
        self.layout.addStretch()
