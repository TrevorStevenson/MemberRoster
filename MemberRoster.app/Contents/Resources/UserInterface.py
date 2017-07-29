import MemberRoster
from PyQt5 import QtWidgets, QtGui, QtCore
from robobrowser import RoboBrowser


class UI(QtWidgets.QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Member Roster")
        self.setGeometry(0, 0, 500, 500)
        self.move(QtWidgets.QApplication.desktop().screen().rect().center() - self.rect().center())
        self.setStyleSheet("background-color: #0074BB")

        self.font = Font()

        self.vertical_layout = QtWidgets.QVBoxLayout()
        self.progress_bar = QtWidgets.QProgressBar()
        self.browser = RoboBrowser()
        self.info_label = QtWidgets.QLabel()
        self.engine_thread = QtCore.QThread()

        self.logo = QtWidgets.QLabel()
        self.logo_layout = QtWidgets.QHBoxLayout()
        self.add_logo()

        self.username = TextField("Username")
        self.username.setFont(self.font)
        self.vertical_layout.addLayout(self.username.layout)

        self.password = TextField("Password")
        self.password.setFont(self.font)
        self.vertical_layout.addLayout(self.password.layout)

        self.run_button = QtWidgets.QPushButton()
        self.run_button.setText("Run")
        self.run_button.setFont(self.font)
        self.run_button.setFixedWidth(200)
        self.run_button.setFixedHeight(60)
        self.run_button.setStyleSheet("background-color: #64A70B; color: white;")
        self.run_button.clicked.connect(self.run)
        self.run_button_layout = QtWidgets.QHBoxLayout()
        self.run_button_layout.addStretch()
        self.run_button_layout.addWidget(self.run_button)
        self.run_button_layout.addStretch()
        self.vertical_layout.addLayout(self.run_button_layout)

        self.help_button = QtWidgets.QPushButton()
        self.help_button.setText("Help")
        self.help_button.setFont(self.font)
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

    def add_logo(self):
        self.logo.setPixmap(QtGui.QPixmap("Logo.png").scaled(150, 150, QtCore.Qt.KeepAspectRatioByExpanding))
        self.logo_layout.addStretch()
        self.logo_layout.addWidget(self.logo)
        self.logo_layout.addStretch()
        self.vertical_layout.addLayout(self.logo_layout)

    def run(self):
        self.username.setStyleSheet("background-color: white")
        self.password.setStyleSheet("background-color: white")
        self.engine_thread = Thread(self.username.text(), self.password.text(), self)
        self.engine_thread.start()
        self.show_progress_bar()

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

    def show_progress_bar(self):
        QtWidgets.QWidget().setLayout(self.vertical_layout)
        self.vertical_layout = QtWidgets.QVBoxLayout()

        self.logo = QtWidgets.QLabel()
        self.logo_layout = QtWidgets.QHBoxLayout()
        self.add_logo()

        self.progress_bar.setStyleSheet("color: white")
        self.progress_bar.setFont(self.font)
        progress_layout = QtWidgets.QHBoxLayout()
        progress_layout.addStretch()
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addStretch()
        self.vertical_layout.addLayout(progress_layout)

        self.info_label.setText("Logging in...")
        self.info_label.setStyleSheet("color: white")
        self.info_label.setFont(self.font)
        label_layout = QtWidgets.QHBoxLayout()
        label_layout.addStretch()
        label_layout.addWidget(self.info_label)
        label_layout.addStretch()
        self.vertical_layout.addLayout(label_layout)

        self.setLayout(self.vertical_layout)

        self.progress_bar.setValue(0)

    def show_welcome_message(self):
        welcome_alert = QtWidgets.QMessageBox()
        welcome_alert.setIcon(QtWidgets.QMessageBox.Information)
        welcome_alert.setFont(self.font)
        welcome_alert.setWindowTitle("Welcome")
        welcome_alert.setText("Enter your Premier Connect login information and press run.")
        welcome_alert.setInformativeText("It may take a few minutes depending on your network speed. The excel spreadsheet will open automatically, and then you can save it to any desired location using File->Save As.")
        welcome_alert.setStandardButtons(QtWidgets.QMessageBox.Ok)
        welcome_alert.exec_()


class Font(QtGui.QFont):

    def __init__(self):
        super().__init__()
        self.setFamily("Avenir")


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


class Thread(QtCore.QThread):

    def __init__(self, username, password, ui):
        super().__init__()

        self.username = username
        self.password = password
        self.ui = ui

    def run(self):
        MemberRoster.run_script(self.username, self.password, self.ui)
        self.exit()
