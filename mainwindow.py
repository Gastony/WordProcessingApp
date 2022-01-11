from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtGui import QPixmap
from PyQt5 import QtCore, QtWidgets
from PyQt5 import QtGui



class Ui_MainWindow(QObject):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(1000, 800)
        # MainWindow.setWindowIcon(QtGui.QIcon(r"Resources/naot.PNG"))

        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r"Resources/naot.PNG"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 200, 200))
        self.label.setPixmap(QPixmap(r'Resources/nembo.jpg'))

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(830, 10, 200, 200))
        self.label_2.setPixmap(QPixmap(r'Resources/naot.jpg'))

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(10, 240, 400, 50))
        self.lineEdit.setText("")
        self.lineEdit.setObjectName("lineEdit")

        self.Browse_Button = QtWidgets.QPushButton(self.centralwidget)
        self.Browse_Button.setGeometry(QtCore.QRect(411, 240, 131, 50))
        self.Browse_Button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.Browse_Button.setObjectName("Browse_Button")

        self.Format_button = QtWidgets.QPushButton(self.centralwidget)
        self.Format_button.setGeometry(QtCore.QRect(150, 350, 200, 50))
        self.Format_button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.Format_button.setObjectName("Format_button")

        self.Remove_button = QtWidgets.QPushButton(self.centralwidget)
        self.Remove_button.setGeometry(QtCore.QRect(350, 350, 200, 50))
        self.Remove_button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.Remove_button.setObjectName("Remove_button")

        self.Open_file_button = QtWidgets.QPushButton(self.centralwidget)
        self.Open_file_button.setGeometry(QtCore.QRect(550, 350, 200, 50))
        self.Open_file_button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.Open_file_button.setObjectName("Open_file_button")

        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(542, 240, 400, 50))
        self.progressBar.setObjectName("progressBar")
        self.progressBar.setValue(0)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.Browse_Button.clicked.connect(self.browseSlot)
        self.lineEdit.returnPressed.connect(self.writeDocSlot)
        self.Format_button.clicked.connect(self.formatDoc)
        self.Remove_button.clicked.connect(self.removeFormat)
        self.Open_file_button.clicked.connect(self.openFile)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Style Guide"))
        self.lineEdit.setPlaceholderText(_translate("MainWindow", "File Name"))
        self.Browse_Button.setText(_translate("MainWindow", "Browse"))
        self.Format_button.setText(_translate("MainWindow", "Apply"))
        self.Remove_button.setText(_translate("MainWindow", "Remove formatting"))
        self.Open_file_button.setText(_translate("MainWindow", "Open file"))

    def update(self, progress):
        progress = 0


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
