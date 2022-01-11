from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from mainwindow import Ui_MainWindow
import sys
from model import Model


class MainWindowUIClass(Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.model = Model()


    def setupUi(self, MW):
        super().setupUi(MW)
        self.Remove_button.setEnabled(False)
        self.Open_file_button.setEnabled(False)

    def refreshAll(self):
        self.lineEdit.setText(self.model.getFileName())

    def browseSlot( self ):
        self.progressBar.setValue(0)
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            None,
            "File Broweser",
            "",
            "Word document (*.doc *.docx)",
            options=options)
        if fileName:
            self.model.setFileName(fileName)
            self.refreshAll()

    def writeDocSlot( self ):
        fileName = self.lineEdit.text()
        if self.model.isValid(fileName):
            self.model.setFileName(self.lineEdit.text())
            self.refreshAll()
        else:
            m = QtWidgets.QMessageBox()
            m.setText("Invalid file name!\n" + fileName)
            m.setIcon(QtWidgets.QMessageBox.Warning)
            m.setStandardButtons(QtWidgets.QMessageBox.Ok
                                 | QtWidgets.QMessageBox.Cancel)
            m.setDefaultButton(QtWidgets.QMessageBox.Cancel)
            ret = m.exec_()
            self.lineEdit.setText("")
            self.refreshAll()

    def formatDoc(self):
        self.model.writeDoc(self.progressBar, self.Remove_button, self.Open_file_button)

    def removeFormat(self):
        self.lineEdit.setText("")
        self.model.remove()

    def openFile(self):
        self.model.openDoc()



def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MainWindowUIClass()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

main()