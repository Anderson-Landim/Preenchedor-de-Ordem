from PyQt5 import QtWidgets, QtCore

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setWindowTitle("Exemplo PyUIBuilder")
        MainWindow.resize(400, 200)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        MainWindow.setCentralWidget(self.centralwidget)

        self.label = QtWidgets.QLabel("Mensagem:", self.centralwidget)
        self.label.setGeometry(20, 20, 100, 30)

        self.input = QtWidgets.QLineEdit(self.centralwidget)
        self.input.setGeometry(120, 20, 240, 30)

        self.button = QtWidgets.QPushButton("Mostrar", self.centralwidget)
        self.button.setGeometry(150, 80, 100, 40)

        self.result = QtWidgets.QLabel("", self.centralwidget)
        self.result.setGeometry(20, 140, 360, 30)

        self.button.clicked.connect(self.show_text)

    def show_text(self):
        txt = self.input.text()
        self.result.setText(f"Você digitou: {txt}")

app = QtWidgets.QApplication([])
window = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(window)
window.show()
app.exec_()
