import sys
import random
from PySide6 import QtCore, QtGui
import os
from PySide6.QtWidgets import QWidget, QApplication, QFileDialog, QGroupBox, QGridLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QHBoxLayout, QVBoxLayout, QCheckBox
from PySide6.QtCore import QProcess

from PySide6.QtCore import QCoreApplication, QIODevice, QByteArray
from PySide6.QtNetwork import QTcpSocket
import json

import time
# from tarrafa_server import Tarrafa
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # Configurações
        self.grid = QGridLayout()
        self.inputLabel = QLabel("Diretorio de Busca:")
        self.inputLineEdit = QLineEdit()
        self.inputButton = QPushButton("Procurar")
        self.outputLabel = QLabel("Diretório de Saída:")
        self.outputLineEdit = QLineEdit()
        self.outputButton = QPushButton("Procurar")
        self.regexTextEdit = QTextEdit() # Descobrir como colocar texto sugestivo e regular as dimensões
        # Check boxes
        self.motivoRevisaoCheckBox = QCheckBox("Ignorar Motivo da Revisão")
        self.listaDistribuicaoCheckBox = QCheckBox("Ignorar Lista de distribuição")
        # Message Label
        self.messageLabel = QLabel("")

        self.inputButton.clicked.connect(self.getInputDir)
        self.outputButton.clicked.connect(self.getoutputDir)
        self.grid.addWidget(self.inputLabel, 0,0)
        self.grid.addWidget(self.inputLineEdit, 0,1)
        self.grid.addWidget(self.inputButton, 0,2)
        self.grid.addWidget(self.outputLabel, 1,0)
        self.grid.addWidget(self.outputLineEdit, 1,1)
        self.grid.addWidget(self.outputButton, 1,2)
        self.grid.addWidget(self.regexTextEdit, 2,0, 1, 3)
        self.grid.addWidget(self.motivoRevisaoCheckBox, 3,0, 1,3)
        self.grid.addWidget(self.listaDistribuicaoCheckBox, 4,0, 1,3)
        self.grid.addWidget(self.messageLabel, 5,0, 1,3)
        # Botões
        self.actionLayout = QHBoxLayout()
        self.gerarExcel_btn = QPushButton("Salvar Excel")
        self.gerarTxt_btn = QPushButton("Criar txt's")
        self.run_btn = QPushButton("Rodar!")
        self.actionLayout.addWidget(self.gerarExcel_btn)
        self.actionLayout.addWidget(self.gerarTxt_btn)
        self.actionLayout.addWidget(self.run_btn)

        self.gerarExcel_btn.clicked.connect(self.salvar_excel)
        self.gerarTxt_btn.clicked.connect(self.gerarTxt)
        self.run_btn.clicked.connect(self.rodar_tarrafa)

       

        # Vertical Layout
        self.VLayout = QVBoxLayout()
        self.VLayout.addLayout(self.grid)
        self.VLayout.addLayout(self.actionLayout)

        self.setLayout(self.VLayout)
        self.setWindowTitle("Tarrafa")
        # self.hello = ["Hallo Welt", "Hei maailma", "Hola Mundo", "Привет мир"]

        # self.button = QtWidgets.QPushButton("Click me!")
        # self.text = QtWidgets.QLabel("Hello World",
        #                              alignment=QtCore.Qt.AlignCenter)

        # self.layout = QtWidgets.QVBoxLayout(self)
        # self.layout.addWidget(self.text)
        # self.layout.addWidget(self.button)

        # self.button.clicked.connect(self.magic)

    @QtCore.Slot()
    def getInputDir(self):
        self.inputDir = QFileDialog.getExistingDirectory(self, "Abrir Diretório", os.getcwd())
        self.inputLineEdit.setText(self.inputDir)
    @QtCore.Slot()
    def getoutputDir(self):
        self.outputDir = QFileDialog.getExistingDirectory(self, "Abrir Diretório", os.getcwd())
        self.outputLineEdit.setText(self.outputDir)
    @QtCore.Slot()
    def salvar_excel(self):
        pass
    @QtCore.Slot()
    def gerarTxt(self):
        # tarrafa.convertAll(self.outputDir)
        pass

    @QtCore.Slot()
    def rodar_tarrafa(self):
        text = self.regexTextEdit.toPlainText()
        if text != "":
            dict2send = {"comando": "search_regex", "args" : (text,)}
            json_str = json.dumps(dict2send)
            json_byte_array = QByteArray(json_str.encode("utf-8"))
            client.write(json_byte_array)


class MyTcpClient(QTcpSocket):
    def __init__(self):
        super().__init__()
        # self.p = None
        # self.start_process()

        self.connected.connect(self.on_connected)
        self.readyRead.connect(self.read_data)
        self.errorOccurred.connect(self.on_error)


    def connect_to_server(self, host, port):
        self.connectToHost(host, port)
        if not self.waitForConnected(3000):
            print(f"Connection failed: {self.errorString()}")
            QCoreApplication.instance().quit()

    def on_connected(self):
        print("Connected to server")
        self.write(b'Hello, Server!')
        # self.write(b"print('Hello, Server!')")

    def read_data(self):
        while self.bytesAvailable():
            data = self.readAll()
            print(f"Received from server: {data.data().decode('utf-8')}")

    def on_error(self, socket_error):
        print(f"Socket error: {self.errorString()}")
        QCoreApplication.instance().quit()


    def start_process(self):
        if self.p is None:
            # self.message("Executing process")
            print("Executing Process")
            self.p = QProcess()  # Keep a reference to the QProcess (e.g. on self) while it's running.
            self.p.finished.connect(self.process_finished)  # Clean up once complete.
            self.p.readyReadStandardOutput.connect(self.handle_stdout)
            self.p.readyReadStandardError.connect(self.handle_stderr)
            self.p.stateChanged.connect(self.handle_state)
            self.p.finished.connect(self.process_finished)  # Clean up once complete.
            self.p.start("python", [r'.\\tarrafa_server.py'])
    
    def handle_stderr(self):
        data = self.p.readAllStandardError()
        stderr = bytes(data).decode("utf8")
        # self.message(stderr)
        print(stderr)

    def handle_stdout(self):
        data = self.p.readAllStandardOutput()
        stdout = bytes(data).decode("utf8")
        # self.message(stdout)
        print(stdout)

    def handle_state(self, state):
        states = {
            QProcess.NotRunning: 'Not running',
            QProcess.Starting: 'Starting',
            QProcess.Running: 'Running',
        }
        state_name = states[state]
        # self.message(f"State changed: {state_name}") 
        print(f"State changed: {state_name}")

    def process_finished(self):
        # self.message("Process finished.")
        print("Process Finished")
        self.p = None
if __name__ == "__main__":
    app = QApplication([])

    client = MyTcpClient()
    client.connect_to_server("127.0.0.1", 12345)
    # tarrafa = Tarrafa()
    widget = MainWindow()
    widget.resize(400, 300)
    widget.show()

sys.exit(app.exec())