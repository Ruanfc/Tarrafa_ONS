import sys
from PySide6 import QtCore
import os
import re
from PySide6.QtWidgets import QWidget, QApplication, QFileDialog, QGridLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QHBoxLayout, QVBoxLayout, QCheckBox
from PySide6.QtCore import QProcess
from PySide6.QtGui import QCloseEvent
from PySide6.QtCore import QCoreApplication, QByteArray, Signal
from PySide6.QtNetwork import QTcpSocket
import json

# from tarrafa_server import Tarrafa
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # regex Matches
        self.re_readinputDir = re.compile(r"(?<=inputDir:).*")
        self.re_readoutputDir = re.compile(r"(?<=outputDir:).*")
        self.re_regexText = re.compile(r"(?<=regexText:).*")
        self.re_igMotivo = re.compile(r"(?<=igMotivo:).*")
        self.re_igTabDist = re.compile(r"(?<=igTabDist:).*")
        self.re_readLog = re.compile(r"(?<=---\n).*", re.DOTALL)

        # Configurações
        self.grid = QGridLayout()
        self.inputLabel = QLabel("Diretorio de Busca:")
        self.inputLineEdit = QLineEdit()
        self.inputButton = QPushButton("Procurar")
        self.outputLabel = QLabel("Diretório de Saída:")
        self.outputLineEdit = QLineEdit()
        self.outputButton = QPushButton("Procurar")
        self.regexLabel = QLabel("Inserir padrão Regex:")
        self.clearLogBtn = QPushButton("Limpar log")
        self.regexLineEdit = QLineEdit()
        # Check boxes
        self.motivoRevisaoCheckBox = QCheckBox("Ignorar Motivo da Revisão")
        self.listaDistribuicaoCheckBox = QCheckBox("Ignorar Lista de distribuição")
        # Result Text Edit
        self.resultLabel = QLabel("Log de resultados:")
        self.resultTextEdit = QTextEdit()
        self.resultTextEdit.setReadOnly(True)
        # Message Label
        self.messageLabel = QLabel("")

        # Grid Layout
        self.inputButton.clicked.connect(self.getInputDir)
        self.outputButton.clicked.connect(self.getoutputDir)
        self.grid.addWidget(self.inputLabel, 0,0)
        self.grid.addWidget(self.inputLineEdit, 0,1)
        self.grid.addWidget(self.inputButton, 0,2)
        self.grid.addWidget(self.outputLabel, 1,0)
        self.grid.addWidget(self.outputLineEdit, 1,1)
        self.grid.addWidget(self.outputButton, 1,2)
        self.grid.addWidget(self.regexLabel, 2,0)
        self.grid.addWidget(self.regexLineEdit, 3,0, 1,3)
        self.grid.addWidget(self.motivoRevisaoCheckBox, 4,0, 1,3)
        self.grid.addWidget(self.listaDistribuicaoCheckBox, 5,0, 1,3)
        self.grid.addWidget(self.resultLabel, 6,0)
        self.grid.addWidget(self.clearLogBtn, 6,2)
        self.grid.addWidget(self.resultTextEdit, 7,0, 1,3)
        self.grid.addWidget(self.messageLabel, 8,0, 1,3)
        # Botões inferiores
        self.actionLayout = QHBoxLayout()
        self.gerarExcel_btn = QPushButton("Salvar Excel")
        self.gerarTxt_btn = QPushButton("Criar txt's")
        self.run_btn = QPushButton("Rodar!")
        self.actionLayout.addWidget(self.gerarExcel_btn)
        self.actionLayout.addWidget(self.gerarTxt_btn)
        self.actionLayout.addWidget(self.run_btn)

        # Associar slots
        self.gerarExcel_btn.clicked.connect(self.salvar_excel)
        self.gerarTxt_btn.clicked.connect(self.gerarTxt)
        self.run_btn.clicked.connect(self.rodar_tarrafa)
        self.regexLineEdit.returnPressed.connect(self.rodar_tarrafa)
        self.clearLogBtn.clicked.connect(self.clearLog)

        # Vertical Layout
        self.VLayout = QVBoxLayout()
        self.VLayout.addLayout(self.grid)
        self.VLayout.addLayout(self.actionLayout)

        self.setLayout(self.VLayout)
        self.setWindowTitle("Tarrafa")

        self.client = MyTcpClient()
        self.client.connect_to_server("127.0.0.1", 12345)

        self.client.message_receive.connect(self.setResultTextEdit) # Conecta mensagem ao textedit

        # Verifica se tem arquivo de log e carrega as variáveis
        self.readLog()


    def readLog(self):
        try:
            # Try to open the file in read mode
            with open(os.path.join(os.getcwd(), 'log.txt'), 'r', encoding='utf-8') as file:
                content = file.read()
                self.inputLineEdit.setText(self.re_readinputDir.findall(content)[0])
                self.outputLineEdit.setText(self.re_readoutputDir.findall(content)[0])
                self.resultTextEdit.setText(self.re_readLog.findall(content)[0])
                self.regexLineEdit.setText(self.re_regexText.findall(content)[0])
                self.motivoRevisaoCheckBox.setChecked(True if (self.re_igMotivo.findall(content)[0]) == "True" else False)
                self.listaDistribuicaoCheckBox.setChecked(True if (self.re_igTabDist.findall(content)[0]) == "True" else False)
        except FileNotFoundError:
            # Handle the case where the file does not exist
            print("The file was not found.")
        except IsADirectoryError:
            # Handle the case where the path is a directory, not a file
            print("Expected a file but found a directory.")
        except PermissionError:
            # Handle the case where the user does not have permission to read the file
            print("Permission denied.")
        except Exception as e:
            # Handle any other exceptions
            print(f"An unexpected error occurred: {e}")

    def fillDirs(self):
        self.inputDir = self.inputLineEdit.text().replace("/", os.path.sep)
        self.outputDir = self.outputLineEdit.text().replace("/", os.path.sep)
    @QtCore.Slot()
    def getInputDir(self):
        self.inputDir = QFileDialog.getExistingDirectory(self, "Abrir Diretório", os.getcwd()).replace("/", os.path.sep)
        self.inputLineEdit.setText(self.inputDir)
    @QtCore.Slot()
    def getoutputDir(self):
        self.outputDir = QFileDialog.getExistingDirectory(self, "Abrir Diretório", os.getcwd()).replace("/", os.path.sep)
        self.outputLineEdit.setText(self.outputDir)
    @QtCore.Slot()
    def salvar_excel(self):
        self.fillDirs()
        if(self.outputDir != ""):
            self.send2server("save_results_to_excel", self.outputDir)
    @QtCore.Slot()
    def gerarTxt(self):
        self.fillDirs()
        if (self.inputDir != "") and (self.outputDir !=""):
            self.send2server("convertAll", self.inputDir, self.outputDir)

    @QtCore.Slot()
    def rodar_tarrafa(self):
        self.fillDirs()
        # Passar como kwargs se motivo da revisão e tabela de distribuição deve ser ignorado
        ig = {"igM" : self.motivoRevisaoCheckBox.isChecked(),
              "igD" : self.listaDistribuicaoCheckBox.isChecked()}
        text = str(self.regexLineEdit.text())
        if text != "":
            self.send2server("search_regex", text, self.outputDir, kwargs=ig)

    @QtCore.Slot()
    def clearLog(self):
        self.resultTextEdit.clear()
    
    def send2server(self, comando, *args, **kwargs):
        dict2send = {"comando": comando, "args": args}
        if kwargs:
            dict2send.update(kwargs)
        json_str = json.dumps(dict2send)
        json_byte_array = QByteArray(json_str.encode("utf-8"))
        self.client.write(json_byte_array)

    def setResultTextEdit(self, message):
        message_dict = eval(message)
        received_list = message_dict['received']
        received_joined  = '\n'.join(received_list)
        if message_dict["status"] == "onload":
            self.resultTextEdit.append(received_joined)
        if message_dict["status"] == "message":
            self.messageLabel.setStyleSheet("")
            self.messageLabel.setText("Carregando..." if received_joined == "" else received_joined)
        if message_dict["status"] == "success":
            self.messageLabel.setStyleSheet("color: green")
            self.messageLabel.setText("Sucesso!" if received_joined == "" else received_joined)
        if message_dict["status"] == "warning":
            self.messageLabel.setStyleSheet("color: orange")
            self.messageLabel.setText(received_joined)
        if message_dict["status"] == "error":
            self.messageLabel.setStyleSheet("color: red")
            self.messageLabel.setText(received_joined)

    def composeLog(self) -> str:
        self.fillDirs()
        logtemplate = open(os.path.join(self.client.base_path, "logtemplate.txt"), "r", encoding="utf-8")
        logtempStr = logtemplate.read()
        logtemplate.close()
        logtempStr = logtempStr.replace(r"{{inputDir}}", self.inputDir)
        logtempStr = logtempStr.replace(r"{{outputDir}}", self.outputDir)
        logtempStr = logtempStr.replace(r"{{regexText}}", self.regexLineEdit.text())
        logtempStr = logtempStr.replace(r"{{igMotivo}}", str(self.motivoRevisaoCheckBox.isChecked()))
        logtempStr = logtempStr.replace(r"{{igTabDist}}", str(self.listaDistribuicaoCheckBox.isChecked()))
        # append log
        logtempStr = logtempStr + "\n" + self.resultTextEdit.toPlainText()
        return logtempStr

    def updatelogfile(self):
        logfile = open(os.path.join(os.getcwd(), "log.txt"), "w+", encoding="utf-8")
        logfile.write(self.composeLog())
        logfile.close()

    def closeEvent(self, event:QCloseEvent):
        self.updatelogfile()
        event.ignore

class MyTcpClient(QTcpSocket):
    message_receive = Signal(str)
    def __init__(self):
        super().__init__()
        # Set basepath for _internal files
        try:
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")
        
        # Keep this uncommented for initiating background server
        self.p = None
        self.start_process()

        self.connected.connect(self.on_connected)
        self.readyRead.connect(self.read_data)
        self.errorOccurred.connect(self.on_error)
        
        # Regex matches
        self.re_separador = re.compile(r"\{.*?\}") # Lazy match


    def connect_to_server(self, host, port):
        self.connectToHost(host, port)
        if not self.waitForConnected(3000):
            print(f"Connection failed: {self.errorString()}")
            QCoreApplication.instance().quit()

    def on_connected(self):
        print("Connected to server")

    def read_data(self):
        while self.bytesAvailable() > 0:
            data = self.readAll()
            decodedData = data.data().decode("utf-8")
            for match in self.re_separador.findall(decodedData):
                self.message_receive.emit(match)

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
            try:
                self.p.start(os.path.join(os.getcwd(),"tarrafa_server.exe"))
            except:
                print(f"Arquivo tarrafa_server.exe não encontrado.\npython .\\tarrafa_server")
                path = os.path.join(self.base_path, "tarrafa_server.py")
                self.p.start("python", [path])
                
    
    def handle_stderr(self):
        data = self.p.readAllStandardError()
        stderr = data.data().decode("latin-1")
        print(stderr)

    def handle_stdout(self):
        data = self.p.readAllStandardOutput()
        stdout = data.data().decode("latin-1")
        print(stdout)

    def handle_state(self, state):
        states = {
            QProcess.NotRunning: 'Not running',
            QProcess.Starting: 'Starting',
            QProcess.Running: 'Running',
        }
        state_name = states[state]
        print(f"State changed: {state_name}")

    def process_finished(self):
        print("Process Finished")
        self.p = None
if __name__ == "__main__":
    app = QApplication([])

    widget = MainWindow()
    widget.resize(400, 600)
    widget.show()

sys.exit(app.exec())