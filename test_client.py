import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QTextEdit
from PySide6.QtCore import QThread, Signal
import socketio

class SocketClient(QThread):
    message_received = Signal(str)

    def __init__(self):
        super().__init__()
        self.sio = socketio.Client()

    def run(self):
        @self.sio.event
        def connect():
            print("Connection established")
            self.sio.send("User has connected!")

        @self.sio.event
        def message(data):
            print("Message received:", data)
            self.message_received.emit(data)
        
        @self.sio.event
        def response(data):
            print("Response from server:", data)
            self.message_received.emit(data)

        @self.sio.event
        def disconnect():
            print("Disconnected from server")

        self.sio.connect('http://localhost:5000')

    def send_message(self, message):
        self.sio.send(message)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.socket_client = SocketClient()
        self.socket_client.message_received.connect(self.display_message)

        self.initUI()
        self.socket_client.start()

    def initUI(self):
        self.setWindowTitle('SocketIO PySide6 Client')
        self.setGeometry(300, 300, 400, 300)

        self.text_edit = QTextEdit(self)
        self.text_edit.setReadOnly(True)

        self.send_button = QPushButton('Send Message', self)
        self.send_button.clicked.connect(self.send_message)

        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addWidget(self.send_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def display_message(self, message):
        self.text_edit.append(f"Server: {message}")

    def send_message(self):
        self.socket_client.send_message("Hello from PySide6 Client!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())