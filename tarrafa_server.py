import os
from glob import glob
import re
import time

import docx2txt
import multiprocessing as mp

import sys
# from PySide6.QtCore import QCoreApplication, QIODevice, QByteArray
# from PySide6.QtNetwork import QTcpServer, QTcpSocket, QHostAddress
import socket
import threading
import json

class Tarrafa():
# class Tarrafa(QTcpServer):
    def __init__(self, host='127.0.0.1', port = 12345, cores = mp.cpu_count()-2):
        # pdf ou docx?
        # numero de cores
        self.cores = cores
        # TCP Server
    #     super().__init__()
    #     self.newConnection.connect(self.on_new_connection)
    #     if not self.listen(QHostAddress(host), int(port)):
    #         print(f"Não foi possível iniciar o servidor; {self.errorString()}")
    #         sys.exit(-1)
    #     print(f"O servidor está em {self.serverAddress().toString()}:{self.serverPort()}")

    # def on_new_connection(self):
    #     client_connection = self.nextPendingConnection()
    #     client_connection.readyRead.connect(self.read_data)
    #     client_connection.disconnected.connect(client_connection.deleteLater)

    # def read_data(self):
    #     client_connection = self.sender()
    #     while client_connection.bytesAvailable():
    #         data = client_connection.readAll()
    #         incoming_message = data.data().decode()
    #         print(f"Received: {incoming_message}")
    #         #threading.Thread(target=self.handle_long_computation, args=(incoming_message,client_connection)).start()
    #         result = self.search_regex(incoming_message) 
    #         # result = "Hello, Client"
    #         response = QByteArray(bytes(result, 'utf-8'))
    #         client_connection.write(response)

    #  def handle_long_computation(self, message, client_connection):
    #     self.search_regex(r"{}".format(message))
    #     response = QByteArray(bytes(message, 'utf-8'))
    #     client_connection.write(response)

        self.host = host
        self.port = port
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server_socket.bind((self.host, self.port))
        self.server_socket.listen(5)
        print(f'Server started on {self.host}:{self.port}')
    
    def start_server(self):
        while True:
            client_socket, client_address = self.server_socket.accept()
            print(f'Nova conexão de {client_address}')
            # threading.Thread(target=self.handle_client, args=(client_socket,)).start()
            self.handle_client(client_socket)
    
    def handle_client(self, client_socket):
        with client_socket:
            while True:
                try:
                    message = client_socket.recv(1024).decode('utf-8')
                    if not message:
                        break
                    # data = json.load(message)
                    data = message
                    print(f'Dados recebidos: {data}')
                    # Processa os dados e prepara resposta
                    if data != "Hello, Server!":
                        data = self.search_regex(r"{}".format(data))
                    response = {"status": "success", "received": data}
                    response_message = json.dumps(response)
                    client_socket.sendall(response_message.encode('utf-8'))
                except json.JSONDecodeError:
                    print('Received invalid JSON data')
                    response = {"status": "error", "message": "Invalid JSON"}
                    client_socket.sendall(json.dumps(response).encode("utf-8"))
                except Exception as e:
                    print(f'Error: {e}')
                    break

###
    def find_ext(self, dr, ext):
        return glob(os.path.join(dr, "**/[A-Z]*.{}".format(ext)), recursive=True)

    def convertWorker(self, filename):
        print(filename)
        text = docx2txt.process(filename)
        f = open(os.path.splitext(filename)[0] + ".txt", "w", encoding="utf-8")
        f.write(text)
        f.close()
        return filename


    def convertAll(self, output_dir = os.getcwd()):
        pool = mp.Pool(10)
        results =  pool.imap_unordered(self.convertWorker, self.find_ext(output_dir, "docx"))
        for result in results:
            print(result)
        pool.close()
        pool.join()

    ###

    def regexWorker(self, txtfilename, re_input):
    #def regexWorker(self, txtfilename, re_input, igM=False, igD=False):
        txt_file = open(txtfilename, "r", encoding="utf-8")
        text = txt_file.read()
        txt_file.close()
        # Utilizar ignorador
        x = re_input.findall(text)
        if x:
            txtfilename = txtfilename.split('\\')[-1]
            return txtfilename
            
    def confirmaMatch(self, x):
        if x != None:
            print(x)

    def search_regex(self, input_string):
    # def search_regex(self, input_string, ignoraMotivoRevisao = False, ignoraListaDistribuicao = False):
        re_input = re.compile(input_string, re.I)
        results = []
        pool = mp.Pool(processes=self.cores)
        for txtfilename in self.find_ext(os.getcwd(), "txt"):
            result = pool.apply_async(self.regexWorker, (txtfilename, re_input), callback= self.confirmaMatch)
            results.append(result)
        # Mais código aqui
        listaFinal = []
        for result in results:
            elemento = result.get()
            if elemento:
                listaFinal.append(elemento)

        pool.close()
        pool.join()
        print(listaFinal)
        return listaFinal


if __name__ == "__main__":
    # app = QCoreApplication(sys.argv)
    start_time = time.time()
    # convertAll()
    tarrafa = Tarrafa()
    tarrafa.start_server()
    # tarrafa.search_regex(r"Recife II")

    # sys.exit(app.exec())
    print("--- %s seconds ---" % (time.time() - start_time))