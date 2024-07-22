import os
from glob import glob
import re
import time

import openpyxl
import docx2txt
import multiprocessing as mp

import socket
import json

class Tarrafa():
    def __init__(self, host='127.0.0.1', port = 12345, cores = mp.cpu_count()-2):
        # pdf ou docx?
        # numero de cores
        self.cores = cores

        self.host = host
        self.port = port
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server_socket.bind((self.host, self.port))
        self.server_socket.listen(5)
        print(f'Server started on {self.host}:{self.port}')

        # Padrões regex previamente compilados para ganhos de performance
        ig_x = lambda start_marker, end_marker : f"(?<={start_marker}).*?(?={end_marker})"
        self.re_MD = re.compile(ig_x("MOTIVO DA REVISÃO", "ÍNDICE"), re.DOTALL)
        self.re_M = re.compile(ig_x("MOTIVO DA REVISÃO", "TABELA DE DISTRIBUIÇÃO"), re.DOTALL)
        self.re_D = re.compile(ig_x("TABELA DE DISTRIBUIÇÃO", "ÍNDICE"), re.DOTALL)
    
    def start_server(self):
        while True:
            client_socket, client_address = self.server_socket.accept()
            print(f'Nova conexão de {client_address}')
            """
            Normalmente, o ideal é que o servidor seja criado em paralelo para deixar a main thread livre,
            mas como este processo já é o processo em background, então não precisa. Aninhar multiprocessing
            em threading trás problemas imprevisíveis.
            """
            # threading.Thread(target=self.handle_client, args=(client_socket,)).start()
            self.handle_client(client_socket)
    
    """
    Conecta cliente ao servidor e espera por novas mensagens
    """
    def handle_client(self, client_socket):
        with client_socket:
            while True:
                try:
                    message = client_socket.recv(1024).decode('utf-8')
                    if not message:
                        break
                    data = json.loads(message)
                    # data = dict(message)
                    print(f'Dados recebidos: {data} no tipo {type(data)}')
                    # Processa os dados e prepara resposta
                    return_obj = eval("self." + data["comando"] + f"(*{data['args']})")
                    response = {"status": "success", "received": return_obj}
                    response_message = json.dumps(response)
                    client_socket.sendall(response_message.encode('utf-8'))
                except json.JSONDecodeError:
                    print('Received invalid JSON data')
                    response = {"status": "error", "message": "Invalid JSON"}
                    client_socket.sendall(json.dumps(response).encode("utf-8"))
                except Exception as e:
                    print(f'Error: {e}')
                    break

    # Acha arquivos em um diretório recursivamente pela sua extensão
    def find_ext(self, dr, ext):
        return glob(os.path.join(dr, "**/[A-Z]*.{}".format(ext)), recursive=True)

    # Worker utilizado para conversão dos documentos normativos em txt
    def convertWorker(self, dirs):
        input_filename = dirs[0]
        text = docx2txt.process(input_filename)
        output_filename = os.path.splitext(dirs[1])[0] + ".txt"
        os.makedirs(os.path.dirname(output_filename), exist_ok = True)
        f = open(output_filename, "w+", encoding="utf-8")
        f.write(text)
        f.close()
        return output_filename

    # [FRONTEND] Converte todos os documentos docx em txt.
    def convertAll(self, *args):
        input_dir = args[0]
        output_dir = args[1]
        pool = mp.Pool(self.cores)
        list_input_files = self.find_ext(input_dir, "docx")
        list_output_files = [ x.replace(input_dir, output_dir) for x in list_input_files]
        results =  pool.imap_unordered(self.convertWorker, tuple(zip(list_input_files, list_output_files)))
        listaFinal = []
        for result in results:
            print(result)
            listaFinal.append(result)
        pool.close()
        pool.join()
        return listaFinal

    # Worker utilizado para captura de padrão regex
    # def regexWorker(self, txtfilename, re_input):
    def regexWorker(self, txtfilename, re_input, igM=False, igD=False):
        txt_file = open(txtfilename, "r", encoding="utf-8")
        text = txt_file.read()
        txt_file.close()
        # Utilizar ignorador
        if igM and igD:
            text = self.remove_ignored_sections(text, self.re_MD)
        elif igM:
            text = self.remove_ignored_sections(text, self.re_M)
        elif igD:
            text = self.remove_ignored_sections(text, self.re_D)
        x = re_input.findall(text)
        if x:
            txtfilename = txtfilename.split('\\')[-1]
            return txtfilename
            
    # Método usado como callback para confirmar match do regex sempre que o regexWorker finaliza em cada arquivo txt
    def confirmaMatch(self, x):
        if x != None:
            print(x)

    # [FRONTEND] Método para fazer a busca em todos os arquivos.
    def search_regex(self, input_string):
    # def search_regex(self, input_string, ignoraMotivoRevisao = False, ignoraListaDistribuicao = False, searchDir=os.getcwd()):
        re_input = re.compile(input_string, re.I)
        results = []
        pool = mp.Pool(processes=self.cores)
        for txtfilename in self.find_ext(os.getcwd(), "txt"):
            result = pool.apply_async(self.regexWorker, (txtfilename, re_input), callback= self.confirmaMatch)
            results.append(result)
        # Espaço para mais código
        listaFinal = []
        for result in results:
            elemento = result.get()
            if elemento:
                listaFinal.append(elemento)
        pool.close()
        pool.join()
        return listaFinal

    def remove_ignored_sections(self, text, re_ignored):
        return re_ignored.sub("", text)

    # [FRONTEND] Salva resultados tanto em excel como em txt
    def save_results(self):
        self.save_results_to_excel()
        self.save_results_to_txt()

    # [FRONTEND] Adiciona os resultados de pesquisa a um novo arquivo em excel
    def save_results_to_excel(self):
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Resultados da Busca"
            sheet.append(["Nome do Arquivo", "Caminho do Arquivo"])

            for file, path in self.doclist:
                sheet.append([file, path])

            output_path = os.path.join(self.output_directory, "resultados_busca.xlsx")
            workbook.save(output_path)
            self.report_message(f"Resultados salvos em {output_path}")
        except Exception as e:
            self.report_error(f"Falha ao salvar os resultados no Excel: {e}")
    # [FRONTEND] Salva o resultado das buscas em um arquivo txt
    def save_results_to_txt(self):
        try:
            output_path = os.path.join(self.output_directory, "resultados_busca.txt")
            with open(output_path, 'w') as file:
                file.write(', '.join([file[0] for file in self.doclist]))
            self.report_message(f"Resultados salvos em {output_path}")
        except Exception as e:
            self.report_error(f"Falha ao salvar os resultados no arquivo de texto: {e}")

    def report_error(self, message):
        pass
    def report_message(self, message):
        pass


if __name__ == "__main__":
    start_time = time.time()
    # convertAll()
    tarrafa = Tarrafa()
    tarrafa.start_server()
    # tarrafa.search_regex(r"Recife II")

    print("--- %s seconds ---" % (time.time() - start_time))