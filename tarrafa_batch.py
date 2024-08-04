import os
from glob import glob
import re
import time

import pdfminer
from pdfminer.layout import LAParams
from pdfminer.high_level import extract_text
import docx2txt
import multiprocessing as mp

class Tarrafa:
    def __init__(self, input_dir = os.getcwd(), output_dir = os.getcwd(), cores = mp.cpu_count() -2):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.cores = cores

    def find_ext(self, dr, ext):
        return glob(os.path.join(dr, "**/[A-Z]*.{}".format(ext)), recursive=True)

    def convertWorker(self, dirs):
        input_filename = dirs[0]
        extension = input_filename.split(".")[-1]
        if extension == "pdf":
            lap = LAParams(detect_vertical=True)
            text = extract_text(input_filename, laparams= lap)
        else:
            text = docx2txt.process(input_filename)
        output_filename = os.path.splitext(dirs[1])[0] + ".txt"
        os.makedirs(os.path.dirname(output_filename), exist_ok = True)
        f = open(output_filename, "w+", encoding="utf-8")
        f.write(text)
        f.close()
        return output_filename

    def convertAll(self, extension = "docx"):
        pool = mp.Pool(self.cores)
        list_input_files = self.find_ext(self.input_dir, extension)
        list_output_files = [ x.replace(self.input_dir, self.output_dir) for x in list_input_files]
        results =  pool.imap_unordered(self.convertWorker, tuple(zip(list_input_files, list_output_files)))
        listaFinal = []
        for result in results:
            result2send = result.replace("\\", "/")
            listaFinal.append(result2send)
        pool.close()
        pool.join()
        return listaFinal
    ###

    # Worker utilizado para captura de padrão regex
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
            
    def confirmaMatch(self,x):
        if x != None:
            print(x)
            # pass

    def search_regex(self, input_string, igM = False, igD = False):
        re_input = re.compile(input_string, re.I)
        results = []
        pool = mp.Pool(processes=self.cores)
        for txtfilename in self.find_ext(self.output_dir, "txt"):
            result = pool.apply_async(self.regexWorker,
                                      args=(txtfilename, re_input),
                                      kwds={"igM" : igM, "igD" : igD},
                                      callback= self.confirmaMatch)
            results.append(result)
        # Espaço para mais código
        listaFinal = []
        for result in results:
            elemento = result.get()
            if elemento:
                listaFinal.append(elemento)
        pool.close()
        pool.join()
        self.doclist = listaFinal
        return listaFinal

    def remove_ignored_sections(self, text, re_ignored):
        return re_ignored.sub("", text)

if __name__ == "__main__":
    start_time = time.time()
    print(Tarrafa.search_regex(r"Hélio Valgas")) 
    print("--- %s seconds ---" % (time.time() - start_time))