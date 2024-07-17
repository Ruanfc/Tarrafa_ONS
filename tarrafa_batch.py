import os
from glob import glob
import re
import time

# from spire.doc import *
# from spire.doc.common import *
import docx2txt
import multiprocessing as mp

def find_ext(dr, ext):
    return glob(os.path.join(dr, "**/[A-Z]*.{}".format(ext)), recursive=True)

def convertWorker(filename):
    print(filename)
    text = docx2txt.process(filename)
    f = open(os.path.splitext(filename)[0] + ".txt", "w", encoding="utf-8")
    f.write(text)
    f.close()
    return filename


def convertAll(output_dir = os.getcwd()):
    pool = mp.Pool(10)
    results =  pool.imap_unordered(convertWorker, find_ext(output_dir, "docx"))
    for result in results:
        print(result)
    pool.close()
    pool.join()

###

def regexWorker(txtfilename, re_input):
    txt_file = open(txtfilename, "r", encoding="utf-8")
    text = txt_file.read()
    txt_file.close()
    x = re_input.findall(text)
    if x:
        txtfilename = txtfilename.split('\\')[-1]
        return txtfilename
        
def confirmaMatch(x):
    if x != None:
        print(x)

def search_regex(input_string):
    re_input = re.compile(input_string, re.I)
    results = []
    # with mp.Pool(processes=10) as pool:
    pool = mp.Pool(processes=10)
    for txtfilename in find_ext(os.getcwd(), "txt"):
        result = pool.apply_async(regexWorker, (txtfilename, re_input), callback= confirmaMatch)
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
    start_time = time.time()
    # convertAll()
    search_regex(r"TBE[\s-]*Sul")
    print("--- %s seconds ---" % (time.time() - start_time))