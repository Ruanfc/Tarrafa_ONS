from tarrafa_batch import Tarrafa as Tf
if __name__ == "__main__":
    tf = Tf(r"C:\\Users\\User\\OneDrive\\Documentos\\GitHub\\Tarrafa_ONS\\Documento Normativo",
            r"C:\\Users\\User\\Desktop\\saida_tarrafa\\Documento Normativo")
    # tf.convertAll("pdf")
    lt345kv = r"LT\s*345\s*kV\s*"
    SEs = [r"Furnas",
          r"Itumbiara",
          r"Jaguara",
          r"Luiz Carlos Barreto",
          r"Marimbondo",
          r"Mascarenhas de Moraes",
          r"Porto Colômbia",
          r"Estreito",
          r"Poços de Caldas",
          r"Volta Grande"]
    writefile = open("output_sample.txt", "w+", encoding="utf-8")
    for se in SEs:
        newse = SEs.copy()
        for se2 in newse:
            regex_string = lt345kv + se + r".*" + se2
            w_str = "\n\n" + se + "/" + se2 + "\n"
            tf_list = tf.search_regex(regex_string)
            w_str = w_str + "\n".join(tf_list)
            writefile.write(w_str)
    writefile.close()
