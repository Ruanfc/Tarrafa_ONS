import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import openpyxl
import docx2txt

class DocxSearchApp(tk.Tk):
    def __init__(self):
        """
        Initialize the DocxSearchApp application.

        This sets up the main window and initializes the user interface components.
        """
        super().__init__()
        self.title("Tarrafa PDP - Busca em Documentos .docx")
        self.geometry("800x600")
        self.init_ui()

    def init_ui(self):
        """
        Initialize the user interface components.

        This method creates and places all the widgets in the main window.
        """
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate")
        self.search_label = tk.Label(self, text="Texto para Buscar:")
        self.search_input = tk.Entry(self, width=50)
        self.directory_label = tk.Label(self, text="Diretório de Busca:")
        self.directory_input = tk.Entry(self, width=50)
        self.browse_button = tk.Button(self, text="Procurar", command=self.browse_directory)
        self.output_label = tk.Label(self, text="Diretório de Saída:")
        self.output_input = tk.Entry(self, width=50)
        self.output_browse_button = tk.Button(self, text="Procurar", command=self.browse_output_directory)
        self.ignore_sections_var = tk.BooleanVar()
        self.ignore_sections_check = tk.Checkbutton(self, text="Ignorar seções entre 'MOTIVO DA REVISÃO' e 'LISTA DE DISTRIBUIÇÃO'", variable=self.ignore_sections_var)
        self.ignore_sections_var2 = tk.BooleanVar()
        self.ignore_sections_check2 = tk.Checkbutton(self, text="Ignorar seções entre 'LISTA DE DISTRIBUIÇÃO' e 'ÍNDICE'", variable=self.ignore_sections_var2)
        self.search_button = tk.Button(self, text="Buscar", command=self.start_search)
        self.exit_button = tk.Button(self, text="Sair", command=self.quit)
        self.message_label = tk.Label(self, text="", fg="red")
        self.results_list = tk.Listbox(self, width=80, height=15)

        # Layout configuration
        self.progress_bar.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky='we')
        self.search_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.search_input.grid(row=1, column=1, padx=5, pady=5, columnspan=2, sticky='we')
        self.directory_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.directory_input.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        self.browse_button.grid(row=2, column=2, padx=5, pady=5, sticky='w')
        self.output_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.output_input.grid(row=3, column=1, padx= 5, pady=5, sticky='we')
        self.output_browse_button.grid(row=3, column=2, padx=5, pady=5, sticky='w')
        self.ignore_sections_check.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.ignore_sections_check2.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.search_button.grid(row=6, column=1, padx=10, pady=5, sticky='e')
        self.exit_button.grid(row=6, column=2, padx=10, pady=5, sticky='w')
        self.message_label.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky='we')
        self.results_list.grid(row=8, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

        # Cheat Sheet for regex
        self.cheat_sheet_label = tk.Label(self, text="Cheat Sheet de Expressões Regulares:")
        self.cheat_sheet_text = tk.Text(self, width=60, height=15, wrap='word', state='normal')
        self.cheat_sheet_text.insert(tk.END, self.get_cheat_sheet())
        self.cheat_sheet_text.configure(state='disabled')

        self.cheat_sheet_label.grid(row=0, column=3, padx=10, pady=10, sticky='nw')
        self.cheat_sheet_text.grid(row=1, column=3, rowspan=8, padx=10, pady=10, sticky='nsew')

        # Grid configuration to make the UI responsive
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(3, weight=1)
        self.grid_rowconfigure(8, weight=1)

    def get_cheat_sheet(self):
        """
        Provides a cheat sheet for regular expressions.

        Returns:
            str: A string containing the cheat sheet.
        """
        return (
            "Cheat Sheet de Expressões Regulares\n"
            "\n"
            "TLDR:\n"
            "Procurar agentes: [\\W]Agente[\\W]\n"
            "Procurar LT: [\\W]LT[\\W*]230[\\W*]kV[\\W*]Fulano[\\W*]de[\\W*]Tal[\\W*]/[\\W*]Sicrano[\\W*]de[\\W*]Tal[\\W]\n"
            "\n"
            "Metacaracteres:\n"
            ".        Qualquer caractere, exceto nova linha\n"
            "^        Início da string\n"
            "$        Fim da string\n"
            "*        0 ou mais repetições\n"
            "+        1 ou mais repetições\n"
            "?        0 ou 1 repetição\n"
            "{n}      Exatamente n repetições\n"
            "{n,}     n ou mais repetições\n"
            "{n,m}    Entre n e m repetições\n"
            "[abc]    Qualquer caractere dentro do conjunto\n"
            "[^abc]   Qualquer caractere fora do conjunto\n"
            "\\d       Qualquer dígito\n"
            "\\D       Qualquer não-dígito\n"
            "\\s       Qualquer espaço em branco\n"
            "\\S       Qualquer não-espaço em branco\n"
            "\\w       Qualquer caractere alfanumérico\n"
            "\\W       Qualquer não-alfanumérico\n"
            "(...)    Captura o grupo\n"
            "(?:...)  Grupo sem captura\n"
            "(?=...)  Lookahead positivo\n"
            "(?!...)  Lookahead negativo\n"
            "\\/      Barra\n"
            "\\(      Parêntese esquerdo\n"
            "\\)      Parêntese direito\n"
            "\n"
            "Exemplos:\n"
            "a.*b           Encontra 'a' seguido de qualquer caractere (0 ou mais) seguido de 'b'\n"
            "^Hello         Encontra 'Hello' no início da string\n"
            "world$         Encontra 'world' no fim da string\n"
            "\\d{2,4}       Encontra sequências de 2 a 4 dígitos\n"
            "a{3}           Encontra exatamente três 'a'\n"
            "[A-Za-z]       Encontra qualquer letra maiúscula ou minúscula\n"
            "\\w+            Encontra uma ou mais palavras\n"
            "\\s*            Encontra zero ou mais espaços em branco\n"
            "a|b            Encontra 'a' ou 'b'\n"
            "(a|b)c         Encontra 'ac' ou 'bc'\n"
            "\\bword\\b      Encontra a palavra 'word' com limites de palavra\n"
            "(?i)abc        Encontra 'abc' case insensitive\n"
            "(?<!foo)bar    Encontra 'bar' que não é precedido por 'foo'\n"
            "(?<=foo)bar    Encontra 'bar' que é precedido por 'foo'\n"
            "abc(?=def)     Encontra 'abc' que é seguido por 'def'\n"
            "abc(?!def)     Encontra 'abc' que não é seguido por 'def'\n"
            "\\W+            Encontra uma ou mais caracteres não-alfanuméricos\n"
            "\\s+            Encontra um ou mais espaços em branco\n"
            ".*             Encontra qualquer sequência de caracteres (0 ou mais)\n"
            "\\S*            Encontra qualquer sequência de caracteres que não sejam espaços em branco (0 ou mais)\n"
            "\\d*            Encontra qualquer sequência de dígitos (0 ou mais)\n"
        )

    def browse_directory(self):
        """
        Opens a dialog to select the directory for searching .docx files.

        This method updates the directory_input Entry widget with the selected directory path.
        """
        directory = filedialog.askdirectory()
        self.directory_input.delete(0, tk.END)
        self.directory_input.insert(0, directory)

    def browse_output_directory(self):
        """
        Opens a dialog to select the output directory for saving search results.

        This method updates the output_input Entry widget with the selected output directory path.
        """
        output_directory = filedialog.askdirectory()
        self.output_input.delete(0, tk.END)
        self.output_input.insert(0, output_directory)

    def start_search(self):
        """
        Initiates the search process based on user inputs.

        This method retrieves the search string, search directory, and output directory from the user inputs,
        and starts the search process if all inputs are provided. It also handles the UI updates for search progress.
        """
        search_string = self.search_input.get()
        search_directory = self.directory_input.get()
        self.output_directory = self.output_input.get()
        self.ignore_sections = self.ignore_sections_var.get()
        self.ignore_sections2 = self.ignore_sections_var2.get()
        if search_string and search_directory and self.output_directory:
            self.message_label.config(text="Buscando...", fg="black")
            self.results_list.delete(0, tk.END)
            self.progress_bar['value'] = 0
            self.doclist = []
            self.total_files = sum([len(files) for _, _, files in os.walk(search_directory)])
            self.processed_files = 0
            self.files = []
            for root, _, files in os.walk(search_directory):
                for file in files:
                    if file.endswith(".docx"):
                        self.files.append(os.path.join(root, file))
            self.search_files(search_string)
        else:
            messagebox.showwarning("Erro de Entrada", "Por favor, forneça o texto para buscar, o diretório de busca e o diretório de saída.")

    def search_files(self, input_string):
        """
        Searches through the .docx files for the given input string.

        Args:
            input_string (str): The string or regex pattern to search for.

        This method processes each .docx file to find matches for the input string and updates the results list.
        It also handles the progress bar updates and error reporting.
        """
        if self.files:
            filename = self.files.pop(0)
            try:
                text = docx2txt.process(filename)
                if self.ignore_sections & self.ignore_sections2:
                    text = self.remove_ignored_sections(text, "MOTIVO DA REVISÃO", "ÍNDICE")
                elif self.ignore_sections:
                    text = self.remove_ignored_sections(text, "MOTIVO DA REVISÃO", "LISTA DE DISTRIBUIÇÃO")
                elif self.ignore_sections2:
                    text = self.remove_ignored_sections(text, "LISTA DE DISTRIBUIÇÃO", "ÍNDICE")
                if re.findall(input_string, text, re.I):
                    base_name = re.sub(r"_Rev\.\d+\..+", "", os.path.basename(filename))
                    self.doclist.append((base_name, filename))
                    self.results_list.insert(tk.END, base_name)
                self.report_progress(filename)
            except Exception as e:
                self.report_error(f"Falha ao processar {filename}: {e}")
            finally:
                self.processed_files += 1
                self.update_progress()
                self.after(10, self.search_files, input_string)
        else:
            self.save_results()
            self.update_results([file for file, _ in self.doclist])

    def remove_ignored_sections(self, text, start_marker, end_marker):
        """
        Removes sections of text between the given start and end markers.

        Args:
            text (str): The text to process.
            start_marker (str): The starting marker of the section to ignore.
            end_marker (str): The ending marker of the section to ignore.

        Returns:
            str: The processed text with specified sections removed.
        """
        pattern = re.compile(f"(?={start_marker}).*?(?<={end_marker})", re.DOTALL)
        return re.sub(pattern, "", text)

    def save_results(self):
        """
        Saves the search results to both Excel and text files.

        This method calls the save_results_to_excel and save_results_to_txt methods to save the search results.
        """
        self.save_results_to_excel()
        self.save_results_to_txt()

    def save_results_to_excel(self):
        """
        Saves the search results to an Excel file.

        This method creates an Excel workbook, adds the search results, and saves it to the specified output directory.
        """
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

    def save_results_to_txt(self):
        """
        Saves the search results to a text file.

        This method creates a text file, adds the search results, and saves it to the specified output directory.
        """
        try:
            output_path = os.path.join(self.output_directory, "resultados_busca.txt")
            with open(output_path, 'w') as file:
                file.write(', '.join([file[0] for file in self.doclist]))
            self.report_message(f"Resultados salvos em {output_path}")
        except Exception as e:
            self.report_error(f"Falha ao salvar os resultados no arquivo de texto: {e}")

    def update_results(self, filenames):
        """
        Updates the UI to indicate that the search is complete.

        Args:
            filenames (list): List of filenames that matched the search.

        This method updates the message label and progress bar to indicate the completion of the search.
        """
        self.message_label.config(text="Busca concluída. Resultados salvos no Excel e no arquivo de texto.", fg="green")
        self.progress_bar['value'] = 100

    def update_progress(self):
        """
        Updates the progress bar based on the number of files processed.

        This method calculates the current progress and updates the progress bar accordingly.
        """
        progress = (self.processed_files / self.total_files) * 100
        self.progress_bar['value'] = progress

    def report_progress(self, filename):
        """
        Reports the progress of the current file being processed.

        Args:
            filename (str): The name of the file being processed.

        This method updates the message label to show the current file being analyzed.
        """
        self.message_label.config(text=f"Analisando: {os.path.basename(filename)}", fg="black")

    def report_error(self, message):
        """
        Reports an error message to the UI.

        Args:
            message (str): The error message to display.

        This method updates the message label to display the error message.
        """
        self.message_label.config(text=message, fg="red")

    def report_message(self, message):
        """
        Reports a general message to the UI.

        Args:
            message (str): The message to display.

        This method updates the message label to display the provided message.
        """
        self.message_label.config(text=message, fg="black")

if __name__ == '__main__':
    app = DocxSearchApp()
    app.mainloop()
