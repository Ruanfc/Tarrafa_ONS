import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import openpyxl
import docx2txt
from concurrent.futures import ThreadPoolExecutor, as_completed

class DocxSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tarrafa PDP - Busca em Documentos .docx")
        self.geometry("800x600")
        self.init_ui()

    def init_ui(self):
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate")
        self.search_label = tk.Label(self, text="Texto para Buscar:")
        self.search_input = tk.Entry(self, width=50)
        self.directory_label = tk.Label(self, text="Diretório de Busca:")
        self.directory_input = tk.Entry(self, width=50)
        self.browse_button = tk.Button(self, text="Procurar", command=self.browse_directory)
        self.output_label = tk.Label(self, text="Diretório de Saída:")
        self.output_input = tk.Entry(self, width=50)
        self.output_browse_button = tk.Button(self, text="Procurar", command=self.browse_output_directory)
        self.convert_button = tk.Button(self, text="Criar Base de Dados", command=self.convert_docx_to_txt)
        self.ignore_sections_var = tk.BooleanVar()
        self.ignore_sections_check = tk.Checkbutton(self, text="Ignorar seções entre 'MOTIVO DA REVISÃO' e 'LISTA DE DISTRIBUIÇÃO'", variable=self.ignore_sections_var)
        self.ignore_sections_var2 = tk.BooleanVar()
        self.ignore_sections_check2 = tk.Checkbutton(self, text="Ignorar seções entre 'LISTA DE DISTRIBUIÇÃO' e 'ÍNDICE'", variable=self.ignore_sections_var2)
        self.search_button = tk.Button(self, text="Buscar na Base de Dados", command=self.start_search)
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
        self.output_input.grid(row=3, column=1, padx=5, pady=5, sticky='we')
        self.output_browse_button.grid(row=3, column=2, padx=5, pady=5, sticky='w')
        self.convert_button.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.ignore_sections_check.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.ignore_sections_check2.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.search_button.grid(row=7, column=1, padx=10, pady=5, sticky='e')
        self.exit_button.grid(row=7, column=2, padx=10, pady=5, sticky='w')
        self.message_label.grid(row=8, column=0, columnspan=3, padx=10, pady=5, sticky='we')
        self.results_list.grid(row=9, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

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
        self.grid_rowconfigure(9, weight=1)

    def get_cheat_sheet(self):
        return (
            "Cheat Sheet de Expressões Regulares\n"
            "\n"
            "TLDR:\n"
            "Procurar agentes: [\\W]Agente[\\W]\n"
            "Procurar LT: [\\W]LT[\\W*]230[\\W*]kV[\\W*]Fulano[\\W*]de[\\W*]Tal[\\W*]\/[\\W*]Sicrano[\\W*]de[\\W*]Tal[\\W]\n"
            "\n"
        )

    def browse_directory(self):
        directory = filedialog.askdirectory()
        self.directory_input.delete(0, tk.END)
        self.directory_input.insert(0, directory)

    def browse_output_directory(self):
        output_directory = filedialog.askdirectory()
        self.output_input.delete(0, tk.END)
        self.output_input.insert(0, output_directory)

    def convert_docx_to_txt(self):
        input_directory = self.directory_input.get()
        output_directory = self.output_input.get()
        if input_directory and output_directory:
            self.message_label.config(text="Convertendo .docx para .txt...", fg="black")
            self.progress_bar['value'] = 0
            self.total_files = sum([len(files) for _, _, files in os.walk(input_directory)])
            self.processed_files = 0
            self.update_progress()  # Initialize progress bar
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = []
                for root, _, files in os.walk(input_directory):
                    for file in files:
                        if file.endswith(".docx"):
                            input_path = os.path.join(root, file)
                            relative_path = os.path.relpath(root, input_directory)
                            output_path = os.path.join(output_directory, relative_path)
                            if not os.path.exists(output_path):
                                os.makedirs(output_path)
                            futures.append(executor.submit(self.convert_file, input_path, os.path.join(output_path, os.path.splitext(file)[0] + ".txt")))
                for future in as_completed(futures):
                    future.result()  # To raise exceptions if any occurred
            self.message_label.config(text="Conversão concluída.", fg="green")
        else:
            messagebox.showwarning("Erro de Entrada", "Por favor, forneça o diretório de entrada e o diretório de saída.")

    def convert_file(self, input_path, output_path):
        try:
            text = docx2txt.process(input_path)
            with open(output_path, 'w', encoding='utf-8') as file:
                file.write(text)
            self.processed_files += 1
            self.report_progress(f"Convertido: {output_path}")
            self.update_progress()
        except Exception as e:
            self.report_error(f"Falha ao converter {input_path}: {e}")

    def start_search(self):
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
            self.update_progress()  # Initialize progress bar
            self.files = []
            for root, _, files in os.walk(search_directory):
                for file in files:
                    if file.endswith(".txt"):
                        self.files.append(os.path.join(root, file))
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = [executor.submit(self.process_file, file, search_string) for file in self.files]
                for future in as_completed(futures):
                    future.result()  # To raise exceptions if any occurred
            self.save_results()
            self.update_results([file for file, _ in self.doclist])
        else:
            messagebox.showwarning("Erro de Entrada", "Por favor, forneça o texto para buscar, o diretório de busca e o diretório de saída.")

    def process_file(self, filename, input_string):
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                text = file.read()
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
            self.report_progress(f"Analisado: {filename}")
        except Exception as e:
            self.report_error(f"Falha ao processar {filename}: {e}")
        finally:
            self.processed_files += 1
            self.update_progress()

    def remove_ignored_sections(self, text, start_marker, end_marker):
        pattern = re.compile(f"(?={start_marker}).*?(?<={end_marker})", re.DOTALL)
        return re.sub(pattern, "", text)

    def save_results(self):
        self.save_results_to_excel()
        self.save_results_to_txt()

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

    def save_results_to_txt(self):
        try:
            output_path = os.path.join(self.output_directory, "resultados_busca.txt")
            with open(output_path, 'w', encoding='utf-8') as file:
                file.write(', '.join([file[0] for file in self.doclist]))
            self.report_message(f"Resultados salvos em {output_path}")
        except Exception as e:
            self.report_error(f"Falha ao salvar os resultados no arquivo de texto: {e}")

    def update_results(self, filenames):
        self.message_label.config(text="Busca concluída. Resultados salvos no Excel e no arquivo de texto.", fg="green")
        self.progress_bar['value'] = 100

    def update_progress(self):
        progress = (self.processed_files / self.total_files) * 100
        self.progress_bar['value'] = progress
        self.update_idletasks()

    def report_progress(self, message):
        self.message_label.config(text=message, fg="black")
        self.update_idletasks()

    def report_error(self, message):
        self.message_label.config(text=message, fg="red")
        self.update_idletasks()

    def report_message(self, message):
        self.message_label.config(text=message, fg="black")
        self.update_idletasks()

if __name__ == '__main__':
    app = DocxSearchApp()
    app.mainloop()
