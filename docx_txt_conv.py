import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import docx2txt
from concurrent.futures import ThreadPoolExecutor, as_completed

class DocxToTxtConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Converter .docx para .txt")
        self.geometry("600x400")
        self.init_ui()

    def init_ui(self):
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate")
        self.input_label = tk.Label(self, text="Diretório de Entrada:")
        self.input_entry = tk.Entry(self, width=50)
        self.input_browse_button = tk.Button(self, text="Procurar", command=self.browse_input_directory)
        self.output_label = tk.Label(self, text="Diretório de Saída:")
        self.output_entry = tk.Entry(self, width=50)
        self.output_browse_button = tk.Button(self, text="Procurar", command=self.browse_output_directory)
        self.convert_button = tk.Button(self, text="Converter", command=self.convert_docx_to_txt)
        self.message_label = tk.Label(self, text="", fg="red")

        # Layout configuration
        self.progress_bar.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky='we')
        self.input_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.input_entry.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        self.input_browse_button.grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.output_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.output_entry.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        self.output_browse_button.grid(row=2, column=2, padx=5, pady=5, sticky='w')
        self.convert_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky='we')
        self.message_label.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky='we')

        # Grid configuration to make the UI responsive
        self.grid_columnconfigure(1, weight=1)

    def browse_input_directory(self):
        directory = filedialog.askdirectory()
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, directory)

    def browse_output_directory(self):
        output_directory = filedialog.askdirectory()
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, output_directory)

    def convert_docx_to_txt(self):
        input_directory = self.input_entry.get()
        output_directory = self.output_entry.get()
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

if __name__ == '__main__':
    app = DocxToTxtConverterApp()
    app.mainloop()
