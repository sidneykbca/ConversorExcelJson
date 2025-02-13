import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os
from pathlib import Path
import subprocess
import platform


class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor Excel para JSON")
        self.root.geometry("800x600")
        self.df = None
        self.selected_file = None
        self.excel_file = None
        self.output_path = None

        # Configuração do layout principal
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.post_conversion_frame = tk.Frame(main_frame)
        self.post_conversion_frame.pack(pady=5)


        # Título
        title_label = tk.Label(
            main_frame, 
            text="Conversor Excel para JSON", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Frame superior para botões e seleção de planilha
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=10)
        
        # Botões
        select_button = tk.Button(
            top_frame,
            text="Selecionar arquivo Excel",
            command=self.select_file,
            width=20,
            height=2
        )
        select_button.pack(side=tk.LEFT, padx=5)
        
        # Label para mostrar o arquivo selecionado
        self.file_label = tk.Label(
            top_frame,
            text="Nenhum arquivo selecionado",
            wraplength=500
        )
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Frame para seleção de planilha
        sheet_frame = tk.Frame(main_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(sheet_frame, text="Selecione a planilha:").pack(side=tk.LEFT, padx=5)
        self.sheet_var = tk.StringVar()
        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.sheet_var)
        self.sheet_combobox.pack(side=tk.LEFT, padx=5)
        self.sheet_combobox.bind('<<ComboboxSelected>>', self.load_selected_sheet)
        
        # Frame para visualização dos dados
        preview_frame = tk.LabelFrame(main_frame, text="Visualização dos Dados", padx=10, pady=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Treeview para mostrar os dados
        self.tree = ttk.Treeview(preview_frame)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Scrollbar para o Treeview
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Frame para seleção de chaves e valores
        selection_frame = tk.LabelFrame(main_frame, text="Seleção de Chaves e Valores", padx=10, pady=10)
        selection_frame.pack(fill=tk.X, pady=10)
        
        # Frame para chaves
        key_frame = tk.Frame(selection_frame)
        key_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(key_frame, text="Selecione a coluna para chave:").pack(side=tk.LEFT, padx=5)
        self.key_var = tk.StringVar()
        self.key_combobox = ttk.Combobox(key_frame, textvariable=self.key_var)
        self.key_combobox.pack(side=tk.LEFT, padx=5)
        
        # Frame para valores
        values_frame = tk.Frame(selection_frame)
        values_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(values_frame, text="Selecione as colunas para valores:").pack(side=tk.LEFT, padx=5)
        self.values_listbox = tk.Listbox(values_frame, selectmode=tk.MULTIPLE, height=4)
        self.values_listbox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Frame para opções
        options_frame = tk.LabelFrame(main_frame, text="Opções", padx=10, pady=10)
        options_frame.pack(fill=tk.X, pady=10)
        
        # Checkbox para manter índice
        self.keep_index = tk.BooleanVar()
        index_check = tk.Checkbutton(
            options_frame,
            text="Manter índice no JSON",
            variable=self.keep_index
        )
        index_check.pack()
        
        # Botão de conversão
        convert_button = tk.Button(
            main_frame,
            text="Converter para JSON",
            command=self.convert_file,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.convert_button = convert_button
        convert_button.pack(pady=10)
        
        self.open_folder_button = tk.Button(
            self.post_conversion_frame,
            text="Abrir Pasta do Arquivo",
            command=self.open_output_folder,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.open_folder_button.pack(pady=5)


        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="",
            wraplength=500
        )
        self.status_label.pack(pady=10)

    
    def load_sheet_names(self):
        try:
            # Carrega todas as planilhas do arquivo Excel
            self.excel_file = pd.ExcelFile(self.selected_file)
            sheet_names = self.excel_file.sheet_names
            
            # Atualiza o combobox com os nomes das planilhas
            self.sheet_combobox['values'] = sheet_names
            
            # Seleciona a primeira planilha por padrão
            if sheet_names:
                self.sheet_combobox.set(sheet_names[0])
                self.load_selected_sheet(None)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilhas: {str(e)}")
    
    def load_selected_sheet(self, event):
        try:
            sheet_name = self.sheet_var.get()
            if sheet_name:
                self.df = pd.read_excel(self.selected_file, sheet_name=sheet_name)
                self.update_preview()
                self.convert_button.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha: {str(e)}")
    
    def update_preview(self):
        if self.df is not None:
            # Limpar a Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Configurar colunas
            columns = list(self.df.columns)
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            
            for column in columns:
                self.tree.heading(column, text=column)
                self.tree.column(column, width=100)  # Largura padrão
            
            # Inserir dados
            for index, row in self.df.head(100).iterrows():  # Mostrar apenas as primeiras 100 linhas
                self.tree.insert("", tk.END, values=list(row))
            
            # Atualizar combobox e listbox
            self.key_combobox['values'] = columns
            self.values_listbox.delete(0, tk.END)
            for column in columns:
                self.values_listbox.insert(tk.END, column)
    
    def select_file(self):
        filetypes = (
            ('Arquivos Excel', '*.xlsx *.xls'),
            ('Todos os arquivos', '*.*')
        )
        
        filename = filedialog.askopenfilename(
            title='Selecione um arquivo Excel',
            filetypes=filetypes
        )
        
        if filename:
            try:
                self.selected_file = filename
                self.file_label.config(text=f"Arquivo selecionado: {filename}")
                self.load_sheet_names()
                self.status_label.config(text="")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler o arquivo: {str(e)}")
    
    def open_output_folder(self):
        """Abre a pasta onde o arquivo JSON foi salvo"""
        if self.output_path:
            folder_path = os.path.dirname(self.output_path)
            try:
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", folder_path])
                else:  # Linux
                    subprocess.run(["xdg-open", folder_path])
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir pasta: {str(e)}")
    
    def convert_file(self):
        try:
            if not self.df is not None:
                raise ValueError("Nenhum arquivo carregado")
            
            # Obter a chave selecionada
            key_column = self.key_var.get()
            if not key_column:
                raise ValueError("Selecione uma coluna para chave")
            
            # Obter os valores selecionados
            selected_indices = self.values_listbox.curselection()
            if not selected_indices:
                raise ValueError("Selecione pelo menos uma coluna para valores")
            
            value_columns = [self.values_listbox.get(i) for i in selected_indices]
            
            # Criar o JSON personalizado
            json_data = {}
            for _, row in self.df.iterrows():
                key = str(row[key_column])
                values = {col: row[col] for col in value_columns}
                json_data[key] = values
            
            # Prepara o nome do arquivo de saída
            self.output_path = Path(self.selected_file).with_suffix('.json')
            
            # Salva o arquivo JSON
            with open(self.output_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            self.status_label.config(
                text=f"Conversão concluída com sucesso!\nArquivo salvo em: {self.output_path}",
                fg="green"
            )
            
            messagebox.showinfo(
                "Sucesso",
                f"Arquivo convertido com sucesso!\nSalvo em: {self.output_path}"
            )
            
            self.open_folder_button.config(state=tk.NORMAL)
            

        except Exception as e:
            self.status_label.config(
                text=f"Erro na conversão: {str(e)}",
                fg="red"
            )
            messagebox.showerror("Erro", f"Erro ao converter arquivo: {str(e)}")   

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToJsonConverter(root)
    root.mainloop()