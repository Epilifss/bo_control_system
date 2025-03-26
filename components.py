import tkinter as tk
from tkinter import ttk

# Cabeçalho dos módulos
class Header:
    def __init__(self, parent, user, caller_id=None):
        self.user = user
        self.parent = parent
        from telas import Embarcados

        self.ultimo_modulo = caller_id

        header = ttk.Frame(self.parent)
        header.pack(fill=tk.X)

        ttk.Button(header, text="Embarcados", command= lambda: Embarcados(user, caller_id=self.ultimo_modulo)).pack(side=tk.LEFT)
        ttk.Button(header, text="Estatísticas").pack(side=tk.LEFT)
        ttk.Button(header, text="Gerar Relatório").pack(side=tk.LEFT)
        ttk.Button(header, text="Logoff",
                   command=self.logoff).pack(side=tk.RIGHT)
        ttk.Button(header, text="Atualizar",
                   command=lambda: self.carregar_bos()).pack(side=tk.RIGHT)
        ttk.Button(header, text="Buscar BO",
                   command= self.buscar_bo).pack(side=tk.RIGHT)
        
    def logoff(self):
        self.parent.destroy()
        from telas import LoginWindow
        LoginWindow()

    def buscar_bo(self):
        from telas import buscarBo

        obj = buscarBo(parent=self.parent, caller_id=self.ultimo_modulo)
        resultado = obj.identificar_chamador()
        print(resultado)

class SearchBar:
    def __init__(self, parent, search_command, clear_command):
        self.parent = parent
        self.search_command = search_command
        self.clear_command = clear_command

        self.frame = ttk.Frame(self.parent)
        self.frame.pack(fill=tk.X)

        self.search_entry = ttk.Entry(self.frame)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        self.search_button = ttk.Button(
            self.frame, text="Buscar", command=self.search_command)
        self.search_button.pack(side=tk.LEFT)

        self.clear_button = ttk.Button(
            self.frame, text="Limpar", command=self.clear_command)
        self.clear_button.pack_forget()
        
        self.search_entry.bind('<KeyRelease>', self.update_buttons)
        self.search_entry.bind('<Return>', self.search_command)

    def update_buttons(self, event=None):
        termo = self.search_entry.get()
        if termo:
            self.search_button.pack_forget()  # Oculta o botão "Pesquisar"
            self.clear_button.pack(side=tk.LEFT)  # Exibe o botão "Limpar"
        else:
            self.clear_button.pack_forget()  # Oculta o botão "Limpar"
            self.search_button.pack(side=tk.LEFT)

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.update_buttons()  # Atualiza os botões após limpar
        self.search_command()  # Recarrega os BOs