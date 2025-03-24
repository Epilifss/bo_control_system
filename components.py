import tkinter as tk
from tkinter import ttk

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