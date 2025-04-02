import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
from telas import *
from components import *
from database import create_connection


class VarejoModule:
    instance = None # Var de class que armazena a instância atual

    def __init__(self, user):
        self.user= user
        self.root = tk.Tk()
        self.root.title("Módulo Varejo")

        self.root.state('zoomed')

        VarejoModule.instance = self     


        # Lista de BOs
        self.tree = ttk.Treeview(self.root, columns=(
            "BO", "OP", "Status", "tipo_ocorrencia", "setor_responsavel", "motivo"), show="headings")
        self.tree.heading("BO", text="BO")
        self.tree.heading("OP", text="OP")
        self.tree.heading("Status", text="Status")
        self.tree.heading("tipo_ocorrencia", text="Tipo de Ocorrência")
        self.tree.heading("setor_responsavel", text="Setor Responsável")
        self.tree.heading("motivo", text="Motivo")

        # Cabeçalho
        self.header = Header(self.root, self.user, caller_id="Varejo", tree=self.tree)

        # Barra de pesquisa
        self.search_bar = SearchBar(self.root, self.pesquisar_bo, self.clear_search)

        self.tree.pack(expand=True, fill=tk.BOTH)

        self.header.carregar_bos()
        self.search_bar.search_entry.bind('<Return>', self.pesquisar_bo)
        self.root.mainloop()

    def pesquisar_bo(self, event=None):
        termo = self.search_bar.search_entry.get()
        if not termo:
            return

        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT bo_number, op, status, tipo_ocorrencia, motivo FROM bo_records WHERE modulo LIKE 'varejo' AND status NOT LIKE 'Embarcado' AND bo_number LIKE ?", (f"%{termo}%",))
            rows = cursor.fetchall()

            self.tree.delete(*self.tree.get_children())
            for row in rows:
                # Converte cada valor para string (se não for None) e aplica strip para remover caracteres indesejados.
                bo = str(row[0]).strip("(' ,)") if row[0] is not None else ""
                op = str(row[1]).strip("(' ,)") if row[1] is not None else ""
                status = str(row[2]).strip(
                    "(' ,)") if row[2] is not None else ""
                tipo_ocorrencia = str(row[3]).strip(
                    "(' ,)") if row[3] is not None else ""
                motivo = str(row[4]).strip(
                    "(' ,)") if row[4] is not None else ""

                self.tree.insert("", tk.END, values=(
                    bo, op, status, tipo_ocorrencia, motivo))
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao pesquisar BOs: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()

    def clear_search(self):
        self.search_bar.search_entry.delete(0, tk.END)
        self.search_bar.update_buttons()  # Atualiza os botões após limpar
        self.carregar_bos()  # Recarrega os BOs