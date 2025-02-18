import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
from telas import *
from modulos.embarcados import Embarcados
from database import create_connection


class CorporativoModule:
    def __init__(self, user):
        self.user= user
        self.root = tk.Tk()
        self.root.title("Módulo Corporativo")

        self.root.state('zoomed')
     

        # Cabeçalho
        header = ttk.Frame(self.root)
        header.pack(fill=tk.X)

        ttk.Button(header, text="Embarcados", command= lambda: Embarcados(user)).pack(side=tk.LEFT)
        ttk.Button(header, text="Estatísticas").pack(side=tk.LEFT)
        ttk.Button(header, text="Gerar Relatório").pack(side=tk.LEFT)
        ttk.Button(header, text="Logoff",
                   command=self.logoff).pack(side=tk.RIGHT)
        ttk.Button(header, text="Atualizar",
                   command=lambda: self.carregar_bos()).pack(side=tk.RIGHT)
        ttk.Button(header, text="Buscar BO",
                   command= self.buscar_bo).pack(side=tk.RIGHT)

        # Barra de pesquisa
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill=tk.X)
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Button(search_frame, text="Pesquisar",
                   command=self.pesquisar_bo).pack(side=tk.LEFT)

        # Lista de BOs
        self.tree = ttk.Treeview(self.root, columns=(
            "BO", "OP", "Status", "tipo_ocorrencia", "motivo"), show="headings")
        self.tree.heading("BO", text="BO")
        self.tree.heading("OP", text="OP")
        self.tree.heading("Status", text="Status")
        self.tree.heading("tipo_ocorrencia", text="Tipo de Ocorrência")
        self.tree.heading("motivo", text="Motivo")

        self.tree.pack(expand=True, fill=tk.BOTH)

        self.carregar_bos()
        self.root.mainloop()

    def carregar_bos(self):
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT bo_number, op, status, tipo_ocorrencia, motivo FROM bo_records WHERE status NOT LIKE 'Embarcado' AND modulo LIKE 'corporativo'")
            rows = cursor.fetchall()

            # Remove todos os itens atuais da Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Insere os dados "limpos" na Treeview
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
            messagebox.showerror("Erro", f"Erro ao carregar BOs: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()

    def pesquisar_bo(self):
        termo = self.search_entry.get()
        if not termo:
            return

        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT bo_number, op, status, tipo_ocorrencia, motivo FROM bo_records WHERE modulo LIKE 'corporativo' AND status NOT LIKE 'Embarcado' AND bo_number LIKE ?", (f"%{termo}%",))
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

    def buscar_bo(self):
        from modulos.buscarBo import buscarBo
        obj = buscarBo(self.root)
        resultado = obj.identificar_chamador("corporativo")
        print(resultado)

    def logoff(self):
        self.root.destroy()
        from telas import LoginWindow
        LoginWindow()