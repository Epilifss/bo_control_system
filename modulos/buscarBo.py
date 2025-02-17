import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
from database import create_connection_mikonos


class buscarBo:
    def __init__(self, user):
        self.user = user
        self.root = tk.Tk()
        self.root.title("Buscar BO")
        self.root.geometry("1000x600")

        # Cabeçalho
        header = ttk.Frame(self.root)
        header.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(header, text="Fechar", command=lambda: self.root.destroy()).pack(side=tk.RIGHT)

        # Barra de pesquisa
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        ttk.Button(search_frame, text="Pesquisar", command=self.pesquisar_bo).pack(side=tk.LEFT)

        # Frame para o Treeview e a barra de rolagem
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))

        # Treeview (lista de BOs)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("C5_PEDREPR", "C6_NUM", "C6_DESCRI", "C6_DATFAT", "C5_EMISSAO", "C6_FILIAL"),
            show="headings"
        )
        self.tree.heading("C5_PEDREPR", text="BO")
        self.tree.heading("C6_NUM", text="OP")
        self.tree.heading("C6_DESCRI", text="DESCRIÇÃO")
        self.tree.heading("C6_DATFAT", text="DATA DE FATURAMENTO")
        self.tree.heading("C5_EMISSAO", text="EMISSÃO")
        self.tree.heading("C6_FILIAL", text="FILIAL")

        # Barra de rolagem vertical
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configura o Treeview para usar a barra de rolagem
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        self.tree.bind("<Double-1>", self.exibir_detalhes)


        self.center_window()
        self.carregar_bos()
        self.root.mainloop()

    def center_window(self):
        # Centraliza a janela na tela
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def carregar_bos(self):
        conn = create_connection_mikonos()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT	C5_PEDREPR, C6_NUM, C6_DESCRI, IIF(C6_DATFAT='',C6_DATFAT,CONVERT(VARCHAR,CONVERT(DATETIME,C6_DATFAT),103)) as FATURAMENTO,
                        CONVERT(VARCHAR,CONVERT(DATETIME,C5_EMISSAO),103) as EMISSAO, IIF(C6_FILIAL='0101','CMT','NAUTICA') AS FILIAL
                FROM	SC6010 SC6
                INNER JOIN SC5010 SC5 ON (SC5.C5_NUM = SC6.C6_NUM AND SC5.C5_FILIAL = SC6.C6_FILIAL AND SC5.D_E_L_E_T_ <> '*')
                WHERE	SC6.D_E_L_E_T_ <> '*'
                        AND SC6.C6_BLQ <> 'R'
                        AND SC5.C5_PEDREPR LIKE 'BO%'
                        AND C6_FILIAL IN ('0101','0201')
                ORDER BY C5_EMISSAO DESC
            """)
            rows = cursor.fetchall()

            # Remove todos os itens atuais da Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Insere os dados na Treeview
            for row in rows:
                bo = str(row[0]).strip() if row[0] is not None else ""
                op = str(row[1]).strip() if row[1] is not None else ""
                descricao = str(row[2]).strip() if row[2] is not None else ""
                data_faturamento = str(row[3]).strip() if row[3] is not None else ""
                emissao = str(row[4]).strip() if row[4] is not None else ""
                filial = str(row[5]).strip() if row[5] is not None else ""

                self.tree.insert("", tk.END, values=(bo, op, descricao, data_faturamento, emissao, filial))
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

        conn = create_connection_mikonos()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            query = """
                SELECT	C5_PEDREPR, C6_NUM, C6_DESCRI, IIF(C6_DATFAT='',C6_DATFAT,CONVERT(VARCHAR,CONVERT(DATETIME,C6_DATFAT),103)) as FATURAMENTO,
                        CONVERT(VARCHAR,CONVERT(DATETIME,C5_EMISSAO),103) as EMISSAO, IIF(C6_FILIAL='0101','CMT','NAUTICA') AS FILIAL
                FROM	SC6010 SC6
                INNER JOIN SC5010 SC5 ON (SC5.C5_NUM = SC6.C6_NUM AND SC5.C5_FILIAL = SC6.C6_FILIAL AND SC5.D_E_L_E_T_ <> '*')
                WHERE	SC6.D_E_L_E_T_ <> '*'
                        AND SC6.C6_BLQ <> 'R'
                        AND SC5.C5_PEDREPR LIKE 'BO%'
                        AND C6_FILIAL IN ('0101','0201')
                        AND C6_NUM LIKE ?
                """

            cursor.execute(query, (f"%{termo}%",))
            rows = cursor.fetchall()

            self.tree.delete(*self.tree.get_children())
            for row in rows:
                bo = str(row[0]).strip() if row[0] is not None else ""
                op = str(row[1]).strip() if row[1] is not None else ""
                descricao = str(row[2]).strip() if row[2] is not None else ""
                data_faturamento = str(row[3]).strip() if row[3] is not None else ""
                emissao = str(row[4]).strip() if row[4] is not None else ""
                filial = str(row[5]).strip() if row[5] is not None else ""

                self.tree.insert("", tk.END, values=(bo, op, descricao, data_faturamento, emissao, filial))
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao pesquisar BOs: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()

    def exibir_detalhes(self, event):
        # Obtém o item selecionado no Treeview
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            return

        # Obtém os valores do item selecionado
        valores = self.tree.item(item_selecionado, "values")

        # Cria uma nova janela para exibir os detalhes
        detalhes_janela = tk.Toplevel(self.root)
        detalhes_janela.title("Detalhes do BO")
        detalhes_janela.geometry("400x300")

        # Exibe os detalhes na nova janela
        tk.Label(detalhes_janela, text=f"BO: {valores[0]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"OP: {valores[1]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"DESCRIÇÃO: {valores[2]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"DATA DE FATURAMENTO: {valores[3]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"EMISSÃO: {valores[4]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"FILIAL: {valores[5]}").pack(pady=5)