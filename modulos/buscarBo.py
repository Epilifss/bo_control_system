import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
from tkcalendar import DateEntry
from database import create_connection_mikonos
from database import create_connection


class buscarBo:
    def __init__(self, parent):
        self.parent = parent
        self.root = tk.Toplevel(parent)
        self.root.title("Buscar BO")
        self.root.geometry("1000x600")
        self.ultimo_modulo = None

        # Cabeçalho
        header = ttk.Frame(self.root)
        header.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(header, text="Fechar",
                   command=lambda: self.root.destroy()).pack(side=tk.RIGHT)

        # Barra de pesquisa
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, expand=True,
                               fill=tk.X, padx=(0, 10))
        ttk.Button(search_frame, text="Pesquisar",
                   command=self.pesquisar_bo).pack(side=tk.LEFT)

        # Frame para o Treeview e a barra de rolagem
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))

        # Treeview (lista de BOs)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("C5_PEDREPR", "C5_NUM", "C5_NOME",
                     "C5_FILIAL", "C5_EMISSAO", "C5_ENTREGA"),
            show="headings"
        )
        self.tree.heading("C5_PEDREPR", text="BO")
        self.tree.heading("C5_NUM", text="OP")
        self.tree.heading("C5_NOME", text="CLIENTE")
        self.tree.heading("C5_FILIAL", text="FILIAL")
        self.tree.heading("C5_EMISSAO", text="EMISSÃO")
        self.tree.heading("C5_ENTREGA", text="PREVISÃO DE ENTREGA")

        # Barra de rolagem vertical
        v_scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configura o Treeview para usar a barra de rolagem
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        self.tree.bind("<Double-1>", self.exibir_detalhes)

        self.center_window()
        self.carregar_bos()

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

    def identificar_chamador(self, caller_id=None):
        if caller_id is None:
            raise ValueError(
                "É necessário informar o identificador do módulo que chamou a função")
        self.ultimo_modulo = caller_id
        return f"Buscar BO chamado pelo módulo: {caller_id}"

    def carregar_bos(self):
        conn = create_connection_mikonos()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT C5_PEDREPR, C5_NUM, C5_NOME, IIF(C5_FILIAL='0101','CMT','NAUTICA') AS FILIAL, CONVERT(VARCHAR,CONVERT(DATETIME,C5_EMISSAO),103) as EMISSAO, CONVERT(VARCHAR,CONVERT(DATETIME,C5_ENTREGA),103) as PREVISAO_ENTREGA
                FROM SC5010 SC5
                WHERE SC5.D_E_L_E_T_ <> '*'
                        AND SC5.C5_PEDREPR LIKE 'BO%'
                        AND C5_FILIAL IN ('0101','0201')
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
                cliente = str(row[2]).strip() if row[2] is not None else ""
                filial = str(row[3]).strip() if row[3] is not None else ""
                emissao = str(row[4]).strip() if row[4] is not None else ""
                previsao_entrega = str(
                    row[5]).strip() if row[5] is not None else ""

                self.tree.insert("", tk.END, values=(
                    bo, op, cliente, filial, emissao, previsao_entrega))
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
                SELECT C5_PEDREPR, C5_NUM, C5_NOME, IIF(C5_FILIAL='0101','CMT','NAUTICA') AS FILIAL, CONVERT(VARCHAR,CONVERT(DATETIME,C5_EMISSAO),103) as EMISSAO, CONVERT(VARCHAR,CONVERT(DATETIME,C5_ENTREGA),103) as PREVISAO_ENTREGA
                FROM SC5010 SC5
                WHERE SC5.D_E_L_E_T_ <> '*'
                        AND SC5.C5_PEDREPR LIKE 'BO%'
                        AND C5_FILIAL IN ('0101','0201')
                        AND C5_NUM LIKE ?
                """

            cursor.execute(query, (f"%{termo}%",))
            rows = cursor.fetchall()

            self.tree.delete(*self.tree.get_children())
            for row in rows:
                bo = str(row[0]).strip() if row[0] is not None else ""
                op = str(row[1]).strip() if row[1] is not None else ""
                cliente = str(row[2]).strip() if row[2] is not None else ""
                filial = str(row[3]).strip() if row[3] is not None else ""
                emissao = str(row[4]).strip() if row[4] is not None else ""
                previsao_entrega = str(
                    row[5]).strip() if row[5] is not None else ""

                self.tree.insert("", tk.END, values=(
                    bo, op, cliente, filial, emissao, previsao_entrega))
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
        detalhes_janela.geometry("600x400")

        # Exibe os detalhes na nova janela
        tk.Label(detalhes_janela, text=f"BO: {valores[0]}").pack(
            pady=5, side=tk.LEFT)
        tk.Label(detalhes_janela, text=f"OP: {valores[1]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"CLIENTE: {valores[2]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"FILIAL: {valores[3]}").pack(pady=5)
        tk.Label(detalhes_janela, text=f"EMISSÃO: {valores[4]}").pack(pady=5)
        tk.Label(detalhes_janela,
                 text=f"PREVISÃO DE ENTREGA: {valores[5]}").pack(pady=5)

        ttk.Button(detalhes_janela, text="Acompanhar BO",
                   command=self.acompanhar_bo).pack(side=tk.LEFT)

    def acompanhar_bo(self):
        obj = acompanhar_Bo(self.root, self.tree)
        resultado = obj.identificar_chamador("corporativo")
        print(resultado)


class acompanhar_Bo:
    def __init__(self, parent, tree):
        self.window = tk.Toplevel(parent)
        self.window.title("Nova BO")
        self.window.grab_set()

        self.tree = tree
        self.ultimo_modulo = None
        self.bo_dados = self.obter_dados_bo()

        self.criar_secao_dados_gerais()
        self.criar_secao_ocorrencia()
        self.criar_secao_transporte()
        self.criar_botao_salvar()

    def criar_secao_dados_gerais(self):
        """Cria a seção de dados gerais da BO."""
        frame_dados_gerais = ttk.LabelFrame(self.window, text="Dados Gerais", padding=10)
        frame_dados_gerais.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        # BO e OP
        ttk.Label(frame_dados_gerais, text="BO:").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[0]).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text="OP:").grid(row=1, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[1]).grid(row=1, column=1, sticky=tk.W)

        # Cliente
        ttk.Label(frame_dados_gerais, text="CLIENTE:").grid(row=2, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[2]).grid(row=2, column=1, sticky=tk.W)

    def criar_secao_ocorrencia(self):
        """Cria a seção de detalhes da ocorrência."""
        frame_ocorrencia = ttk.LabelFrame(self.window, text="Detalhes da Ocorrência", padding=10)
        frame_ocorrencia.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        campos_ocorrencia = [
            ("Tipo de Ocorrência", ttk.Entry),
            ("Motivo", ttk.Combobox, self.motivos()),
            ("Descrição", ttk.Entry),
        ]

        self.entries_ocorrencia = {}
        for i, (campo, widget, *args) in enumerate(campos_ocorrencia):
            ttk.Label(frame_ocorrencia, text=f"{campo}:").grid(row=i, column=0, sticky=tk.W)
            if widget == ttk.Combobox:
                entry = widget(frame_ocorrencia, values=args[0])
            else:
                entry = widget(frame_ocorrencia)
            entry.grid(row=i, column=1, sticky="ew")
            self.entries_ocorrencia[campo] = entry

    def criar_secao_transporte(self):
        """Cria a seção de detalhes do transporte."""
        frame_transporte = ttk.LabelFrame(self.window, text="Transporte", padding=10)
        frame_transporte.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        campos_transporte = [
            ("Frete", ttk.Combobox, ["CIF", "FOB"]),
            ("Previsão de Embarque", DateEntry),
            ("NF de Envio", ttk.Entry),
            ("NF de Devolução", ttk.Entry),
        ]

        self.entries_transporte = {}
        for i, (campo, widget, *args) in enumerate(campos_transporte):
            ttk.Label(frame_transporte, text=f"{campo}:").grid(row=i, column=0, sticky=tk.W)
            if widget == ttk.Combobox:
                entry = widget(frame_transporte, values=args[0])
            elif widget == DateEntry:
                entry = widget(frame_transporte)
            else:
                entry = widget(frame_transporte)
            entry.grid(row=i, column=1, sticky="ew")
            self.entries_transporte[campo] = entry

    def criar_botao_salvar(self):
        """Cria o botão de salvar."""
        frame_botoes = ttk.Frame(self.window, padding=10)
        frame_botoes.grid(row=3, column=0, sticky="ew", padx=10, pady=10)

        ttk.Button(frame_botoes, text="Salvar", command=self.salvar).pack(side=tk.RIGHT)

    def motivos(self):
        """Retorna uma lista de motivos."""
        return [
            "Pé quebrado",
            "Tela rasgada",
            "Outro"
        ]
    
    def identificar_chamador(self, caller_id=None):

        if caller_id is None:
            raise ValueError("É necessário informar o identificador do módulo que chamou a função")
        self.ultimo_modulo = caller_id
        return f"Acompanhar BO chamado pelo módulo: {caller_id}"

    def obter_dados_bo(self):
        """Obtém os dados da BO selecionada na Treeview."""
        selecionado = self.tree.selection()
        if not selecionado:
            return None
        
        valores = self.tree.item(selecionado, "values")
        if valores:
            return valores
        return None

    def salvar(self):
        """Salva a BO no banco e só então incrementa o número da sequência."""
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()

            # Coleta os valores dos campos
            valores = [
                self.bo_dados[0],  # BO
                self.bo_dados[1],  # OP
                self.bo_dados[2],  # Cliente
                self.entries_ocorrencia["Tipo de Ocorrência"].get(),  # Tipo de Ocorrência
                self.entries_ocorrencia["Motivo"].get(),  # Motivo
                self.entries_ocorrencia["Descrição"].get(),  # Descrição
                self.entries_transporte["Frete"].get(),  # Frete
                self.entries_transporte["Previsão de Embarque"].get(),  # Previsão de Embarque
                self.entries_transporte["NF de Envio"].get(),  # NF de Envio
                self.entries_transporte["NF de Devolução"].get(),  # NF de Devolução
                self.ultimo_modulo  # Módulo que chamou a função
            ]

            # Insere a BO no banco de dados
            cursor.execute('''
                IF NOT EXISTS (SELECT 1 FROM bo_records WHERE bo_number = ?)
                BEGIN
                    INSERT INTO bo_records (
                        bo_number, op, loja, tipo_ocorrencia, motivo, descricao,
                        frete, previsao_embarque, nf_envio, nf_devolucao, modulo, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Em Andamento')
                END
            ''', valores)
            conn.commit()

            # Atualiza a sequência APÓS salvar
            cursor.execute("UPDATE bo_sequence SET last_number = last_number + 1")
            conn.commit()

            messagebox.showinfo("Sucesso", "BO salva com sucesso!")
            self.window.destroy()

            # Atualiza a lista de BOs no módulo corporativo, se necessário
            if self.ultimo_modulo == "corporativo":
                from modulos.corporativo import CorporativoModule
                corporativo_module = CorporativoModule(self.window.master)  # Passa a janela principal como parent
                corporativo_module.carregar_bos()

        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao salvar BO: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()