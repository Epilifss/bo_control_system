import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
from PIL import Image, ImageTk
from database import create_connection_mikonos
from database import create_connection

class buscarBo:
    def __init__(self, parent, caller_id=None):
        self.parent = parent
        self.root = tk.Toplevel(parent)
        self.root.title("Buscar BO")
        self.root.geometry("1000x600")
        self.root.state('zoomed')

        self.ultimo_modulo = caller_id

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

        # Frame para Treeview e barra de rolagem
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))

        # lista de BOs Treeview
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

        self.tree.configure(yscrollcommand=v_scrollbar.set)
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        self.tree.bind("<Double-1>", lambda event: exibir_detalhes(self.root, self.tree, caller_id=self.ultimo_modulo))

        self.center_window()
        self.carregar_bos()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def identificar_chamador(self):
        if self.ultimo_modulo is None:
            raise ValueError(
                "É necessário informar o identificador do módulo que chamou a função")
        return f"Buscar BO chamado pelo módulo: {self.ultimo_modulo}"

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

class exibir_detalhes():
    def __init__(self, parent, tree, caller_id=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Detalhes BO")
        self.window.geometry("600x350")
        self.window.grab_set()

        self.tree = tree
        self.ultimo_modulo = caller_id

        self.center_window()
        self.sc5_detalhes()
        self.itens_bo()
        
    def sc5_detalhes(self):
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        frame_sc5Detalhes = ttk.LabelFrame(self.window, text="Detalhes da Ocorrência", padding=10)
        frame_sc5Detalhes.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        frame_sc5Detalhes.grid_rowconfigure(0, weight=1)
        frame_sc5Detalhes.grid_columnconfigure(1, weight=1)  # Permite expansão da segunda coluna

        # Obtém o item selecionado no Treeview
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            return

        # Obtém os valores do item selecionado
        valores = self.tree.item(item_selecionado, "values")

        # Criando as labels
        labels = ["BO:", "OP:", "CLIENTE:", "FILIAL:", "EMISSÃO:", "PREVISÃO DE ENTREGA:"]
        for i, label_text in enumerate(labels):
            ttk.Label(frame_sc5Detalhes, text=label_text).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(frame_sc5Detalhes, text=valores[i]).grid(row=i, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(frame_sc5Detalhes, text="Acompanhar BO", command=self.acompanhar_bo).grid(row=i, column=2, sticky="e", padx=5, pady=2)

        self.window.update_idletasks()
        self.window.minsize(400, self.window.winfo_height())

    def itens_bo(self):
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        frame_itensBo = ttk.LabelFrame(self.window, text="Itens da BO", padding=10)
        frame_itensBo.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        frame_itensBo.grid_rowconfigure(0, weight=1)
        frame_itensBo.grid_columnconfigure(1, weight=1)

        labels= ["CÓDIGO", "DESCRIÇÃO", "LINHA"]

        item_selecionado2 = self.tree.selection()
        if not item_selecionado2:
            return

        valores2 = self.tree.item(item_selecionado2, "values")

        conn = create_connection_mikonos()
        if conn is None:
            messagebox.showerror("Erro", "Falha ao conectar ao banco de dados")
            return

        cursor = conn.cursor()

        query = """SELECT SC6.C6_CODTIDI, SC6.C6_DESCRI, SC6.C6_LINHA
        FROM SC6010 SC6
        INNER JOIN SC5010 SC5 ON (SC5.C5_NUM = SC6.C6_NUM AND SC5.C5_FILIAL = SC6.C6_FILIAL AND SC5.D_E_L_E_T_ <> '*')
        WHERE SC6.C6_NUM LIKE ? AND SC5.C5_PEDREPR LIKE ? AND SC5.C5_PEDREPR LIKE ?"""

        try:
            cursor.execute(query, (f"%{valores2[1]}%", "%BO%", f"%{valores2[0]}%"))
            itens_bo = cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao executar a consulta: {e}")
            return
        finally:
            cursor.close()
            conn.close()


        for i, label_txt in enumerate(labels):
            label = ttk.Label(frame_itensBo, text=label_txt)
            label.grid(row=0, column=i, sticky="w", padx=5, pady=2)
            frame_itensBo.grid_columnconfigure(i, weight=1)

        for idx, item in enumerate(itens_bo):
            for col, value in enumerate(item):
                label = ttk.Label(frame_itensBo, text=value)
                label.grid(row=idx + 1, column=col, sticky="w", padx=5, pady=2)
                frame_itensBo.grid_columnconfigure(col, weight=1)

        self.window.update_idletasks()
        self.window.minsize(400, self.window.winfo_height())

    def center_window(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def acompanhar_bo(self):
        obj = acompanhar_Bo(self.window, self.tree, caller_id=self.ultimo_modulo)
        resultado = obj.identificar_chamador()
        print(resultado)

if __name__ == "__main__":
    root = tk.Tk()
    app = exibir_detalhes(root)
    root.mainloop()


class acompanhar_Bo:
    def __init__(self, parent, tree, caller_id=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Acompanhar BO")
        self.window.grab_set()

        self.tree = tree
        self.ultimo_modulo = caller_id
        self.anexos = []

        self.bo_dados = self.obter_dados_bo()
        self.secao_dados_gerais()
        self.secao_ocorrencia()
        self.secao_transporte()
        self.secao_anexo()
        self.botao_salvar()

    def secao_dados_gerais(self):
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

    def secao_ocorrencia(self):
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

    def secao_transporte(self):
        """Cria a seção de detalhes do transporte."""
        frame_transporte = ttk.LabelFrame(self.window, text="Transporte", padding=10)
        frame_transporte.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        ttk.Label(frame_transporte, text="Previsão de entrega:").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(frame_transporte, text=self.bo_dados[5]).grid(row=0, column=1, sticky=tk.W)

        campos_transporte = [
            ("Frete", ttk.Combobox, ["CIF", "FOB"]),
            ("NF de Envio", ttk.Entry),
            ("NF de Devolução", ttk.Entry),
        ]

        self.entries_transporte = {}
        for i, (campo, widget, *args) in enumerate(campos_transporte, start=1):
            ttk.Label(frame_transporte, text=f"{campo}:").grid(row=i, column=0, sticky=tk.W)
            if widget == ttk.Combobox:
                entry = widget(frame_transporte, values=args[0])
            else:
                entry = widget(frame_transporte)
            entry.grid(row=i, column=1, sticky="ew")
            self.entries_transporte[campo] = entry

    def secao_anexo(self):
        """Cria a seção de anexos."""
        frame_anexo = ttk.LabelFrame(self.window, text="Documentos", padding=10)
        frame_anexo.grid(row=3, column=0, sticky="ew", padx=10, pady=10)

        self.anexo_container = ttk.Frame(frame_anexo)
        self.anexo_container.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        self.anexo_button = ttk.Button(frame_anexo, text="Anexar Arquivos", command=self.anexar_arquivo)
        self.anexo_button.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        self.remove_button = ttk.Button(frame_anexo, text="Remover Arquivo", command=self.remover_anexo)
        self.remove_button.grid(row=2, column=0, sticky="ew", padx=5, pady=5)


    def anexar_arquivo(self):
        """Abre a janela para anexar um arquivo."""
        file_patchs = filedialog.askopenfilenames(filetypes=[("Imagens", "*.jpg;*.jpeg;*.png"), ("Vídeos", ".mp4;*.avi"), ("Todos os Arquivos", "*.*")])

        if file_patchs:
            self.anexos.extend(file_patchs)
            self.atualizar_anexo_exibicao()

    def remover_anexo(self):
        """Remove o anexo atual."""
        if self.anexos:
            self.anexos.pop()
            self.atualizar_anexo_exibicao()

    def atualizar_anexo_exibicao(self):
        """Atualiza a exibição dos anexos."""
        for widget in self.anexo_container.winfo_children():
            widget.destroy()
        
        if self.anexos:
            for i, file_path in enumerate(self.anexos):
                if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                    self.exibir_imagem(file_path, i)
                elif file_path.lower().endswith(('.mp4', '.avi')):
                    self.exibir_video(file_path, i)
        else:
            label = ttk.Label(self.anexo_container, text="Nenhum Arquivo Anexado.")

            label.pack()

    def exibir_imagem(self, file_patch, index):
        img = Image.open(file_patch)
        img.thumbnail((100, 100))
        img = ImageTk.PhotoImage(img)
        
        label = ttk.Label(self.anexo_container, image=img)
        label.image = img
        label.grid(row=0, column=index, padx=5, pady=5)

        self.lixeira_icone = ImageTk.PhotoImage(Image.open("./images/lixeira.png").resize((20, 20)))
        remove_button = ttk.Button(self.anexo_container, image=self.lixeira_icone, command=lambda img=file_patch: self.remover_imagem(img))
        remove_button.image = self.lixeira_icone
        remove_button.grid(row=0, column=index+1, padx=5, pady=5)

    def remover_imagem(self, file_path):
        """Remove a imagem e o botão de remoção do container."""
        for widget in self.anexo_container.winfo_children():
            if file_path in widget.winfo_children():
                widget.destroy()
        self.atualizar_anexo_exibicao()

    def exibir_video(self, file_path):
        """Exibe um vídeo (requer biblioteca adicional como tkvideo)."""
        # Aqui você pode adicionar lógica para exibir o vídeo
        messagebox.showinfo("Aviso", "Exibição de vídeo não implementada. Arquivo selecionado: " + file_path)

    def botao_salvar(self):
        """Cria o botão de salvar."""
        frame_botoes = ttk.Frame(self.window, padding=10)
        frame_botoes.grid(row=4, column=0, sticky="ew", padx=10, pady=10)

        ttk.Button(frame_botoes, text="Salvar", command=self.salvar).pack(side=tk.RIGHT)

    def motivos(self):
        """Retorna uma lista de motivos."""
        return [
            "Pé quebrado",
            "Tela rasgada",
            "Outro"
        ]
    
    def identificar_chamador(self):
        if self.ultimo_modulo is None:
            raise ValueError(
                "É necessário informar o identificador do módulo que chamou a função")
        return f"Acompanhar BO chamado pelo módulo: {self.ultimo_modulo}"

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
        # Salva a BO no banco e só então incrementa o número da sequência.
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
                self.bo_dados[5],  # Previsão de Embarque
                self.entries_transporte["NF de Envio"].get(),  # NF de Envio
                self.entries_transporte["NF de Devolução"].get(),  # NF de Devolução
                self.ultimo_modulo,  # Módulo que chamou a função
                'Em Andamento' # Status
            ]

            # Primeiro, verifica se o BO já existe
            cursor.execute("SELECT 1 FROM bo_records WHERE bo_number = ?", (self.bo_dados[0],))
            exists = cursor.fetchone()

            if not exists:
                # Agora, insere os dados
                cursor.execute('''
                    INSERT INTO bo_records (
                        bo_number, op, loja, tipo_ocorrencia, motivo, descricao,
                        frete, previsao_embarque, nf_envio, nf_devolucao, modulo, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                if CorporativoModule.instance is not None:
                    CorporativoModule.instance.carregar_bos()
                else:
                    pass
            elif self.ultimo_modulo == "varejo":
                print('tentou atualizar varejo')
                from modulos.varejo import VarejoModule
                if VarejoModule.instance is not None:
                    VarejoModule.instance.carregar_bos()

        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao salvar BO: {e}")
            print(e)
        finally:
            if conn:
                cursor.close()
                conn.close()