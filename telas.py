import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
from moviepy import VideoFileClip
from PIL import Image, ImageTk
import pyodbc
import hashlib
import os
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from modulos.corporativo import CorporativoModule
from modulos.varejo import VarejoModule
# from modulos.exportacao import ExportacaoModule
from components import *
from database import create_connection_mikonos
from database import create_connection


class LoginWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Login")
        self.root.geometry("300x200")

        self.frame = ttk.Frame(self.root, padding=10)
        self.frame.pack(expand=True)

        ttk.Label(self.frame, text="Usuário:").grid(
            row=0, column=0, sticky=tk.W)
        self.username = ttk.Entry(self.frame)
        self.username.grid(row=0, column=1)

        ttk.Label(self.frame, text="Senha:").grid(row=1, column=0, sticky=tk.W)
        self.password = ttk.Entry(self.frame, show="*")
        self.password.grid(row=1, column=1)

        self.login_btn = ttk.Button(
            self.frame, text="Login", command=self.login)
        self.login_btn.grid(row=2, column=0, columnspan=2, pady=5)

        self.center_window()
        self.username.focus_set()
        self.root.bind('<Return>', self.login)
        self.root.mainloop()

    def center_window(self):
        # Força o Tkinter a calcular os tamanhos
        self.root.update_idletasks()

        # Obtém as dimensões da janela
        width = self.root.winfo_width()
        height = self.root.winfo_height()

        # Obtém as dimensões da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calcula as coordenadas para centralizar
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Aplica a geometria centralizada
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def admin_login(self):
        self.username.delete(0, tk.END)
        self.password.delete(0, tk.END)

    def login(self, event=None):
        username = self.username.get()
        password = self.password.get()
        hashed_pw = hashlib.sha256(password.encode()).hexdigest()
        

        conn = create_connection()
        if conn is None:
            messagebox.showerror("Erro", "Falha ao conectar ao banco de dados")
            return

        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM users WHERE username COLLATE Latin1_General_BIN=? AND password_hash COLLATE Latin1_General_BIN=?", (username, hashed_pw))
            user = cursor.fetchone()
            

            if user:
                printUser = (f"Usuário {user[1]} logou no ")

                if user[4]:  # is_admin
                    self.root.destroy()
                    print(printUser + "painel de administração")
                    AdminPanel()
                else:
                    module = user[3]
                    if (module) == '0':
                        self.root.destroy()
                        print(printUser + "módulo Corporativo")
                        CorporativoModule(user)
                    elif (module) == '1':
                        self.root.destroy()
                        print(printUser + "módulo Varejo")
                        VarejoModule(user)
                    # elif (module) == '2':
                    #     self.root.destroy()
                    #     print(printUser + "no módulo Exportação")
                    #     ExportacaoModule(user)
                    else:
                        messagebox.showerror(
                            "Erro", "Módulo do usuário não encontrado. Contate o administrador.")
            else:
                messagebox.showerror("Erro", "Credenciais inválidas")
        except pyodbc.Error as e:
            messagebox.showerror(
                "Erro", f"Erro ao consultar o banco de dados: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()


class AdminPanel:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Painel de Administração")

        self.root.state('zoomed')

        # Abas
        self.notebook = ttk.Notebook(self.root)

        # Aba de Usuários
        self.user_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.user_frame, text="Usuários")

        # Lista de usuários
        self.tree = ttk.Treeview(self.user_frame, columns=(
            "ID", "Usuário", "Módulo", "Admin"), show="headings")
        self.tree.heading("ID", text="ID")
        self.tree.heading("Usuário", text="Usuário")
        self.tree.heading("Módulo", text="Módulo")
        self.tree.heading("Admin", text="Admin")
        self.tree.column("ID", width=0, stretch=False)
        self.tree.pack(expand=True, fill=tk.BOTH)

        self.tree.bind("<<TreeviewSelect>>", self.atualizar_selecao)

        header = ttk.Frame(self.root)
        header.pack(fill=tk.X)

        ttk.Button(header, text="Novo Usuário",
                   command=self.novo_usuario).pack(side=tk.LEFT)
        ttk.Button(header, text="Excluir Usuário",
                   command=self.excluir_usuario).pack(side=tk.LEFT)
        ttk.Button(header, text="Editar Usuário",
                   command=self.editar_usuario).pack(side=tk.LEFT)
        ttk.Button(header, text="Logoff",
                   command=self.logoff).pack(side=tk.RIGHT)

        # Aba de Usuários
        self.user_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.user_frame, text="Banco")

        self.notebook.pack(expand=True, fill=tk.BOTH)
        self.carregar_usuarios()
        self.root.mainloop()

    def carregar_usuarios(self):
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, username, IIF(module='0','Corporativo',IIF(module='1','Varejo',IIF(module='2', 'Exportação', ''))), IIF(is_admin='True', 'Sim', 'Não') FROM users")
            rows = cursor.fetchall()

            for item in self.tree.get_children():
                self.tree.delete(item)

            for row in rows:
                valores_formatados = [str(item) for item in row]
                self.tree.insert("", tk.END, values=valores_formatados)
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar usuários: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()

    def atualizar_selecao(self, event=None):
        selected_items = self.tree.selection()
        if selected_items:
            item_id = selected_items[0]
            valores = self.tree.item(item_id, "values")
            self.usuario_selecionado = valores[0]
            print("Usuário selecionado:", self.usuario_selecionado)
        else:
            self.usuario_selecionado = None

    def novo_usuario(self):
        NovoUsuarioWindow(self.root, self)

    def excluir_usuario(self):
        """Exclui o usuário selecionado"""
        if not hasattr(self, 'usuario_selecionado') or not self.usuario_selecionado:
            messagebox.showerror("Erro", "Nenhum usuário selecionado")
            return

        resposta = messagebox.askyesno(
            "Confirmar", f"Tem certeza que deseja excluir {self.usuario_selecionado}?")
        if not resposta:
            return

        conn = create_connection()
        if conn is None:
            messagebox.showerror("Erro", "Falha ao conectar ao banco de dados")
            return

        try:
            cursor = conn.cursor()

            # Debug: Verifique se o usuário realmente existe antes da exclusão
            cursor.execute("SELECT * FROM users WHERE id=?",
                           (self.usuario_selecionado,))
            usuario_existente = cursor.fetchone()

            if not usuario_existente:
                messagebox.showerror(
                    "Erro", f"O usuário '{self.usuario_selecionado}' não foi encontrado no banco de dados.")
                return

            usuario_formatado = self.usuario_selecionado.strip()

            print(f"Excluindo usuário: {usuario_formatado}")

            cursor.execute("DELETE FROM users WHERE id=?",
                           (self.usuario_selecionado,))
            conn.commit()

            # Debug: Verifique se algum registro foi realmente excluído
            if cursor.rowcount == 0:
                messagebox.showerror(
                    "Erro", "Nenhum registro foi excluído. Verifique se o usuário existe.")
            else:
                messagebox.showinfo("Sucesso", "Usuário excluído com sucesso!")

            self.carregar_usuarios()  # Atualiza a lista
            self.usuario_selecionado = None  # Resetar seleção

        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao excluir usuário: {e}")

        finally:
            if conn:
                cursor.close()
                conn.close()

    def editar_usuario(self):
        """Abre a janela para editar o usuário selecionado."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning(
                "Atenção", "Selecione um usuário para editar.")
            return

        user_data = self.tree.item(selected_items[0], "values")
        EditarUsuarioWindow(self.root, self, user_data)

    def logoff(self):
        self.root.destroy()
        LoginWindow()


class NovoUsuarioWindow:
    def __init__(self, parent, admin_panel):
        self.window = tk.Toplevel(parent)
        self.window.title("Novo Usuário")

        self.admin_panel = admin_panel

        campos = ["Usuário", "Senha", "Módulo", "Admin"]
        self.entries = {}

        for i, campo in enumerate(campos):
            ttk.Label(self.window, text=f"{campo}:").grid(
                row=i, column=0, sticky=tk.W)
            if campo == "Admin":
                entry = ttk.Combobox(self.window, values=["Sim", "Não"])
            elif campo == "Módulo":
                entry = ttk.Combobox(self.window, values=[
                                     "Corporativo", "Varejo", "Exportação"])
            else:
                entry = ttk.Entry(self.window)

            entry.grid(row=i, column=1)
            self.entries[campo] = entry

        ttk.Button(self.window, text="Salvar", command=self.salvar).grid(
            row=len(campos), columnspan=2)

    def salvar(self):
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            username = self.entries["Usuário"].get()
            password = self.entries["Senha"].get()
            hashed_pw = hashlib.sha256(password.encode()).hexdigest()
            module = self.entries["Módulo"].get()
            is_admin = 1 if self.entries["Admin"].get() == "Sim" else 0
            module = 0 if self.entries["Módulo"].get(
            ) == "Corporativo" else 1 if self.entries["Módulo"].get() == "Varejo" else 2

            cursor.execute('''
                INSERT INTO users (username, password_hash, module, is_admin)
                VALUES (?, ?, ?, ?)
            ''', (username, hashed_pw, module, is_admin))

            conn.commit()
            messagebox.showinfo("Sucesso", "Usuário cadastrado com sucesso!")
            self.window.destroy()

            self.admin_panel.carregar_usuarios()
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao salvar usuário: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()


class EditarUsuarioWindow:
    def __init__(self, parent, admin_panel, user_data):
        self.window = tk.Toplevel(parent)
        self.window.title("Editar Usuário")
        self.admin_panel = admin_panel
        self.user_data = user_data  # user_data deve ser uma tupla

        campos = ["Usuário", "Senha", "Módulo", "Admin"]
        self.entries = {}

        for i, campo in enumerate(campos):
            ttk.Label(self.window, text=f"{campo}:").grid(
                row=i, column=0, sticky=tk.W)

            if campo == "Admin":
                entry = ttk.Combobox(self.window, values=["Sim", "Não"])
                # is_admin está no índice 3
                entry.set("Sim" if self.user_data[3] == 0 else "Não")
            elif campo == "Senha":
                entry = ttk.Entry(self.window, show="*")
            else:
                entry = ttk.Entry(self.window)
                if campo == "Usuário":
                    # username está no índice 1
                    entry.insert(0, self.user_data[1])
                elif campo == "Módulo":
                    entry = ttk.Combobox(self.window, values=[
                                         "Corporativo", "Varejo", "Exportação"])
            entry.grid(row=i, column=1)
            self.entries[campo] = entry

        ttk.Button(self.window, text="Salvar", command=self.salvar).grid(
            row=len(campos), columnspan=2)

    def salvar(self):
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            username = self.entries["Usuário"].get()
            new_password = self.entries["Senha"].get()
            module = self.entries["Módulo"].get()
            is_admin = 1 if self.entries["Admin"].get() == "Sim" else 0
            module = 0 if self.entries["Módulo"].get(
            ) == "Corporativo" else 1 if self.entries["Módulo"].get() == "Varejo" else 2

            if new_password:
                hashed_pw = hashlib.sha256(new_password.encode()).hexdigest()
                cursor.execute('''
                    UPDATE users
                    SET username = ?, password_hash = ?, module = ?, is_admin = ?
                    WHERE id = ?
                ''', (username, hashed_pw, module, is_admin, self.user_data[0]))  # ID está no índice 0
            else:
                cursor.execute('''
                    UPDATE users
                    SET username = ?, module = ?, is_admin = ?
                    WHERE id = ?
                ''', (username, module, is_admin, self.user_data[0]))  # ID está no índice 0

            conn.commit()
            messagebox.showinfo("Sucesso", "Usuário editado com sucesso!")
            self.window.destroy()

            self.admin_panel.carregar_usuarios()
        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao editar usuário: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()


if __name__ == "__main__":
    LoginWindow()


class Embarcados:
    def __init__(self, user, caller_id=None):
        self.user = user
        self.root = tk.Tk()
        self.root.title("Embarcados")
        self.root.geometry("1000x600")

        self.ultimo_modulo = caller_id
        print(self.identificar_chamador())

        # Cabeçalho
        header = ttk.Frame(self.root)
        header.pack(fill=tk.X)

        ttk.Button(header, text="Fechar",
                   command=lambda: self.root.destroy()).pack(side=tk.RIGHT)

        # Barra de pesquisa
        self.search_bar = SearchBar(
            self.root, self.pesquisar_bo, self.clear_search)

        # Lista de BOs
        self.tree = ttk.Treeview(self.root, columns=(
            "BO", "OP", "Status", "tipo_ocorrencia", "dt_embarque"), show="headings")
        self.tree.heading("BO", text="BO")
        self.tree.heading("OP", text="OP")
        self.tree.heading("Status", text="Status")
        self.tree.heading("tipo_ocorrencia", text="Tipo de Ocorrência")
        self.tree.heading("dt_embarque", text="Data de Embarque")

        self.tree.pack(expand=True, fill=tk.BOTH)

        self.center_window()
        self.carregar_bos()
        self.root.mainloop()

    def center_window(self):
        # Força o Tkinter a calcular os tamanhos
        self.root.update_idletasks()

        # Obtém as dimensões da janela
        width = self.root.winfo_width()
        height = self.root.winfo_height()

        # Obtém as dimensões da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calcula as coordenadas para centralizar
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Aplica a geometria centralizada
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def identificar_chamador(self):
        if self.ultimo_modulo is None:
            raise ValueError(
                "É necessário informar o identificador do módulo que chamou a função")
        return f"Tela de embarcados chamada pelo módulo: {self.ultimo_modulo}"

    def carregar_bos(self):
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            query = """
                SELECT bo_number, op, status, tipo_ocorrencia, dt_embarque FROM bo_records WHERE status LIKE 'Embarcado' AND modulo LIKE ?
            """
            cursor.execute(query, self.ultimo_modulo)
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
                "SELECT bo_number, op, status, tipo_ocorrencia, motivo FROM bo_records WHERE bo_number LIKE ?", (f"%{termo}%",))
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
                dt_embarque = str(row[4]).strip(
                    "(' ,)") if row[4] is not None else ""

                self.tree.insert("", tk.END, values=(
                    bo, op, status, tipo_ocorrencia, dt_embarque))
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


class Estatisticas:
    def __init__(self, user, caller_id=None):
        self.user = user
        self.root = tk.Tk()
        self.root.title("Estatísticas")
        self.root.geometry("1000x600")
        self.root.state('zoomed')

        self.ultimo_modulo = caller_id
        print(self.identificar_chamador())
        
        def obter_anos():
            conn = create_connection_mikonos()
            cursor = conn.cursor()

            cursor.execute(f"SELECT DISTINCT YEAR(C5_EMISSAO) AS ano FROM SC5010 ORDER BY ano DESC")
            anos = [row[0] for row in cursor.fetchall()]

            cursor.close()
            conn.close()

            return anos
        
        def obter_setores():
            conn = create_connection()
            cursor = conn.cursor()

            cursor.execute(f"SELECT setor_responsavel FROM [bo_system].[dbo].[bo_records]")
            setores = [row[0] for row in cursor.fetchall()]

            cursor.close()
            conn.close()

            return setores
        
        def obter_contagens_por_setor():
            conn = create_connection()
            cursor = conn.cursor()

            cursor.execute(f"SELECT COALESCE(setor_responsavel, 'Não especificado'), COUNT(*) FROM [bo_system].[dbo].[bo_records] GROUP BY setor_responsavel")
            contagens = {row[0]: row[1] for row in cursor.fetchall()}
            cursor.close()
            conn.close()

            return contagens

        # Cabeçalho
        header = ttk.Frame(self.root)
        header.pack(fill=tk.X, padx=10, pady=10)

        anos = obter_anos()
        setores = obter_setores()
        contagem = obter_contagens_por_setor()        
        ttk.Label(header, text="Ano: ").pack(side=tk.LEFT)

        listaAnos = ttk.Combobox(header, values=anos, state="readonly")
        listaAnos.pack(side=tk.LEFT)
        listaAnos.set(anos[0])
        
        ttk.Button(header, text="Fechar",
                   command=lambda: self.root.destroy()).pack(side=tk.RIGHT)
        
        # Gráfico

        frame = ttk.Frame(self.root, padding="3 3 12 12")
        frame.pack(fill=tk.BOTH, expand=True)

        fig = Figure(figsize=(8, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        contagens_completo = [contagem.get(sector, 0) for sector in setores]
        ax.bar(setores, contagens_completo, color='green')
        ax.set_title("Contagem de BO's por setor")
        ax.set_xlabel("Setor")
        ax.set_ylabel("Contagem")
        
        plt.xticks(rotation=45)
        
        
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        root.mainloop()

        
    def identificar_chamador(self):
        if self.ultimo_modulo is None:
            raise ValueError(
                "É necessário informar o identificador do módulo que chamou a função")
        return f"Tela de estatísticas chamada pelo módulo: {self.ultimo_modulo}"


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
        self.search_bar = SearchBar(
            self.root, self.pesquisar_bo, self.clear_search)

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
        self.tree.heading("C5_FILIAL", text="EMPRESA")
        self.tree.heading("C5_EMISSAO", text="EMISSÃO")
        self.tree.heading("C5_ENTREGA", text="PREVISÃO DE ENTREGA")

        # Barra de rolagem vertical
        v_scrollbar = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.configure(yscrollcommand=v_scrollbar.set)
        self.tree.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)

        self.tree.bind("<Double-1>", lambda event: exibir_detalhes(self.root,
                       self.tree, caller_id=self.ultimo_modulo))

        self.center_window()
        self.carregar_bos()
        self.search_bar.search_entry.bind('<Return>', self.pesquisar_bo)

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
        return f"Tela de Buscar BO chamada pelo módulo: {self.ultimo_modulo}"

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
                        AND SC5.C5_NOTA = ''
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

    def pesquisar_bo(self, event=None):
        termo = self.search_bar.search_entry.get()
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
                        AND SC5.C5_NOTA = ''
                        AND SC5.C5_PEDREPR LIKE 'BO%'
                        AND C5_FILIAL IN ('0101','0201')
                        AND (UPPER(C5_NUM) LIKE UPPER(?) OR UPPER(C5_NOME) LIKE UPPER(?)) OR UPPER(C5_PEDREPR) LIKE UPPER(?)
                ORDER BY C5_EMISSAO DESC
                """

            cursor.execute(
                query, (f"%{termo}%", f"%{termo}%", "BO" + f"%{termo}%"))
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

    def clear_search(self):
        self.search_bar.search_entry.delete(0, tk.END)
        self.search_bar.update_buttons()  # Atualiza os botões após limpar
        self.carregar_bos()  # Recarrega os BOs

    def update_buttons(self, event=None):
        termo = self.search_entry.get()
        if termo:
            self.search_button.pack_forget()  # Oculta o botão "Pesquisar"
            self.clear_button.pack(side=tk.LEFT)  # Exibe o botão "Limpar"
        else:
            self.clear_button.pack_forget()  # Oculta o botão "Limpar"
            self.search_button.pack(side=tk.LEFT)


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

        frame_sc5Detalhes = ttk.LabelFrame(
            self.window, text="Detalhes da Ocorrência", padding=10)
        frame_sc5Detalhes.grid(
            row=0, column=0, sticky="nsew", padx=10, pady=10)

        frame_sc5Detalhes.grid_rowconfigure(0, weight=1)
        # Permite expansão da segunda coluna
        frame_sc5Detalhes.grid_columnconfigure(1, weight=1)

        # Obtém o item selecionado no Treeview
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            return

        # Obtém os valores do item selecionado
        valores = self.tree.item(item_selecionado, "values")

        # Criando as labels
        labels = ["BO:", "OP:", "CLIENTE:", "FILIAL:",
                  "EMISSÃO:", "PREVISÃO DE ENTREGA:"]
        for i, label_text in enumerate(labels):
            ttk.Label(frame_sc5Detalhes, text=label_text).grid(
                row=i, column=0, sticky="w", padx=5, pady=2)
            ttk.Label(frame_sc5Detalhes, text=valores[i]).grid(
                row=i, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(frame_sc5Detalhes, text="Acompanhar BO", command=self.acompanhar_bo).grid(
            row=i, column=2, sticky="e", padx=5, pady=2)

        self.window.update_idletasks()
        self.window.minsize(400, self.window.winfo_height())

    def itens_bo(self):
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        frame_itensBo = ttk.LabelFrame(
            self.window, text="Itens da BO", padding=10)
        frame_itensBo.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        frame_itensBo.grid_rowconfigure(0, weight=1)
        frame_itensBo.grid_columnconfigure(1, weight=1)

        labels = ["CÓDIGO", "DESCRIÇÃO", "LINHA"]

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
            cursor.execute(
                query, (f"%{valores2[1]}%", "%BO%", f"%{valores2[0]}%"))
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
        obj = acompanhar_Bo(self.window, self.tree,
                            caller_id=self.ultimo_modulo)
        resultado = obj.identificar_chamador()
        print(resultado)

class acompanhar_Bo:
    def __init__(self, parent, tree, caller_id=None):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Acompanhar BO")
        self.window.grab_set()

        self.tree = tree
        self.ultimo_modulo = caller_id
        self.anexos = []
        self.loading_label = None
        self.max_anexos = 5
        self.max_tamanho_anexo = 50 * 1024 * 1024

        self.bo_dados = self.obter_dados_bo()
        self.secao_dados_gerais()
        self.secao_ocorrencia()
        self.secao_transporte()
        self.secao_anexo()
        self.botao_salvar()

    def secao_dados_gerais(self):
        """Cria a seção de dados gerais da BO."""
        frame_dados_gerais = ttk.LabelFrame(
            self.window, text="Dados Gerais", padding=10)
        frame_dados_gerais.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        # BO e OP
        ttk.Label(frame_dados_gerais, text="BO:").grid(
            row=0, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[0]).grid(
            row=0, column=1, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text="OP:").grid(
            row=1, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[1]).grid(
            row=1, column=1, sticky=tk.W)

        # Cliente
        ttk.Label(frame_dados_gerais, text="CLIENTE:").grid(
            row=2, column=0, sticky=tk.W)
        ttk.Label(frame_dados_gerais, text=self.bo_dados[2]).grid(
            row=2, column=1, sticky=tk.W)
        
    def obter_setores(self):
        conn = create_connection()
        cursor = conn.cursor()
    
        cursor.execute("SELECT setor_responsavel FROM bo_records")
        setores = [row[0] for row in cursor.fetchall()]

        cursor.close()
        conn.close()

        return setores

    def secao_ocorrencia(self):
        # Cria a seção de detalhes da ocorrência.
        frame_ocorrencia = ttk.LabelFrame(
            self.window, text="Detalhes da Ocorrência", padding=10)
        frame_ocorrencia.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        

        campos_ocorrencia = [
            ("Tipo de Ocorrência", ttk.Entry),
            ("Motivo", ttk.Combobox, self.motivos()),
            ("Descrição", ttk.Entry),
            ("Setor Responsável", ttk.Combobox, self.obter_setores())
        ]

        self.entries_ocorrencia = {}
        for i, (campo, widget, *args) in enumerate(campos_ocorrencia):
            ttk.Label(frame_ocorrencia, text=f"{campo}:").grid(
                row=i, column=0, sticky=tk.W)
            if widget == ttk.Combobox:
                entry = widget(frame_ocorrencia, values=args[0])
            else:
                entry = widget(frame_ocorrencia)
            entry.grid(row=i, column=1, sticky="ew")
            self.entries_ocorrencia[campo] = entry

    def secao_transporte(self):
        """Cria a seção de detalhes do transporte."""
        frame_transporte = ttk.LabelFrame(
            self.window, text="Transporte", padding=10)
        frame_transporte.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        ttk.Label(frame_transporte, text="Previsão de entrega:").grid(
            row=0, column=0, sticky=tk.W)
        ttk.Label(frame_transporte, text=self.bo_dados[5]).grid(
            row=0, column=1, sticky=tk.W)

        campos_transporte = [
            ("Frete", ttk.Combobox, ["CIF", "FOB"])
        ]

        self.entries_transporte = {}
        for i, (campo, widget, *args) in enumerate(campos_transporte, start=1):
            ttk.Label(frame_transporte, text=f"{campo}:").grid(
                row=i, column=0, sticky=tk.W)
            if widget == ttk.Combobox:
                entry = widget(frame_transporte, values=args[0])
            else:
                entry = widget(frame_transporte)
            entry.grid(row=i, column=1, sticky="ew")
            self.entries_transporte[campo] = entry

    def secao_anexo(self):
        """Cria a seção de anexos."""
        frame_anexo = ttk.LabelFrame(
            self.window, text="Documentos", padding=10)
        frame_anexo.grid(row=3, column=0, sticky="ew", padx=10, pady=10)

        self.anexo_container = ttk.Frame(frame_anexo)
        self.anexo_container.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        self.anexo_button = ttk.Button(
            frame_anexo, text="Anexar Arquivos", command=self.anexar_arquivo)
        self.anexo_button.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

    def anexar_arquivo(self):
        file_paths = filedialog.askopenfilenames(filetypes=[
            ("Compatível", "*.jpg;*.jpeg;*.png;*.mp4;*.avi"), ("Todos os Arquivos", "*.*")])

        total_size = sum(os.path.getsize(path) for path in self.anexos)
        for file_path in file_paths:
            total_size += os.path.getsize(file_path)
            if total_size > self.max_tamanho_anexo:
                messagebox.showerror(
                    "Erro", "Você excedeu o limite de 50MB total de anexo.")
                return

        if len(self.anexos) + len(file_paths) > self.max_anexos:
            messagebox.showerror(
                "Erro", f"Você não pode anexar mais do que {self.max_anexos} arquivos.")
            return

        self.mostrar_texto_carregamento()
        threading.Thread(target=self.processar_anexos,
                         args=(file_paths,)).start()

    def mostrar_texto_carregamento(self):
        # Remover o frame de carregamento, se já existir
        for widget in self.anexo_container.winfo_children():
            widget.destroy()

        # Criar um widget para o texto de carregamento
        self.loading_widget = ttk.Label(
            self.anexo_container, text="Carregando anexo...", font=("Arial", 11))
        self.loading_widget.grid(row=0, column=0, padx=5, pady=5)
        print("Texto de carregamento sendo exibido")

    def processar_anexos(self, file_paths):
        for file_path in file_paths:
            time.sleep(2)  # Simula o tempo de processamento
            self.anexos.append(file_path)

        self.parent.after(100, self.finalizar_carregamento)

    def finalizar_carregamento(self):
        if self.loading_widget and self.loading_widget.winfo_exists():
            self.loading_widget.destroy()
        self.atualizar_anexo_exibicao()
        print("Texto de carregamento sendo removido")

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
            label = ttk.Label(self.anexo_container,
                              text="Nenhum Arquivo Anexado.")

            label.pack()

    def exibir_imagem(self, file_patch, index):
        img = Image.open(file_patch)
        img.thumbnail((100, 100))
        img = ImageTk.PhotoImage(img)

        label = ttk.Label(self.anexo_container, image=img)
        label.image = img
        label.grid(row=index * 2, column=0, padx=5, pady=5)

        self.lixeira_icone = ImageTk.PhotoImage(
            Image.open("./images/lixeira.png").resize((20, 20)))
        remove_button = ttk.Button(self.anexo_container, image=self.lixeira_icone,
                                   command=lambda img=file_patch: self.remover_imagem(img))
        remove_button.image = self.lixeira_icone
        remove_button.grid(row=index * 2, column=1, padx=5, pady=5)

    def remover_imagem(self, file_path):
        """Remove a imagem e o botão de remoção do container."""
        if file_path in self.anexos:
            self.anexos.remove(file_path)
            self.atualizar_anexo_exibicao()
            print(f"Arquivo removido: {file_path}")
        else:
            print("Arquivo não encontrado na lista de anexos.")

    def exibir_video(self, file_path, index):
        try:
            clip = VideoFileClip(file_path)
            frame = clip.get_frame(0)  # Pega o primeiro frame do vídeo
            clip.close()

            # Converta o frame para um formato exibível
            img = Image.fromarray(frame)
            img.thumbnail((100, 100))
            img = ImageTk.PhotoImage(img)

            # Crie um rótulo para exibir a miniatura
            label = ttk.Label(self.anexo_container, image=img)
            label.image = img
            label.grid(row=index * 2, column=0, padx=5, pady=5)

            # Adicione o botão de remoção
            self.lixeira_icone = ImageTk.PhotoImage(
                Image.open("./images/lixeira.png").resize((20, 20)))
            remove_button = ttk.Button(self.anexo_container, image=self.lixeira_icone,
                                       command=lambda img=file_path: self.remover_imagem(img))
            remove_button.image = self.lixeira_icone
            remove_button.grid(row=index * 2, column=1, padx=5, pady=5)

        except Exception as e:
            print(f"Erro ao exibir vídeo: {e}")
            label = ttk.Label(self.anexo_container,
                              text="Vídeo não pode ser exibido.")
            label.grid(row=index * 2, column=0, padx=5, pady=5)

    def fechar_janela_video(self, window, player):
        player.stop()  # Pare o vídeo
        window.destroy()  # Destrua a janela

    def botao_salvar(self):
        """Cria o botão de salvar."""
        frame_botoes = ttk.Frame(self.window, padding=10)
        frame_botoes.grid(row=4, column=0, sticky="ew", padx=10, pady=10)

        ttk.Button(frame_botoes, text="Salvar",
                   command=self.salvar).pack(side=tk.RIGHT)

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
        return f"Tela de acompanhar BO chamada pelo módulo: {self.ultimo_modulo}"

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
                # Tipo de Ocorrência
                self.entries_ocorrencia["Tipo de Ocorrência"].get(),
                self.entries_ocorrencia["Motivo"].get(),  # Motivo
                self.entries_ocorrencia["Descrição"].get(),  # Descrição
                self.entries_ocorrencia["Setor Responsável"].get(),  # Descrição
                self.entries_transporte["Frete"].get(),  # Frete
                self.bo_dados[5],  # Previsão de Embarque
                self.ultimo_modulo,  # Módulo que chamou a função
                'Em Andamento'  # Status
            ]

            # Primeiro, verifica se o BO já existe
            cursor.execute(
                "SELECT 1 FROM bo_records WHERE bo_number = ?", (self.bo_dados[0],))
            exists = cursor.fetchone()

            if not exists:
                # Agora, insere os dados
                cursor.execute('''
                    INSERT INTO bo_records (
                        bo_number, op, loja, tipo_ocorrencia, motivo, descricao, setor_responsavel,
                        frete, previsao_embarque, modulo, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', valores)
                conn.commit()

            # Atualiza a sequência APÓS salvar
            cursor.execute(
                "UPDATE bo_sequence SET last_number = last_number + 1")
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

if __name__ == "__main__":
    root = tk.Tk()
    app = exibir_detalhes(root)
    root.mainloop()