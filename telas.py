import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pyodbc
import datetime
import hashlib
from modulos.corporativo import CorporativoModule
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

    def login(self):
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
                "SELECT * FROM users WHERE username=? AND password_hash=?", (username, hashed_pw))
            user = cursor.fetchone()
            print(user)

            if user:
                if user[4]:  # is_admin
                    self.root.destroy()
                    AdminPanel()
                else:
                    module = user[3]
                    if (module) == '0':
                        self.root.destroy()
                        CorporativoModule(user)
                    else:
                        messagebox.showerror("Erro", "Módulo do usuário não encontrado. Contate o administrador.")
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
        self.tree = ttk.Treeview(self.user_frame, columns=("ID", "Usuário", "Módulo", "Admin"), show="headings")
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
            cursor.execute("SELECT id, username, module, is_admin FROM users")
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

class NovaBOWindow:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Nova BO")
        self.window.grab_set()

        self.ultimo_modulo = None
        # Obtém o próximo número disponível sem alterar no banco
        self.bo_number = self.gerar_numero_bo()

        ttk.Label(self.window, text=f"BO: {self.bo_number}").grid(
            row=0, column=0, sticky=tk.W)

        campos = [
            "OP", "Loja", "NF de envio", "NF de devolução", "OP de venda",
            "Tipo de Ocorrência", "Motivo", "Frete", "Setor responsável", "Status",
            "Previsão de embarque", "Descrição", "Filial"
        ]

        self.entries = {}
        for i, campo in enumerate(campos, start=1):
            ttk.Label(self.window, text=f"{campo}:").grid(
                row=i, column=0, sticky=tk.W)
            if campo == "Frete":
                entry = ttk.Combobox(self.window, values=["CIF", "FOB"])
            elif campo == "Status":
                entry = ttk.Combobox(self.window, values=[
                                     "Pronto", "Em andamento", "Embarcado", "Cancelado", "Devolvido"])
            elif campo == "Motivo":
                entry = ttk.Combobox(self.window, values=self.motivos())
            elif campo == "Previsão de embarque":
                entry = DateEntry(self.window)
            else:
                entry = ttk.Entry(self.window)
            entry.grid(row=i, column=1)
            self.entries[campo] = entry

        ttk.Button(self.window, text="Salvar", command=self.salvar).grid(
            row=len(campos)+1, columnspan=2)

    def motivos(self):
        return [
            "Pé quebrado",
            "Tela rasgada"
        ]
    
    def identificar_chamador(self, caller_id=None):

        if caller_id is None:
            raise ValueError("É necessário informar o identificador do módulo que chamou a função")
        self.ultimo_modulo = caller_id
        return f"Nova BO chamado pelo módulo: {caller_id}"


    def gerar_numero_bo(self):
        """Obtém o próximo número de BO disponível, mas não o atualiza no banco."""
        conn = create_connection()
        if conn is None:
            return "BO000/00"

        try:
            cursor = conn.cursor()

            # Buscar o último número gerado
            cursor.execute("SELECT last_number FROM bo_sequence")
            result = cursor.fetchone()

            if result:
                new_num = result[0] + 1  # Calcula o próximo número
            else:
                new_num = 1  # Se não houver registros, inicia em 1

            # Retorna o número da BO sem atualizar a sequência no banco
            current_year = datetime.datetime.now().strftime("%y")
            bo_number = f"BO{new_num:03d}/{current_year}"
            return bo_number

        except pyodbc.Error as e:
            print(f"Erro ao gerar número de BO: {e}")
            return "BO000/00"
        finally:
            if conn:
                cursor.close()
                conn.close()

    def salvar(self):
        """Salva a BO no banco e só então incrementa o número da sequência."""
        conn = create_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()
            valores = [self.bo_number]
            campos = [
                "OP", "Loja", "NF de envio", "NF de devolução", "OP de venda",
                "Tipo de Ocorrência", "Motivo", "Frete", "Setor responsável", "Status",
                "Previsão de embarque", "Descrição", "Filial"
            ]

            for campo in campos:
                valores.append(self.entries[campo].get())

            valores.append(self.ultimo_modulo)

            # Insere a BO no banco de dados
            cursor.execute('''
                INSERT INTO bo_records (
                    bo_number, op, loja, nf_envio, nf_devolucao, op_venda,
                    tipo_ocorrencia, motivo, frete, setor_responsavel, status,
                    previsao_embarque, descricao, modulo, filial
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', valores)
            conn.commit()

            # Atualiza a sequência APÓS salvar
            cursor.execute(
                "UPDATE bo_sequence SET last_number = last_number + 1")
            conn.commit()

            messagebox.showinfo("Sucesso", "BO salva com sucesso!")
            self.window.destroy()
            if self.ultimo_modulo == "corporativo":
                self.CorporativoModule.carregar_bos()  # Atualiza a lista de BOs

        except pyodbc.Error as e:
            messagebox.showerror("Erro", f"Erro ao salvar BO: {e}")
        finally:
            if conn:
                cursor.close()
                conn.close()

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
                entry = ttk.Combobox(self.window, values=["Corporativo", "Varejo", "Exportação"])
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
            module = 0 if self.entries["Módulo"].get() == "Corporativo" else 1 if self.entries["Módulo"].get() == "Varejo" else 2

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
            ttk.Label(self.window, text=f"{campo}:").grid(row=i, column=0, sticky=tk.W)

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
                    entry = ttk.Combobox(self.window, values=["Corporativo", "Varejo", "Exportação"])
            entry.grid(row=i, column=1)
            self.entries[campo] = entry

        ttk.Button(self.window, text="Salvar", command=self.salvar).grid(row=len(campos), columnspan=2)

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
            module = 0 if self.entries["Módulo"].get() == "Corporativo" else 1 if self.entries["Módulo"].get() == "Varejo" else 2

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