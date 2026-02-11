import pandas as pd
import tkinter as tk
from tkinter import messagebox
import pyperclip # Biblioteca mágica do Ctrl+C
import os

class LoginHelperApp:
    def __init__(self, root, excel_path):
        self.root = root
        self.root.title("Assistente de Login 🤖")
        self.root.geometry("400x350")
        
        # MANTÉM A JANELA SEMPRE NO TOPO (Pulo do gato!)
        self.root.attributes('-topmost', True)
        
        self.dados = []
        self.index_atual = 0
        
        self.carregar_excel(excel_path)
        self.montar_interface()
        self.atualizar_tela()

    def carregar_excel(self, path):
        if not os.path.exists(path):
            messagebox.showerror("Erro", "Planilha não encontrada!")
            self.root.destroy()
            return
        
        try:
            df = pd.read_excel(path)
            # Converte para lista de dicionários para facilitar
            self.dados = df.to_dict('records')
            print(f"Carregados {len(self.dados)} usuários.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler Excel: {e}")

    def montar_interface(self):
        # Estilo
        font_lbl = ("Arial", 10, "bold")
        font_data = ("Consolas", 12)
        bg_color = "#f0f0f0"
        self.root.configure(bg=bg_color)

        # --- CONTADOR ---
        self.lbl_contador = tk.Label(self.root, text="0/0", bg=bg_color)
        self.lbl_contador.pack(pady=5)

        # --- UC ---
        tk.Label(self.root, text="CONTA CONTRATO (UC):", bg=bg_color, font=font_lbl).pack()
        self.lbl_uc = tk.Label(self.root, text="---", font=("Arial", 14, "bold"), fg="blue", bg=bg_color)
        self.lbl_uc.pack(pady=5)

        # --- LOGIN ---
        frame_login = tk.Frame(self.root, bg=bg_color)
        frame_login.pack(pady=10, fill="x", padx=20)
        
        tk.Label(frame_login, text="LOGIN (CPF/CNPJ):", bg=bg_color, font=font_lbl).pack(anchor="w")
        self.entry_login = tk.Entry(frame_login, font=font_data)
        self.entry_login.pack(fill="x", pady=2)
        
        btn_copy_login = tk.Button(frame_login, text="COPIAR LOGIN 📋", bg="#dddddd", 
                                   command=self.copiar_login)
        btn_copy_login.pack(fill="x")

        # --- SENHA ---
        frame_senha = tk.Frame(self.root, bg=bg_color)
        frame_senha.pack(pady=10, fill="x", padx=20)
        
        tk.Label(frame_senha, text="SENHA:", bg=bg_color, font=font_lbl).pack(anchor="w")
        self.entry_senha = tk.Entry(frame_senha, font=font_data)
        self.entry_senha.pack(fill="x", pady=2)
        
        btn_copy_senha = tk.Button(frame_senha, text="COPIAR SENHA 🔑", bg="#dddddd", 
                                   command=self.copiar_senha)
        btn_copy_senha.pack(fill="x")

        # --- NAVEGAÇÃO ---
        frame_nav = tk.Frame(self.root, bg=bg_color)
        frame_nav.pack(pady=20, fill="x", padx=20)

        btn_prev = tk.Button(frame_nav, text="<< Anterior", command=self.voltar)
        btn_prev.pack(side="left", expand=True, fill="x", padx=5)

        btn_next = tk.Button(frame_nav, text="Próximo >>", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
                             command=self.avancar)
        btn_next.pack(side="right", expand=True, fill="x", padx=5)

    def atualizar_tela(self):
        if not self.dados: return
        
        item = self.dados[self.index_atual]
        
        # Formata dados
        uc = str(item.get('Conta Contrato', '')).replace('.0', '')
        login = str(item.get('CNPJ/CPF', '')).replace('.0', '').replace('.', '').replace('-', '').replace('/', '')
        senha = str(item.get('Acesso equatorial', '')).strip()

        # Atualiza Interface
        self.lbl_contador.config(text=f"Usuário {self.index_atual + 1} de {len(self.dados)}")
        self.lbl_uc.config(text=uc)
        
        self.entry_login.delete(0, tk.END)
        self.entry_login.insert(0, login)
        
        self.entry_senha.delete(0, tk.END)
        self.entry_senha.insert(0, senha)

    def copiar_login(self):
        texto = self.entry_login.get()
        pyperclip.copy(texto)
        # Pisca a cor para dar feedback
        self.root.config(bg="#d1e7dd")
        self.root.after(200, lambda: self.root.config(bg="#f0f0f0"))

    def copiar_senha(self):
        texto = self.entry_senha.get()
        pyperclip.copy(texto)
        self.root.config(bg="#d1e7dd")
        self.root.after(200, lambda: self.root.config(bg="#f0f0f0"))

    def avancar(self):
        if self.index_atual < len(self.dados) - 1:
            self.index_atual += 1
            self.atualizar_tela()
        else:
            messagebox.showinfo("Fim", "Você chegou no último usuário!")

    def voltar(self):
        if self.index_atual > 0:
            self.index_atual -= 1
            self.atualizar_tela()

if __name__ == "__main__":
    # Caminho do arquivo
    base_dir = os.getcwd()
    arquivo_excel = os.path.join(base_dir, "output", "Cad_RateioConsumo_Final.xlsx")

    root = tk.Tk()
    app = LoginHelperApp(root, arquivo_excel)
    root.mainloop()