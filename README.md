# Botzap



"""
Automacao WhatsApp via Selenium - Edge
Interface grafica com CustomTkinter
"""

import sys
import os
import json
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import math
import threading
import pandas as pd
import time
import pyperclip
import customtkinter as ctk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ============================================================
# CONFIGURACOES
# ============================================================

PAUSA_ENTRE_MENSAGENS = 10

OPERADORES_PADRAO = {
    "Ana Paula": "11900000001",
}

def _pasta_app():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

ARQUIVO_OPERADORES = os.path.join(_pasta_app(), "operadores.json")

def carregar_operadores():
    try:
        if os.path.exists(ARQUIVO_OPERADORES):
            with open(ARQUIVO_OPERADORES, "r", encoding="utf-8") as f:
                dados = json.load(f)
                if dados:
                    return dados
    except Exception as e:
        print(f"[AVISO] Erro ao carregar operadores: {e}")
    salvar_operadores(dict(OPERADORES_PADRAO))
    return dict(OPERADORES_PADRAO)

def salvar_operadores(ops):
    try:
        caminho = os.path.join(_pasta_app(), "operadores.json")
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(ops, f, ensure_ascii=False, indent=2)
        print(f"[OK] Operadores salvos em: {caminho}")
        return True
    except Exception as e:
        print(f"[ERRO] Falha ao salvar operadores: {e}")
        return False

# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

parar_envio   = threading.Event()
driver_global = None


def extrair_numeros_lead(row):
    """Extrai tel1 a tel5 do lead, so o numero puro sem adicionar 55."""
    numeros = []
    for i in range(1, 6):
        tel = row.get(f"tel{i}")
        if pd.isna(tel) or str(tel).strip() in ("", "0", "nan"):
            continue
        tel_str = "".join(filter(str.isdigit, str(tel)))
        if tel_str:
            numeros.append(tel_str)
    return numeros


def distribuir_leads(df, operadores):
    nomes = list(operadores.keys())
    total = len(df)
    qtd_por_op = math.ceil(total / len(nomes))
    distribuicao = {}
    for i, nome in enumerate(nomes):
        inicio = i * qtd_por_op
        fim = min(inicio + qtd_por_op, total)
        distribuicao[nome] = df.iloc[inicio:fim]
    return distribuicao


def montar_mensagem(nome_operador, leads_operador):
    """Monta mensagem apenas com os numeros, sem nome do cliente."""
    linhas = [f"Ola {nome_operador}! Aqui estao seus leads para hoje:\n"]
    for _, row in leads_operador.iterrows():
        numeros = extrair_numeros_lead(row)
        for num in numeros:
            linhas.append(num)
    linhas.append("\nBom trabalho!")
    return "\n".join(linhas)


def iniciar_driver():
    pasta_script = _pasta_app()
    pasta_sessao = os.path.join(pasta_script, "whatsapp_session")
    os.makedirs(pasta_sessao, exist_ok=True)
    opcoes = Options()
    opcoes.add_argument(f"--user-data-dir={pasta_sessao}")
    opcoes.add_argument("--profile-directory=Default")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Edge(options=opcoes)
    driver.minimize_window()
    return driver


def enviar_mensagem(driver, numero_whatsapp, mensagem):
    """Envia mensagem para o operador. Adiciona 55 apenas no numero do operador."""
    numero_fmt = "55" + "".join(filter(str.isdigit, numero_whatsapp))
    url = f"https://web.whatsapp.com/send?phone={numero_fmt}"
    driver.get(url)
    try:
        campo = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
            )
        )
        time.sleep(2)
        pyperclip.copy(mensagem)
        campo.click()
        campo.send_keys(Keys.CONTROL, "v")
        time.sleep(1)
        campo.send_keys(Keys.ENTER)
        return True
    except Exception as e:
        print(f"[ERRO enviar_mensagem] {e}")
        return False


# ============================================================
# JANELA DE OPERADORES
# ============================================================

class JanelaOperadores(ctk.CTkToplevel):
    def __init__(self, master, callback_atualizar):
        super().__init__(master)
        self.title("Gerenciar Operadores")
        self.geometry("520x750")
        self.resizable(True, True)
        self.grab_set()
        self.callback_atualizar = callback_atualizar
        self.operadores = carregar_operadores()
        self.scroll = None
        self._build_ui()
        self._atualizar_lista()

    def _build_ui(self):
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Gerenciar Operadores", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 2))
        ctk.CTkLabel(self, text="Adicione, edite ou remova operadores", text_color="gray", font=ctk.CTkFont(size=12)).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 8))

        # Formulario
        frame_form = ctk.CTkFrame(self, corner_radius=12)
        frame_form.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 10))
        frame_form.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame_form, text="Nome do Operador", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 2))
        self.entry_nome = ctk.CTkEntry(frame_form, placeholder_text="Ex: Ana Paula", height=36)
        self.entry_nome.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 8))

        ctk.CTkLabel(frame_form, text="WhatsApp (DDD + numero, sem +55)", font=ctk.CTkFont(size=12, weight="bold")).grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 2))
        self.entry_tel = ctk.CTkEntry(frame_form, placeholder_text="Ex: 11999990001", height=36)
        self.entry_tel.grid(row=3, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 10))

        ctk.CTkButton(frame_form, text="Adicionar / Atualizar", height=36, command=self._adicionar).grid(row=4, column=0, sticky="ew", padx=(14, 6), pady=(0, 14))
        ctk.CTkButton(frame_form, text="Limpar", height=36, width=90, fg_color="gray30", hover_color="gray20", command=self._limpar_form).grid(row=4, column=1, sticky="e", padx=(0, 14), pady=(0, 14))

        # Lista
        frame_lista = ctk.CTkFrame(self, corner_radius=12)
        frame_lista.grid(row=3, column=0, sticky="nsew", padx=24, pady=(0, 8))
        frame_lista.grid_rowconfigure(1, weight=1)
        frame_lista.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame_lista, text="Operadores Cadastrados", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 4))
        self.scroll = ctk.CTkScrollableFrame(frame_lista, height=180)
        self.scroll.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))

        # Botao fixo na base
        ctk.CTkButton(self, text="Salvar e Fechar", height=44,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._salvar_fechar).grid(row=4, column=0, sticky="ew", padx=24, pady=(0, 20))

    def _atualizar_lista(self):
        if self.scroll is None:
            return
        for w in self.scroll.winfo_children():
            w.destroy()

        if not self.operadores:
            ctk.CTkLabel(self.scroll, text="Nenhum operador cadastrado.", text_color="gray").pack(pady=10)
            return

        for nome, tel in self.operadores.items():
            row = ctk.CTkFrame(self.scroll, corner_radius=8, fg_color=("gray85", "gray20"))
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=nome, font=ctk.CTkFont(size=12, weight="bold"), anchor="w").pack(side="left", padx=10, pady=8, fill="x", expand=True)
            ctk.CTkLabel(row, text=tel, text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left", padx=6)
            ctk.CTkButton(row, text="Editar", width=60, height=28, fg_color="gray40", hover_color="gray30",
                command=lambda n=nome: self._editar(n)).pack(side="left", padx=4, pady=6)
            ctk.CTkButton(row, text="Remover", width=70, height=28, fg_color="#c0392b", hover_color="#a93226",
                command=lambda n=nome: self._remover(n)).pack(side="left", padx=(0, 8), pady=6)

    def _adicionar(self):
        nome = self.entry_nome.get().strip()
        tel  = "".join(filter(str.isdigit, self.entry_tel.get().strip()))
        if not nome or not tel:
            messagebox.showwarning("Atencao", "Preencha nome e telefone.", parent=self)
            return
        self.operadores[nome] = tel
        self._limpar_form()
        self._atualizar_lista()

    def _editar(self, nome):
        self.entry_nome.delete(0, "end")
        self.entry_nome.insert(0, nome)
        self.entry_tel.delete(0, "end")
        self.entry_tel.insert(0, self.operadores[nome])

    def _remover(self, nome):
        if messagebox.askyesno("Confirmar", f"Remover '{nome}'?", parent=self):
            del self.operadores[nome]
            self._atualizar_lista()

    def _limpar_form(self):
        self.entry_nome.delete(0, "end")
        self.entry_tel.delete(0, "end")

    def _salvar_fechar(self):
        if not self.operadores:
            messagebox.showwarning("Atencao", "Nenhum operador cadastrado.", parent=self)
            return
        ok = salvar_operadores(self.operadores)
        if ok:
            messagebox.showinfo("Salvo", f"{len(self.operadores)} operador(es) salvos!", parent=self)
            self.callback_atualizar(self.operadores)
            self.destroy()
        else:
            messagebox.showerror("Erro", "Nao foi possivel salvar. Verifique as permissoes da pasta.", parent=self)


# ============================================================
# JANELA PRINCIPAL
# ============================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Bot WhatsApp - Distribuicao de Leads")
        self.geometry("700x740")
        self.resizable(False, False)
        self.arquivo_excel = None
        self.operadores = carregar_operadores()
        self._build_ui()

    def _build_ui(self):
        frame_header = ctk.CTkFrame(self, fg_color="transparent")
        frame_header.pack(fill="x", padx=30, pady=(24, 0))

        ctk.CTkLabel(frame_header, text="Bot WhatsApp", font=ctk.CTkFont(size=26, weight="bold")).pack(side="left")
        ctk.CTkButton(frame_header, text="Operadores", width=130, height=34,
            fg_color="gray30", hover_color="gray20",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.abrir_operadores).pack(side="right")

        ctk.CTkLabel(self, text="Distribuicao automatica de leads para operadores", font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(4, 16))

        frame_arquivo = ctk.CTkFrame(self, corner_radius=12)
        frame_arquivo.pack(fill="x", padx=30, pady=(0, 12))
        ctk.CTkLabel(frame_arquivo, text="Arquivo de Leads", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 4))
        row_arquivo = ctk.CTkFrame(frame_arquivo, fg_color="transparent")
        row_arquivo.pack(fill="x", padx=16, pady=(0, 12))
        self.label_arquivo = ctk.CTkLabel(row_arquivo, text="Nenhum arquivo selecionado", text_color="gray", anchor="w")
        self.label_arquivo.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row_arquivo, text="Selecionar Excel", width=150, command=self.selecionar_arquivo).pack(side="right")

        frame_info = ctk.CTkFrame(self, corner_radius=12)
        frame_info.pack(fill="x", padx=30, pady=(0, 12))
        self.label_leads  = ctk.CTkLabel(frame_info, text="Leads: --", font=ctk.CTkFont(size=13))
        self.label_ops    = ctk.CTkLabel(frame_info, text=f"Operadores: {len(self.operadores)}", font=ctk.CTkFont(size=13))
        self.label_por_op = ctk.CTkLabel(frame_info, text="Leads por operador: --", font=ctk.CTkFont(size=13))
        self.label_leads.pack(side="left", padx=20, pady=12)
        self.label_ops.pack(side="left", padx=20, pady=12)
        self.label_por_op.pack(side="left", padx=20, pady=12)

        frame_prog = ctk.CTkFrame(self, corner_radius=12)
        frame_prog.pack(fill="x", padx=30, pady=(0, 12))
        ctk.CTkLabel(frame_prog, text="Progresso", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 4))
        self.barra = ctk.CTkProgressBar(frame_prog, height=18, corner_radius=8)
        self.barra.pack(fill="x", padx=16, pady=(0, 8))
        self.barra.set(0)
        self.label_progresso = ctk.CTkLabel(frame_prog, text="Aguardando inicio...", text_color="gray", font=ctk.CTkFont(size=12))
        self.label_progresso.pack(anchor="w", padx=16, pady=(0, 12))

        frame_log = ctk.CTkFrame(self, corner_radius=12)
        frame_log.pack(fill="both", expand=True, padx=30, pady=(0, 12))
        ctk.CTkLabel(frame_log, text="Log de Atividades", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 4))
        self.log_box = ctk.CTkTextbox(frame_log, height=160, font=ctk.CTkFont(size=12), state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        frame_btns.pack(fill="x", padx=30, pady=(0, 20))
        self.btn_iniciar = ctk.CTkButton(frame_btns, text="Iniciar Envio", height=42,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.iniciar_envio)
        self.btn_iniciar.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.btn_parar = ctk.CTkButton(frame_btns, text="Parar", height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#e74c3c", hover_color="#c0392b",
            state="disabled", command=self.parar)
        self.btn_parar.pack(side="left", fill="x", expand=True, padx=(8, 0))

    def log(self, texto):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", texto + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def abrir_operadores(self):
        JanelaOperadores(self, self._atualizar_operadores)

    def _atualizar_operadores(self, novos_ops):
        self.operadores = novos_ops
        self.label_ops.configure(text=f"Operadores: {len(self.operadores)}")
        self.log(f"[OK] Operadores atualizados: {len(self.operadores)} cadastrados.")

    def selecionar_arquivo(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.arquivo_excel = path
            nome = os.path.basename(path)
            self.label_arquivo.configure(text=nome, text_color="white")
            try:
                df = pd.read_excel(path)
                total = len(df)
                por_op = math.ceil(total / len(self.operadores)) if self.operadores else 0
                self.label_leads.configure(text=f"Leads: {total}")
                self.label_por_op.configure(text=f"Leads por operador: {por_op}")
                self.log(f"[OK] Arquivo carregado: {nome} ({total} leads)")
            except Exception as e:
                self.log(f"[ERRO] Nao foi possivel ler o arquivo: {e}")

    def parar(self):
        parar_envio.set()
        self.log("[PARADO] Envio interrompido pelo usuario.")
        self.btn_parar.configure(state="disabled")
        self.btn_iniciar.configure(state="normal")

    def iniciar_envio(self):
        if not self.arquivo_excel:
            messagebox.showwarning("Atencao", "Selecione o arquivo Excel antes de iniciar.")
            return
        if not self.operadores:
            messagebox.showwarning("Atencao", "Cadastre ao menos um operador antes de iniciar.")
            return
        parar_envio.clear()
        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.barra.set(0)
        self.label_progresso.configure(text="Iniciando...")
        t = threading.Thread(target=self.executar, daemon=True)
        t.start()

    def executar(self):
        global driver_global
        try:
            self.log("=" * 45)
            self.log("   Iniciando envio de leads")
            self.log("=" * 45)

            self.log(f"\n[1/4] Lendo arquivo: {os.path.basename(self.arquivo_excel)}")
            df = pd.read_excel(self.arquivo_excel)
            self.log(f"      {len(df)} leads encontrados.")

            self.log(f"\n[2/4] Distribuindo entre {len(self.operadores)} operadores...")
            distribuicao = distribuir_leads(df, self.operadores)
            for nome, leads in distribuicao.items():
                self.log(f"      {nome}: {len(leads)} leads")

            self.log("\n[3/4] Iniciando Edge (minimizado)...")
            driver_global = iniciar_driver()

            self.log("[LOGIN] Abrindo WhatsApp Web...")
            driver_global.get("https://web.whatsapp.com")
            self.log("       Aguardando 60s para login/verificacao...")
            for i in range(60):
                if parar_envio.is_set():
                    break
                time.sleep(1)

            if parar_envio.is_set():
                driver_global.quit()
                return

            self.log("\n[4/4] Enviando mensagens...")
            total_ops = len(self.operadores)
            enviados  = 0
            falhas    = 0

            for i, (nome, whatsapp) in enumerate(self.operadores.items(), start=1):
                if parar_envio.is_set():
                    self.log("\n[PARADO] Envio cancelado.")
                    break

                leads_op = distribuicao.get(nome)
                if leads_op is None or len(leads_op) == 0:
                    self.log(f"\n[{i}/{total_ops}] {nome} - sem leads, pulando.")
                    continue

                mensagem = montar_mensagem(nome, leads_op)
                self.log(f"\n[{i}/{total_ops}] Enviando para: {nome} (WA: {whatsapp})")
                self.log(f"      Total numeros na mensagem: {len(leads_op)}")
                self.label_progresso.configure(text=f"Enviando para {nome} ({i}/{total_ops})")

                sucesso = enviar_mensagem(driver_global, whatsapp, mensagem)

                if sucesso:
                    enviados += 1
                    self.log(f"      [OK] Enviado com sucesso!")
                else:
                    falhas += 1
                    self.log(f"      [ERRO] Falha no envio.")

                self.barra.set(i / total_ops)
                self.label_progresso.configure(text=f"{i}/{total_ops} operadores concluidos")

                if i < total_ops and not parar_envio.is_set():
                    self.log(f"      Aguardando {PAUSA_ENTRE_MENSAGENS}s...")
                    for _ in range(PAUSA_ENTRE_MENSAGENS):
                        if parar_envio.is_set():
                            break
                        time.sleep(1)

            self.barra.set(1)
            self.log("\n" + "=" * 45)
            self.log(f"   Concluido! Enviados: {enviados}  Falhas: {falhas}")
            self.log("=" * 45)
            self.label_progresso.configure(text=f"Concluido! {enviados} enviados, {falhas} falhas.")
            driver_global.quit()

        except Exception as e:
            self.log(f"\n[ERRO GERAL] {e}")
        finally:
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
