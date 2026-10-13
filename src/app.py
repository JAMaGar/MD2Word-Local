# -*- coding: utf-8 -*-
"""
MD2Word Local - interface grafica para converter Markdown em Word, 100% offline.

Uso:
  MD2Word.exe                         abre a janela
  MD2Word.exe a.md b.md               abre a janela com os arquivos ja na lista
  MD2Word.exe --rapido docx a.md      converte direto (usado pelo menu de contexto do Explorer)
"""
from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
from dataclasses import asdict
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import conversor as cv

APP_NOME = "MD2Word Local"
EXT_MD = (".md", ".markdown", ".mdown", ".mkd", ".txt")

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # arrastar-e-soltar (opcional)
except Exception:  # pragma: no cover
    TkinterDnD = None
    DND_FILES = None


# --------------------------------------------------------------------------------------
# Configuracao persistente (%APPDATA%\MD2Word\config.json)
# --------------------------------------------------------------------------------------
def pasta_config() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home() / ".config")
    p = Path(base) / "MD2Word"
    p.mkdir(parents=True, exist_ok=True)
    return p


CONFIG_PADRAO = {
    "formatos": ["docx"],
    "saida_mesma_pasta": True,
    "pasta_saida": "",
    "abrir_apos": True,
    "sobrescrever": True,
    "ultima_pasta": "",
    "opcoes": asdict(cv.Opcoes()),
}


def carregar_config() -> dict:
    cfg = json.loads(json.dumps(CONFIG_PADRAO))
    try:
        dados = json.loads((pasta_config() / "config.json").read_text(encoding="utf-8"))
        cfg.update({k: v for k, v in dados.items() if k in cfg and k != "opcoes"})
        cfg["opcoes"].update({k: v for k, v in dados.get("opcoes", {}).items()
                              if k in cfg["opcoes"]})
    except Exception:
        pass
    return cfg


def salvar_config(cfg: dict):
    try:
        (pasta_config() / "config.json").write_text(json.dumps(cfg, indent=2, ensure_ascii=False),
                                                    encoding="utf-8")
    except Exception:
        pass


def recurso(nome: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    for c in (base / "assets" / nome, Path(__file__).resolve().parent.parent / "assets" / nome):
        if c.exists():
            return c
    return base / "assets" / nome


def abrir_arquivo(path):
    try:
        if os.name == "nt":
            os.startfile(str(path))  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


# --------------------------------------------------------------------------------------
# Modo rapido (menu de contexto): converte sem abrir a janela
# --------------------------------------------------------------------------------------
def modo_rapido(formato: str, arquivos: list[str]):
    cfg = carregar_config()
    op = cv.Opcoes(**cfg["opcoes"])
    mensagens, erros, gerados = [], [], []
    for f in arquivos:
        try:
            gerados += cv.converter_arquivo(f, None, [formato], op, sobrescrever=cfg["sobrescrever"],
                                            log=mensagens.append)
        except Exception as e:
            erros.append(f"{Path(f).name}: {e}")
    erros += [m.strip() for m in mensagens if m.strip().startswith("X")]
    if gerados and cfg.get("abrir_apos", True):
        for g in gerados:
            abrir_arquivo(g)
    if erros or not gerados:
        root = tk.Tk(); root.withdraw()
        texto = "\n".join(erros + [m.strip() for m in mensagens if m.strip().startswith("!")])
        (messagebox.showwarning if gerados else messagebox.showerror)(
            APP_NOME, texto or "Nenhum arquivo foi gerado.")
        root.destroy()
    elif not cfg.get("abrir_apos", True):
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo(APP_NOME, "Arquivo gerado:\n" + "\n".join(str(g) for g in gerados))
        root.destroy()


# --------------------------------------------------------------------------------------
# Janela principal
# --------------------------------------------------------------------------------------
class App:
    def __init__(self, root: tk.Tk, arquivos_iniciais=()):
        self.root = root
        self.cfg = carregar_config()
        self.fila: queue.Queue = queue.Queue()
        self.trabalhando = False
        self.arquivos: list[str] = []

        root.title(f"{APP_NOME}  {cv.__version__}  -  Markdown para Word (offline)")
        root.minsize(760, 640)
        try:
            if os.name == "nt":
                root.iconbitmap(default=str(recurso("md2word.ico")))
        except Exception:
            pass
        self._estilo()
        self._montar()
        for a in arquivos_iniciais:
            self._adicionar(a)
        self._status_ferramentas()
        root.after(100, self._processar_fila)
        root.protocol("WM_DELETE_WINDOW", self._fechar)

    # ------------------------------------------------------------------ layout
    def _estilo(self):
        st = ttk.Style()
        if os.name == "nt" and "vista" in st.theme_names():
            st.theme_use("vista")
        st.configure("Grande.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        st.configure("Dica.TLabel", foreground="#666666")

    def _montar(self):
        r = self.root
        op = self.cfg["opcoes"]
        principal = ttk.Frame(r, padding=10)
        principal.pack(fill="both", expand=True)

        self.nb = ttk.Notebook(principal)
        self.nb.pack(fill="both", expand=True)

        # --- aba 1: arquivos
        aba1 = ttk.Frame(self.nb, padding=8)
        self.nb.add(aba1, text="  Arquivos .md  ")
        dica = ("Arraste arquivos .md para a lista ou use os botoes ao lado."
                if TkinterDnD else "Use os botoes ao lado para adicionar arquivos .md.")
        ttk.Label(aba1, text=dica, style="Dica.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        lf = ttk.Frame(aba1)
        lf.grid(row=1, column=0, sticky="nsew")
        self.lista = tk.Listbox(lf, selectmode="extended", activestyle="none", height=10,
                                font=("Segoe UI", 10))
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=sb.set)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        if TkinterDnD:
            self.lista.drop_target_register(DND_FILES)
            self.lista.dnd_bind("<<Drop>>", self._soltar)
        bts = ttk.Frame(aba1)
        bts.grid(row=1, column=1, sticky="n", padx=(8, 0))
        for txt, cmd in (("Adicionar arquivos...", self._escolher_arquivos),
                         ("Adicionar pasta...", self._escolher_pasta),
                         ("Remover selecionados", self._remover),
                         ("Limpar lista", self._limpar)):
            ttk.Button(bts, text=txt, command=cmd, width=22).pack(fill="x", pady=2)
        aba1.columnconfigure(0, weight=1)
        aba1.rowconfigure(1, weight=1)
        self.lista.bind("<Delete>", lambda e: self._remover())

        # --- aba 2: colar texto
        aba2 = ttk.Frame(self.nb, padding=8)
        self.nb.add(aba2, text="  Colar Markdown  ")
        ttk.Label(aba2, text="Cole aqui o texto em Markdown (Ctrl+V) e clique em Converter.",
                  style="Dica.TLabel").pack(anchor="w", pady=(0, 4))
        self.texto = ScrolledText(aba2, wrap="word", font=("Consolas", 10), undo=True, height=12)
        self.texto.pack(fill="both", expand=True)
        b2 = ttk.Frame(aba2)
        b2.pack(fill="x", pady=(6, 0))
        ttk.Button(b2, text="Colar da area de transferencia", command=self._colar).pack(side="left")
        ttk.Button(b2, text="Limpar", command=lambda: self.texto.delete("1.0", "end")).pack(side="left", padx=6)

        # --- opcoes
        of = ttk.LabelFrame(principal, text="Opcoes", padding=8)
        of.pack(fill="x", pady=(8, 0))

        linha = ttk.Frame(of); linha.pack(fill="x")
        ttk.Label(linha, text="Gerar:").pack(side="left")
        self.v_docx = tk.BooleanVar(value="docx" in self.cfg["formatos"])
        self.v_doc = tk.BooleanVar(value="doc" in self.cfg["formatos"])
        self.v_pdf = tk.BooleanVar(value="pdf" in self.cfg["formatos"])
        ttk.Checkbutton(linha, text=".docx", variable=self.v_docx).pack(side="left", padx=4)
        ttk.Checkbutton(linha, text=".doc (Word 97-2003)", variable=self.v_doc).pack(side="left", padx=4)
        ttk.Checkbutton(linha, text=".pdf", variable=self.v_pdf).pack(side="left", padx=4)
        self.lbl_ferr = ttk.Label(linha, text="", style="Dica.TLabel")
        self.lbl_ferr.pack(side="right")

        linha = ttk.Frame(of); linha.pack(fill="x", pady=(6, 0))
        self.v_mesma = tk.BooleanVar(value=self.cfg["saida_mesma_pasta"])
        ttk.Radiobutton(linha, text="Salvar na mesma pasta do .md", variable=self.v_mesma,
                        value=True).pack(side="left")
        ttk.Radiobutton(linha, text="Salvar em:", variable=self.v_mesma, value=False).pack(side="left", padx=(12, 4))
        self.v_pasta = tk.StringVar(value=self.cfg["pasta_saida"])
        ttk.Entry(linha, textvariable=self.v_pasta).pack(side="left", fill="x", expand=True)
        ttk.Button(linha, text="Escolher...", command=self._escolher_saida).pack(side="left", padx=4)

        linha = ttk.Frame(of); linha.pack(fill="x", pady=(6, 0))
        ttk.Label(linha, text="Papel:").pack(side="left")
        self.v_papel = tk.StringVar(value=op["papel"])
        ttk.Combobox(linha, textvariable=self.v_papel, values=["A4", "Carta"], width=6,
                     state="readonly").pack(side="left", padx=(4, 12))
        ttk.Label(linha, text="Fonte:").pack(side="left")
        self.v_fonte = tk.StringVar(value=op["fonte"])
        ttk.Combobox(linha, textvariable=self.v_fonte, width=16,
                     values=["Calibri", "Aptos", "Arial", "Times New Roman", "Cambria", "Georgia",
                             "Segoe UI", "Verdana", "Garamond"]).pack(side="left", padx=(4, 12))
        ttk.Label(linha, text="Tamanho:").pack(side="left")
        self.v_tam = tk.StringVar(value=str(op["tamanho_pt"]).rstrip("0").rstrip("."))
        ttk.Spinbox(linha, from_=8, to=16, increment=0.5, textvariable=self.v_tam, width=5).pack(side="left", padx=(4, 12))
        ttk.Label(linha, text="Margens (cm):").pack(side="left")
        self.v_margem = tk.StringVar(value=str(op["margem_cm"]))
        ttk.Spinbox(linha, from_=1, to=4, increment=0.5, textvariable=self.v_margem, width=5).pack(side="left", padx=4)

        linha = ttk.Frame(of); linha.pack(fill="x", pady=(6, 0))
        self.v_just = tk.BooleanVar(value=op["justificar"])
        self.v_eq = tk.BooleanVar(value=op["equacoes_nativas"])
        self.v_abrir = tk.BooleanVar(value=self.cfg["abrir_apos"])
        self.v_sobre = tk.BooleanVar(value=not self.cfg["sobrescrever"])
        ttk.Checkbutton(linha, text="Texto justificado", variable=self.v_just).pack(side="left")
        ttk.Checkbutton(linha, text="Equacoes editaveis do Word", variable=self.v_eq).pack(side="left", padx=10)
        ttk.Checkbutton(linha, text="Abrir ao terminar", variable=self.v_abrir).pack(side="left", padx=10)
        ttk.Checkbutton(linha, text="Nao sobrescrever (criar copia)", variable=self.v_sobre).pack(side="left", padx=10)

        linha = ttk.Frame(of); linha.pack(fill="x", pady=(6, 0))
        ttk.Label(linha, text="Modelo .docx (opcional):").pack(side="left")
        self.v_modelo = tk.StringVar(value=op["modelo_docx"])
        ttk.Entry(linha, textvariable=self.v_modelo).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(linha, text="Escolher...", command=self._escolher_modelo).pack(side="left")
        ttk.Button(linha, text="X", width=3, command=lambda: self.v_modelo.set("")).pack(side="left", padx=(4, 0))

        # --- acao + log
        af = ttk.Frame(principal)
        af.pack(fill="x", pady=(8, 0))
        self.bt = ttk.Button(af, text="Converter", style="Grande.TButton", command=self._converter)
        self.bt.pack(side="left")
        self.prog = ttk.Progressbar(af, mode="determinate")
        self.prog.pack(side="left", fill="x", expand=True, padx=10)
        ttk.Button(af, text="Abrir pasta de saida", command=self._abrir_pasta_saida).pack(side="right")

        self.log = ScrolledText(principal, height=7, font=("Consolas", 9), state="disabled",
                                background="#FAFAFA")
        self.log.pack(fill="both", expand=False, pady=(8, 0))
        self.ultima_saida = None
        root = self.root
        root.bind("<Control-Return>", lambda e: self._converter())

    # ------------------------------------------------------------------ status
    def _status_ferramentas(self):
        def checar():
            word = cv.word_disponivel()
            lo = cv.localizar_libreoffice() is not None
            xsl = cv.localizar_mml2omml() is not None
            partes = []
            partes.append("Word: detectado" if word else ("LibreOffice: detectado" if lo else
                                                          "Word/LibreOffice: nao encontrado (.doc/.pdf indisponiveis)"))
            partes.append("Equacoes: nativas" if xsl else "Equacoes: como texto")
            self.fila.put(("status", "   |   ".join(partes)))
        threading.Thread(target=checar, daemon=True).start()

    # ------------------------------------------------------------------ lista
    def _adicionar(self, caminho):
        p = Path(caminho)
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.suffix.lower() in EXT_MD[:-1]:
                    self._adicionar(f)
            return
        if not p.is_file():
            return
        s = str(p.resolve())
        if s not in self.arquivos:
            self.arquivos.append(s)
            self.lista.insert("end", s)
        self.nb.select(0)

    def _soltar(self, evento):
        for item in self.root.tk.splitlist(evento.data):
            self._adicionar(item)

    def _escolher_arquivos(self):
        fs = filedialog.askopenfilenames(
            title="Escolha os arquivos Markdown", initialdir=self.cfg.get("ultima_pasta") or None,
            filetypes=[("Markdown", "*.md *.markdown *.mdown *.mkd"), ("Texto", "*.txt"), ("Todos", "*.*")])
        for f in fs:
            self._adicionar(f)
        if fs:
            self.cfg["ultima_pasta"] = str(Path(fs[0]).parent)

    def _escolher_pasta(self):
        d = filedialog.askdirectory(title="Escolha uma pasta (inclui subpastas)",
                                    initialdir=self.cfg.get("ultima_pasta") or None)
        if d:
            self._adicionar(d)
            self.cfg["ultima_pasta"] = d

    def _remover(self):
        for i in reversed(self.lista.curselection()):
            self.lista.delete(i)
            del self.arquivos[i]

    def _limpar(self):
        self.lista.delete(0, "end")
        self.arquivos.clear()

    def _colar(self):
        try:
            self.texto.insert("insert", self.root.clipboard_get())
        except tk.TclError:
            pass

    def _escolher_saida(self):
        d = filedialog.askdirectory(title="Pasta de saida")
        if d:
            self.v_pasta.set(d)
            self.v_mesma.set(False)

    def _escolher_modelo(self):
        f = filedialog.askopenfilename(title="Modelo .docx (os estilos dele serao usados)",
                                       filetypes=[("Documento do Word", "*.docx *.dotx")])
        if f:
            self.v_modelo.set(f)

    def _abrir_pasta_saida(self):
        alvo = self.ultima_saida
        if not alvo:
            if not self.v_mesma.get() and self.v_pasta.get():
                alvo = self.v_pasta.get()
            elif self.arquivos:
                alvo = str(Path(self.arquivos[0]).parent)
        if alvo:
            abrir_arquivo(alvo)

    # ------------------------------------------------------------------ conversao
    def _opcoes(self) -> cv.Opcoes:
        def num(v, padrao):
            try:
                return float(str(v).replace(",", "."))
            except ValueError:
                return padrao
        return cv.Opcoes(papel=self.v_papel.get(), margem_cm=num(self.v_margem.get(), 2.5),
                         fonte=self.v_fonte.get().strip() or "Calibri",
                         tamanho_pt=num(self.v_tam.get(), 11.0),
                         modelo_docx=self.v_modelo.get().strip(),
                         equacoes_nativas=self.v_eq.get(), justificar=self.v_just.get())

    def _formatos(self):
        f = [x for x, v in (("docx", self.v_docx), ("doc", self.v_doc), ("pdf", self.v_pdf)) if v.get()]
        return f

    def _salvar_estado(self):
        self.cfg["formatos"] = self._formatos() or ["docx"]
        self.cfg["saida_mesma_pasta"] = bool(self.v_mesma.get())
        self.cfg["pasta_saida"] = self.v_pasta.get()
        self.cfg["abrir_apos"] = bool(self.v_abrir.get())
        self.cfg["sobrescrever"] = not self.v_sobre.get()
        self.cfg["opcoes"] = asdict(self._opcoes())
        salvar_config(self.cfg)

    def _escrever(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _converter(self):
        if self.trabalhando:
            return
        self._salvar_estado()
        op = self._opcoes()
        if op.modelo_docx and not Path(op.modelo_docx).is_file():
            messagebox.showwarning(APP_NOME, "O arquivo de modelo .docx nao foi encontrado.")
            return
        if self.nb.index("current") == 1:
            self._converter_texto(op)
            return
        formatos = self._formatos()
        if not formatos:
            messagebox.showwarning(APP_NOME, "Marque pelo menos um formato de saida.")
            return
        if not self.arquivos:
            messagebox.showinfo(APP_NOME, "Adicione pelo menos um arquivo .md.")
            return
        pasta = None if self.v_mesma.get() else (self.v_pasta.get().strip() or None)
        self.trabalhando = True
        self.bt.configure(state="disabled")
        self.prog.configure(maximum=len(self.arquivos), value=0)
        threading.Thread(target=self._trabalho, args=(list(self.arquivos), pasta, formatos, op,
                                                      not self.v_sobre.get(), self.v_abrir.get()),
                         daemon=True).start()

    def _trabalho(self, arquivos, pasta, formatos, op, sobrescrever, abrir):
        log = lambda m: self.fila.put(("log", m))  # noqa: E731
        total_ok, gerados_todos = 0, []
        for i, f in enumerate(arquivos, 1):
            log(f"> {Path(f).name}")
            try:
                gerados = cv.converter_arquivo(f, pasta, formatos, op, sobrescrever=sobrescrever, log=log)
                for g in gerados:
                    log(f"  OK {g}")
                gerados_todos += gerados
                total_ok += 1 if gerados else 0
            except Exception as e:
                log(f"  X Erro: {e}")
                log("    " + traceback.format_exc().strip().splitlines()[-1])
            self.fila.put(("prog", i))
        if abrir and gerados_todos:
            if len(gerados_todos) <= 3:
                for g in gerados_todos:
                    abrir_arquivo(g)
            else:
                abrir_arquivo(Path(gerados_todos[0]).parent)
        self.fila.put(("fim", (total_ok, len(arquivos), gerados_todos)))

    def _converter_texto(self, op):
        texto = self.texto.get("1.0", "end").strip()
        if not texto:
            messagebox.showinfo(APP_NOME, "Cole algum texto em Markdown primeiro.")
            return
        m = re.search(r"^\s*#\s+(.+)$", texto, flags=re.M)
        nome = re.sub(r'[\\/:*?"<>|]+', "", m.group(1)).strip()[:80] if m else "documento"
        destino = filedialog.asksaveasfilename(
            title="Salvar como", defaultextension=".docx", initialfile=nome or "documento",
            initialdir=self.cfg.get("ultima_pasta") or None,
            filetypes=[("Documento do Word", "*.docx"), ("Word 97-2003", "*.doc"), ("PDF", "*.pdf")])
        if not destino:
            return
        self.cfg["ultima_pasta"] = str(Path(destino).parent)
        self.trabalhando = True
        self.bt.configure(state="disabled")
        self.prog.configure(maximum=1, value=0)
        abrir = bool(self.v_abrir.get())  # ler variaveis Tk apenas na thread principal

        def job():
            log = lambda m: self.fila.put(("log", m))  # noqa: E731
            log("> texto colado")
            try:
                g = cv.converter_texto(texto, destino, op, log=log)
                log(f"  OK {g}")
                if abrir:
                    abrir_arquivo(g)
                self.fila.put(("prog", 1))
                self.fila.put(("fim", (1, 1, [g])))
            except Exception as e:
                log(f"  X Erro: {e}")
                self.fila.put(("fim", (0, 1, [])))
        threading.Thread(target=job, daemon=True).start()

    def _processar_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(dado)
                elif tipo == "prog":
                    self.prog.configure(value=dado)
                elif tipo == "status":
                    self.lbl_ferr.configure(text=dado)
                elif tipo == "fim":
                    ok, total, gerados = dado
                    self._escrever(f"Concluido: {ok} de {total} convertido(s).\n")
                    if gerados:
                        self.ultima_saida = str(Path(gerados[-1]).parent)
                    self.trabalhando = False
                    self.bt.configure(state="normal")
        except queue.Empty:
            pass
        self.root.after(100, self._processar_fila)

    def _fechar(self):
        self._salvar_estado()
        self.root.destroy()


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # janela nitida em telas com escala
        except Exception:
            pass
    args = sys.argv[1:]
    if len(args) >= 2 and args[0] in ("--rapido", "--quick"):
        fmt = args[1].lower().lstrip(".")
        if fmt in ("docx", "doc", "pdf"):
            modo_rapido(fmt, args[2:])
            return
    global TkinterDnD
    root = None
    if TkinterDnD:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            TkinterDnD = None
    if root is None:
        root = tk.Tk()
    App(root, [a for a in args if not a.startswith("--")])
    root.mainloop()


if __name__ == "__main__":
    main()
