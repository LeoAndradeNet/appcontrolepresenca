"""Interface desktop — UX como prioridade.

Princípios:
- Tudo em no máximo 2 cliques. Nada escondido em menu.
- Feedback imediato: totais recalculam ao digitar, erros mostram onde.
- Teclado primeiro: Enter avança, Ctrl+S salva, Ctrl+F busca.
- Nunca perde dado: dirty-flag, backup, erro de lock explicado.
"""

from __future__ import annotations

import calendar
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.calculo import calc_mes, fmt_horas, parse_hora, validar_hora
from core.planilha import MESES_PT, PlanilhaDB, app_base_dir, localizar_planilha

ANO = 2026
SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]

# Paleta — contraste alto, sem cinza-sobre-cinza
C_BG = "#EEF1F5"
C_CARD = "#FFFFFF"
C_SIDE = "#1E293B"
C_SIDE_SEL = "#334155"
C_TXT = "#0F172A"
C_MUT = "#64748B"
C_ACC = "#2563EB"
C_ACC_D = "#1D4ED8"
C_OK = "#047857"
C_OK_BG = "#D1FAE5"
C_WARN_BG = "#FEF3C7"
C_ERR = "#B91C1C"
C_ERR_BG = "#FEE2E2"
C_LINE = "#E2E8F0"
C_FDS = "#F8FAFC"
C_FALTA = "#B91C1C"
C_EXTRA = "#1D4ED8"


class Toast(ttk.Frame):
    """Faixa de mensagem no rodapé (sucesso/erro/aviso)."""

    def __init__(self, master):
        super().__init__(master)
        self.lbl = ttk.Label(self, text="", font=("Segoe UI", 10), wraplength=900)
        self.lbl.pack(side="left", fill="x", expand=True, padx=10, pady=6)
        self._job = None

    def show(self, texto: str, kind: str = "ok"):
        cores = {"ok": (C_OK_BG, C_OK), "warn": (C_WARN_BG, "#92400E"),
                 "err": (C_ERR_BG, C_ERR), "info": ("#DBEAFE", C_ACC_D)}
        bg, fg = cores.get(kind, cores["info"])
        self.configure(style=f"Toast{kind}.TFrame")
        try:
            s = ttk.Style()
            s.configure(f"Toast{kind}.TFrame", background=bg)
        except Exception:
            pass
        self.lbl.configure(background=bg, foreground=fg, text=texto)
        self.place(relx=0.5, rely=1.0, anchor="s", relwidth=1.0)
        self.lift()
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(6000, self.hide)

    def hide(self):
        self.place_forget()


class HoraEntry(ttk.Entry):
    """Entry HH:MM tolerante: aceita 730, 7:30, 7h30. Normaliza ao sair."""

    def __init__(self, master, on_change=None, width=8, **kw):
        super().__init__(master, width=width, font=("Segoe UI", 11), **kw)
        self.on_change = on_change
        self._ok = True
        self.bind("<FocusOut>", self._norm)
        self.bind("<KeyRelease>", self._tick)
        self.bind("<Return>", self._next)
        self.bind("<KP_Enter>", self._next)

    def _tick(self, _e=None):
        if self.on_change:
            self.on_change()

    def _norm(self, _e=None):
        txt = self.get().strip()
        if txt == "":
            self._set_ok(True)
            if self.on_change:
                self.on_change()
            return
        ok, mins, _msg = validar_hora(txt)
        if ok and mins is not None:
            h, m = divmod(mins, 60)
            self.delete(0, "end")
            self.insert(0, f"{h:02d}:{m:02d}")
            self._set_ok(True)
        else:
            self._set_ok(False)
        if self.on_change:
            self.on_change()

    def _set_ok(self, ok: bool):
        self._ok = ok
        self.configure(style="TEntry" if ok else "Err.TEntry")

    def _next(self, _e=None):
        # Enter avança para o próximo campo (UX de digitação rápida)
        self.tk_focusNext().focus_set()
        return "break"

    @property
    def minutos(self):
        txt = self.get().strip()
        if txt == "":
            return None
        try:
            return parse_hora(txt)
        except ValueError:
            return "ERRO"

    def invalido(self) -> bool:
        return self.minutos == "ERRO"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Controle de Faltas & Horas Extras — 2026")
        self.geometry("1360x820")
        self.minsize(1100, 680)
        self.configure(bg=C_BG)
        try:
            self.state("zoomed")
        except Exception:
            pass

        self._style()
        self.base = app_base_dir()
        self.cfg_path = os.path.join(self.base, "config.json")
        self.db: PlanilhaDB | None = None
        self.arquivo: str | None = self._descobrir_arquivo()

        self.func_atual: str | None = None
        self.aba_atual: str | None = None
        self.dirty = False
        self._calc_job = None
        self._entradas: dict[tuple[int, int], HoraEntry] = {}
        self._calc_lbls: dict[int, tuple[ttk.Label, ttk.Label, ttk.Label]] = {}
        self._row_frames: dict[int, ttk.Frame] = {}
        self._restore_btns: dict[int, ttk.Button] = {}
        self._overrides: list[tuple] = [(None, None)] * 31
        self._dias_cache: list[tuple] = [(None, None, None, None)] * 31

        self._montar_layout()
        self._binds_globais()
        if self.arquivo:
            self._abrir_arquivo(self.arquivo)
        else:
            self._pedir_arquivo(inicial=True)
        self.after(4000, self._vigiar_disco)

    # -- setup ------------------------------------------------------------
    def _style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("TFrame", background=C_BG)
        s.configure("Card.TFrame", background=C_CARD)
        s.configure("TLabel", background=C_BG, foreground=C_TXT, font=("Segoe UI", 10))
        s.configure("Card.TLabel", background=C_CARD, foreground=C_TXT)
        s.configure("Mut.TLabel", background=C_BG, foreground=C_MUT, font=("Segoe UI", 9))
        s.configure("CardMut.TLabel", background=C_CARD, foreground=C_MUT, font=("Segoe UI", 9))
        s.configure("H1.TLabel", background=C_BG, foreground=C_TXT,
                    font=("Segoe UI", 13, "bold"))
        s.configure("H2.TLabel", background=C_CARD, foreground=C_TXT,
                    font=("Segoe UI", 11, "bold"))
        s.configure("Big.TLabel", background=C_CARD, font=("Segoe UI", 16, "bold"))
        s.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        s.configure("Primary.TButton", background=C_ACC, foreground="white",
                    font=("Segoe UI", 11, "bold"), padding=(18, 9))
        s.map("Primary.TButton", background=[("active", C_ACC_D), ("disabled", "#93C5FD")])
        s.configure("Ghost.TButton", background=C_CARD, padding=(10, 6))
        s.configure("Err.TEntry", fieldbackground=C_ERR_BG, foreground=C_ERR)
        s.configure("Side.TFrame", background=C_SIDE)
        s.configure("Side.TLabel", background=C_SIDE, foreground="#E2E8F0")
        s.configure("SideMut.TLabel", background=C_SIDE, foreground="#94A3B8",
                    font=("Segoe UI", 9))
        s.configure("Side.TButton", background=C_SIDE_SEL, foreground="white",
                    font=("Segoe UI", 10), padding=(10, 7))
        s.map("Side.TButton", background=[("active", "#475569")])
        s.configure("TNotebook", background=C_BG, borderwidth=0)
        s.configure("TNotebook.Tab", font=("Segoe UI", 11), padding=(18, 8))

    def _descobrir_arquivo(self) -> str | None:
        try:
            if os.path.exists(self.cfg_path):
                with open(self.cfg_path, encoding="utf-8") as f:
                    p = json.load(f).get("planilha")
                if p and os.path.exists(p):
                    return p
        except Exception:
            pass
        return localizar_planilha(self.base)

    def _salvar_cfg(self):
        try:
            with open(self.cfg_path, "w", encoding="utf-8") as f:
                json.dump({"planilha": self.arquivo}, f)
        except Exception:
            pass

    # -- layout ------------------------------------------------------------
    def _montar_layout(self):
        # Barra superior
        top = ttk.Frame(self, padding=(14, 10))
        top.pack(fill="x")
        ttk.Label(top, text="🕒 Controle de Faltas & Horas Extras",
                  style="H1.TLabel").pack(side="left")
        ttk.Label(top, text="  grava direto na planilha",
                  style="Mut.TLabel").pack(side="left", pady=(4, 0))
        self.btn_salvar = ttk.Button(top, text="💾 Salvar na planilha  (Ctrl+S)",
                                     style="Primary.TButton", command=self.salvar)
        self.btn_salvar.pack(side="right", padx=(8, 0))
        ttk.Button(top, text="↻ Recarregar", command=self.recarregar,
                   style="Ghost.TButton").pack(side="right", padx=(4, 0))
        ttk.Button(top, text="📂 Trocar planilha", command=self._pedir_arquivo,
                   style="Ghost.TButton").pack(side="right")

        # Corpo: lateral + principal
        corpo = ttk.Frame(self)
        corpo.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        corpo.columnconfigure(1, weight=1)
        corpo.rowconfigure(0, weight=1)

        # --- lateral ---
        side = ttk.Frame(corpo, style="Side.TFrame", padding=12, width=250)
        side.grid(row=0, column=0, sticky="nsw")
        side.grid_propagate(False)
        ttk.Label(side, text="Funcionários", style="Side.TLabel",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(side, text="digite para filtrar  (Ctrl+F)",
                  style="SideMut.TLabel").pack(anchor="w", pady=(0, 6))
        self.busca_var = tk.StringVar()
        self.busca_var.trace_add("write", lambda *_: self._filtrar_func())
        self.busca = ttk.Entry(side, textvariable=self.busca_var, font=("Segoe UI", 11))
        self.busca.pack(fill="x", pady=(0, 8))
        self.lista = tk.Listbox(side, font=("Segoe UI", 11), bg=C_SIDE, fg="white",
                                selectbackground=C_ACC, selectforeground="white",
                                highlightthickness=0, borderwidth=0, activestyle="none",
                                height=20)
        self.lista.pack(fill="both", expand=True)
        self.lista.bind("<<ListboxSelect>>", lambda _e: self._trocou_func())
        ttk.Button(side, text="＋ Novo funcionário", style="Side.TButton",
                   command=self._dlg_novo_func).pack(fill="x", pady=(10, 4))
        ttk.Button(side, text="＋ Novo mês", style="Side.TButton",
                   command=self._dlg_novo_mes).pack(fill="x")
        self.lbl_arq = ttk.Label(side, text="", style="SideMut.TLabel", wraplength=220)
        self.lbl_arq.pack(fill="x", pady=(10, 0))

        # --- principal ---
        main = ttk.Frame(corpo)
        main.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        # seletor de mês
        barra_mes = ttk.Frame(main, style="Card.TFrame", padding=10)
        barra_mes.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(barra_mes, text="◀", width=3,
                   command=lambda: self._mudar_mes(-1)).pack(side="left")
        self.combo_mes = ttk.Combobox(barra_mes, state="readonly",
                                      font=("Segoe UI", 12, "bold"), width=26)
        self.combo_mes.pack(side="left", padx=10)
        self.combo_mes.bind("<<ComboboxSelected>>", lambda _e: self._trocou_mes())
        ttk.Button(barra_mes, text="▶", width=3,
                   command=lambda: self._mudar_mes(1)).pack(side="left")
        self.lbl_hj = ttk.Label(barra_mes, text="", style="CardMut.TLabel")
        self.lbl_hj.pack(side="left", padx=10)
        ttk.Button(barra_mes, text="⚡ Preencher dias úteis vazios",
                   command=self._preenchimento_rapido).pack(side="right")
        ttk.Button(barra_mes, text="🧹 Zerar fins de semana",
                   command=self._limpar_fds).pack(side="right", padx=(0, 8))

        # notebook
        self.nb = ttk.Notebook(main)
        self.nb.grid(row=1, column=0, sticky="nsew")
        self.tab_lanc = ttk.Frame(self.nb, padding=10)
        self.tab_res = ttk.Frame(self.nb, padding=10)
        self.nb.add(self.tab_lanc, text="  📝 Lançamentos  ")
        self.nb.add(self.tab_res, text="  📊 Resumo geral  ")
        self.nb.bind("<<NotebookTabChanged>>", lambda _e: self._atualizar_resumo())

        self._montar_lancamentos(self.tab_lanc)
        self._montar_resumo(self.tab_res)

        # rodapé
        roda = ttk.Frame(self, padding=(14, 0, 14, 10))
        roda.pack(fill="x")
        self.lbl_status = ttk.Label(roda, text="Pronto.", style="Mut.TLabel")
        self.lbl_status.pack(side="left")
        self.lbl_dirty = ttk.Label(roda, text="", style="Mut.TLabel",
                                   font=("Segoe UI", 10, "bold"))
        self.lbl_dirty.pack(side="right")
        self.toast = Toast(self)

    def _montar_lancamentos(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=0)
        tab.rowconfigure(0, weight=1)

        # tabela rolável
        outer = ttk.Frame(tab, style="Card.TFrame")
        outer.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        head = ttk.Frame(outer, style="Card.TFrame", padding=(10, 10, 10, 4))
        head.grid(row=0, column=0, sticky="ew")
        # Larguras compartilhadas entre cabeçalho e linhas (alinhamento pixel a pixel)
        larg = [64, 108, 108, 108, 108, 66, 66, 66, 104]
        cols = ["Dia", "Manhã Entrada", "Manhã Saída",
                "Tarde Entrada", "Tarde Saída",
                "Trab.", "Falta", "Extra", ""]
        for i, (t, w) in enumerate(zip(cols, larg)):
            head.columnconfigure(i, minsize=w)
            ttk.Label(head, text=t, style="CardMut.TLabel",
                      font=("Segoe UI", 9, "bold"),
                      anchor="center" if i else "w").grid(row=0, column=i, padx=3, sticky="ew")
        self._larg_cols = larg

        cv = tk.Canvas(outer, bg=C_CARD, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        cv.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")
        self._corpo_tab = ttk.Frame(cv, style="Card.TFrame")
        self._win = cv.create_window((0, 0), window=self._corpo_tab, anchor="nw")
        self._corpo_tab.bind("<Configure>",
                             lambda _e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(self._win, width=e.width))
        # roda do mouse
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(-1 * (e.delta // 120), "units"),
                    add="+")

        for dia in range(1, 32):
            fr = ttk.Frame(self._corpo_tab, style="Card.TFrame", padding=(10, 2))
            fr.pack(fill="x")
            for i, w in enumerate(self._larg_cols):
                fr.columnconfigure(i, minsize=w)
            self._row_frames[dia] = fr
            ttk.Label(fr, text=f"{dia:02d}", font=("Segoe UI", 11, "bold"),
                      background=C_CARD, width=3).grid(row=0, column=0, padx=3, sticky="w")
            self.lbl_sem = getattr(self, "lbl_sem", {})
            sem = ttk.Label(fr, text="", font=("Segoe UI", 9), background=C_CARD,
                            foreground=C_MUT, width=4)
            sem.grid(row=0, column=0, padx=(34, 0), sticky="w")
            self.lbl_sem[dia] = sem
            for ci in range(4):
                ent = HoraEntry(fr, on_change=self._agendar_calc, width=8,
                                justify="center")
                ent.grid(row=0, column=1 + ci, padx=3, pady=2, sticky="ew")
                ent.bind("<FocusIn>", lambda e: e.widget.select_range(0, "end"))
                self._entradas[(dia, ci)] = ent
            f_lbl = ttk.Label(fr, text="—", background=C_CARD, font=("Segoe UI", 10),
                              anchor="center")
            g_lbl = ttk.Label(fr, text="—", background=C_CARD, font=("Segoe UI", 10, "bold"),
                              foreground=C_FALTA, anchor="center")
            h_lbl = ttk.Label(fr, text="—", background=C_CARD, font=("Segoe UI", 10, "bold"),
                              foreground=C_EXTRA, anchor="center")
            f_lbl.grid(row=0, column=5, padx=3, sticky="ew")
            g_lbl.grid(row=0, column=6, padx=3, sticky="ew")
            h_lbl.grid(row=0, column=7, padx=3, sticky="ew")
            self._calc_lbls[dia] = (f_lbl, g_lbl, h_lbl)
            row_btns = ttk.Frame(fr, style="Card.TFrame")
            row_btns.grid(row=0, column=8, padx=2)
            ttk.Button(row_btns, text="↺", width=2,
                       command=lambda d=dia: self._repetir_ontem(d),
                       ).pack(side="left")
            ttk.Button(row_btns, text="✕", width=2,
                       command=lambda d=dia: self._limpar_dia(d)).pack(side="left")
            rb = ttk.Button(row_btns, text="ƒx", width=3,
                            command=lambda d=dia: self._restaurar_dia(d))
            rb.pack(side="left")
            self._restore_btns[dia] = rb

        # painel lateral direito: parâmetros + totais
        lat = ttk.Frame(tab, style="Card.TFrame", padding=14, width=300)
        lat.grid(row=0, column=1, sticky="ns")
        lat.grid_propagate(False)
        ttk.Label(lat, text="⚙ Parâmetros do mês", style="H2.TLabel").pack(anchor="w")
        ttk.Label(lat, text="valem para todos os dias desta aba",
                  style="CardMut.TLabel").pack(anchor="w", pady=(0, 8))
        self.ent_jor = self._campo_param(lat, "Jornada esperada (ex 08:48)", "08:48")
        self.ent_tol = self._campo_param(lat, "Tolerância (ex 00:10)", "00:10")
        self.ent_lan = self._campo_param(lat, "Desconto lanche (ex 00:10)", "00:10")
        self.ent_sal = self._campo_param(lat, "Saldo do mês passado (ex 02:16)", "")
        ttk.Separator(lat).pack(fill="x", pady=10)
        ttk.Label(lat, text="📊 Totais calculados", style="H2.TLabel").pack(anchor="w")
        self.cards = {}
        for chave, titulo in [("falta", "🔴 Faltas no mês"), ("extra", "🔵 Extras no mês"),
                              ("credito", "💰 Crédito (extras + anterior)"),
                              ("debito", "🧾 Débito (faltas)"),
                              ("saldo", "⚖️ Saldo final")]:
            box = ttk.Frame(lat, style="Card.TFrame", padding=4)
            box.pack(fill="x", pady=2)
            ttk.Label(box, text=titulo, style="CardMut.TLabel").pack(anchor="w")
            v = ttk.Label(box, text="—", style="Big.TLabel")
            v.pack(anchor="w")
            self.cards[chave] = v
        ttk.Label(lat, text="↺ repete o dia anterior • ✕ limpa o dia • ƒx reverte ajuste manual ✎.\n"
                            "✎ = valor ajustado à mão no Excel (abono/atestado), preservado pelo app.\n"
                            "Fórmulas e estilos da planilha são preservados.\nSomente batidas e parâmetros são gravados.",
                  style="CardMut.TLabel", wraplength=260).pack(pady=(10, 0))

    def _campo_param(self, master, rotulo, placeholder):
        ttk.Label(master, text=rotulo, style="CardMut.TLabel").pack(anchor="w", pady=(6, 0))
        ent = HoraEntry(master, on_change=self._agendar_calc, width=12)
        ent.pack(fill="x", pady=2)
        if placeholder:
            ent.insert(0, placeholder)
        return ent

    def _montar_resumo(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        fr = ttk.Frame(tab)
        fr.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(fr, text="Mês:").pack(side="left")
        self.combo_res = ttk.Combobox(fr, state="readonly", width=24,
                                      font=("Segoe UI", 11))
        self.combo_res.pack(side="left", padx=8)
        self.combo_res.bind("<<ComboboxSelected>>", lambda _e: self._atualizar_resumo())
        ttk.Button(fr, text="Atualizar", command=self._atualizar_resumo).pack(side="left")
        cols = ("func", "mes", "falta", "extra", "saldo")
        self.tree = ttk.Treeview(tab, columns=cols, show="headings", height=18)
        for c, t, w in [("func", "Funcionário", 200), ("mes", "Mês", 150),
                        ("falta", "Faltas", 110), ("extra", "Extras", 110),
                        ("saldo", "Saldo", 110)]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "func" else "w")
        self.tree.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=1, sticky="ns")
        self.tree.bind("<Double-1>", self._ir_para_linha)

    # -- arquivo ------------------------------------------------------------
    def _pedir_arquivo(self, inicial=False):
        path = filedialog.askopenfilename(
            title="Escolha a planilha (.xlsx)",
            filetypes=[("Planilha Excel", "*.xlsx")],
            initialdir=self.base)
        if path:
            self._abrir_arquivo(path)
        elif inicial and not self.db:
            self.lbl_status.configure(text="Nenhuma planilha aberta. Clique em 📂 Trocar planilha.")
            self.toast.show("Escolha a planilha .xlsx para começar.", "info")

    def _abrir_arquivo(self, path: str):
        try:
            if self.db:
                self.db.fechar()
            self.db = PlanilhaDB(path).abrir()
            self.arquivo = os.path.abspath(path)
            self._salvar_cfg()
            self.lbl_arq.configure(text=f"📄 {os.path.basename(path)}")
            self._carregar_funcionarios()
            self.toast.show(f"Planilha aberta: {os.path.basename(path)} — {len(self.db.funcionarios)} funcionários, {len(self.db.abas)} meses.", "ok")
            self.lbl_status.configure(text=f"Aberto: {path}")
        except Exception as e:
            messagebox.showerror("Não consegui abrir", f"Verifique se o arquivo é um .xlsx válido.\n\n{e}")
            self.toast.show(f"Erro ao abrir: {e}", "err")

    def recarregar(self):
        if not self.db:
            return
        try:
            func, aba = self.func_atual, self.aba_atual
            self.db.recarregar()
            self._carregar_funcionarios(manter=(func, aba))
            self._carregar_aba()
            self._marcar_limpo()
            self.toast.show("Recarregado da planilha.", "ok")
        except Exception as e:
            self.toast.show(f"Falha ao recarregar: {e}", "err")

    # -- funcionários / meses -------------------------------------------------
    def _carregar_funcionarios(self, manter=None):
        self.lista.delete(0, "end")
        self._funcs_visiveis = list(self.db.funcionarios)
        for f in self._funcs_visiveis:
            n = len(self.db.meses_de(f))
            self.lista.insert("end", f"  {f}   ({n})")
        alvo = (manter[0] if manter else None) or self.func_atual or (self._funcs_visiveis[0] if self._funcs_visiveis else None)
        if alvo and alvo in self._funcs_visiveis:
            self.lista.selection_clear(0, "end")
            self.lista.selection_set(self._funcs_visiveis.index(alvo))
            self.func_atual = alvo
        elif self._funcs_visiveis:
            self.lista.selection_set(0)
            self.func_atual = self._funcs_visiveis[0]
        self._carregar_meses_combo(manter_aba=manter[1] if manter else None)
        if self.aba_atual:
            self._carregar_aba()
        self._atualizar_resumo()

    def _filtrar_func(self):
        q = self.busca_var.get().strip().lower()
        self.lista.delete(0, "end")
        self._funcs_visiveis = [f for f in self.db.funcionarios if q in f.lower()] if self.db else []
        for f in self._funcs_visiveis:
            self.lista.insert("end", f"  {f}   ({len(self.db.meses_de(f))})")
        if self._funcs_visiveis:
            self.lista.selection_set(0)
            self._trocou_func()

    def _trocou_func(self):
        sel = self.lista.curselection()
        if not sel or not self.db:
            return
        novo = self._funcs_visiveis[sel[0]]
        if novo == self.func_atual and self.aba_atual:
            return
        if self.dirty and not self._confirmar_perda():
            # volta seleção
            try:
                self.lista.selection_clear(0, "end")
                self.lista.selection_set(self._funcs_visiveis.index(self.func_atual))
            except Exception:
                pass
            return
        self.func_atual = novo
        self._carregar_meses_combo()
        self._carregar_aba()

    def _carregar_meses_combo(self, manter_aba=None):
        meses = self.db.meses_de(self.func_atual) if self.db and self.func_atual else []
        self._meses = meses
        self.combo_mes["values"] = [
            f"{a.mes_label}  —  {a.sheet.replace(chr(0xFFFD), '–')}" for a in meses]
        if not meses:
            self.aba_atual = None
            return
        alvo = manter_aba or self.aba_atual
        idx = next((i for i, a in enumerate(meses) if a.sheet == alvo), len(meses) - 1)
        self.combo_mes.current(idx)
        self.aba_atual = meses[idx].sheet

    def _trocou_mes(self):
        if self.dirty and not self._confirmar_perda():
            self._carregar_meses_combo(manter_aba=self.aba_atual)
            return
        idx = self.combo_mes.current()
        if 0 <= idx < len(self._meses):
            self.aba_atual = self._meses[idx].sheet
            self._carregar_aba()

    def _mudar_mes(self, delta: int):
        if not getattr(self, "_meses", None):
            return
        idx = self.combo_mes.current()
        novo = max(0, min(len(self._meses) - 1, idx + delta))
        if novo != idx:
            if self.dirty and not self._confirmar_perda():
                return
            self.combo_mes.current(novo)
            self.aba_atual = self._meses[novo].sheet
            self._carregar_aba()

    def _confirmar_perda(self) -> bool:
        return messagebox.askyesno("Alterações não salvas",
                                   "Você tem alterações NÃO salvas.\n\nTrocar de tela descarta o que foi digitado.\nDeseja descartar?")

    # -- carregar aba na tela --------------------------------------------------
    def _carregar_aba(self):
        if not self.db or not self.aba_atual:
            return
        try:
            dias = self.db.ler_dias(self.aba_atual)
            par = self.db.ler_parametros(self.aba_atual)
            try:
                detalhe = self.db.ler_detalhe_dias(self.aba_atual)
                self._overrides = [(d["g_manual"], d["h_manual"]) for d in detalhe]
            except Exception:
                self._overrides = [(None, None)] * 31
        except KeyError as e:
            self.toast.show(str(e), "err")
            return
        n_man = sum(1 for g, h in self._overrides if g is not None or h is not None)
        self._dias_cache = dias
        for dia in range(1, 32):
            b, c, d, e = dias[dia - 1]
            for ci, v in enumerate((b, c, d, e)):
                ent = self._entradas[(dia, ci)]
                ent.delete(0, "end")
                if v is not None:
                    ent.insert(0, f"{v // 60:02d}:{v % 60:02d}")
                ent._set_ok(True)
        for ent, v in [(self.ent_jor, par["jornada"]), (self.ent_tol, par["tolerancia"]),
                       (self.ent_lan, par["lanche"])]:
            ent.delete(0, "end")
            ent.insert(0, f"{v // 60:02d}:{v % 60:02d}")
            ent._set_ok(True)
        self.ent_sal.delete(0, "end")
        if par["saldo_anterior"]:
            self.ent_sal.insert(0, fmt_horas(par["saldo_anterior"]))
        self.ent_sal._set_ok(True)
        # dias da semana + destaque fds
        info = next((a for a in self.db.abas if a.sheet == self.aba_atual), None)
        mes_num = info.mes_num if info and 1 <= info.mes_num <= 12 else None
        ndias = calendar.monthrange(ANO, mes_num)[1] if mes_num else 31
        for dia in range(1, 32):
            sem_lbl = self.lbl_sem[dia]
            if mes_num:
                try:
                    wd = calendar.weekday(ANO, mes_num, min(dia, ndias))
                    sem_lbl.configure(text=SEMANA[wd] if dia <= ndias else "·")
                    fds = wd >= 5
                except ValueError:
                    sem_lbl.configure(text=""); fds = False
            else:
                sem_lbl.configure(text=""); fds = False
            invalido = dia > ndias
            for ci in range(4):
                self._entradas[(dia, ci)].configure(state="disabled" if invalido else "normal")
            if invalido:
                sem_lbl.configure(text="·")
        self.lbl_hj.configure(text=f"• {ndias} dias no mês" if mes_num else "")
        self._marcar_limpo()
        self._recalcular()
        if n_man:
            self.toast.show(
                f"⚠ {n_man} dia(s) com ajuste manual feito no Excel (selo ✎) — "
                "preservados e somados nos totais. Passe o mouse / use ƒx para reverter.",
                "warn")
        # rola para o topo
        try:
            self._corpo_tab.master.yview_moveto(0)
        except Exception:
            pass

    # -- edição ------------------------------------------------------------------
    def _agendar_calc(self):
        if self._calc_job:
            try:
                self.after_cancel(self._calc_job)
            except Exception:
                pass
        self._calc_job = self.after(120, self._recalcular)
        self._marcar_sujo()

    def _ler_estado_tela(self):
        """Lê UI -> (dias, params, erros). Erros = lista de 'Dia 5 Manhã Entrada'."""
        nomes = ["Manhã Entrada", "Manhã Saída", "Tarde Entrada", "Tarde Saída"]
        dias, erros = [], []
        for dia in range(1, 32):
            vals = []
            for ci in range(4):
                ent = self._entradas[(dia, ci)]
                if str(ent.cget("state")) == "disabled":
                    vals.append(None)
                    continue
                m = ent.minutos
                if m == "ERRO":
                    erros.append(f"Dia {dia} • {nomes[ci]} = '{ent.get().strip()}'")
                    vals.append(None)
                else:
                    vals.append(m)
            dias.append(tuple(vals))
        params, perr = {}, []
        for ent, chave, rot in [(self.ent_jor, "jornada", "Jornada"),
                                (self.ent_tol, "tolerancia", "Tolerância"),
                                (self.ent_lan, "lanche", "Lanche"),
                                (self.ent_sal, "saldo_anterior", "Saldo anterior")]:
            m = ent.minutos if ent.get().strip() != "" else 0
            if m == "ERRO":
                perr.append(f"{rot} = '{ent.get().strip()}'")
                params[chave] = 0
            else:
                params[chave] = m or 0
        return dias, params, erros + perr

    def _recalcular(self):
        dias, params, erros = self._ler_estado_tela()
        # batidas editadas invalidam o ajuste manual? NÃO: o ajuste manual
        # pertence à célula G/H e é preservado até o usuário revertê-lo
        # com ƒx. Apenas mostramos o selo ✎.
        res = calc_mes(dias, params["jornada"], params["tolerancia"],
                       params["lanche"], params["saldo_anterior"],
                       overrides=self._overrides)
        manuais = res.get("manuais", [False] * 31)
        for dia in range(1, 32):
            f, g, h = res["linhas"][dia - 1]
            fl, gl, hl = self._calc_lbls[dia]
            fl.configure(text=fmt_horas(f, vazio="—") if f else "—")
            if manuais[dia - 1]:
                gm, hm = self._overrides[dia - 1]
                gl.configure(text=(fmt_horas(g) if g else "0:00") + " ✎")
                hl.configure(text=(fmt_horas(h) if h else "0:00") + (" ✎" if hm is not None else ""))
                try:
                    self._restore_btns[dia].configure(state="normal")
                except Exception:
                    pass
            else:
                gl.configure(text=fmt_horas(g) if g else "—")
                hl.configure(text=fmt_horas(h) if h else "—")
                try:
                    self._restore_btns[dia].configure(state="disabled")
                except Exception:
                    pass
        self.cards["falta"].configure(text=fmt_horas(res["tot_falta"], vazio="0:00"),
                                      foreground=C_FALTA)
        self.cards["extra"].configure(text=fmt_horas(res["tot_extra"], vazio="0:00"),
                                      foreground=C_EXTRA)
        self.cards["credito"].configure(text=fmt_horas(res["credito"], vazio="0:00"))
        self.cards["debito"].configure(text=fmt_horas(res["debito"], vazio="0:00"))
        self.cards["saldo"].configure(text=fmt_horas(res["saldo"], vazio="0:00"))
        if erros:
            self.lbl_status.configure(text=f"⚠ {len(erros)} campo(s) com formato inválido — use HH:MM (ex 07:30). Primeiro: {erros[0]}")
        elif self.dirty:
            self.lbl_status.configure(text="● Alterações não salvas — Ctrl+S para gravar na planilha.")
        else:
            self.lbl_status.configure(text=f"✔ Em dia com a planilha • {self.aba_atual or ''}")

    def _marcar_sujo(self):
        if not self.dirty:
            self.dirty = True
            self.lbl_dirty.configure(text="● NÃO SALVO", foreground=C_ERR)
            self.btn_salvar.state(["!disabled"])

    def _marcar_limpo(self):
        self.dirty = False
        self.lbl_dirty.configure(text="✔ salvo", foreground=C_OK)

    # -- ações de linha ------------------------------------------------------------
    def _limpar_dia(self, dia: int):
        for ci in range(4):
            ent = self._entradas[(dia, ci)]
            if str(ent.cget("state")) != "disabled":
                ent.delete(0, "end")
        self._agendar_calc()
        self._entradas[(dia, 0)].focus_set()

    def _repetir_ontem(self, dia: int):
        if dia <= 1:
            return
        for ci in range(4):
            src = self._entradas[(dia - 1, ci)].get()
            dst = self._entradas[(dia, ci)]
            if str(dst.cget("state")) != "disabled":
                dst.delete(0, "end")
                dst.insert(0, src)
                dst._set_ok(True)
        self._agendar_calc()

    def _restaurar_dia(self, dia: int):
        """Reverte o ajuste manual do dia para a fórmula padrão."""
        gm, hm = self._overrides[dia - 1] if dia - 1 < len(self._overrides) else (None, None)
        if gm is None and hm is None:
            return
        if not messagebox.askyesno("Reverter ajuste manual",
                                    f"Dia {dia}: voltar a calcular falta/extra pela fórmula?\n\n"
                                    "O valor digitado manualmente no Excel será substituído."):
            return
        try:
            self.db.restaurar_formulas_dia(self.aba_atual, dia)
        except Exception as e:
            self.toast.show(f"Não consegui reverter: {e}", "err")
            return
        self._overrides[dia - 1] = (None, None)
        self._marcar_sujo()
        self._recalcular()
        self.toast.show(f"Dia {dia}: fórmula restaurada. Clique em Salvar para gravar.", "ok")

    def _preenchimento_rapido(self):
        """Preenche dias úteis vazios com jornada padrão. Confirmado antes."""
        if not messagebox.askyesno("Preenchimento rápido",
                                    "Preencher todos os dias ÚTEIS vazios com\n07:30 – 12:00 / 13:00 – 17:48?\n\nDias já preenchidos NÃO são alterados."):
            return
        n = 0
        for dia in range(1, 32):
            ents = [self._entradas[(dia, ci)] for ci in range(4)]
            if any(str(e.cget("state")) == "disabled" for e in ents):
                continue
            lbl = self.lbl_sem[dia].cget("text")
            if lbl in ("sáb", "dom", "·"):
                continue
            if any(e.get().strip() for e in ents):
                continue
            for e, v in zip(ents, ["07:30", "12:00", "13:00", "17:48"]):
                e.delete(0, "end")
                e.insert(0, v)
                e._set_ok(True)
            n += 1
        self._agendar_calc()
        self.toast.show(f"{n} dia(s) preenchidos. Confira e clique em Salvar.", "ok")

    def _limpar_fds(self):
        n = 0
        for dia in range(1, 32):
            if self.lbl_sem[dia].cget("text") in ("sáb", "dom"):
                if any(self._entradas[(dia, ci)].get().strip() for ci in range(4)):
                    self._limpar_dia(dia)
                    n += 1
        self._recalcular()
        self.toast.show(f"{n} dia(s) de fim de semana limpos." if n else "Nenhum fim de semana preenchido.", "info")

    # -- salvar ----------------------------------------------------------------------
    def salvar(self, _e=None):
        if not self.db or not self.aba_atual:
            self.toast.show("Nenhuma aba aberta para salvar.", "warn")
            return
        dias, params, erros = self._ler_estado_tela()
        if erros:
            primeiro = erros[0]
            if not messagebox.askyesno("Campos inválidos",
                                       f"Há {len(erros)} campo(s) com horário inválido.\nPrimeiro: {primeiro}\n\nSalvar mesmo assim (campos inválidos serão ignorados)?"):
                return
        try:
            for dia in range(1, 32):
                b, c, d, e = dias[dia - 1]
                self.db.escrever_dia(self.aba_atual, dia, b, c, d, e)
            self.db.escrever_parametros(self.aba_atual, params["jornada"],
                                        params["tolerancia"], params["lanche"],
                                        params["saldo_anterior"])
            ok, msg = self.db.salvar()
        except Exception as e:
            self.toast.show(f"Falha ao salvar: {e}", "err")
            return
        if ok:
            self._marcar_limpo()
            self._recalcular()
            self.toast.show("💾 " + msg + " Backup criado em /backups.", "ok")
            self._atualizar_resumo()
        else:
            messagebox.showwarning("Feche o Excel e tente de novo", msg)
            self.toast.show(msg, "warn")

    # -- resumo -------------------------------------------------------------------------
    def _atualizar_resumo(self, *_a):
        try:
            if self.nb.select() != str(self.tab_res):
                return
        except Exception:
            pass
        if not self.db:
            return
        meses_labels = sorted({a.mes_label for a in self.db.abas})
        self.combo_res["values"] = ["Todos os meses"] + meses_labels
        if not self.combo_res.get():
            # pré-seleciona o mês da aba atual
            info = next((a for a in self.db.abas if a.sheet == self.aba_atual), None)
            self.combo_res.set(info.mes_label if info else "Todos os meses")
        filtro = self.combo_res.get()
        for i in self.tree.get_children():
            self.tree.delete(i)
        for a in self.db.abas:
            if filtro != "Todos os meses" and a.mes_label != filtro:
                continue
            try:
                dias = self.db.ler_dias(a.sheet)
                par = self.db.ler_parametros(a.sheet)
                try:
                    det = self.db.ler_detalhe_dias(a.sheet)
                    ov = [(d["g_manual"], d["h_manual"]) for d in det]
                except Exception:
                    ov = None
                r = calc_mes(dias, par["jornada"], par["tolerancia"],
                             par["lanche"], par["saldo_anterior"], overrides=ov)
                self.tree.insert("", "end", values=(a.funcionario, a.mes_label,
                                                    fmt_horas(r["tot_falta"], vazio="0:00"),
                                                    fmt_horas(r["tot_extra"], vazio="0:00"),
                                                    fmt_horas(r["saldo"], vazio="0:00")),
                                 tags=(a.sheet,))
            except Exception:
                continue

    def _ir_para_linha(self, _e):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals:
            return
        func, mes = vals[0], vals[1]
        if func in (self.db.funcionarios if self.db else []):
            if self.dirty and not self._confirmar_perda():
                return
            self.func_atual = func
            self.lista.selection_clear(0, "end")
            try:
                self.lista.selection_set(self._funcs_visiveis.index(func))
            except ValueError:
                pass
            self._carregar_meses_combo()
            for i, a in enumerate(self._meses):
                if a.mes_label == mes:
                    self.combo_mes.current(i)
                    self.aba_atual = a.sheet
                    break
            self._carregar_aba()
            self.nb.select(self.tab_lanc)

    # -- criação ----------------------------------------------------------------------------
    def _dlg_novo_mes(self):
        if not self.db:
            return
        win = tk.Toplevel(self)
        win.title("Novo mês")
        win.geometry("420x330")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Criar novo mês", font=("Segoe UI", 13, "bold")).pack(pady=(14, 4))
        ttk.Label(win, text="Copia a última aba do funcionário (mantém tolerância e lanche) e zera as batidas.",
                  wraplength=360, style="Mut.TLabel").pack(pady=(0, 10))
        ttk.Label(win, text="Funcionário").pack(anchor="w", padx=20)
        cb_f = ttk.Combobox(win, values=self.db.funcionarios, state="readonly", width=32)
        cb_f.pack(padx=20, pady=4)
        cb_f.set(self.func_atual or "")
        ttk.Label(win, text="Mês").pack(anchor="w", padx=20)
        cb_m = ttk.Combobox(win, values=MESES_PT, state="readonly", width=32)
        cb_m.pack(padx=20, pady=4)
        var_saldo = tk.BooleanVar(value=True)
        ttk.Checkbutton(win, text="Puxar saldo anterior automaticamente (recomendado)",
                        variable=var_saldo).pack(padx=20, pady=8, anchor="w")

        def ok():
            f, m = cb_f.get().strip(), cb_m.get().strip()
            if not f or not m:
                messagebox.showwarning("Faltou dado", "Escolha funcionário e mês.")
                return
            try:
                novo = self.db.criar_aba(f, m)
                if var_saldo.get():
                    # saldo anterior = O37 da última aba (calculado em Python)
                    ant = [a for a in self.db.meses_de(f) if a.sheet != novo]
                    if ant:
                        d = self.db.ler_dias(ant[-1].sheet)
                        p = self.db.ler_parametros(ant[-1].sheet)
                        try:
                            det0 = self.db.ler_detalhe_dias(ant[-1].sheet)
                            ov0 = [(x["g_manual"], x["h_manual"]) for x in det0]
                        except Exception:
                            ov0 = None
                        r = calc_mes(d, p["jornada"], p["tolerancia"], p["lanche"], p["saldo_anterior"], overrides=ov0)
                        self.db.escrever_parametros(novo, saldo_anterior=r["saldo"])
                ok2, msg = self.db.salvar()
                if not ok2:
                    messagebox.showwarning("Atenção", msg)
                    return
                win.destroy()
                self._carregar_funcionarios(manter=(f.title(), novo))
                self._carregar_aba()
                self.toast.show(f"Aba '{novo}' criada e salva.", "ok")
            except ValueError as e:
                messagebox.showwarning("Já existe", str(e))
            except Exception as e:
                messagebox.showerror("Erro", str(e))
        ttk.Button(win, text="Criar e salvar", style="Primary.TButton", command=ok).pack(pady=12)
        win.bind("<Escape>", lambda _e: win.destroy())

    def _dlg_novo_func(self):
        if not self.db:
            return
        win = tk.Toplevel(self)
        win.title("Novo funcionário")
        win.geometry("420x280")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Novo funcionário", font=("Segoe UI", 13, "bold")).pack(pady=(14, 4))
        ttk.Label(win, text="Cria a primeira aba a partir do MODELO, com jornada 08:48.",
                  wraplength=360, style="Mut.TLabel").pack(pady=(0, 10))
        ttk.Label(win, text="Nome").pack(anchor="w", padx=20)
        ent_n = ttk.Entry(win, font=("Segoe UI", 12), width=30)
        ent_n.pack(padx=20, pady=4)
        ent_n.focus_set()
        ttk.Label(win, text="Primeiro mês").pack(anchor="w", padx=20)
        cb_m = ttk.Combobox(win, values=MESES_PT, state="readonly", width=28)
        cb_m.pack(padx=20, pady=4)

        def ok():
            f, m = ent_n.get().strip(), cb_m.get().strip()
            if not f or not m:
                messagebox.showwarning("Faltou dado", "Digite o nome e escolha o mês.")
                return
            try:
                novo = self.db.criar_aba(f, m, origem="MODELO")
                ok2, msg = self.db.salvar()
                if not ok2:
                    messagebox.showwarning("Atenção", msg)
                    return
                win.destroy()
                self._carregar_funcionarios(manter=(f.title(), novo))
                self._carregar_aba()
                self.toast.show(f"'{f.title()}' criado com a aba '{novo}'.", "ok")
            except ValueError as e:
                messagebox.showwarning("Já existe", str(e))
            except Exception as e:
                messagebox.showerror("Erro", str(e))
        ttk.Button(win, text="Criar e salvar", style="Primary.TButton", command=ok).pack(pady=12)
        win.bind("<Escape>", lambda _e: win.destroy())
        win.bind("<Return>", lambda _e: ok())

    # -- sistema ----------------------------------------------------------------------------------
    def _binds_globais(self):
        self.bind("<Control-s>", self.salvar)
        self.bind("<Control-S>", self.salvar)
        self.bind("<Control-r>", lambda _e: self.recarregar())
        self.bind("<Control-f>", lambda _e: self.busca.focus_set())
        self.bind("<Control-n>", lambda _e: self._dlg_novo_mes())
        self.protocol("WM_DELETE_WINDOW", self._sair)

    def _sair(self):
        if self.dirty:
            r = messagebox.askyesnocancel("Sair sem salvar?",
                                          "Há alterações NÃO salvas.\n\nSim = salvar e sair\nNão = sair sem salvar\nCancelar = voltar")
            if r is None:
                return
            if r:
                self.salvar()
                if self.dirty:
                    return  # salvamento falhou (lock) — não sai para não perder
        try:
            if self.db:
                self.db.fechar()
        finally:
            self.destroy()

    def _vigiar_disco(self):
        try:
            if self.db and not self.dirty and self.db.mudou_no_disco():
                if messagebox.askyesno("Planilha mudou",
                                       "A planilha foi alterada fora do app (provavelmente no Excel).\n\nRecarregar agora?"):
                    self.recarregar()
                else:
                    self.db.mtime = os.path.getmtime(self.db.path)
        except Exception:
            pass
        self.after(4000, self._vigiar_disco)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
