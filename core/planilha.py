"""Camada de acesso à planilha (a planilha É o banco de dados).

Regras:
- Lê tudo com openpyxl preservando fórmulas e estilos.
- Só ESCREVE em B4:E34 (batidas), K3/K4/K5 (parâmetros) e L35 (saldo anterior).
  Nunca toca em F/G/H e totais — o Excel recalcula.
- Backup automático antes de cada salvamento.
- Salvamento seguro com mensagem amigável se o Excel estiver com lock.
"""

from __future__ import annotations

import copy
import datetime
import glob
import os
import re
import shutil
import unicodedata
from dataclasses import dataclass, field

import openpyxl
from openpyxl.utils import get_column_letter

from .calculo import timedelta_to_min, min_to_timedelta


MODELO_SHEET = "MODELO"

MESES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
            "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
MESES_NUM = {"01": 1, "02": 2, "03": 3, "04": 4, "05": 5, "06": 6,
             "07": 7, "08": 8, "09": 9, "10": 10, "11": 11, "12": 12}

# Colunas de batidas: B, C, D, E nas linhas 4..34 (dias 1..31)
BAT_COLS = [2, 3, 4, 5]
FIRST_ROW = 4
LAST_ROW = 34


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper().strip()


@dataclass
class AbaInfo:
    sheet: str
    funcionario: str
    mes_num: int
    mes_label: str
    ordem: tuple = field(default_factory=tuple)


def parse_sheet_name(name: str) -> AbaInfo | None:
    """Interpreta 'ALBERTO - SETEMBRO' e 'ALBERTO_01-01 – 31-01'."""
    if not name or _norm(name) == "MODELO":
        return None
    # Padrão novo: "NOME - MES"
    m = re.match(r"^\s*(.+?)\s*[-–—]\s*([A-Za-zÇÃÕÉÊÍÓÚÂÔÀ ]+)\s*$", name)
    if m:
        func = m.group(1).strip()
        mes_txt = _norm(m.group(2))
        for i, mes in enumerate(MESES_PT, start=1):
            if _norm(mes) == mes_txt:
                return AbaInfo(name, _titulo_func(func), i, MESES_PT[i - 1].capitalize(), (2026, i))
        # mês desconhecido (ex: nome com hífen duplo) -> tenta extrair número
    # Padrão antigo: "NOME 01-03 – 31-03" (dia-mês – dia-mês).
    # O mês é o SEGUNDO número (group 2); group 1 é o dia (sempre 01).
    matches = list(re.finditer(r"(\d{2})\s*[-–—\ufffd]\s*(\d{2})", name))
    m2 = matches[-1] if matches else None
    if m2:
        func = re.split(r"[\d]", name, maxsplit=1)[0]
        func = re.sub(r"[_\-\s.]+$", "", func).strip(" _-.") or name
        mes_num = MESES_NUM.get(m2.group(2), 1)
        return AbaInfo(name, _titulo_func(func), mes_num,
                       f"{MESES_PT[mes_num-1].capitalize()}",
                       (2026, mes_num))
    # fallback: trata como funcionário sem mês
    return AbaInfo(name, _titulo_func(name), 99, name, (2026, 99))


def _titulo_func(nome: str) -> str:
    nome = re.sub(r"\s+", " ", (nome or "").strip(" _-."))
    nome = nome.replace("_", " ")
    return nome.strip().title() or "Sem nome"


class PlanilhaDB:
    """Wrapper fino sobre o workbook."""

    def __init__(self, path: str):
        self.path = os.path.abspath(path)
        self.wb: openpyxl.Workbook | None = None
        self.mtime: float = 0.0
        self.abas: list[AbaInfo] = []
        self.funcionarios: list[str] = []

    # -- ciclo de vida -----------------------------------------------------
    def abrir(self):
        self.wb = openpyxl.load_workbook(self.path, data_only=False)
        self.mtime = os.path.getmtime(self.path)
        self._indexar()
        return self

    def recarregar(self):
        self.fechar()
        return self.abrir()

    def fechar(self):
        try:
            if self.wb is not None:
                self.wb.close()
        except Exception:
            pass
        self.wb = None

    def _indexar(self):
        abas = []
        for ws in self.wb.worksheets:
            info = parse_sheet_name(ws.title)
            if info:
                abas.append(info)
        abas.sort(key=lambda a: (_norm(a.funcionario), a.ordem, a.sheet))
        self.abas = abas
        vistos, funcs = set(), []
        for a in abas:
            k = _norm(a.funcionario)
            if k not in vistos:
                vistos.add(k)
                funcs.append(a.funcionario)
        funcs.sort(key=_norm)
        self.funcionarios = funcs

    # -- consultas ----------------------------------------------------------
    def meses_de(self, funcionario: str) -> list[AbaInfo]:
        k = _norm(funcionario)
        return [a for a in self.abas if _norm(a.funcionario) == k]

    def mudou_no_disco(self) -> bool:
        try:
            return os.path.getmtime(self.path) != self.mtime
        except OSError:
            return False

    def _ws(self, sheet: str):
        if sheet not in self.wb.sheetnames:
            raise KeyError(f"Aba '{sheet}' não encontrada (talvez renomeada no Excel?).")
        return self.wb[sheet]

    def ler_dias(self, sheet: str) -> list[tuple]:
        """Retorna 31 tuplas (b,c,d,e) em minutos/None."""
        ws = self._ws(sheet)
        dias = []
        for r in range(FIRST_ROW, LAST_ROW + 1):
            vals = []
            for ci in BAT_COLS:
                vals.append(timedelta_to_min(ws.cell(r, ci).value))
            dias.append(tuple(vals))
        return dias

    def ler_parametros(self, sheet: str) -> dict:
        ws = self._ws(sheet)
        j = timedelta_to_min(ws["K3"].value) or 528  # 8:48
        t = timedelta_to_min(ws["K4"].value) or 0
        l = timedelta_to_min(ws["K5"].value) or 0
        s = timedelta_to_min(ws["L35"].value) or 0
        return {"jornada": j, "tolerancia": t, "lanche": l, "saldo_anterior": s}

    def ler_totais_cache(self, sheet: str) -> dict:
        import openpyxl as _oxl
        try:
            wb2 = _oxl.load_workbook(self.path, data_only=True)
            ws = wb2[sheet]
            out = {c: timedelta_to_min(ws[c].value) or 0
                   for c in ("G35", "H35", "G36", "L39", "N39", "O37")}
            wb2.close()
            return out
        except Exception:
            return {}

    @staticmethod
    def formulas_dia(r: int) -> tuple[str, str, str]:
        """Fórmulas padrão de F/G/H para a linha r (iguais às do MODELO)."""
        f = f"=IF((C{r}-B{r})+(E{r}-D{r})=0,0,((C{r}-B{r})+(E{r}-D{r}))-$K$5)"
        g = f"=IF(F{r}=0,0,IF(($K$3-F{r})>$K$4,ABS($K$3-F{r}),0))"
        h = f"=IF((F{r}-$K$3)>$K$4,ABS(F{r}-$K$3),0)"
        return f, g, h

    def ler_detalhe_dias(self, sheet: str) -> list[dict]:
        """Para cada dia: batidas + se G/H são fórmula ou ajuste manual.

        Retorna 31 dicts {b,c,d,e, g_formula, h_formula, g_manual, h_manual}.
        g_manual/h_manual = minutos do valor estático (quando a célula NÃO
        é fórmula) ou None (quando é fórmula — vale o cálculo).
        """
        ws = self._ws(sheet)
        det = []
        for r in range(FIRST_ROW, LAST_ROW + 1):
                b = timedelta_to_min(ws.cell(r, 2).value)
                c = timedelta_to_min(ws.cell(r, 3).value)
                d = timedelta_to_min(ws.cell(r, 4).value)
                e = timedelta_to_min(ws.cell(r, 5).value)
                gcell, hcell = ws.cell(r, 7), ws.cell(r, 8)
                g_form = (gcell.data_type == "f" or
                          (isinstance(gcell.value, str) and str(gcell.value).startswith("=")))
                h_form = (hcell.data_type == "f" or
                          (isinstance(hcell.value, str) and str(hcell.value).startswith("=")))
                g_man = h_man = None
                if not g_form:
                    # Estática: valor digitado (abono/zeramento) ou célula
                    # esvaziada (fórmula apagada = 0 no SUM do Excel).
                    # Nunca usa cache aqui: cache pode estar obsoleto.
                    g_man = timedelta_to_min(gcell.value)
                    if g_man is None:
                        g_man = 0
                if not h_form:
                    h_man = timedelta_to_min(hcell.value)
                    if h_man is None:
                        h_man = 0
                det.append({"b": b, "c": c, "d": d, "e": e,
                            "g_formula": g_form, "h_formula": h_form,
                            "g_manual": g_man, "h_manual": h_man})
        return det

    def restaurar_formulas_dia(self, sheet: str, dia: int):
        """Reverte G/H (e garante F) do dia para as fórmulas padrão."""
        ws = self._ws(sheet)
        r = FIRST_ROW + dia - 1
        f, g, h = self.formulas_dia(r)
        for col, form, fmt in ((6, f, "[h]:mm:ss;@"), (7, g, "[h]:mm:ss;@"),
                               (8, h, "[h]:mm:ss;@")):
            cell = ws.cell(r, col)
            cell.value = form
            cell.number_format = fmt

    # -- escrita (somente colunas permitidas) --------------------------------
    def escrever_dia(self, sheet: str, dia: int, b, c, d, e):
        ws = self._ws(sheet)
        r = FIRST_ROW + dia - 1
        for ci, minutos in zip(BAT_COLS, (b, c, d, e)):
            cell = ws.cell(r, ci)
            cell.value = min_to_timedelta(minutos)
            cell.number_format = "[h]:mm:ss;@"

    def escrever_parametros(self, sheet: str, jornada=None, tolerancia=None,
                            lanche=None, saldo_anterior=None):
        ws = self._ws(sheet)
        if jornada is not None:
            ws["K3"].value = min_to_timedelta(jornada)
            ws["K3"].number_format = "[h]:mm:ss;@"
        if tolerancia is not None:
            ws["K4"].value = min_to_timedelta(tolerancia)
            ws["K4"].number_format = "[h]:mm:ss;@"
        if lanche is not None:
            ws["K5"].value = min_to_timedelta(lanche)
            ws["K5"].number_format = "[h]:mm:ss;@"
        if saldo_anterior is not None:
            ws["L35"].value = min_to_timedelta(saldo_anterior)
            ws["L35"].number_format = "[h]:mm:ss;@"

    def criar_aba(self, funcionario: str, mes_nome: str, origem: str | None = None) -> str:
        """Cria 'FUNCIONARIO - MES' copiando a última aba do funcionário
        (herda tolerância/lanche) ou o MODELO. Zera batidas. Retorna o nome."""
        funcionario = re.sub(r"\s+", " ", funcionario.strip()).upper()
        mes_nome = mes_nome.strip().upper()
        novo = f"{funcionario} - {mes_nome}"
        if novo in self.wb.sheetnames:
            raise ValueError(f"A aba '{novo}' já existe.")
        src_nome = origem or self._origem_para(funcionario)
        src = self.wb[src_nome]
        nova = self.wb.copy_worksheet(src)
        nova.title = novo
        # zera batidas
        for r in range(FIRST_ROW, LAST_ROW + 1):
            for ci in BAT_COLS:
                nova.cell(r, ci).value = None
        # garante fórmulas limpas em F/G/H (a origem pode conter ajustes
        # manuais ✎ — o mês novo sempre começa zerado, sem herdar ajustes)
        for dia in range(1, 32):
            r = FIRST_ROW + dia - 1
            f, g, h = self.formulas_dia(r)
            for col, form in ((6, f), (7, g), (8, h)):
                cell = nova.cell(r, col)
                cell.value = form
                cell.number_format = "[h]:mm:ss;@"
        # saldo anterior começa zerado; a tela oferece puxar o saldo final
        # do mês anterior automaticamente
        try:
            nova["L35"].value = min_to_timedelta(0)
            nova["L35"].number_format = "[h]:mm:ss;@"
        except Exception:
            pass
        self._indexar()
        return novo

    def _origem_para(self, funcionario_up: str) -> str:
        meses = [a for a in self.abas if _norm(a.funcionario) == _norm(funcionario_up)]
        if meses:
            return meses[-1].sheet
        if MODELO_SHEET in self.wb.sheetnames:
            return MODELO_SHEET
        raise ValueError("Aba MODELO não encontrada e funcionário sem histórico.")

    # -- persistência ----------------------------------------------------------
    def backup(self, pasta: str | None = None) -> str:
        base = os.path.dirname(self.path)
        pasta_bk = pasta or os.path.join(base, "backups")
        os.makedirs(pasta_bk, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome = os.path.basename(self.path)
        dest = os.path.join(pasta_bk, f"{stamp}_{nome}")
        shutil.copy2(self.path, dest)
        # mantém só os 20 mais recentes
        try:
            arqs = sorted(glob.glob(os.path.join(pasta_bk, f"*_{nome}")))
            for velho in arqs[:-20]:
                os.remove(velho)
        except OSError:
            pass
        return dest

    def salvar(self) -> tuple[bool, str]:
        """Salva com backup prévio. Retorna (ok, mensagem)."""
        try:
            self.backup()
        except Exception:
            pass  # backup é melhor-esforço; não bloqueia o salvamento
        try:
            # save para temp + replace = escrita mais segura
            tmp = self.path + ".tmp"
            self.wb.save(tmp)
            os.replace(tmp, self.path)
            self.mtime = os.path.getmtime(self.path)
            return True, "Salvo na planilha com sucesso."
        except PermissionError:
            return False, ("Não consegui gravar porque a planilha está ABERTA no Excel "
                           "com bloqueio. Feche o Excel e clique em Salvar de novo — "
                           "nada foi perdido, suas edições continuam na tela.")
        except Exception as e:
            return False, f"Falha ao salvar: {e}"


def localizar_planilha(pasta: str) -> str | None:
    """Encontra o .xlsx na pasta do app (prefere Calc*). Ignora ~$ temporários."""
    cands = [p for p in glob.glob(os.path.join(pasta, "*.xlsx"))
             if not os.path.basename(p).startswith("~$")]
    if not cands:
        return None
    cands.sort(key=lambda p: (0 if os.path.basename(p).lower().startswith("calc") else 1, p))
    return cands[0]


def app_base_dir() -> str:
    import sys
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
