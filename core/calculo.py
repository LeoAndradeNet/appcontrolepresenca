"""Utilitários de tempo + espelho das fórmulas da planilha.

Replica exatamente a lógica das colunas F/G/H e dos totais para que o
preview na tela seja idêntico ao que o Excel calcula após salvar.
"""

from __future__ import annotations

import re
from datetime import timedelta


# --------------------------------------------------------------------------
# Conversão tempo <-> minutos
# --------------------------------------------------------------------------

def timedelta_to_min(val) -> int | None:
    """Converte valor lido do Excel (timedelta/time/float/str) em minutos."""
    if val is None or val == "":
        return None
    if isinstance(val, timedelta):
        return int(round(val.total_seconds() / 60))
    if hasattr(val, "hour"):  # datetime.time / datetime.datetime
        try:
            return int(val.hour * 60 + val.minute + val.second / 60)
        except Exception:
            return None
    if isinstance(val, (int, float)):
        # Excel armazena hora como fração de dia (ex: 0.366 = 08:48)
        if 0 < float(val) < 1:
            return int(round(float(val) * 24 * 60))
        return int(val)
    if isinstance(val, str):
        return parse_hora(val)
    return None


def min_to_timedelta(minutos: int | None):
    if minutos is None:
        return None
    return timedelta(minutes=int(minutos))


def fmt_horas(minutos: int | None, vazio="—") -> str:
    """Formata minutos como H:MM (aceita >24h, ex: 82:38)."""
    if minutos is None:
        return vazio
    neg = minutos < 0
    m = abs(int(minutos))
    h, mm = divmod(m, 60)
    return f"{'-' if neg else ''}{h}:{mm:02d}"


def fmt_horas_curto(minutos: int | None) -> str:
    if minutos is None or minutos == 0:
        return "—"
    return fmt_horas(minutos)


_H_SUFIX = re.compile(r"[hH\.:]")


def parse_hora(texto: str) -> int | None:
    """Parser tolerante. Aceita: '', '7', '73', '730', '0730', '7:30',
    '07:30', '7h30', '7h', '12:45:00'. Retorna minutos ou None se vazio.
    Levanta ValueError se inválido."""
    if texto is None:
        return None
    t = str(texto).strip().lower().replace(" ", "")
    if t in ("", "—", "-", "--", ":", "x", "f"):
        return None
    # remove segundos :ss
    m = re.match(r"^(\d{1,3})[:h](\d{1,2})(?::\d{1,2})?$", t)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if mi > 59:
            raise ValueError(f"Minutos inválidos em '{texto}'")
        if h > 99:
            raise ValueError(f"Hora inválida em '{texto}'")
        return h * 60 + mi
    # só dígitos: 7 -> 07:00 | 730 -> 07:30 | 0730 -> 07:30 | 1245 -> 12:45
    if re.fullmatch(r"\d{1,4}", t):
        d = t
        if len(d) <= 2:  # "7" ou "12" = horas cheias
            return int(d) * 60
        if len(d) == 3:  # "730" = 7:30
            return int(d[0]) * 60 + int(d[1:])
        # 4 dígitos HHMM
        h, mi = int(d[:2]), int(d[2:])
        if mi > 59 or h > 23:
            raise ValueError(f"Horário inválido: '{texto}'")
        return h * 60 + mi
    raise ValueError(f"Use HH:MM — ex: 07:30 (recebido: '{texto}')")


def validar_hora(texto: str) -> tuple[bool, int | None, str]:
    """Retorna (ok, minutos, mensagem). Nunca levanta exceção."""
    try:
        v = parse_hora(texto)
        return True, v, ""
    except ValueError as e:
        return False, None, str(e)


# --------------------------------------------------------------------------
# Espelho das fórmulas (por dia e totais)
# --------------------------------------------------------------------------

def calc_dia(b, c, d, e, jornada_min: int, toler_min: int, lanche_min: int) -> tuple[int, int, int]:
    """Replica F/G/H da planilha.

    F = IF((C-B)+(E-D)=0, 0, ((C-B)+(E-D)) - lanche)
    G = IF(F=0, 0, IF((jornada-F) > toler, |jornada-F|, 0))
    H = IF((F-jornada) > toler, |F-jornada|, 0)
    Entradas None = célula vazia = 0 (semântica do Excel).
    """
    bv = b if b is not None else 0
    cv = c if c is not None else 0
    dv = d if d is not None else 0
    ev = e if e is not None else 0
    soma = (cv - bv) + (ev - dv)
    if soma == 0:
        return 0, 0, 0
    f = soma - (lanche_min or 0)
    if f < 0:
        f = 0
    if f == 0:
        return 0, 0, 0
    j = jornada_min or 0
    t = toler_min or 0
    g = abs(j - f) if (j - f) > t else 0
    h = abs(f - j) if (f - j) > t else 0
    return f, g, h


def calc_mes(dias: list[tuple], jornada_min: int, toler_min: int, lanche_min: int,
             saldo_anterior_min: int = 0, overrides: list[tuple] | None = None) -> dict:
    """Calcula linhas + totais. dias = [(b,c,d,e), ...].

    overrides (opcional): lista de (g_manual, h_manual) por dia. Quando o
    valor não é None, ele representa um AJUSTE MANUAL feito no Excel
    (célula G/H estática) e é usado no lugar do cálculo — e sinalizado
    em 'manuais' para a tela exibir o selo ✎.
    """
    linhas, manuais = [], []
    for i, (b, c, d, e) in enumerate(dias):
        f, g, h = calc_dia(b, c, d, e, jornada_min, toler_min, lanche_min)
        manual = False
        if overrides and i < len(overrides):
            gm, hm = overrides[i]
            if gm is not None:
                g, manual = int(gm), True
            if hm is not None:
                h, manual = int(hm), True
        linhas.append((f, g, h))
        manuais.append(manual)
    tot_falta = sum(g for _, g, _ in linhas)
    tot_extra = sum(h for _, _, h in linhas)
    dif = abs(tot_falta - tot_extra)  # G36
    credito = tot_extra + (saldo_anterior_min or 0)  # L39
    debito = tot_falta  # N39
    saldo = abs(credito - debito) if credito != 0 else debito  # O37: IF(L39=0,N35,|L39-N35|)
    if credito == 0:
        saldo = debito
    return {
        "linhas": linhas,
        "manuais": manuais,
        "tot_falta": tot_falta,
        "tot_extra": tot_extra,
        "dif": dif,
        "credito": credito,
        "debito": debito,
        "saldo": saldo,
        "saldo_anterior": saldo_anterior_min or 0,
    }
