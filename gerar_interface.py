"""
gerar_interface.py - deck de vazoes -> interface HTML

Uso:
    python gerar_interface.py                  (procura 'vazoes'/'vazoes.txt'; senao pergunta)
    python gerar_interface.py vazoes_2026.txt
    python gerar_interface.py vazoes_2026.txt --saida minha_interface.html
    python gerar_interface.py --ana-usos Resolucao-92-2021_..._v2.xlsx   (ativa aba Usos Consuntivos)
    python gerar_interface.py --corr-estruturais 1,2        (trata Classe 2 tambem como estrutural)
    python gerar_interface.py --corr-excluir-especiais      (nao corrige casos especiais do projeto)

Entradas (mesma pasta):
    vazoes / vazoes.txt              CodPosto Ano Jan..Dez
    usinas_atualizadas.xlsx          CodUsina | Usina | Jusante | CodPosto
    dados_hidroterm_completo.xlsx    abas Evap, Polynomials_PAC, PhysicalData (opcional)
    <resolucao ANA usos consuntivos>.xlsx   aba AHEs_series1931_2021 (opcional, --ana-usos)

Saidas:
    interface_usinas_<ini>-<fim>.html
    saidas/  (txt/csv intermediarios, diagnostico.txt, log)
    saidas/vazoes_corrigidas_massa.txt       deck corrigido (mesmo formato do original)
    saidas/correcao_massa_por_usina.csv      status e indicadores da correcao por usina
    saidas/evaporacao_por_usina.csv          perdas por evaporacao por usina (ranking)
    saidas/evaporacao_por_bacia.csv          perdas por evaporacao por bacia

Convencoes:
    Qinc      = Qi - SUM(Qm)     (vazao natural oficial do deck, sem nenhum ajuste)
    Evaporacao: DIAGNOSTICO de perdas, nao correcao. A vazao natural oficial ja incorpora a
                  evaporacao liquida dos reservatorios na reconstituicao (ONS, 2011), entao
                  somar ou subtrair a evaporacao do Qinc seria contar o termo duas vezes.
                  Calcula-se apenas a perda por usina e por bacia:
                    Evap (m3/s) = EvapMen (mm) x Area (km2) x 1e3 / segundos do mes
                  Area pelo polinomio area-cota (PAC) nas cotas min / media / max; EvapMen < 0
                  (condensacao) -> 0. Area Media e o valor de referencia; Min e Max sao a faixa
                  de incerteza pela cota. Media anual = volume anual / segundos do ano.
                  Indicadores: Evap/Qnat (perda relativa a vazao natural do posto) e
                  Evap/Qinc (peso do termo de evaporacao frente a vazao incremental).
                  Bacia = conjunto de usinas que drenam para a mesma usina de jusante final;
                  perda da bacia = soma da evap das usinas / Qnat da usina de jusante final.
                  Relacao com Qinc < 0: apenas descritiva (correlacao de Spearman).
    Anual = media dos meses
    Classes (% de meses com Qinc < 0): 1 > 30 % | 2 = 2-30 % | 3 = 0-2 % | 0 = nenhum
    Estacionariedade: MK + Sen + Pettitt + Yue-Wang, >= 30 anos completos.
                  Resumo separado: postos de usinas (base do relatorio) x todos os postos do deck.
    Janela movel: MK + Sen em janelas de 10/20/30 anos deslizantes (sem correcao Yue-Wang -
                  janelas curtas tornam a correcao instavel; ressalva mantida no proprio texto da interface)
    Permanencia: Weibull m/(n+1), vazao natural mensal
    Topologia isolada: usinas cuja cascata tem tamanho 1 (sem montante cadastrada) - a maioria
                  sao cabecas de bacia reais; casos citados no projeto original (Belo Monte-
                  Pimental etc.) sao marcados a parte
    Correcao (conservacao de massa): Qinc < 0 -> 0 e o mesmo volume e retirado dos
                  meses positivos do mesmo ano (fator 1 - D/P). Natural corrigida reconstruida
                  de montante p/ jusante: Qi' = SUM(Qm') + Qinc'_i (volume anual preservado em
                  todos os postos). Anos com Qinc anual <= 0 nao sao corrigiveis. Classe 1
                  = estrutural (nao alterada; recomenda-se corrigir topologia/posto).
    Usos consuntivos: DIAGNOSTICO, nao correcao - a vazao natural oficial ja soma usos
                  consuntivos de volta (Qnat = Qobs + Qconsumo), entao o termo incremental
                  recalculado aqui ja esta embutido no Qinc oficial. Compara a magnitude do
                  consumo no trecho incremental com o deficit medio dos meses de Qinc
                  negativo - nunca para subtrair de novo do Qinc.
                  Correspondencia nome-usina <-> nome-ANA fica embutida neste arquivo
                  (NOME_ANA_POR_USINA / CEG_POR_USINA / NOMES_SOMA_POR_USINA), resolvida
                  manualmente uma vez; casamento automatico (nome exato + apelido "Antiga X")
                  cobre usinas novas sem precisar editar nada.

Requisitos: pip install pandas numpy scipy openpyxl
"""
import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path
from collections import defaultdict

import numpy as np
import pandas as pd
from scipy.stats import norm, spearmanr

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
SEG_MES = [d * 86400 for d in [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]]
SEG_ANO = float(sum(SEG_MES))

MIN_ANOS_ESTAC = 30
ALPHA = 0.05
AREA_MAX_PLAUSIVEL_KM2 = 6000.0   # limite p/ area suspeita
NA_TOKENS = ["(NA)", "NA", "na", "Na", "N/A", "n/a", "<NA>", "None", ""]
JANELAS_MOVEIS = [10, 20, 30]      # tamanhos de janela p/ Mann-Kendall movel
LIM_CLASSE1 = 30.0                 # % meses negativos
LIM_CLASSE2 = 2.0

# nomes citados no projeto original como casos especiais sensiveis (topologia/transposicao)
CASOS_ESPECIAIS_TOPO = [
    "BELO MONTE", "PIMENTAL", "PAULO AFONSO", "MOXOTO", "PAF-MOX", "ILHA SOLTEIRA", "I. SOLTEIRA",
    "TRES IRMAOS", "NILO PECANHA", "FONTES", "JORDAO", "HENRY BORDEN", "LAJES", "P. PASSOS",
]

try:  # UTF-8 no terminal Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

LOG = []
def log(msg):
    print(msg)
    LOG.append(msg)


# ============================================================
# 1) LEITURA
# ============================================================
def ler_deck(path: Path) -> pd.DataFrame:
    """Le deck -> CodPosto, Ano, Jan..Dez."""
    try:
        df = pd.read_csv(path, sep=r"\s+", engine="python", header=None)
        if df.shape[1] != 14:
            raise ValueError
    except Exception:
        # fallback largura fixa (I3, I5, 12xI6)
        df = pd.read_fwf(path, widths=[3, 5] + [6] * 12, header=None)
    if df.shape[1] != 14:
        raise ValueError(f"{path}: esperado 14 colunas, encontrei {df.shape[1]}")
    df.columns = ["CodPosto", "Ano"] + MESES
    df = df.apply(pd.to_numeric, errors="coerce")
    df = df.dropna(subset=["CodPosto", "Ano"])
    df["CodPosto"] = df["CodPosto"].astype(int)
    df["Ano"] = df["Ano"].astype(int)
    df[MESES] = df[MESES].astype(float)
    return df.sort_values(["CodPosto", "Ano"]).reset_index(drop=True)


def marcar_meses_futuros(deck: pd.DataFrame) -> pd.DataFrame:
    """Ultimo ano: meses finais zerados em todos os postos -> NaN."""
    ult = deck["Ano"].max()
    sub = deck[deck["Ano"] == ult]
    futuros = [m for m in MESES if (sub[m].fillna(0) == 0).all()]
    # so meses contiguos ate Dez
    fut_final = []
    for m in reversed(MESES):
        if m in futuros:
            fut_final.insert(0, m)
        else:
            break
    if fut_final:
        deck.loc[deck["Ano"] == ult, fut_final] = np.nan
        log(f"AVISO: Ano {ult}: meses {fut_final[0]}–{fut_final[-1]} zerados em todos os postos "
            f"→ tratados como ausentes.")
    return deck


def ler_usinas(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path).replace(NA_TOKENS, pd.NA)
    ren = {}
    for c in df.columns:
        k = str(c).strip().lower()
        if k in {"codusina", "códusina", "codigo", "código"}: ren[c] = "CodUsina"
        elif k in {"usina", "nome"}:                        ren[c] = "Usina"
        elif k in {"jusante", "codusinaajusante", "jus"}:   ren[c] = "Jusante"
        elif k in {"codposto", "posto", "idposto"}:         ren[c] = "CodPosto"
    df = df.rename(columns=ren)
    falt = [c for c in ["CodUsina", "Usina", "Jusante", "CodPosto"] if c not in df.columns]
    if falt:
        raise ValueError(f"Colunas ausentes em {path}: {falt}")
    for c in ["CodUsina", "Jusante", "CodPosto"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
    df["Usina"] = df["Usina"].astype("string").str.strip()
    df = df.dropna(subset=["CodUsina"]).drop_duplicates("CodUsina").reset_index(drop=True)

    dup = df.dropna(subset=["CodPosto"]).groupby("CodPosto")["CodUsina"].apply(list)
    dup = dup[dup.apply(len) > 1]
    for p, us in dup.items():
        log(f"INFO: Posto {p} compartilhado pelas usinas {us}")
    return df


def ler_hidroterm(path: Path):
    evap = pd.read_excel(path, sheet_name="Evap")
    pac = pd.read_excel(path, sheet_name="Polynomials_PAC")
    phy = pd.read_excel(path, sheet_name="PhysicalData")
    for d in (evap, pac, phy):
        d.columns = [str(c).strip() for c in d.columns]
        d["CodUsina"] = pd.to_numeric(d["CodUsina"], errors="coerce").astype("Int64")
    pac = pac.rename(columns={f"PAC({i})": f"PAC{i}" for i in range(5)})
    phy = phy.rename(columns={"Cota Máx (m)": "CotaMax", "Cota Mín (m)": "CotaMin"})
    return evap, pac, phy


# ============================================================
# 2) EVAPORACAO (diagnostico de perdas - NAO entra no Qinc)
# ============================================================
def calcular_evap(evap_raw, pac, phy) -> dict:
    """Area via PAC nas cotas min/med/max; evap mm/mes -> m3/s, clamp >= 0."""
    ar = phy.merge(pac[["CodUsina"] + [f"PAC{i}" for i in range(5)]], on="CodUsina", how="left")
    areas = {}
    for _, r in ar.iterrows():
        if pd.isna(r["CodUsina"]):
            continue
        c = [r[f"PAC{i}"] for i in range(5)]
        hmin, hmax = r["CotaMin"], r["CotaMax"]
        if any(pd.isna(x) for x in c + [hmin, hmax]):
            continue
        f = lambda h: c[0] + c[1]*h + c[2]*h**2 + c[3]*h**3 + c[4]*h**4
        areas[int(r["CodUsina"])] = (f(hmin), f((hmin + hmax) / 2), f(hmax))

    cols = [f"EvapMen({i})" for i in range(1, 13)]
    out = {}
    for _, r in evap_raw.iterrows():
        if pd.isna(r["CodUsina"]):
            continue
        u = int(r["CodUsina"])
        if u not in areas:
            continue
        amin, amed, amax = areas[u]
        suspeita = (amin <= 0 or amax <= 0 or amax > AREA_MAX_PLAUSIVEL_KM2 or amin > amax)
        mm = [0.0 if pd.isna(r[cols[i]]) else max(float(r[cols[i]]), 0.0) for i in range(12)]
        d = {"areaMin": amin, "areaMed": amed, "areaMax": amax, "suspeita": bool(suspeita), "mm": mm}
        for tipo, a in [("min", amin), ("med", amed), ("max", amax)]:
            d[tipo] = [max(mm[i] * a * 1e3 / SEG_MES[i], 0.0) for i in range(12)]
        out[u] = d
        if suspeita:
            log(f"AVISO: Usina {u}: área fora do plausível (mín={amin:.1f}, máx={amax:.1f} km²) "
                f"— checar PAC/cotas.")
    return out


def media_volumetrica(mensal_m3s):
    """Vazao media anual (m3/s) = volume anual / segundos do ano."""
    return float(sum(v * s for v, s in zip(mensal_m3s, SEG_MES)) / SEG_ANO)


def vazao_media(df: pd.DataFrame):
    """Media de longo termo: media das medias anuais dos anos completos (ou de todos os meses)."""
    if df is None:
        return None
    comp = df.dropna()
    if len(comp):
        return float(comp.mean(axis=1).mean())
    v = df.values[~np.isnan(df.values)]
    return float(v.mean()) if v.size else None


def usina_terminal(usinas: pd.DataFrame) -> dict:
    """{CodUsina: usina de jusante final} seguindo a cadeia de jusantes."""
    cods = set(int(c) for c in usinas["CodUsina"])
    nxt = {int(r.CodUsina): int(r.Jusante) for r in usinas.itertuples()
           if pd.notna(r.Jusante) and int(r.Jusante) != 0 and int(r.Jusante) in cods}
    term = {}
    for u in cods:
        v, visto = u, {u}
        while v in nxt and nxt[v] not in visto:
            v = nxt[v]
            visto.add(v)
        term[u] = v
    return term


def calc_metricas_evap(usinas: pd.DataFrame, series: dict, qinc: dict, evap: dict) -> dict:
    """Perdas por evaporacao por usina e por bacia + relacao descritiva com Qinc < 0."""
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}
    term = usina_terminal(usinas)

    por_usina = {}
    for u, e in evap.items():
        p = u2p.get(u)
        qnat_m = vazao_media(series.get(p)) if p is not None else None
        qinc_m = vazao_media(qinc[u]) if u in qinc else None
        neg, tot = n_neg(qinc[u]) if u in qinc else (0, 0)
        m = {k: media_volumetrica(e[k]) for k in ["min", "med", "max"]}
        vals_q = qinc[u].values[~np.isnan(qinc[u].values)] if u in qinc else np.array([])
        por_usina[u] = {
            "evapMin": m["min"], "evapMed": m["med"], "evapMax": m["max"],
            "volAnualHm3": m["med"] * SEG_ANO / 1e6,
            "laminaAnualMm": float(sum(e["mm"])),
            "qnatMedia": qnat_m, "qincMedia": qinc_m,
            "pctQnat": (m["med"] / qnat_m * 100) if qnat_m and qnat_m > 0 else None,
            "pctQinc": (m["med"] / qinc_m * 100) if qinc_m and qinc_m > 0 else None,
            "pctNeg": (neg / tot * 100) if tot else None,
            "qincNulo": bool(vals_q.size) and bool(np.all(vals_q == 0)),
            "terminal": term.get(u),
        }
    for k, u in enumerate(sorted(por_usina, key=lambda x: -por_usina[x]["evapMed"]), 1):
        por_usina[u]["rank"] = k
    for k, u in enumerate(sorted([x for x in por_usina if por_usina[x]["pctQnat"] is not None],
                                 key=lambda x: -por_usina[x]["pctQnat"]), 1):
        por_usina[u]["rankPctQnat"] = k

    # ---- bacias (usinas que drenam para a mesma usina de jusante final)
    grupos = defaultdict(list)
    for u in (int(c) for c in usinas["CodUsina"]):
        grupos[term[u]].append(u)
    bacias = []
    for t, membros in grupos.items():
        com = [u for u in membros if u in por_usina]
        if not com:
            continue
        pt = u2p.get(t)
        qnat_t = vazao_media(series.get(pt)) if pt is not None else None
        tot = {k: sum(por_usina[u][f"evap{k}"] for u in com) for k in ["Min", "Med", "Max"]}
        top = sorted(com, key=lambda u: -por_usina[u]["evapMed"])[:3]
        bacias.append({
            "terminal": t, "nome": nome.get(t, str(t)), "nUsinas": len(membros), "nComEvap": len(com),
            "evapMin": tot["Min"], "evapMed": tot["Med"], "evapMax": tot["Max"],
            "volAnualHm3": tot["Med"] * SEG_ANO / 1e6,
            "qnatExutorio": qnat_t,
            "pctQnat": (tot["Med"] / qnat_t * 100) if qnat_t and qnat_t > 0 else None,
            "principais": [nome.get(u, str(u)) for u in top],
        })
    bacias.sort(key=lambda b: -b["evapMed"])

    # ---- relacao com Qinc < 0 (descritiva, Spearman)
    def corr(chave):
        pares = [(v[chave], v["pctNeg"]) for v in por_usina.values()
                 if v[chave] is not None and v["pctNeg"] is not None and not v["qincNulo"]]
        if len(pares) < 5:
            return None
        x, y = np.array(pares).T
        rho, p = spearmanr(x, y)
        if not np.isfinite(rho) or not np.isfinite(p):
            return None
        return {"n": len(pares), "rho": round(float(rho), 3), "p": float(p)}

    total = {k: sum(v[f"evap{k}"] for v in por_usina.values()) for k in ["Min", "Med", "Max"]}
    qnat_bacias = sum(b["qnatExutorio"] for b in bacias if b["qnatExutorio"])
    resumo = {
        "nUsinas": len(por_usina),
        "totalMin": total["Min"], "totalMed": total["Med"], "totalMax": total["Max"],
        "volAnualKm3": total["Med"] * SEG_ANO / 1e9,
        "qnatBacias": qnat_bacias,
        "pctSIN": (total["Med"] / qnat_bacias * 100) if qnat_bacias else None,
        "bacias": bacias,
        "relacao": {"pctQinc": corr("pctQinc"), "pctQnat": corr("pctQnat")},
    }
    log(f"Evaporação: {len(por_usina)} usinas, total {total['Med']:.0f} m³/s (Área Méd; faixa "
        f"{total['Min']:.0f}–{total['Max']:.0f}) ≈ {resumo['volAnualKm3']:.1f} km³/ano.")
    return {"porUsina": por_usina, "resumo": resumo}


# ============================================================
# 3) INCREMENTAIS
# ============================================================
def matriz_posto(deck: pd.DataFrame) -> dict:
    """{posto: DF ano x mes}"""
    return {p: g.set_index("Ano")[MESES] for p, g in deck.groupby("CodPosto")}


def montante_map(usinas: pd.DataFrame) -> dict:
    """{CodUsina: [montantes imediatas]}"""
    mont = {}
    for r in usinas.itertuples():
        if pd.notna(r.Jusante) and int(r.Jusante) != 0:
            mont.setdefault(int(r.Jusante), []).append(int(r.CodUsina))
    return mont


def calcular_incrementais(usinas: pd.DataFrame, series: dict) -> dict:
    """{CodUsina: DF ano x mes} com Qinc = Qi - SUM(Qm)."""
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    mont = montante_map(usinas)

    res = {}
    for u in sorted(u2p):
        p = u2p[u]
        if p not in series:
            continue
        qi = series[p]
        soma = None
        for um in sorted(set(mont.get(u, []))):
            pm = u2p.get(um)
            if pm is None or pm not in series:
                log(f"AVISO: Usina {u}: montante {um} sem posto/série — ignorada no ΣQm "
                    f"(as montantes dela também ficam de fora).")
                continue
            soma = series[pm] if soma is None else soma.add(series[pm], fill_value=0)
        qinc = qi.copy() if soma is None else qi.subtract(soma, fill_value=0)
        # mantem NaN da propria usina
        qinc = qinc.where(qi.notna())
        vals = qinc.values[~np.isnan(qinc.values)]
        if vals.size and np.all(vals == 0):
            log(f"AVISO: Usina {u} (posto {p}): Qinc ≡ 0 (posto igual ao de montante).")
        res[u] = qinc
    return res


# ============================================================
# 4) ESTACIONARIEDADE (+ janela movel)
# ============================================================
def sen_slope(y, x):
    i, j = np.triu_indices(len(y), k=1)
    dx = x[j] - x[i]
    ok = dx != 0
    return float(np.median((y[j] - y[i])[ok] / dx[ok]))


def mann_kendall(y, x):
    n = len(y)
    s = sum(np.sum(np.sign(y[k+1:] - y[k])) for k in range(n - 1))
    _, cnt = np.unique(y, return_counts=True)
    var_s = (n*(n-1)*(2*n+5) - np.sum(cnt*(cnt-1)*(2*cnt+5))) / 18.0
    if var_s <= 0:
        return 1.0, "sem tendência sig."
    z = (s - 1)/np.sqrt(var_s) if s > 0 else ((s + 1)/np.sqrt(var_s) if s < 0 else 0.0)
    p = 2 * (1 - norm.cdf(abs(z)))
    trend = ("crescente" if (p < ALPHA and s > 0) else
             "decrescente" if (p < ALPHA and s < 0) else "sem tendência sig.")
    return float(p), trend


def pettitt(y, anos):
    n = len(y)
    sm = np.sign(y[None, :] - y[:, None])
    cs = np.cumsum(sm, axis=0)
    U = np.array([-np.sum(cs[t, t+1:]) for t in range(n - 1)])
    k = int(np.argmax(np.abs(U)))
    K = float(abs(U[k]))
    p = min(2 * np.exp(-6 * K**2 / (n**3 + n**2)), 1.0)
    return int(anos[k]), float(p), k


def yue_wang(y, x):
    n = len(y)
    beta = sen_slope(y, x)
    yd = y - beta * (x - x[0])
    r1 = float(np.corrcoef(yd[:-1], yd[1:])[0, 1])
    lim = (-1 + 1.96 * np.sqrt(n - 2)) / (n - 1)
    sig = abs(r1) > lim
    if sig:
        yw = yd[1:] - r1 * yd[:-1]
        return yw + beta * (x[1:] - x[0]), x[1:], r1, sig, beta
    return y.copy(), x.copy(), r1, sig, beta


def janela_movel(anos, valores, janela):
    """MK + Sen em janelas deslizantes (sem Yue-Wang - janela curta deixa a correcao instavel)."""
    anos = np.asarray(anos)
    valores = np.asarray(valores, dtype=float)
    n = len(anos)
    if n < janela:
        return None
    anos_final, p_valores, tendencias, sen_slopes = [], [], [], []
    for start in range(0, n - janela + 1):
        end = start + janela
        y, x = valores[start:end], anos[start:end]
        p, trend = mann_kendall(y, x)
        sen = sen_slope(y, x)
        anos_final.append(int(x[-1]))
        p_valores.append(round(p, 4))
        tendencias.append(trend if trend != "sem tendência sig." else "neutro")
        sen_slopes.append(round(sen, 3))
    return {"anosFinal": anos_final, "pValores": p_valores, "tendencias": tendencias, "senSlopes": sen_slopes}


def estacionariedade_posto(df_posto: pd.DataFrame):
    """So anos completos. Retorna teste na serie inteira + janelas moveis (10/20/30 anos)."""
    comp = df_posto.dropna()
    if len(comp) < MIN_ANOS_ESTAC:
        return None
    anos = comp.index.values.astype(int)
    y = comp.mean(axis=1).values.astype(float)
    p_b, t_b = mann_kendall(y, anos)
    y_pw, x_pw, r1, r1sig, beta = yue_wang(y, anos)
    p_c, t_c = mann_kendall(y_pw, x_pw)
    ano_q, p_pt, k = pettitt(y, anos)
    intercept = float(np.median(y - beta * anos))
    pre, pos = float(y[:k+1].mean()), float(y[k+1:].mean())

    jm = {}
    for j in JANELAS_MOVEIS:
        r = janela_movel(anos, y, j)
        if r:
            jm[str(j)] = r

    return {
        "anos": anos.tolist(),
        "natural": [round(v, 2) for v in y],
        "trendLine": [round(intercept + beta * a, 2) for a in anos],
        "senSlope": round(beta, 4),
        "r1": round(r1, 3), "r1Sig": bool(r1sig),
        "mkPBruto": p_b, "mkTrendBruto": t_b,
        "mkPCorr": p_c, "mkTrendCorr": t_c,
        "pettittAno": ano_q, "pettittP": p_pt,
        "mediaPre": round(pre, 1), "mediaPos": round(pos, 1),
        "preLine": [round(pre, 2) if i <= k else None for i in range(len(anos))],
        "posLine": [round(pos, 2) if i > k else None for i in range(len(anos))],
        "janelaMovel": jm,
    }


def resumo_estac(estac_postos: dict) -> dict:
    v = list(estac_postos.values())
    n = len(v)
    nsig = sum(e["mkPCorr"] < ALPHA for e in v)
    nbrk = sum(e["pettittP"] < ALPHA for e in v)
    nr1 = sum(e["r1Sig"] for e in v)
    # quebra sem tendencia significativa: Pettitt tambem acusa quebra quando ha tendencia,
    # entao esse grupo e o mais proximo de "mudanca de patamar" pura
    nbrk_sem_tend = sum(e["pettittP"] < ALPHA and e["mkPCorr"] >= ALPHA for e in v)
    dec = pd.Series([e["pettittAno"] // 10 * 10 for e in v if e["pettittP"] < ALPHA],
                    dtype="int64").value_counts().sort_index()
    pct = lambda a: round(a / n * 100, 1) if n else 0
    return {
        "nTotal": n, "nSigCorr": int(nsig), "pctSigCorr": pct(nsig),
        "nCresc": int(sum(e["mkTrendCorr"] == "crescente" for e in v)),
        "nDecr": int(sum(e["mkTrendCorr"] == "decrescente" for e in v)),
        "nBreakSig": int(nbrk), "pctBreakSig": pct(nbrk),
        "nBreakSemTend": int(nbrk_sem_tend),
        "nR1Sig": int(nr1), "pctR1Sig": pct(nr1),
        "fpEsperados": round(n * ALPHA, 1),
        "decadaLabels": [str(int(d)) for d in dec.index],
        "decadaValores": [int(c) for c in dec.values],
    }


# ============================================================
# 5) CASCATAS
# ============================================================
def montar_cascatas(usinas: pd.DataFrame) -> list:
    # Jusante vazio ou 0 = fim
    nxt = {int(r.CodUsina): int(r.Jusante) for r in usinas.itertuples()
           if pd.notna(r.Jusante) and int(r.Jusante) != 0}
    cods = set(usinas["CodUsina"].astype(int))
    cabecas = sorted(cods - set(nxt.values()))
    cas = []
    for u0 in cabecas:
        cad, visto, u = [u0], {u0}, u0
        while u in nxt:
            v = nxt[u]
            if v in visto:
                log(f"AVISO: Ciclo na topologia a partir da usina {u0} (volta em {v}).")
                break
            cad.append(v); visto.add(v); u = v
        cas.append(cad)
    return cas


def normaliza_nome(s: str) -> str:
    n = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return n.upper()


def eh_caso_especial(nome: str) -> bool:
    nn = normaliza_nome(nome or "").strip()
    return bool(nn) and any(c in nn for c in CASOS_ESPECIAIS_TOPO)


def calc_topologia(usinas: pd.DataFrame, cascatas: list, series: dict) -> dict:
    """Usinas em cascatas de tamanho 1 (sem montante cadastrada) - diagnostico, nao erro."""
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}
    isoladas_cod = {cad[0] for cad in cascatas if len(cad) == 1}

    lista = []
    for u in sorted(isoladas_cod):
        p = u2p.get(u)
        media_geral = 0.0
        if p in series:
            comp = series[p].dropna()
            if len(comp):
                media_geral = float(comp.mean(axis=1).mean())
        lista.append({
            "codUsina": u, "nome": nome.get(u, f"Usina {u}"), "codPosto": p,
            "mediaAnual": round(media_geral, 1), "casoEspecial": eh_caso_especial(nome.get(u, "")),
        })
    lista.sort(key=lambda x: -x["mediaAnual"])
    return {"totalUsinas": len(usinas), "totalIsoladas": len(lista), "lista": lista}


# ============================================================
# 5b) SAZONALIDADE, PERMANENCIA, EXTREMOS (vazao natural)
# ============================================================
PONTOS_PERM = [1, 2, 5, 10, 20, 30, 40, 50, 60, 70, 80, 90, 95, 98, 99]


def curva_weibull(vals, pontos=PONTOS_PERM):
    """Permanencia (Weibull m/(n+1)); interp. linear nos pontos (%)."""
    x = np.sort(np.asarray(vals, dtype=float))[::-1]
    if x.size < 2:
        return None
    pp = np.arange(1, x.size + 1) / (x.size + 1) * 100
    return [round(float(np.interp(p, pp, x)), 2) for p in pontos]


def calc_sazonalidade(df: pd.DataFrame):
    por_mes = {}
    for m in MESES:
        x = df[m].dropna().values
        if x.size:
            por_mes[m] = {"media": round(float(x.mean()), 2), "mediana": round(float(np.median(x)), 2),
                          "p10": round(float(np.percentile(x, 10)), 2), "p90": round(float(np.percentile(x, 90)), 2),
                          "min": float(x.min()), "max": float(x.max())}
    if not por_mes:
        return None
    todos = df.values[~np.isnan(df.values)]
    cheio = max(por_mes, key=lambda m: por_mes[m]["media"])
    seco = min(por_mes, key=lambda m: por_mes[m]["media"])
    media = float(todos.mean()) if todos.size else 0
    ind = round((por_mes[cheio]["media"] - por_mes[seco]["media"]) / media, 3) if media else None
    return {"porMes": por_mes, "mesMaisCheio": cheio, "mesMaisSeco": seco, "indice": ind}


def calc_permanencia(df: pd.DataFrame):
    todos = df.values[~np.isnan(df.values)]
    geral = curva_weibull(todos)
    if not geral:
        return None
    mensal = {}
    for m in MESES:
        c = curva_weibull(df[m].dropna().values)
        if c:
            mensal[m] = c
    q = {p: v for p, v in zip(PONTOS_PERM, geral)}
    return {"pontosGeral": PONTOS_PERM, "curvaGeral": geral,
            "pontosMensal": PONTOS_PERM, "curvaMensal": mensal,
            "q10": q[10], "q50": q[50], "q95": q[95]}


def calc_extremos(df: pd.DataFrame):
    longo = df.stack()
    if longo.empty:
        return None
    (amin, mmin), (amax, mmax) = longo.idxmin(), longo.idxmax()
    out = {"minMensal": {"valor": float(longo.min()), "quando": f"{mmin}/{int(amin)}"},
           "maxMensal": {"valor": float(longo.max()), "quando": f"{mmax}/{int(amax)}"}}
    comp = df.dropna()
    if len(comp):
        an = comp.mean(axis=1)
        out["minAnual"] = {"valor": round(float(an.min()), 2), "ano": int(an.idxmin())}
        out["maxAnual"] = {"valor": round(float(an.max()), 2), "ano": int(an.idxmax())}
    return out


# ============================================================
# 5c) USOS CONSUNTIVOS (diagnostico, opcional - precisa --ana-usos)
# ============================================================
# Correspondencia Usina (nome truncado do cadastro) -> nome oficial na base ANA
# (Resolucao ANA 92/2021, "Series Historicas Mensais de Usos Consuntivos a Montante de AHEs").
# Resolvida manualmente cruzando por rio/CEG quando o nome sozinho era ambiguo ou repetido
# entre usinas diferentes (a base ANA reaproveita nomes iguais em rios distintos).
NOME_ANA_POR_USINA = {
    'FUNIL-GRANDE': 'Funil',
    'M. DE MORAES': 'Marechal Mascarenhas de Moraes (Antiga Peixoto)',
    'P. COLOMBIA': 'Porto Colômbia',
    'E. DA CUNHA': 'Euclides da Cunha',
    'A.S. OLIVEIRA': 'Limoeiro (Armando Salles de Oliveira)',
    'A. VERMELHA': 'Água Vermelha (Antiga José Ermírio de Moraes)',
    'SERRA FACAO': 'Serra do Facão',
    'EMBOCACAO': 'Emborcação',
    'CAPIM BRANC1': 'Amador Aguiar I (Antiga Capim Branco I)',
    'CAPIM BRANC2': 'Amador Aguiar II (Antiga Capim Branco II)',
    'I. SOLTEIRA': 'Ilha Solteira',
    'A.S. LIMA': 'Bariri (Álvaro de Souza Lima)',
    'NAVANHANDAVA': 'Nova Avanhandava (Rui Barbosa)',
    'P. PRIMAVERA': 'Porto Primavera (Eng° Sérgio Motta)',
    'A.A. LAYDNER': 'Jurumirim (Armando Avellanal Laydner)',
    'L.N. GARCEZ': 'Salto Grande (Lucas Nogueira Garcez)',
    'CAPIVARA': 'Capivara (Escola de Engenharia Mackenzie)',
    'STA CLARA PR': 'Santa Clara',
    'JORDAO': 'Derivação do Rio Jordão',
    'G.B. MUNHOZ': 'Governador Bento Munhoz da Rocha Neto (Foz do Areia)',
    'SEGREDO': 'Governador Ney Aminthas de Barros Braga (Segredo)',
    'SLT.SANTIAGO': 'Salto Santiago',
    'SALTO CAXIAS': 'Governador José Richa (Salto Caxias)',
    'PASSO S JOAO': 'Passo São João',
    'FOZ CHAPECO': 'Foz do Chapecó',
    'D. FRANCISCA': 'Dona Francisca',
    'G.P. SOUZA': 'Governador Pedro Viriato Parigot de Souza (Capivari/Cachoeira)',
    'ILHA POMBOS': 'Ilha dos Pombos',
    'P. PASSOS': 'Pereira Passos',
    'P. ESTRELA': 'Porto Estrela',
    'ITAPARICA': 'Luiz Gonzaga (Itaparica)',
    'P. CAVALO': 'Pedra do Cavalo',
    'B. ESPERANCA': 'Boa Esperança (Antiga Castelo Branco)',
    'GUILMLAM-AMOR': 'Guilman Amorim',
    'SLT VERDINHO': 'Salto do Rio Verdinho',
    'SERRA MESA': 'Serra da Mesa',
    'LAJEADO': 'Luís Eduardo Magalhães (Lajeado)',
    'ESTREITO TOC': 'Estreito (Luiz Carlos Barreto de Carvalho)',
    'PONTE PEDRA': 'Ponte de Pedra',
    'STO ANT JARI': 'Santo Antônio do Jari',
    'STO ANTONIO': 'Santo Antônio',
    'JAGUARI': 'Jaguari (UHEPHSP027131-4)',
    'ITIQUIRA I': 'Itiquira (Casas de Forças I e II)',
    'ITIQUIRA II': 'Itiquira (Casas de Forças I e II)',
    'B. COQUEIROS': 'Barra dos Coqueiros',
    'FOZ R. CLARO': 'Engenheiro José Luiz Muller de Godoy Pereira (Antiga Foz do Rio Claro)',
    'BATALHA': 'Batalha (Antiga Paulista)',
    'CACH.DOURADA': 'Cachoeira Dourada',
    'PROMISSAO': 'Promissão (Mário Lopes Leão)',
    'JUPIA': 'Jupiá (Eng° Souza Dias)',
    'TIBAGI MONT': 'Tibagi Montante',
    'TAQUARUCU': 'Taquaruçu (Escola Politécnica)',
    'ITAIPU': 'Itaipu (Parte Brasileira)',
    'QUEBRA QUEIX': 'Quebra Queixo',
    'LAJES': 'Lajes (Fontes Velha)',
    'FONTES': 'Fontes Nova',
    'CACH.CALDEIR': 'Cachoeira Caldeirão',
    'PEIXE ANGIC': 'Peixe Angical',
    'COARACY NUNE': 'Coaracy Nunes',
    'FERREIRA GOM': 'Ferreira Gomes',
}
# casos onde o NOME sozinho na base ANA nao e unico (mesmo nome em rios diferentes) -> usar CEG
CEG_POR_USINA = {
    'STA BRANCA T': 'UHE.PH.PR.035290-0.01',   # Santa Branca, Rio Tibagi (nao confundir c/ Paraiba do Sul)
    'STA CLARA MG': 'UHE.PH.MG.002699-9.01',   # Santa Clara, Rio Mucuri (nao confundir c/ Rio Jordao)
}
# complexo sem registro unico na base ANA -> varios registros individuais. Como a base ANA traz o
# consumo ACUMULADO a montante de cada aproveitamento, usa-se o MAIOR valor do mes entre os registros
# (o mais de jusante); somar os registros contaria a bacia a montante varias vezes.
NOMES_SOMA_POR_USINA = {
    'COMP PAF-MOX': ['Paulo Afonso I', 'Paulo Afonso IV', 'Apolônio Sales (Antiga Moxotó)'],
}


def casar_nome_ana_auto(nome_eric, alias_lookup):
    """Match automatico: nome exato (normalizado) ou apelido 'Antiga X' entre parenteses."""
    return alias_lookup.get(normaliza_nome(nome_eric))


def ler_ana_usos(path: Path):
    """Le a planilha da Resolucao ANA 92/2021 (aba AHEs_series1931_2021)."""
    log(f"Lendo base ANA de usos consuntivos: {path} (pode demorar)...")
    df = pd.read_excel(path, sheet_name="AHEs_series1931_2021", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    log(f"   {df['NOME'].nunique()} nomes distintos, {len(df)} registros mensais")
    return df


def montar_series_usos(usinas: pd.DataFrame, ana_df: pd.DataFrame):
    """{CodUsina: {(ano,mes): consumo_TOTAL_m3_s}} aplicando overrides/CEG/soma."""
    # alias automatico: nome exato + "Antiga X" entre parenteses
    alias_lookup = {}
    for n in ana_df["NOME"].dropna().unique():
        alias_lookup[normaliza_nome(n)] = n
        m = re.search(r"\((?:Antiga|Antigo)\s+([^)]+)\)", n, re.I)
        if m:
            alias_lookup[normaliza_nome(m.group(1))] = n

    por_nome = ana_df.groupby(["NOME", "aa_ref", "mm_ref"])["consumo_TOTAL_m3_s"].sum()
    por_ceg = ana_df.dropna(subset=["CEG"]).groupby(["CEG", "aa_ref", "mm_ref"])["consumo_TOTAL_m3_s"].sum()

    series = {}
    n_resolvidas = 0
    for r in usinas.itertuples():
        u, nome_eric = int(r.CodUsina), str(r.Usina)
        if nome_eric in NOMES_SOMA_POR_USINA:
            serie = defaultdict(lambda: float('-inf'))
            for n in NOMES_SOMA_POR_USINA[nome_eric]:
                if n in por_nome.index.get_level_values(0):
                    for (ano, mes), v in por_nome.loc[n].groupby(level=[0, 1]).sum().items():
                        serie[(ano, mes)] = max(serie[(ano, mes)], v)
            if serie:
                series[u] = dict(serie); n_resolvidas += 1
        elif nome_eric in CEG_POR_USINA:
            ceg = CEG_POR_USINA[nome_eric]
            if ceg in por_ceg.index.get_level_values(0):
                series[u] = {(a, m): v for (a, m), v in por_ceg.loc[ceg].items()}
                n_resolvidas += 1
        else:
            nome_ana = NOME_ANA_POR_USINA.get(nome_eric) or casar_nome_ana_auto(nome_eric, alias_lookup)
            if nome_ana and nome_ana in por_nome.index.get_level_values(0):
                series[u] = {(a, m): v for (a, m), v in por_nome.loc[nome_ana].items()}
                n_resolvidas += 1
    log(f"Usos consuntivos: correspondência resolvida p/ {n_resolvidas} de {len(usinas)} usinas")
    return series


def calc_usos_incremental(usinas: pd.DataFrame, series_usos: dict):
    """Usos consuntivos incremental por trecho (jusante - soma montantes), mensal."""
    mont = montante_map(usinas)
    inc = {}
    for u, serie_jus in series_usos.items():
        montantes = mont.get(u, [])
        chaves = set(serie_jus.keys())
        for m in montantes:
            chaves |= set(series_usos.get(m, {}).keys())
        d = {}
        for k in chaves:
            v_jus = serie_jus.get(k, 0.0)
            v_mont = sum(series_usos.get(m, {}).get(k, 0.0) for m in montantes)
            d[k] = v_jus - v_mont
        inc[u] = d
    return inc


def calc_diagnostico_usos(usinas: pd.DataFrame, qinc: dict, usos_inc: dict) -> dict:
    """Usos consuntivos incremental x meses de Qinc negativo."""
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}
    linhas = []
    for u, inc_usos in usos_inc.items():
        if u not in qinc:
            continue
        q = qinc[u]
        negativos = []
        for ano, row in q.iterrows():
            for i, m in enumerate(MESES, 1):
                v = row[m]
                if pd.notna(v) and v < 0 and (int(ano), i) in inc_usos:
                    negativos.append((v, inc_usos[(int(ano), i)]))
        if not negativos:
            continue
        usos_medio_geral = np.mean(list(inc_usos.values())) if inc_usos else 0.0
        q_neg = np.mean([x for x, _ in negativos])
        u_neg = np.mean([y for _, y in negativos])
        cobertura = round(u_neg / abs(q_neg) * 100, 1) if q_neg != 0 else None
        linhas.append({
            "codUsina": u, "nome": nome.get(u, f"Usina {u}"),
            "nMesesNeg": len(negativos),
            "qincMedioNeg": round(float(q_neg), 2),
            "usosMedioGeral": round(float(usos_medio_geral), 3),
            "usosMedioNeg": round(float(u_neg), 3),
            "coberturaPct": cobertura,
        })
    linhas.sort(key=lambda x: -(x["coberturaPct"] or -1))
    return {"lista": linhas}


# ============================================================
# 5d) CORRECAO POR CONSERVACAO DE MASSA (proposta de correcao)
# ============================================================
# Onde Qinc < 0 num mes, zera o negativo e retira o mesmo VOLUME dos meses positivos do
# mesmo ano, proporcionalmente (fator f = 1 - D/P, com D = deficit e P = volume positivo
# do ano, ambos em m3). A vazao natural corrigida e reconstruida de montante p/ jusante:
#     Qi' = SUM(Qm') + Qinc'_i
# e o volume anual de Qi' e igual ao do Qi original em todos os postos.
# Anos com D >= P (Qinc anual <= 0) nao sao corrigiveis por redistribuicao (o problema e de
# volume, nao de distribuicao mensal) -> ficam como estao e sao listados.
# Usinas "estruturais" (Classe 1 por padrao) nao sao alteradas: Qi' = Qi original.
# Usinas com Qinc = 0 (posto igual ao de montante) passam a acompanhar a montante corrigida.
def ordem_topologica(usinas: pd.DataFrame) -> list:
    """Usinas ordenadas de montante p/ jusante (cada montante antes da sua jusante)."""
    mont = montante_map(usinas)
    cods = sorted(int(c) for c in usinas["CodUsina"])
    ordem, visto, pilha = [], set(), set()

    def visita(u):
        if u in visto:
            return
        if u in pilha:
            log(f"AVISO: ciclo na topologia envolvendo a usina {u} — ordem parcial.")
            return
        pilha.add(u)
        for m in sorted(mont.get(u, [])):
            visita(m)
        pilha.discard(u)
        visto.add(u)
        ordem.append(u)

    for u in cods:
        visita(u)
    return ordem


def redistribuir_ano(q: pd.DataFrame):
    """Zera Qinc < 0 e retira o deficit dos meses positivos do mesmo ano (proporcional ao volume).
    Retorna (Qinc corrigido inteiro, info)."""
    seg = np.array(SEG_MES, dtype=float)
    vals = q.values.astype(float).copy()
    anos_corr, anos_nc = [], []
    meses_corr = 0
    vol_def = vol_pos = 0.0
    for k, ano in enumerate(q.index):
        v = vals[k]
        ok = ~np.isnan(v)
        neg = ok & (v < 0)
        if not neg.any():
            continue
        pos = ok & (v > 0)
        D = float(np.sum(-v[neg] * seg[neg]))
        P = float(np.sum(v[pos] * seg[pos]))
        if P <= D:
            anos_nc.append(int(ano))
            continue
        f = 1.0 - D / P
        v[neg] = 0.0
        v[pos] = v[pos] * f
        vals[k] = v
        anos_corr.append(int(ano))
        meses_corr += int(neg.sum())
        vol_def += D
        vol_pos += P
    # arredondamento p/ inteiro (formato do deck). Nos anos corrigidos usa arredondamento
    # acumulado: preserva a soma anual (erro < 0,5 m3/s no total do ano) e nunca cria negativo.
    for k, ano in enumerate(q.index):
        v = vals[k]
        ok = ~np.isnan(v)
        if int(ano) in anos_corr:
            cum = np.round(np.cumsum(v[ok]))
            v[ok] = np.diff(np.concatenate([[0.0], cum]))
        else:
            v[ok] = np.round(v[ok])
        vals[k] = v
    qc = pd.DataFrame(vals, index=q.index, columns=q.columns)
    delta = (qc - q).abs().values
    info = {
        "mesesCorrigidos": meses_corr,
        "anosCorrigidos": anos_corr,
        "anosNaoCorrigiveis": anos_nc,
        "volRedistPct": round(vol_def / vol_pos * 100, 2) if vol_pos else 0.0,
        "maxDelta": float(np.nanmax(delta)) if np.isfinite(delta).any() else 0.0,
    }
    return qc, info


def corrigir_massa(usinas: pd.DataFrame, series: dict, qinc: dict,
                   classes_estruturais=(1,), excluir_especiais=False) -> dict:
    """Aplica a correcao de montante p/ jusante. Retorna nat_corr por usina, Qinc corrigido e info."""
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}
    mont = montante_map(usinas)
    nat_corr, qinc_corr, info = {}, {}, {}

    for u in ordem_topologica(usinas):
        p = u2p.get(u)
        if p is None or p not in series or u not in qinc:
            continue
        qi = series[p]
        soma = None
        for m in sorted(set(mont.get(u, []))):
            if m in nat_corr:
                s = nat_corr[m].reindex(qi.index).fillna(0)
                soma = s if soma is None else soma + s
        if soma is None:
            soma = pd.DataFrame(0.0, index=qi.index, columns=qi.columns)

        # Qinc de referencia = Qinc original (Qi - SUM(Qm) com as series oficiais). Corrigir o
        # incremental original (e nao Qi - SUM(Qm')) evita que a redistribuicao mensal feita a
        # montante "vaze" para as incrementais de jusante; a natural corrigida e a acumulacao.
        orig = qinc[u].reindex(qi.index)
        neg_orig, tot = n_neg(orig)
        pct = neg_orig / tot * 100 if tot else 0.0
        classe = classe_por_pct(pct)
        vals_o = orig.values[~np.isnan(orig.values)]
        nulo = bool(vals_o.size) and bool(np.all(vals_o == 0))
        especial = eh_caso_especial(nome.get(u, ""))

        base = {"negOrig": neg_orig, "negAntes": neg_orig, "classe": classe,
                "mesesCorrigidos": 0, "anosCorrigidos": [], "anosNaoCorrigiveis": [],
                "volRedistPct": 0.0, "maxDelta": 0.0}

        if (classe in classes_estruturais) or (excluir_especiais and especial and not nulo):
            # nao altera a natural; o incremental passa a ser medido contra a montante corrigida
            nat_corr[u] = qi.copy()
            qinc_corr[u] = (qi - soma).where(qi.notna())
            base.update(status="estrutural", negDepois=n_neg(qinc_corr[u])[0])
        else:
            if nulo or neg_orig == 0:
                qc, inf = orig.copy(), None
            else:
                qc, inf = redistribuir_ano(orig)
                qc = qc.where(qi.notna())
            nat_corr[u] = (soma + qc).where(qi.notna())
            qinc_corr[u] = qc
            if inf:
                base.update(inf)
            base["negDepois"] = n_neg(qc)[0]
            if nulo:
                base["status"] = "acompanha montante"
            elif neg_orig == 0:
                base["status"] = "sem negativos"
            else:
                base["status"] = "corrigida" if not inf["anosNaoCorrigiveis"] else "corrigida parcial"
        base["casoEspecial"] = especial
        info[u] = base

    # validacao: volume anual preservado em cada posto (anos completos)
    # media anual ponderada pelos segundos de cada mes = volume anual / segundos do ano
    seg = np.array(SEG_MES, dtype=float)
    vmed = lambda df: pd.Series((df.values * seg).sum(axis=1) / seg.sum(), index=df.index)
    dif_max = 0.0
    for u, nc in nat_corr.items():
        p = u2p[u]
        a0 = vmed(series[p].dropna())
        a1 = vmed(nc.dropna()).reindex(a0.index)
        d = float((a1 - a0).abs().max()) if len(a0) else 0.0
        info[u]["difVolAnualMax"] = round(d, 3)
        dif_max = max(dif_max, d)

    n_orig = sum(i["negOrig"] for i in info.values())
    n_dep = sum(i["negDepois"] for i in info.values())
    st = defaultdict(int)
    for i in info.values():
        st[i["status"]] += 1
    resumo = {
        "negOrig": n_orig, "negDepois": n_dep,
        "status": dict(st),
        "anosNaoCorrigiveis": sum(len(i["anosNaoCorrigiveis"]) for i in info.values()),
        "difVolAnualMax": round(dif_max, 3),
        "classesEstruturais": list(classes_estruturais),
        "excluiEspeciais": bool(excluir_especiais),
    }
    log(f"Correção por conservação de massa: meses com Qinc<0 {n_orig} -> {n_dep}; "
        f"status {dict(st)}; maior diferença de volume anual por posto (em vazão média) {dif_max:.3f} m³/s.")
    return {"natCorr": nat_corr, "qincCorr": qinc_corr, "info": info, "resumo": resumo}


def deck_corrigido(series: dict, usinas: pd.DataFrame, corr: dict) -> dict:
    """{posto: DF} com a vazao natural corrigida nos postos de usinas; demais postos inalterados."""
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    out = {p: df.copy() for p, df in series.items()}
    dono = {}
    for u, nc in corr["natCorr"].items():
        p = u2p[u]
        if p in dono:
            if not np.allclose(np.nan_to_num(out[p].values), np.nan_to_num(nc.values)):
                log(f"AVISO: posto {p} compartilhado por {dono[p]} e {u} com correções diferentes — "
                    f"mantida a da usina {dono[p]}.")
            continue
        dono[p] = u
        out[p] = nc
    return out


def escrever_deck(path: Path, deck: dict):
    """Formato do deck original: I3, I5, 12 x I6 (meses ausentes = 0)."""
    linhas = []
    for p in sorted(deck):
        df = deck[p]
        for ano, row in df.iterrows():
            vals = "".join(f"{int(round(0 if pd.isna(v) else v)):6d}" for v in row.values)
            linhas.append(f"{p:3d}{int(ano):5d}{vals}")
    path.write_text("\n".join(linhas) + "\n", encoding="utf-8")


# ============================================================
# 6) MONTAGEM DO JSON
# ============================================================
def anuais(df: pd.DataFrame):
    comp = df.dropna()
    return comp.index.astype(int).tolist(), comp.mean(axis=1)


def mensal(df: pd.DataFrame):
    lab, val = [], []
    # inclui meses do ano parcial
    for ano, row in df.iterrows():
        for m in MESES:
            if pd.isna(row[m]):
                continue
            lab.append(f"{m}/{int(ano)}")
            val.append(round(float(row[m]), 2))
    return lab, val


def bloco_serie(df):
    anos, med = anuais(df)
    lab, val = mensal(df)
    return {"anos": anos, "media": [round(float(v), 2) for v in med],
            "mensalLabels": lab, "mensalValores": val}


def n_neg(df):
    a = df.values.ravel()
    a = a[~np.isnan(a)]
    return int((a < 0).sum()), int(a.size)


def classe_por_pct(pct):
    if pct is None:
        return None
    if pct > LIM_CLASSE1:
        return 1
    if pct >= LIM_CLASSE2:
        return 2
    if pct > 0:
        return 3
    return 0


def calc_resumo_geral(qinc: dict, lista: list) -> dict:
    """Numeros agregados p/ a aba Visao Geral: negativos por mes-do-ano e por decada."""
    por_mes = np.zeros(12, dtype=int)
    por_dec = defaultdict(int)
    tot_meses = 0
    for df in qinc.values():
        tot_meses += int(np.sum(~np.isnan(df.values)))
        mask = np.nan_to_num(df.values, nan=0.0) < 0
        por_mes += mask.sum(axis=0)
        for ano, linha in zip(df.index.astype(int), mask):
            por_dec[ano // 10 * 10] += int(linha.sum())
    decadas = sorted(por_dec)
    neg = int(por_mes.sum())
    classes = {str(k): 0 for k in [1, 2, 3, 0]}
    for it in lista:
        if it.get("classe") is not None:
            classes[str(it["classe"])] += 1
    return {
        "totalMeses": tot_meses,
        "neg": {"orig": neg},
        "pct": {"orig": round(neg / tot_meses * 100, 2) if tot_meses else 0},
        "porMes": {"orig": por_mes.tolist()},
        "decadas": [str(d) for d in decadas],
        "porDecada": {"orig": [por_dec.get(d, 0) for d in decadas]},
        "classes": classes,
        "limites": {"c1": LIM_CLASSE1, "c2": LIM_CLASSE2},
        "qincNulas": [it["nome"] for it in lista if it.get("qincNulo")],
    }


def r_(v, n=2):
    return None if v is None else round(float(v), n)


def montar_json(usinas, series, qinc, evap, met_evap, estac_postos, cascatas, diag_usos, correcao=None):
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}
    u2c = {}
    for i, cad in enumerate(cascatas, 1):
        for u in cad:
            u2c.setdefault(u, []).append(i)
    pu = met_evap["porUsina"] if met_evap else {}

    lista = []
    for r in usinas.sort_values("CodUsina").itertuples():
        u = int(r.CodUsina)
        p = int(r.CodPosto) if pd.notna(r.CodPosto) else None
        item = {"codUsina": u, "nome": nome[u], "codPosto": p,
                "casoEspecial": eh_caso_especial(nome[u])}

        item["natural"] = bloco_serie(series[p]) if p in series else None
        item["sazonalidade"] = calc_sazonalidade(series[p]) if p in series else None
        item["permanencia"] = calc_permanencia(series[p]) if p in series else None
        item["extremos"] = calc_extremos(series[p]) if p in series else None

        item["classe"] = None
        item["qincNulo"] = False
        if u in qinc:
            inc = bloco_serie(qinc[u])
            neg, tot = n_neg(qinc[u])
            pct = round(neg / tot * 100, 2) if tot else 0.0
            inc.update({"negativos": neg, "totalMeses": tot, "pctNegativos": pct})
            item["incremental"] = inc
            item["classe"] = classe_por_pct(pct)
            item["qincNulo"] = bool(inc["mensalValores"]) and all(v == 0 for v in inc["mensalValores"])
        else:
            item["incremental"] = None

        e = evap.get(u)
        if e:
            m = pu[u]
            item["evaporacao"] = {
                "meses": MESES,
                "min": [round(v, 4) for v in e["min"]],
                "med": [round(v, 4) for v in e["med"]],
                "max": [round(v, 4) for v in e["max"]],
                "mm": [round(v, 1) for v in e["mm"]],
                "areaMinKm2": round(e["areaMin"], 2),
                "areaMedKm2": round(e["areaMed"], 2),
                "areaMaxKm2": round(e["areaMax"], 2),
                "areaSuspeita": e["suspeita"],
                "evapMin": r_(m["evapMin"], 3), "evapMed": r_(m["evapMed"], 3), "evapMax": r_(m["evapMax"], 3),
                "volAnualHm3": r_(m["volAnualHm3"], 1), "laminaAnualMm": r_(m["laminaAnualMm"], 0),
                "qnatMedia": r_(m["qnatMedia"], 1), "qincMedia": r_(m["qincMedia"], 1),
                "pctQnat": r_(m["pctQnat"], 3), "pctQinc": r_(m["pctQinc"], 3),
                "rank": m["rank"], "rankPctQnat": m.get("rankPctQnat"),
                "terminal": m["terminal"], "terminalNome": nome.get(m["terminal"], ""),
            }
        else:
            item["evaporacao"] = None

        item["cascataIdx"] = u2c.get(u, [])
        item["estacionariedade"] = estac_postos.get(p) if p is not None else None

        ci = correcao["info"].get(u) if correcao else None
        if ci:
            blk = {k: ci[k] for k in ["status", "negOrig", "negAntes", "negDepois", "mesesCorrigidos",
                                       "volRedistPct", "maxDelta", "difVolAnualMax"]}
            blk["anosCorrigidos"] = len(ci["anosCorrigidos"])
            blk["anosNaoCorrigiveis"] = ci["anosNaoCorrigiveis"]
            qc = correcao["qincCorr"][u]
            mudou = u in qinc and not np.allclose(np.nan_to_num(qc.values), np.nan_to_num(qinc[u].values))
            blk["mudou"] = bool(mudou)
            if mudou:
                _, vals = mensal(qc)
                blk["mensalCorr"] = vals
                blk["perfilOrig"] = [round(float(x), 2) for x in qinc[u].mean(axis=0).values]
                blk["perfilCorr"] = [round(float(x), 2) for x in qc.mean(axis=0).values]
            item["correcao"] = blk
        else:
            item["correcao"] = None
        lista.append(item)

    casc_json = []
    por_cod = {x["codUsina"]: x for x in lista}
    for i, cad in enumerate(cascatas, 1):
        sn, si = [], []
        for u in cad:
            it = por_cod.get(u)
            if not it:
                continue
            if it["natural"]:
                sn.append({"codUsina": u, "nome": it["nome"],
                           "anos": it["natural"]["anos"], "media": it["natural"]["media"]})
            if it["incremental"]:
                si.append({"codUsina": u, "nome": it["nome"],
                           "anos": it["incremental"]["anos"], "media": it["incremental"]["media"],
                           "negativos": it["incremental"]["negativos"],
                           "totalMeses": it["incremental"]["totalMeses"]})
        casc_json.append({"indice": i, "usinas": cad,
                          "nomes": [nome.get(u, f"Usina {u}") for u in cad],
                          "seriesNatural": sn, "seriesIncremental": si})

    topologia = calc_topologia(usinas, cascatas, series)

    postos_usinas = {int(p) for p in usinas["CodPosto"].dropna()}
    estac_usinas = {p: e for p, e in estac_postos.items() if p in postos_usinas}

    dados = {"usinas": lista, "cascatas": casc_json,
             "estacionariedadeResumo": {"usinas": resumo_estac(estac_usinas),
                                        "todos": resumo_estac(estac_postos)},
             "topologia": topologia,
             "resumo": calc_resumo_geral(qinc, lista)}
    if met_evap:
        R = met_evap["resumo"]
        dados["evaporacao"] = {
            "nUsinas": R["nUsinas"],
            "totalMin": r_(R["totalMin"], 1), "totalMed": r_(R["totalMed"], 1), "totalMax": r_(R["totalMax"], 1),
            "volAnualKm3": r_(R["volAnualKm3"], 2), "qnatBacias": r_(R["qnatBacias"], 0),
            "pctSIN": r_(R["pctSIN"], 3),
            "relacao": R["relacao"],
            "bacias": [{**b, **{k: r_(b[k], 3) for k in ["evapMin", "evapMed", "evapMax", "pctQnat"]},
                        "volAnualHm3": r_(b["volAnualHm3"], 1), "qnatExutorio": r_(b["qnatExutorio"], 1)}
                       for b in R["bacias"]],
        }
    if diag_usos is not None:
        dados["usosConsuntivos"] = diag_usos
    if correcao is not None:
        dados["correcao"] = correcao["resumo"]
    return dados


# ============================================================
# 7) SAIDAS
# ============================================================
def salvar_intermediarios(pasta: Path, usinas, qinc, met_evap, estac_postos, cascatas, dados,
                          corr=None, series=None):
    pasta.mkdir(exist_ok=True)
    u2p = {int(r.CodUsina): int(r.CodPosto) for r in usinas.itertuples() if pd.notna(r.CodPosto)}
    nome = {int(r.CodUsina): str(r.Usina) for r in usinas.itertuples()}

    def deck_txt(tab, arq):
        linhas = []
        for u in sorted(tab, key=lambda x: u2p[x]):
            for ano, row in tab[u].iterrows():
                vals = " ".join("" if np.isnan(v) else "%.0f" % v for v in row.values)
                linhas.append(f"{u2p[u]} {int(ano)} {vals}")
        (pasta / arq).write_text("\n".join(linhas) + "\n", encoding="utf-8")

    deck_txt(qinc, "vazoes_incrementais_por_posto.txt")

    # ---- evaporacao
    if met_evap:
        rows = []
        for it in dados["usinas"]:
            e = it["evaporacao"]
            if not e:
                continue
            r = {"Rank": e["rank"], "CodUsina": it["codUsina"], "Usina": it["nome"], "CodPosto": it["codPosto"],
                 "Bacia_ate": e["terminalNome"],
                 "AreaMin_km2": e["areaMinKm2"], "AreaMed_km2": e["areaMedKm2"], "AreaMax_km2": e["areaMaxKm2"],
                 "LaminaLiquidaAnual_mm": e["laminaAnualMm"],
                 "EvapMin_m3s": e["evapMin"], "EvapMed_m3s": e["evapMed"], "EvapMax_m3s": e["evapMax"],
                 "VolAnual_hm3": e["volAnualHm3"],
                 "Qnat_media_m3s": e["qnatMedia"], "Evap_sobre_Qnat_pct": e["pctQnat"],
                 "Qinc_media_m3s": e["qincMedia"], "Evap_sobre_Qinc_pct": e["pctQinc"],
                 "Pct_meses_Qinc_neg": it["incremental"]["pctNegativos"] if it["incremental"] else None,
                 "Classe": it["classe"], "AreaSuspeita": e["areaSuspeita"]}
            for m, v in zip(MESES, e["med"]):
                r[f"EvapMed_{m}_m3s"] = v
            rows.append(r)
        pd.DataFrame(rows).sort_values("Rank").to_csv(pasta / "evaporacao_por_usina.csv", index=False)
        pd.DataFrame([{"Bacia_ate": b["nome"], "CodUsinaJusanteFinal": b["terminal"], "nUsinas": b["nUsinas"],
                       "nComEvap": b["nComEvap"], "EvapMin_m3s": b["evapMin"], "EvapMed_m3s": b["evapMed"],
                       "EvapMax_m3s": b["evapMax"], "VolAnual_hm3": b["volAnualHm3"],
                       "Qnat_exutorio_m3s": b["qnatExutorio"], "Evap_sobre_Qnat_pct": b["pctQnat"],
                       "Principais": ", ".join(b["principais"])}
                      for b in dados["evaporacao"]["bacias"]]).to_csv(pasta / "evaporacao_por_bacia.csv", index=False)

    with open(pasta / "cascatas_por_usina.txt", "w", encoding="utf-8") as f:
        for i, cad in enumerate(cascatas, 1):
            f.write(f"[{i:03}] " + "-".join(map(str, cad)) + "\n")

    postos_usinas = {int(p) for p in usinas["CodPosto"].dropna()}
    est = [{"CodPosto": p, "posto_de_usina": p in postos_usinas, "n_anos": len(e["anos"]),
            "r1": e["r1"], "r1_sig": e["r1Sig"],
            "sen_slope": e["senSlope"], "mk_bruto_p": e["mkPBruto"], "mk_bruto_trend": e["mkTrendBruto"],
            "mk_corr_p": e["mkPCorr"], "mk_corr_trend": e["mkTrendCorr"],
            "pettitt_ano": e["pettittAno"], "pettitt_p": e["pettittP"]}
           for p, e in sorted(estac_postos.items())]
    pd.DataFrame(est).to_csv(pasta / "resultado_estacionariedade_corrigido.csv", index=False)

    diag = []
    for it in dados["usinas"]:
        if not it["incremental"]:
            continue
        diag.append({"CodUsina": it["codUsina"], "Usina": it["nome"], "CodPosto": it["codPosto"],
                     "Classe": it["classe"], "totalMeses": it["incremental"]["totalMeses"],
                     "neg": it["incremental"]["negativos"], "pct": it["incremental"]["pctNegativos"]})
    df_diag = pd.DataFrame(diag)
    if not df_diag.empty:
        df_diag = df_diag.sort_values("pct", ascending=False)
    df_diag.to_csv(pasta / "diagnostico_negativos.csv", index=False)

    if dados.get("topologia"):
        pd.DataFrame(dados["topologia"]["lista"]).to_csv(pasta / "topologia_isolada.csv", index=False)
    if dados.get("usosConsuntivos"):
        pd.DataFrame(dados["usosConsuntivos"]["lista"]).to_csv(pasta / "usos_consuntivos_diagnostico.csv", index=False)

    if corr is not None and series is not None:
        escrever_deck(pasta / "vazoes_corrigidas_massa.txt", deck_corrigido(series, usinas, corr))
        deck_txt(corr["qincCorr"], "vazoes_incrementais_corrigidas_massa.txt")
        rows = []
        for u, i in sorted(corr["info"].items()):
            rows.append({"CodUsina": u, "Usina": nome.get(u), "CodPosto": u2p.get(u), "Classe": i["classe"],
                         "Status": i["status"], "neg_Orig": i["negOrig"],
                         "neg_Depois": i["negDepois"], "meses_corrigidos": i["mesesCorrigidos"],
                         "anos_corrigidos": len(i["anosCorrigidos"]),
                         "anos_nao_corrigiveis": " ".join(map(str, i["anosNaoCorrigiveis"])),
                         "vol_redistribuido_pct": i["volRedistPct"], "max_delta_m3s": i["maxDelta"],
                         "dif_vazao_media_anual_max_m3s": i["difVolAnualMax"]})
        pd.DataFrame(rows).to_csv(pasta / "correcao_massa_por_usina.csv", index=False)

    (pasta / "log.txt").write_text("\n".join(LOG) + "\n", encoding="utf-8")


def gerar_analise(dados: dict) -> list:
    """Secao de analise automatica (texto) p/ o diagnostico."""
    U = dados["usinas"]
    RU = dados["estacionariedadeResumo"]["usinas"]
    RT = dados["estacionariedadeResumo"]["todos"]
    RG = dados["resumo"]
    L = []
    top = 10

    def nome(u):
        return f"{u['nome']} ({u['codUsina']})"

    # ---------------- Qinc
    com_inc = [u for u in U if u["incremental"]]
    tot = RG["totalMeses"]
    neg = RG["neg"]["orig"]
    afet = [u for u in com_inc if u["incremental"]["negativos"] > 0]
    L += ["A) VAZOES INCREMENTAIS", ""]
    L.append(f"- {neg} de {tot} meses com Qinc < 0 ({RG['pct']['orig']:.2f} %).")
    L.append(f"- {len(afet)} de {len(com_inc)} usinas ({len(afet) / max(len(com_inc), 1) * 100:.1f} %) "
             f"tem ao menos 1 mes negativo.")
    c = RG["classes"]
    L.append(f"- Classes: 1 (> {LIM_CLASSE1:.0f} %): {c['1']} | 2 ({LIM_CLASSE2:.0f}-{LIM_CLASSE1:.0f} %): {c['2']} "
             f"| 3 (< {LIM_CLASSE2:.0f} %): {c['3']} | sem negativos: {c['0']}.")
    if RG["qincNulas"]:
        L.append(f"- Qinc identicamente nulo (posto igual ao de montante): {', '.join(RG['qincNulas'])}.")
    if neg:
        pm = list(zip(MESES, RG["porMes"]["orig"]))
        ordm = sorted(pm, key=lambda x: -x[1])
        L.append("- Meses com mais negativos: " +
                 ", ".join(f"{m} {n} ({n / neg * 100:.0f} %)" for m, n in ordm[:4]) + ".")
        L.append("- Negativos por decada: " +
                 ", ".join(f"{d}s: {n}" for d, n in zip(RG["decadas"], RG["porDecada"]["orig"])) + ".")
    rk = sorted(com_inc, key=lambda u: -u["incremental"]["pctNegativos"])[:top]
    L.append(f"- Top {top} usinas por % de meses negativos:")
    for u in rk:
        i = u["incremental"]
        L.append(f"    {nome(u):<28} {i['negativos']:>5} meses ({i['pctNegativos']:.2f} %)  classe {u['classe']}")
    L.append("")

    # ---------------- Evaporacao (perdas; ja embutida na vazao natural oficial)
    E = dados.get("evaporacao")
    L += ["B) EVAPORACAO LIQUIDA - PERDAS (diagnostico; ja embutida na vazao natural oficial, "
          "nao entra no Qinc)", ""]
    if E:
        com_ev = [u for u in U if u["evaporacao"]]
        L.append(f"- {E['nUsinas']} usinas: evaporacao total {E['totalMed']:.0f} m3/s (Area Med; faixa "
                 f"{E['totalMin']:.0f}-{E['totalMax']:.0f}) = {E['volAnualKm3']:.1f} km3/ano"
                 + (f" = {E['pctSIN']:.2f} % da vazao natural das bacias." if E['pctSIN'] else "."))
        rk = sorted(com_ev, key=lambda u: u["evaporacao"]["rank"])
        acum, n80 = 0.0, None
        for k, u in enumerate(rk, 1):
            acum += u["evaporacao"]["evapMed"]
            if n80 is None and acum >= 0.8 * E["totalMed"]:
                n80 = k
        if n80:
            L.append(f"- {n80} usinas respondem por 80 % da evaporacao total.")
        L.append(f"- Top {top} usinas por perda absoluta (m3/s, Area Med):")
        for u in rk[:top]:
            e = u["evaporacao"]
            pq = f"{e['pctQnat']:.2f} % da Qnat" if e["pctQnat"] is not None else "-"
            L.append(f"    {e['rank']:>3}. {nome(u):<28} {e['evapMed']:8.2f} m3/s  ({e['evapMin']:.1f}-{e['evapMax']:.1f})"
                     f"  {e['volAnualHm3']:7.0f} hm3/ano  {pq}")
        rk2 = sorted([u for u in com_ev if u["evaporacao"]["pctQnat"] is not None],
                     key=lambda u: -u["evaporacao"]["pctQnat"])[:top]
        L.append(f"- Top {top} usinas por perda relativa (Evap / Qnat do posto):")
        for u in rk2:
            e = u["evaporacao"]
            L.append(f"    {nome(u):<28} {e['pctQnat']:6.2f} %  (evap {e['evapMed']:.2f} m3/s, Qnat {e['qnatMedia']:.1f} m3/s,"
                     f" lamina {e['laminaAnualMm']:.0f} mm/ano)")
        L.append("- Perdas por bacia (soma das usinas / Qnat na usina de jusante final):")
        for b in E["bacias"][:top]:
            pq = f"{b['pctQnat']:.2f} %" if b["pctQnat"] is not None else "-"
            L.append(f"    ate {b['nome']:<18} {b['nUsinas']:>3} usinas  {b['evapMed']:8.1f} m3/s  {pq:>8}  "
                     f"(principais: {', '.join(b['principais'])})")
        rel = E["relacao"]
        L.append("- Relacao com Qinc < 0 (Spearman, descritiva, sem Qinc = 0):")
        for chave, rot in [("pctQinc", "Evap/Qinc x % meses Qinc<0"), ("pctQnat", "Evap/Qnat x % meses Qinc<0")]:
            r = rel.get(chave)
            if r:
                L.append(f"    {rot:<28} rho = {r['rho']:+.3f}  p = {r['p']:.2e}  n = {r['n']}")
        cl12 = [u for u in com_ev if u["classe"] in (1, 2)]
        if cl12:
            L.append("- Usinas das classes 1 e 2 e sua posicao nos rankings de evaporacao:")
            for u in sorted(cl12, key=lambda u: u["evaporacao"]["rank"]):
                e = u["evaporacao"]
                pqi = f"{e['pctQinc']:.1f} %" if e["pctQinc"] is not None else "Qinc medio <= 0"
                L.append(f"    {nome(u):<28} classe {u['classe']}  rank m3/s {e['rank']:>3}  "
                         f"rank %Qnat {e['rankPctQnat'] or '-':>3}  Evap/Qinc {pqi}")
    else:
        L.append("- Sem dados de evaporacao.")
    L.append("")

    # ---------------- Estacionariedade
    L += ["C) ESTACIONARIEDADE (vazao natural anual, MK corrigido + Pettitt)", ""]
    for rot, R in [("postos de usinas", RU), ("todos os postos do deck", RT)]:
        L.append(f"- [{rot}] {R['nSigCorr']} de {R['nTotal']} ({R['pctSigCorr']} %) com tendencia significativa: "
                 f"{R['nCresc']} crescentes, {R['nDecr']} decrescentes "
                 f"(~{R['fpEsperados']} falsos positivos esperados a 5 %).")
        L.append(f"  {R['nBreakSig']} ({R['pctBreakSig']} %) com quebra significativa (Pettitt); "
                 f"{R['nBreakSemTend']} delas sem tendencia MK significativa.")
    if RU["decadaValores"]:
        i = int(np.argmax(RU["decadaValores"]))
        L.append(f"- Decada com mais quebras (usinas): {RU['decadaLabels'][i]}s ({RU['decadaValores'][i]} postos).")
    L.append(f"- {RU['nR1Sig']} postos de usinas ({RU['pctR1Sig']} %) com autocorrelacao lag-1 significativa "
             f"(corrigida por Yue-Wang).")
    vistos, est = set(), []
    for u in U:
        s = u["estacionariedade"]
        if s and u["codPosto"] not in vistos:
            vistos.add(u["codPosto"])
            est.append(u)
    mud = sum(1 for u in est if (u["estacionariedade"]["mkPBruto"] < ALPHA) != (u["estacionariedade"]["mkPCorr"] < ALPHA))
    L.append(f"- Entre os {len(est)} postos das usinas, {mud} mudam de veredito com a correcao de autocorrelacao.")

    def rel_slope(u):
        s = u["estacionariedade"]
        m = sum(s["natural"]) / len(s["natural"])
        return s["senSlope"] / m * 100 if m else 0.0

    sig = [u for u in est if u["estacionariedade"]["mkPCorr"] < ALPHA]
    dec = sorted([u for u in sig if u["estacionariedade"]["mkTrendCorr"] == "decrescente"], key=rel_slope)[:top]
    cre = sorted([u for u in sig if u["estacionariedade"]["mkTrendCorr"] == "crescente"], key=lambda u: -rel_slope(u))[:top]
    for titulo, lst in [("decrescentes", dec), ("crescentes", cre)]:
        if lst:
            L.append(f"- Tendencias {titulo} mais fortes (% da media por ano):")
            for u in lst:
                s = u["estacionariedade"]
                L.append(f"    {nome(u):<28} posto {u['codPosto']:>4}  {rel_slope(u):+6.2f} %/ano  "
                         f"(Sen {s['senSlope']:+.2f} m3/s/ano, quebra {s['pettittAno']})")

    grupos = {}
    for u in U:
        s, i = u["estacionariedade"], u["incremental"]
        if s and i:
            grupos.setdefault(s["mkTrendCorr"], []).append(i["pctNegativos"])
    if grupos:
        L.append("- % media de meses com Qinc < 0, por tendencia da vazao natural da usina "
                 "(descritivo, nao implica causa):")
        for k in ["decrescente", "sem tendência sig.", "crescente"]:
            if k in grupos:
                v = grupos[k]
                L.append(f"    {k:<20} {len(v):>3} usinas  media {sum(v) / len(v):5.2f} %")
    L.append("")

    # ---------------- Topologia isolada
    topo = dados.get("topologia")
    if topo:
        L += ["D) TOPOLOGIA ISOLADA (cascatas de tamanho 1 - sem montante cadastrada)", ""]
        L.append(f"- {topo['totalIsoladas']} de {topo['totalUsinas']} usinas em cascata isolada.")
        esp = [i for i in topo["lista"] if i["casoEspecial"]]
        if esp:
            L.append("- Casos especiais citados no projeto original (ligacao fisica conhecida, "
                     "ausente no cadastro): " + ", ".join(i["nome"] for i in esp) + ".")
        L.append("")

    # ---------------- Usos consuntivos
    usos = dados.get("usosConsuntivos")
    if usos and usos["lista"]:
        L += ["E) USOS CONSUNTIVOS INCREMENTAIS (diagnostico, NAO correcao - ja embutido no Qinc oficial)", ""]
        L.append(f"- {len(usos['lista'])} usinas com meses de Qinc negativo cruzadas com a base ANA.")
        top_cob = [i for i in usos["lista"] if i["coberturaPct"] and i["coberturaPct"] >= 50][:top]
        if top_cob:
            L.append("- Top usinas por cobertura do deficit (usos cons. incremental / deficit medio):")
            for i in top_cob:
                L.append(f"    {i['nome']:<28} {i['nMesesNeg']:>4} meses neg.  "
                         f"cobertura {i['coberturaPct']:>6.1f} %")
        L.append("")

    return L


def salvar_diagnostico_txt(path: Path, dados: dict, f_vaz: Path):
    """Tabelas por usina: negativos, evaporacao, estacionariedade, correcao + analise."""
    U = dados["usinas"]
    RG = dados["resumo"]
    E = dados.get("evaporacao")
    L = []
    sep = "=" * 124
    sub = "-" * 124

    def pct(a, b):
        return f"{a / b * 100:6.2f}" if b else "     -"

    def f_(v, fmt, w):
        return f"{'-':>{w}}" if v is None else format(v, f"{w}{fmt}")

    # ---- resumo geral
    tot = RG["totalMeses"]
    L += [sep, "DIAGNOSTICO - VAZOES INCREMENTAIS, EVAPORACAO E ESTACIONARIEDADE", sep,
          f"Deck: {f_vaz}",
          f"Usinas: {len(U)} | com Qinc: {sum(1 for u in U if u['incremental'])} "
          f"| com evap: {sum(1 for u in U if u['evaporacao'])}",
          f"Total de meses (Qinc): {tot}", "",
          f"Meses com Qinc < 0: {RG['neg']['orig']} ({pct(RG['neg']['orig'], tot).strip()} %)"]
    if E:
        L.append(f"Evaporacao liquida total (Area Med): {E['totalMed']:.1f} m3/s "
                 f"(faixa {E['totalMin']:.1f}-{E['totalMax']:.1f}) = {E['volAnualKm3']:.2f} km3/ano"
                 + (f" = {E['pctSIN']:.2f} % da vazao natural das bacias" if E['pctSIN'] else ""))
    for rot, key in [("postos de usinas", "usinas"), ("todos os postos do deck", "todos")]:
        R = dados["estacionariedadeResumo"][key]
        L += ["",
              f"Estacionariedade ({rot}, >= 30 anos):",
              f"  Postos analisados:              {R['nTotal']}",
              f"  Tendencia sig. (MK corrigido):  {R['nSigCorr']} ({R['pctSigCorr']} %)"
              f"  | crescente {R['nCresc']} | decrescente {R['nDecr']}"
              f"  | falsos positivos esperados ~{R['fpEsperados']}",
              f"  Quebra sig. (Pettitt):          {R['nBreakSig']} ({R['pctBreakSig']} %)"
              f"  | sem tendencia MK: {R['nBreakSemTend']}",
              f"  Autocorrelacao sig. (Yue-Wang): {R['nR1Sig']} ({R['pctR1Sig']} %)",
              "  Quebras por decada: " + ", ".join(f"{d}: {v}" for d, v in zip(R["decadaLabels"], R["decadaValores"]))]
    L.append("")

    # ---- tabela 1: negativos
    L += [sep, "1) MESES COM Qinc NEGATIVO POR USINA  (Qinc = Qi - SUM(Qm))", sep]
    L.append(f"{'Cod':>4} {'Usina':<14} {'Posto':>5} {'Cl':>2} {'Meses':>6} | {'Neg':>6} {'%':>6}  Obs")
    L.append(sub)
    for u in U:
        inc = u["incremental"]
        if not inc:
            L.append(f"{u['codUsina']:>4} {u['nome'][:14]:<14} {str(u['codPosto'] or '-'):>5}      sem Qinc")
            continue
        t = inc["totalMeses"]
        cl = "-" if u["classe"] is None else str(u["classe"])
        obs = []
        if u["qincNulo"]:
            obs.append("Qinc=0")
        if u["casoEspecial"]:
            obs.append("caso especial")
        L.append(f"{u['codUsina']:>4} {u['nome'][:14]:<14} {u['codPosto']:>5} {cl:>2} {t:>6} | "
                 f"{inc['negativos']:>6} {pct(inc['negativos'], t)}" + ("  " + ", ".join(obs) if obs else ""))
    L.append("")

    # ---- tabela 2: evaporacao
    L += [sep, "2) EVAPORACAO LIQUIDA POR USINA  (perdas; ja embutida na vazao natural oficial - NAO entra no Qinc)", sep]
    L.append(f"{'Rk':>3} {'Cod':>4} {'Usina':<14} | {'AreaMed':>8} {'Lamina':>6} | {'EvMin':>7} {'EvMed':>7} {'EvMax':>7} "
             f"{'hm3/ano':>8} | {'Qnat':>8} {'Ev/Qnat%':>8} | {'Qinc':>8} {'Ev/Qinc%':>8} | {'%Qinc<0':>7}  Obs")
    L.append(f"{'':>3} {'':>4} {'':<14} | {'(km2)':>8} {'(mm/a)':>6} | {'(m3/s)':>7} {'':>7} {'':>7} {'':>8} | "
             f"{'(m3/s)':>8} {'':>8} | {'(m3/s)':>8}")
    L.append(sub)
    for u in sorted([u for u in U if u["evaporacao"]], key=lambda u: u["evaporacao"]["rank"]):
        e = u["evaporacao"]
        pn = u["incremental"]["pctNegativos"] if u["incremental"] else None
        obs = []
        if e["areaSuspeita"]:
            obs.append("AREA SUSPEITA")
        if e["areaMinKm2"] == e["areaMaxKm2"]:
            obs.append("area constante")
        if u["qincNulo"]:
            obs.append("Qinc=0")
        L.append(f"{e['rank']:>3} {u['codUsina']:>4} {u['nome'][:14]:<14} | {e['areaMedKm2']:>8.1f} {e['laminaAnualMm']:>6.0f} | "
                 f"{e['evapMin']:>7.2f} {e['evapMed']:>7.2f} {e['evapMax']:>7.2f} {e['volAnualHm3']:>8.0f} | "
                 f"{f_(e['qnatMedia'], '.1f', 8)} {f_(e['pctQnat'], '.3f', 8)} | "
                 f"{f_(e['qincMedia'], '.1f', 8)} {f_(e['pctQinc'], '.2f', 8)} | {f_(pn, '.2f', 7)}"
                 + ("  " + ", ".join(obs) if obs else ""))
    sem = [u for u in U if not u["evaporacao"]]
    if sem:
        L.append("Sem dados de evaporacao: " + ", ".join(u["nome"] for u in sem))
    if E:
        L += ["", "Perdas por bacia (usinas que drenam para a mesma usina de jusante final):",
              f"{'Bacia ate':<18} {'Usinas':>6} {'EvMed':>8} {'hm3/ano':>9} {'Qnat exut.':>10} {'Ev/Qnat%':>8}  Principais"]
        for b in E["bacias"]:
            L.append(f"{b['nome'][:18]:<18} {b['nUsinas']:>6} {b['evapMed']:>8.2f} {b['volAnualHm3']:>9.0f} "
                     f"{f_(b['qnatExutorio'], '.1f', 10)} {f_(b['pctQnat'], '.3f', 8)}  {', '.join(b['principais'])}")
    L.append("")

    # ---- tabela 3: estacionariedade
    L += [sep, "3) ESTACIONARIEDADE POR USINA  (vazao natural anual do posto)", sep]
    L.append(f"{'Cod':>4} {'Usina':<14} {'Posto':>5} {'Anos':>9} | {'MK corrigido':<13} {'p':>9} | "
             f"{'Sen(m3/s/a)':>11} {'r1':>6} | {'Quebra':>6} {'p':>9} | {'MedPre':>8} {'MedPos':>8} {'Var%':>7}")
    L.append(sub)
    for u in U:
        s = u["estacionariedade"]
        if not s:
            L.append(f"{u['codUsina']:>4} {u['nome'][:14]:<14} {str(u['codPosto'] or '-'):>5}   serie < 30 anos")
            continue
        var = (s["mediaPos"] - s["mediaPre"]) / s["mediaPre"] * 100 if s["mediaPre"] else 0
        tr = {"crescente": "crescente", "decrescente": "decrescente"}.get(s["mkTrendCorr"], "sem tend.")
        brk = str(s["pettittAno"]) + ("*" if s["pettittP"] < ALPHA else " ")
        L.append(f"{u['codUsina']:>4} {u['nome'][:14]:<14} {u['codPosto']:>5} "
                 f"{s['anos'][0]}-{s['anos'][-1]} | {tr:<13} {s['mkPCorr']:>9.2e} | "
                 f"{s['senSlope']:>11.3f} {s['r1']:>6.2f}{'*' if s['r1Sig'] else ' '}| "
                 f"{brk:>6} {s['pettittP']:>9.2e} | {s['mediaPre']:>8.1f} {s['mediaPos']:>8.1f} {var:>7.1f}")
    L += ["", "* = significativo a 5% (Pettitt: quebra; r1: autocorrelacao corrigida via Yue-Wang)", ""]

    # ---- tabela 4: correcao por conservacao de massa
    C = dados.get("correcao")
    if C:
        L += [sep, "4) CORRECAO POR CONSERVACAO DE MASSA (redistribuicao anual do deficit)", sep,
              f"Meses com Qinc < 0: {C['negOrig']} -> {C['negDepois']}  | status: "
              + ", ".join(f"{k} {v}" for k, v in sorted(C["status"].items())),
              f"Anos nao corrigiveis (Qinc anual <= 0): {C['anosNaoCorrigiveis']}  | "
              f"maior diferenca de volume anual por posto (em vazao media): {C['difVolAnualMax']} m3/s", ""]
        L.append(f"{'Cod':>4} {'Usina':<14} {'Cl':>2} {'Status':<18} | {'NegOrig':>7} {'NegDep':>6} | "
                 f"{'MesesCorr':>9} {'AnosCorr':>8} {'AnosNC':>6} | {'VolRedist%':>10} {'MaxDelta':>9}")
        L.append(sub)
        for u in U:
            c = u.get("correcao")
            if not c or c["status"] == "sem negativos":
                continue
            cl = "-" if u["classe"] is None else str(u["classe"])
            L.append(f"{u['codUsina']:>4} {u['nome'][:14]:<14} {cl:>2} {c['status']:<18} | {c['negOrig']:>7} "
                     f"{c['negDepois']:>6} | {c['mesesCorrigidos']:>9} {c['anosCorrigidos']:>8} "
                     f"{len(c['anosNaoCorrigiveis']):>6} | {c['volRedistPct']:>10.2f} {c['maxDelta']:>9.0f}")
        L.append("")

    L += [sep, "5) ANALISE", sep, ""] + gerar_analise(dados)

    # ---- avisos do log
    av = [m for m in LOG if m.startswith(("AVISO", "INFO"))]
    if av:
        L += [sep, "AVISOS", sep] + av + [""]

    path.write_text("\n".join(L), encoding="utf-8")


# ============================================================
# 8) MAIN
# ============================================================
def achar(nome_padrao, alternativas):
    for n in [nome_padrao] + alternativas:
        if Path(n).exists():
            return Path(n)
    return Path(nome_padrao)


def main():
    ap = argparse.ArgumentParser(description="Gera a interface HTML a partir do deck de vazões.")
    ap.add_argument("vazoes_pos", nargs="?", default=None, help="deck de vazões (ex.: vazoes.txt)")
    ap.add_argument("--vazoes", default=None)
    ap.add_argument("--usinas", default=None)
    ap.add_argument("--hidroterm", default=None)
    ap.add_argument("--ana-usos", default=None,
                    help="planilha da Resolução ANA 92/2021 (usos consuntivos) — opcional, ativa a aba")
    ap.add_argument("--saida", default=None, help="padrão: interface_usinas_<anoIni>-<anoFim>.html")
    ap.add_argument("--pasta-saidas", default="saidas")
    ap.add_argument("--corr-estruturais", default="1",
                    help="classes tratadas como estruturais (não corrigidas), ex.: '1' ou '1,2'")
    ap.add_argument("--corr-excluir-especiais", action="store_true",
                    help="não corrige os casos especiais do projeto (derivação/transposição/complexos)")
    ap.add_argument("--manter-meses-zerados", action="store_true",
                    help="não trata os meses finais zerados do último ano como ausentes")
    a = ap.parse_args()

    arq = a.vazoes_pos or a.vazoes
    f_vaz = Path(arq) if arq else achar("vazoes", ["vazoes.txt", "vazoes.dat"])
    while not f_vaz.exists():
        resp = input(f"Arquivo de vazões '{f_vaz}' não encontrado. Digite o caminho (ou arraste o arquivo aqui): ")
        f_vaz = Path(resp.strip().strip('"').strip("'"))
    f_usi = Path(a.usinas) if a.usinas else achar("usinas atualizadas.xlsx", ["usinas_atualizadas.xlsx"])
    f_hid = Path(a.hidroterm) if a.hidroterm else achar("dados_hidroterm_completo.xlsx", [])
    f_usos = Path(a.ana_usos) if a.ana_usos else achar(
        "Resolucao-92-2021_Vazoes-Mensais_Series-Historicas_Usos-Consuntivos-Montante-de-AHEs_v2.xlsx", [])
    if not f_usi.exists():
        raise SystemExit(f"ERRO: Planilha de usinas não encontrada: {f_usi}")

    log(f"Deck: {f_vaz} | Usinas: {f_usi} | Hidroterm: {f_hid} | Usos consuntivos: {f_usos}")
    deck = ler_deck(f_vaz)
    ano_ini, ano_fim = int(deck["Ano"].min()), int(deck["Ano"].max())
    log(f"   {deck['CodPosto'].nunique()} postos, anos {ano_ini}–{ano_fim}")
    saida = Path(a.saida) if a.saida else Path(f"interface_usinas_{ano_ini}-{ano_fim}.html")
    if not a.manter_meses_zerados:
        deck = marcar_meses_futuros(deck)
    usinas = ler_usinas(f_usi)
    series = matriz_posto(deck)

    qinc = calcular_incrementais(usinas, series)
    log(f"Qinc calculado p/ {len(qinc)} usinas")

    if f_hid.exists():
        evap = calcular_evap(*ler_hidroterm(f_hid))
        met_evap = calc_metricas_evap(usinas, series, qinc, evap) if evap else None
    else:
        evap, met_evap = {}, None
        log(f"AVISO: {f_hid} não encontrado — interface sai sem evaporação.")

    estac = {}
    for p, df in series.items():
        r = estacionariedade_posto(df)
        if r:
            estac[p] = r
    log(f"Estacionariedade: {len(estac)} postos com ≥{MIN_ANOS_ESTAC} anos completos "
        f"(janelas móveis: {JANELAS_MOVEIS} anos)")

    cascatas = montar_cascatas(usinas)

    diag_usos = None
    if f_usos.exists():
        try:
            ana_df = ler_ana_usos(f_usos)
            series_usos = montar_series_usos(usinas, ana_df)
            usos_inc = calc_usos_incremental(usinas, series_usos)
            diag_usos = calc_diagnostico_usos(usinas, qinc, usos_inc)
            log(f"Usos consuntivos: diagnóstico calculado p/ {len(diag_usos['lista'])} usinas "
                f"com meses de Qinc negativo")
        except Exception as ex:
            log(f"AVISO: falha ao processar {f_usos} — interface sai sem aba de usos consuntivos ({ex}).")
    else:
        log(f"AVISO: {f_usos} não encontrado — interface sai sem aba de usos consuntivos (use --ana-usos).")

    cl_estr = tuple(int(x) for x in a.corr_estruturais.split(",") if x.strip())
    corr = corrigir_massa(usinas, series, qinc, classes_estruturais=cl_estr,
                          excluir_especiais=a.corr_excluir_especiais)

    dados = montar_json(usinas, series, qinc, evap, met_evap, estac, cascatas, diag_usos, corr)
    dados["meta"] = {"deck": f_vaz.name, "anoIni": ano_ini, "anoFim": ano_fim}

    html = TEMPLATE.replace("__APP_DATA__", json.dumps(dados, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
                            .replace("</", "<\\/"))
    saida.write_text(html, encoding="utf-8")
    salvar_intermediarios(Path(a.pasta_saidas), usinas, qinc, met_evap, estac, cascatas, dados, corr, series)
    salvar_diagnostico_txt(Path(a.pasta_saidas) / "diagnostico.txt", dados, f_vaz)
    log(f"OK: {saida} ({len(html)/1e6:.1f} MB) + intermediários em {a.pasta_saidas}/")


# ============================================================
# TEMPLATE HTML (dados em __APP_DATA__)
# ============================================================
TEMPLATE = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Vazões Incrementais &amp; Evaporação — SIN</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  :root{
    --bg-0:#0b1620; --bg-1:#0f1e2b; --bg-2:#132537; --bg-3:#1a3145;
    --line:#24405a; --line-soft:#1c3348;
    --text-hi:#eaf2f6; --text-mid:#9db3c4; --text-dim:#5f7a90;
    --cyan:#4fd8e8; --cyan-dim:#2b8ea3; --amber:#f2b84b; --coral:#f2765b; --green:#6fd6a0;
    --mono:'IBM Plex Mono','SFMono-Regular',Consolas,monospace;
    --sans:'Inter',-apple-system,'Segoe UI',sans-serif;
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;background:var(--bg-0);color:var(--text-hi);font-family:var(--sans);-webkit-font-smoothing:antialiased;}
  body{
    background-image:
      radial-gradient(ellipse 900px 500px at 15% -10%, rgba(79,216,232,0.07), transparent),
      radial-gradient(ellipse 700px 500px at 100% 10%, rgba(242,184,75,0.05), transparent);
    min-height:100vh;
  }
  .app{max-width:1320px;margin:0 auto;padding:20px 24px 60px;}

  header.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;padding-bottom:16px;margin-bottom:16px;border-bottom:1px solid var(--line-soft);flex-wrap:wrap;}
  .brand{display:flex;flex-direction:column;gap:4px;}
  .brand h1{margin:0;font-size:28px;font-weight:700;letter-spacing:-.01em;color:var(--text-hi);}
  .brand .meta{font-family:var(--mono);font-size:14px;color:var(--text-dim);}
  .header-right{display:flex;align-items:center;gap:14px;}
  .search-box{position:relative;}
  .search-box input{background:var(--bg-2);border:1px solid var(--line);color:var(--text-hi);font-family:var(--sans);font-size:17px;padding:8px 30px 8px 32px;border-radius:8px;width:260px;outline:none;transition:border-color .15s;}
  .search-box input:focus{border-color:var(--cyan-dim);}
  .search-box svg{position:absolute;left:10px;top:50%;transform:translateY(-50%);opacity:.5;}
  .search-box kbd{position:absolute;right:8px;top:50%;transform:translateY(-50%);font-family:var(--mono);font-size:12px;color:var(--text-dim);border:1px solid var(--line);border-radius:4px;padding:0 5px;}
  .search-results{position:absolute;top:calc(100% + 6px);left:0;right:0;background:var(--bg-2);border:1px solid var(--line);border-radius:8px;max-height:320px;overflow-y:auto;z-index:50;display:none;box-shadow:0 12px 32px rgba(0,0,0,.4);}
  .search-results.show{display:block;}
  .search-results .item{padding:12px 16px;font-size:17px;cursor:pointer;display:flex;justify-content:space-between;gap:10px;border-bottom:1px solid var(--line-soft);}
  .search-results .item:last-child{border-bottom:none;}
  .search-results .item:hover,.search-results .item.sel{background:var(--bg-3);}
  .search-results .item .code{font-family:var(--mono);color:var(--text-dim);font-size:14px;white-space:nowrap;}

  .tabbar{display:flex;gap:6px;margin-bottom:18px;background:var(--bg-2);border:1px solid var(--line);border-radius:11px;padding:5px;width:fit-content;flex-wrap:wrap;}
  .tabbar button{
    background:transparent;border:none;color:var(--text-mid);font-family:var(--sans);font-weight:600;
    font-size:17px;padding:10px 18px;border-radius:8px;cursor:pointer;transition:all .15s;
    display:flex;align-items:center;gap:8px;
  }
  .tabbar button svg{width:16px;height:16px;}
  .tabbar button.active{background:var(--cyan);color:#04222b;}
  .tabbar button:not(.active):hover{color:var(--text-hi);}
  button:focus:not(:focus-visible){outline:none;}
  .view{display:none;}
  .view.active{display:block;}

  .nav-strip{display:flex;align-items:center;gap:18px;background:linear-gradient(180deg,var(--bg-2),var(--bg-1));border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:22px;}
  .nav-btn{flex:0 0 auto;width:46px;height:46px;border-radius:10px;background:var(--bg-3);border:1px solid var(--line);color:var(--text-hi);display:flex;align-items:center;justify-content:center;cursor:pointer;transition:background .15s,border-color .15s,transform .1s;}
  .nav-btn:hover{background:var(--cyan-dim);border-color:var(--cyan);color:#04222b;}
  .nav-btn:active{transform:scale(.94);}
  .nav-btn:disabled{opacity:.3;cursor:not-allowed;pointer-events:none;}
  .nav-btn svg{width:20px;height:20px;}

  .identity{flex:1;min-width:0;}
  .identity .top-row{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;}
  .identity h2{margin:0;font-size:38px;font-weight:700;letter-spacing:-.02em;color:var(--text-hi);}
  .tag{font-family:var(--mono);font-size:15px;padding:5px 12px;border-radius:5px;border:1px solid var(--line);color:var(--text-mid);white-space:nowrap;background:transparent;cursor:default;}
  .tag.cyan{color:var(--cyan);border-color:var(--cyan-dim);}
  .identity .sub{margin-top:8px;font-size:17px;color:var(--text-dim);font-family:var(--mono);}

  .position-counter{flex:0 0 auto;text-align:right;font-family:var(--mono);color:var(--text-mid);font-size:17px;}
  .position-counter .big{color:var(--cyan);font-weight:600;font-size:22px;}
  .progress-track{width:140px;height:4px;background:var(--bg-3);border-radius:2px;margin-top:8px;overflow:hidden;}
  .progress-fill{height:100%;background:var(--cyan);border-radius:2px;transition:width .2s;}

  .badges{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px;}
  .badge{display:flex;align-items:center;gap:8px;background:var(--bg-2);border:1px solid var(--line);border-radius:10px;padding:12px 18px;font-size:17px;color:var(--text-mid);}
  .badge.clickable{cursor:pointer;transition:border-color .15s;}
  .badge.clickable:hover{border-color:var(--cyan-dim);}
  .badge.selected{border-color:var(--cyan);box-shadow:0 0 0 1px var(--cyan) inset;}
  .badge.critical{border-color:#7a3a2f;}
  .badge .dot{width:7px;height:7px;border-radius:50%;flex:0 0 auto;}
  .badge b{color:var(--text-hi);font-family:var(--mono);}
  .badge.warn b{color:var(--coral);}
  .badge.ok b{color:var(--green);}

  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:18px;}
  .kpi{background:var(--bg-1);border:1px solid var(--line);border-radius:12px;padding:14px 16px;}
  .kpi .lbl{font-size:14px;color:var(--text-dim);font-family:var(--mono);text-transform:uppercase;letter-spacing:.04em;}
  .kpi .val{font-size:30px;font-weight:700;font-family:var(--mono);margin-top:6px;color:var(--text-hi);}
  .kpi .sub{font-size:14px;color:var(--text-mid);margin-top:4px;font-family:var(--mono);}
  .kpi.warn .val{color:var(--coral);} .kpi.amber .val{color:var(--amber);} .kpi.cyan .val{color:var(--cyan);}

  .grid{display:grid;grid-template-columns:repeat(2,1fr);gap:18px;}
  @media (max-width:920px){.grid{grid-template-columns:1fr;} .identity h2{font-size:28px;}}

  .card{background:var(--bg-1);border:1px solid var(--line);border-radius:14px;padding:18px 18px 12px;display:flex;flex-direction:column;}
  .card.wide{grid-column:1 / -1;}
  .card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap;gap:8px;}
  .card-title{display:flex;align-items:center;gap:8px;}
  .card-title .idx{font-family:var(--mono);font-size:15px;color:var(--bg-0);background:var(--cyan);min-width:28px;height:28px;padding:0 4px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-weight:700;flex:0 0 auto;}
  .card-title h3{margin:0;font-size:20px;font-weight:600;color:var(--text-hi);}
  .toggle-group{display:flex;gap:4px;background:var(--bg-2);border-radius:7px;padding:3px;}
  .toggle-group button{background:transparent;border:none;color:var(--text-dim);font-family:var(--mono);font-size:15px;padding:7px 14px;border-radius:5px;cursor:pointer;transition:all .12s;}
  .toggle-group button.active{background:var(--bg-3);color:var(--cyan);}
  .chart-wrap{position:relative;height:280px;margin-top:10px;}
  .chart-wrap.small{height:170px;}
  .chart-wrap.tall{height:340px;}
  .chart-wrap.xtall{height:560px;}
  .empty-note{display:flex;align-items:center;justify-content:center;height:230px;color:var(--text-dim);font-size:17px;font-family:var(--mono);text-align:center;}
  .note{color:var(--text-dim);font-size:14px;margin:12px 2px 6px;line-height:1.55;}
  .note b{color:var(--text-mid);}

  .head-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
  .dl-btn{background:var(--bg-2);border:1px solid var(--line);color:var(--text-mid);font-family:var(--mono);font-size:14px;padding:6px 10px;border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:6px;transition:all .12s;}
  .dl-btn:hover{border-color:var(--cyan);color:var(--cyan);}
  .dl-btn svg{width:14px;height:14px;}

  .filter-row{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:10px;}
  .inp{background:var(--bg-2);border:1px solid var(--line);color:var(--text-hi);font-family:var(--mono);font-size:14px;padding:7px 10px;border-radius:7px;outline:none;}
  .inp:focus{border-color:var(--cyan-dim);}
  .filter-row .count{font-family:var(--mono);font-size:14px;color:var(--text-dim);margin-left:auto;}

  .tbl-wrap{overflow:auto;max-height:560px;margin-top:10px;border:1px solid var(--line-soft);border-radius:10px;}
  .tbl{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:14px;}
  .tbl th{padding:9px 10px;text-align:left;color:var(--text-dim);border-bottom:1px solid var(--line);white-space:nowrap;position:sticky;top:0;background:var(--bg-2);z-index:1;font-weight:600;}
  .tbl th.sortable{cursor:pointer;user-select:none;}
  .tbl th.sortable:hover{color:var(--cyan);}
  .tbl th.num,.tbl td.num{text-align:right;}
  .tbl td{padding:7px 10px;border-bottom:1px solid var(--line-soft);white-space:nowrap;}
  .tbl tr.clickrow{cursor:pointer;}
  .tbl tr.clickrow:hover{background:var(--bg-3);}
  .pill{display:inline-block;padding:1px 8px;border-radius:4px;font-size:13px;border:1px solid;}

  footer{margin-top:26px;text-align:center;color:var(--text-dim);font-size:15px;font-family:var(--mono);}
  footer kbd{background:var(--bg-2);border:1px solid var(--line);border-radius:4px;padding:1px 6px;color:var(--text-mid);font-family:var(--mono);}
</style>
</head>
<body>
<div class="app">

  <header class="topbar">
    <div class="brand">
      <h1>Painel Hidrológico</h1>
      <div class="meta" id="metaLinha">—</div>
    </div>
    <div class="header-right">
      <div class="search-box">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input id="searchInput" type="text" placeholder="Buscar usina ou posto…" autocomplete="off">
        <kbd>/</kbd>
        <div id="searchResults" class="search-results"></div>
      </div>
    </div>
  </header>

  <div class="tabbar" id="tabbar">
    <button data-tab="geral" class="active">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>
      Visão Geral
    </button>
    <button data-tab="usinas">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="10" width="4" height="10"/><rect x="10" y="4" width="4" height="16"/><rect x="16" y="7" width="4" height="13"/></svg>
      Usinas
    </button>
    <button data-tab="cascatas">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 4v6h6M4 10l6-6M14 4v16M20 4v16M4 20h4M4 15h4"/></svg>
      Cascatas
    </button>
    <button data-tab="estac">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 17l6-6 4 4 8-8M15 5h6v6"/></svg>
      Estacionariedade
    </button>
    <button data-tab="saz">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="10" width="4" height="10"/><rect x="10" y="4" width="4" height="16"/><rect x="17" y="13" width="4" height="7"/></svg>
      Sazonalidade
    </button>
    <button data-tab="perm">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 5c4 0 5 14 9 14s5-14 9-14"/></svg>
      Permanência
    </button>
    <button data-tab="evap">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3c3 4 6 7 6 11a6 6 0 0 1-12 0c0-4 3-7 6-11z"/><path d="M9 14a3 3 0 0 0 3 3"/></svg>
      Evaporação
    </button>
    <button data-tab="corr">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12l5 5L20 6"/></svg>
      Correção
    </button>
    <button data-tab="usos">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v6M8 5l4 3 4-3M4 12h16M4 12c0 5 4 9 8 9s8-4 8-9"/></svg>
      Usos Consuntivos
    </button>
    <button data-tab="comp">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 3v18M16 3v18M4 8h4M16 8h4M4 16h4M16 16h4"/></svg>
      Comparar
    </button>
  </div>

  <!-- ============ VIEW: VISAO GERAL ============ -->
  <section id="viewGeral" class="view active">
    <div class="kpis" id="geralKpis"></div>

    <div class="grid">
      <div class="card">
        <div class="card-head">
          <div class="card-title"><span class="idx">G1</span><h3>Meses com Qinc &lt; 0 por mês do ano</h3></div>
        </div>
        <div class="chart-wrap" data-for="chartGeralMes"><canvas id="chartGeralMes"></canvas></div>
      </div>
      <div class="card">
        <div class="card-head">
          <div class="card-title"><span class="idx">G2</span><h3>Meses com Qinc &lt; 0 por década</h3></div>
        </div>
        <div class="chart-wrap" data-for="chartGeralDec"><canvas id="chartGeralDec"></canvas></div>
      </div>

      <div class="card wide" data-nodl="1">
        <div class="card-head">
          <div class="card-title"><span class="idx">G3</span><h3>Tabela de usinas — incrementais negativas, evaporação e estacionariedade</h3></div>
          <button class="dl-btn" id="btnCsvGeral" title="Baixar tabela filtrada em CSV">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>CSV
          </button>
        </div>
        <div class="badges" id="geralClasses" style="margin:10px 0 0;"></div>
        <div class="filter-row">
          <input id="geralFiltro" class="inp" type="text" placeholder="Filtrar por nome/código..." style="width:240px;">
          <label class="note" style="margin:0;display:flex;align-items:center;gap:6px;"><input type="checkbox" id="geralSoNeg"> só com Qinc &lt; 0</label>
          <label class="note" style="margin:0;display:flex;align-items:center;gap:6px;"><input type="checkbox" id="geralSoEsp"> só casos especiais</label>
          <span class="count" id="geralCount"></span>
        </div>
        <div class="tbl-wrap">
          <table class="tbl" id="tabelaGeral">
            <thead><tr id="geralHead"></tr></thead>
            <tbody id="geralTbody"></tbody>
          </table>
        </div>
        <p class="note">Clique no cabeçalho para ordenar e na linha para abrir a usina. Classes pelo % de meses com Qinc &lt; 0:
          <b>1</b> &gt; <span class="limC1"></span>% · <b>2</b> entre <span class="limC2"></span>% e <span class="limC1"></span>% · <b>3</b> abaixo de <span class="limC2"></span>%.
          <b>Evap.</b> = evaporação líquida média do reservatório (Área Méd) — só diagnóstico de perda, já embutida na vazão natural e não descontada do Qinc.</p>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: USINAS ============ -->
  <section id="viewUsinas" class="view">

    <div class="nav-strip">
      <button class="nav-btn" id="uPrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="usinaNome">—</h2>
          <span class="tag cyan" id="tagUsina">USINA —</span>
          <span class="tag" id="tagPosto">POSTO —</span>
          <span class="tag" id="tagClasse" style="display:none;"></span>
        </div>
        <div class="sub" id="usinaSub">carregando…</div>
      </div>
      <div class="position-counter">
        <span class="big" id="uPosAtual">–</span> / <span id="uPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="uProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="uNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="badgesRow"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">1</span><h3>Vazão Natural — Qnat</h3></div>
          <div class="toggle-group" data-chart="natural">
            <button data-mode="anual" class="active">Anual</button>
            <button data-mode="mensal">Mensal</button>
          </div>
        </div>
        <div class="chart-wrap" data-for="chartNatural"><canvas id="chartNatural"></canvas></div>
      </div>

      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">2</span><h3>Vazão Incremental — Qinc</h3></div>
          <div class="toggle-group" data-chart="incremental">
            <button data-mode="anual" class="active">Anual</button>
            <button data-mode="mensal">Mensal</button>
          </div>
        </div>
        <div class="chart-wrap" data-for="chartIncremental"><canvas id="chartIncremental"></canvas></div>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: CASCATAS ============ -->
  <section id="viewCascatas" class="view">

    <div class="nav-strip">
      <button class="nav-btn" id="cPrevBtn" title="Cascata anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="cascataNomeH2">Cascata —</h2>
          <span class="tag cyan" id="cascataTagCount">— usinas</span>
        </div>
        <div class="sub" id="cascataSub">carregando…</div>
      </div>
      <div class="position-counter">
        <span class="big" id="cPosAtual">–</span> / <span id="cPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="cProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="cNextBtn" title="Próxima cascata (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="card wide">
      <div class="card-head">
        <div class="card-title"><span class="idx">C</span><h3 id="cascataChartTitle">Série em Cascata — Anual</h3></div>
        <div class="toggle-group" data-chart="cascata">
          <button data-mode="natural" class="active">Natural</button>
          <button data-mode="incremental">Incremental</button>
        </div>
      </div>
      <div class="chart-wrap tall" data-for="chartCascata"><canvas id="chartCascata"></canvas></div>
    </div>

    <div class="badges" id="cascataBadges" style="margin-top:18px;"></div>

    <div class="card wide" style="margin-top:18px;" data-nodl="1">
      <div class="card-head">
        <div class="card-title"><span class="idx">T</span><h3>Usinas em cascatas isoladas — sem montante cadastrada</h3></div>
        <button class="dl-btn" id="btnCsvTopo" title="Baixar tabela em CSV">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>CSV
        </button>
      </div>
      <div class="badges" id="topoResumoRow" style="margin-top:6px;"></div>
      <p class="note">
        Usinas cuja cascata cadastrada tem tamanho 1 — sem nenhuma usina imediatamente a montante no banco de dados.
        A maioria são cabeceiras reais. As marcadas como <b>caso especial</b> são citadas no projeto original
        como tendo ligação física conhecida (derivação, transposição) que não aparece no cadastro — ex.: Belo Monte / Pimental.
      </p>
      <div class="tbl-wrap">
        <table class="tbl">
          <thead><tr><th>Usina</th><th class="num">Posto</th><th class="num">Vazão natural média (m³/s)</th><th>Observação</th></tr></thead>
          <tbody id="topoTbody"></tbody>
        </table>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: ESTACIONARIEDADE ============ -->
  <section id="viewEstac" class="view">

    <div class="card wide" style="margin-bottom:18px;">
      <div class="card-head">
        <div class="card-title"><span class="idx">E</span><h3>Resumo — Mann-Kendall (corrigido por autocorrelação) &amp; Pettitt</h3></div>
        <div class="toggle-group" data-chart="estacEscopo">
          <button data-mode="usinas" class="active">Postos de usinas</button>
          <button data-mode="todos">Todos os postos</button>
        </div>
        <button class="dl-btn" id="btnCsvResumoEstac" title="Baixar tabela em CSV">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>CSV
        </button>
      </div>
      <div class="badges" id="estacResumoRow" style="margin-top:6px;"></div>
      <div class="chart-wrap small" style="margin-top:4px;" data-for="chartEstacDecadas"><canvas id="chartEstacDecadas"></canvas></div>
      <p class="note">Com α = 5%, espera-se cerca de 5% de resultados significativos por acaso (falsos positivos). O teste de Pettitt também acusa quebra quando há tendência monotônica,
        por isso "quebra sem tendência MK" é o grupo mais próximo de uma mudança de patamar pura.</p>
    </div>

    <div class="nav-strip">
      <button class="nav-btn" id="ePrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="estacUsinaNome">—</h2>
          <span class="tag cyan" id="estacTagUsina">USINA —</span>
          <span class="tag" id="estacTagPosto">POSTO —</span>
        </div>
        <div class="sub" id="estacUsinaSub">carregando…</div>
      </div>
      <div class="position-counter">
        <span class="big" id="ePosAtual">–</span> / <span id="ePosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="eProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="eNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="estacBadgesRow"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">1</span><h3>Vazão Natural Anual — Tendência (Sen) &amp; Ponto de Quebra (Pettitt)</h3></div>
        </div>
        <div class="chart-wrap tall" data-for="chartEstac"><canvas id="chartEstac"></canvas></div>
      </div>
    </div>

    <div class="grid" style="margin-top:18px;">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">2</span><h3 id="janelaMovelTitulo">Tendência em Janela Móvel (30 anos)</h3></div>
          <div class="toggle-group" data-chart="janela">
            <button data-mode="10">10 anos</button>
            <button data-mode="20">20 anos</button>
            <button data-mode="30" class="active">30 anos</button>
          </div>
        </div>
        <div class="chart-wrap tall" data-for="chartJanelaMovel"><canvas id="chartJanelaMovel"></canvas></div>
        <p class="note">Cada ponto é a declividade de Sen da janela que termina naquele ano. Verde = crescente, vermelho = decrescente (MK, p &lt; 0,05), cinza = não significativo.
          Janelas móveis não usam a correção de Yue-Wang (instável em janelas curtas) — use como leitura exploratória.</p>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: SAZONALIDADE ============ -->
  <section id="viewSaz" class="view">
    <div class="nav-strip">
      <button class="nav-btn" id="sPrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="sazUsinaNome">—</h2>
          <span class="tag cyan" id="sazTagUsina">USINA —</span>
          <span class="tag" id="sazTagPosto">POSTO —</span>
        </div>
        <div class="sub" id="sazUsinaSub">carregando…</div>
      </div>
      <div class="position-counter">
        <span class="big" id="sPosAtual">–</span> / <span id="sPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="sProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="sNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="sazBadgesRow"></div>
    <div class="badges" id="extremosBadgesRow" style="margin-top:6px;"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">1</span><h3>Sazonalidade Mensal — Vazão Natural (faixa P10–P90, mediana e média)</h3></div>
        </div>
        <div class="chart-wrap tall" data-for="chartSaz"><canvas id="chartSaz"></canvas></div>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: PERMANÊNCIA ============ -->
  <section id="viewPerm" class="view">
    <div class="nav-strip">
      <button class="nav-btn" id="pPrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="permUsinaNome">—</h2>
          <span class="tag cyan" id="permTagUsina">USINA —</span>
          <span class="tag" id="permTagPosto">POSTO —</span>
        </div>
        <div class="sub" id="permUsinaSub">carregando…</div>
      </div>
      <div class="position-counter">
        <span class="big" id="pPosAtual">–</span> / <span id="pPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="pProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="pNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="permBadgesRow"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">1</span><h3>Curva de Permanência — Vazão Natural</h3></div>
          <select id="permMesSelect" class="inp">
            <option value="">+ sobrepor mês…</option>
          </select>
        </div>
        <div class="chart-wrap tall" data-for="chartPerm"><canvas id="chartPerm"></canvas></div>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: EVAPORACAO ============ -->
  <section id="viewEvap" class="view">
    <div class="kpis" id="evapKpis"></div>
    <p class="note" style="margin:-4px 2px 16px;">
      A evaporação líquida dos reservatórios <b>já está incorporada na vazão natural oficial</b> (reconstituição do ONS), por isso
      <b>não é somada nem subtraída do Qinc</b>. Esta aba mede o tamanho dessa perda por usina e por bacia:
      Evap = EvapMen (mm) × Área (km²) × 10³ ÷ segundos do mês, com a área pelo polinômio área-cota na cota média (<b>Área Méd</b>, valor de referência)
      e a faixa entre as cotas mínima e máxima. Meses com evaporação líquida negativa (condensação) = 0. Média anual = volume anual ÷ segundos do ano.
    </p>
    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">V1</span><h3>Ranking — usinas que mais perdem por evaporação</h3></div>
          <div class="toggle-group" data-chart="evapMetrica">
            <button data-mode="m3s" class="active">m³/s</button>
            <button data-mode="pctQnat">% da Qnat</button>
            <button data-mode="pctQinc">% da Qinc</button>
          </div>
          <select id="evapTopN" class="inp">
            <option value="15">Top 15</option><option value="20" selected>Top 20</option><option value="30">Top 30</option>
          </select>
        </div>
        <div class="chart-wrap xtall" data-for="chartEvapRank"><canvas id="chartEvapRank"></canvas></div>
        <p class="note" id="evapRankNota"></p>
      </div>

      <div class="card">
        <div class="card-head">
          <div class="card-title"><span class="idx">V2</span><h3>Perdas por bacia — % da vazão natural na usina de jusante final</h3></div>
        </div>
        <div class="chart-wrap xtall" data-for="chartEvapBacia"><canvas id="chartEvapBacia"></canvas></div>
        <p class="note"><b>Bacia</b> = usinas que drenam para a mesma usina de jusante final do cadastro. Perda da bacia = soma da evaporação dessas usinas ÷ vazão natural média
          na usina de jusante final (que já acumula toda a bacia a montante).</p>
      </div>

      <div class="card">
        <div class="card-head">
          <div class="card-title"><span class="idx">V3</span><h3>Relação com as vazões incrementais negativas</h3></div>
          <div class="toggle-group" data-chart="evapRelX">
            <button data-mode="pctQinc" class="active">Evap/Qinc</button>
            <button data-mode="pctQnat">Evap/Qnat</button>
          </div>
        </div>
        <div class="chart-wrap xtall" data-for="chartEvapRel"><canvas id="chartEvapRel"></canvas></div>
        <p class="note" id="evapRelNota"></p>
      </div>

      <div class="card wide" data-nodl="1">
        <div class="card-head">
          <div class="card-title"><span class="idx">V4</span><h3>Tabela — perdas por evaporação por usina</h3></div>
          <button class="dl-btn" id="btnCsvEvap" title="Baixar tabela em CSV">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>CSV
          </button>
        </div>
        <div class="filter-row">
          <input id="evapFiltro" class="inp" type="text" placeholder="Filtrar por usina ou bacia..." style="width:240px;">
          <span class="count" id="evapCount"></span>
        </div>
        <div class="tbl-wrap">
          <table class="tbl">
            <thead><tr id="evapHead"></tr></thead>
            <tbody id="evapTbody"></tbody>
          </table>
        </div>
        <p class="note">Clique no cabeçalho para ordenar e na linha para ver o perfil mensal da usina. <b>Evap/Qinc</b> fica em branco quando o Qinc médio é ≤ 0.</p>
      </div>
    </div>

    <div class="nav-strip" id="evapUsinaNav" style="margin-top:18px;">
      <button class="nav-btn" id="vPrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="evapUsinaNome">—</h2>
          <span class="tag cyan" id="evapTagUsina">USINA —</span>
          <span class="tag" id="evapTagPosto">POSTO —</span>
        </div>
        <div class="sub" id="evapUsinaSub">—</div>
      </div>
      <div class="position-counter">
        <span class="big" id="vPosAtual">–</span> / <span id="vPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="vProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="vNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="evapBadgesRow"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">V5</span><h3>Evaporação líquida — perfil mensal</h3></div>
        </div>
        <div class="chart-wrap tall" data-for="chartEvap"><canvas id="chartEvap"></canvas></div>
        <p class="note">Evaporação de cada mês do calendário com a área nas cotas mínima, média e máxima. O perfil é o mesmo em todos os anos:
          a lâmina de evaporação líquida (EvapMen) é um vetor climatológico de 12 valores por usina.</p>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: CORRECAO ============ -->
  <section id="viewCorr" class="view">
    <div class="card wide" data-nodl="1" style="margin-bottom:18px;">
      <div class="card-head">
        <div class="card-title"><span class="idx">K</span><h3>Proposta de correção — conservação de massa (redistribuição anual do déficit)</h3></div>
      </div>
      <div class="kpis" id="corrKpis" style="margin-top:10px;"></div>
      <p class="note">
        Para cada usina, os meses com <b>Qinc &lt; 0</b> são zerados e o mesmo <b>volume</b> é retirado dos meses positivos do mesmo ano,
        proporcionalmente (fator f = 1 − D/P, com D = déficit e P = volume positivo do ano). A vazão natural corrigida é reconstruída de montante
        para jusante como <b>Qi' = ΣQm' + Qinc'ᵢ</b>, o que preserva o volume anual em todos os postos e elimina os negativos nos anos corrigíveis.
        Anos com Qinc anual ≤ 0 <b>não são corrigíveis</b> por redistribuição (o problema é de volume, não de distribuição mensal) e ficam listados.
        Usinas <b>estruturais</b> (Classe 1 — derivação/transposição/topologia) não são alteradas: a recomendação para elas é corrigir o cadastro ou o posto.
        A correção parte do Qinc do deck oficial. Deck corrigido em <code>saidas/vazoes_corrigidas_massa.txt</code>.
      </p>
      <div class="filter-row">
        <label class="note" style="margin:0;display:flex;align-items:center;gap:6px;"><input type="checkbox" id="corrTodas"> mostrar também usinas sem negativos</label>
        <span class="count" id="corrCount"></span>
      </div>
      <div class="tbl-wrap" style="max-height:420px;">
        <table class="tbl">
          <thead><tr>
            <th>Usina</th><th>Classe</th><th>Status</th><th class="num">Neg. orig.</th><th class="num">Neg. após</th>
            <th class="num">Meses corrigidos</th><th class="num">Anos corrigidos</th><th class="num">Anos não corrigíveis</th>
            <th class="num">Vol. redistribuído (%)</th><th class="num">Máx |ΔQ| (m³/s)</th>
          </tr></thead>
          <tbody id="corrTbody"></tbody>
        </table>
      </div>
      <p class="note"><b>Vol. redistribuído</b> = déficit retirado ÷ volume positivo dos anos corrigidos. Valores altos (ex.: &gt; 20%) indicam que a correção
        altera muito a sazonalidade da incremental — nesses casos a correção é numérica, não uma explicação física.</p>
    </div>

    <div class="nav-strip">
      <button class="nav-btn" id="kPrevBtn" title="Usina anterior (←)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="identity">
        <div class="top-row">
          <h2 id="corrUsinaNome">—</h2>
          <span class="tag cyan" id="corrTagUsina">USINA —</span>
          <span class="tag" id="corrTagPosto">POSTO —</span>
        </div>
        <div class="sub" id="corrUsinaSub">—</div>
      </div>
      <div class="position-counter">
        <span class="big" id="kPosAtual">–</span> / <span id="kPosTotal">–</span>
        <div class="progress-track"><div class="progress-fill" id="kProgressFill" style="width:0%"></div></div>
      </div>
      <button class="nav-btn" id="kNextBtn" title="Próxima usina (→)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>

    <div class="badges" id="corrBadgesRow"></div>

    <div class="grid">
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">1</span><h3>Qinc mensal — original vs. corrigido</h3></div>
        </div>
        <div class="chart-wrap tall" data-for="chartCorrMensal"><canvas id="chartCorrMensal"></canvas></div>
      </div>
      <div class="card wide">
        <div class="card-head">
          <div class="card-title"><span class="idx">2</span><h3>Perfil sazonal do Qinc — média por mês do ano</h3></div>
        </div>
        <div class="chart-wrap" data-for="chartCorrPerfil"><canvas id="chartCorrPerfil"></canvas></div>
      </div>
    </div>
  </section>

  <!-- ============ VIEW: COMPARAR ============ -->
  <section id="viewComp" class="view">
    <div class="card wide">
      <div class="card-head">
        <div class="card-title"><span class="idx">⇄</span><h3 id="compTitulo">Comparar duas usinas — Vazão Natural Anual</h3></div>
        <div class="toggle-group" data-chart="compVar">
          <button data-mode="natural" class="active">Natural</button>
          <button data-mode="incremental">Incremental</button>
        </div>
      </div>
      <div style="display:flex; gap:16px; flex-wrap:wrap; margin:10px 2px 16px;">
        <select id="compSelectA" class="inp" style="flex:1; min-width:220px;"></select>
        <select id="compSelectB" class="inp" style="flex:1; min-width:220px;"></select>
        <label class="note" style="margin:0;display:flex;align-items:center;gap:6px;"><input type="checkbox" id="compMesmoEixo"> mesmo eixo Y</label>
      </div>
      <div class="badges" id="compBadgesRow"></div>
      <div class="chart-wrap tall" style="margin-top:12px;" data-for="chartComp"><canvas id="chartComp"></canvas></div>
    </div>
  </section>

  <!-- ============ VIEW: USOS CONSUNTIVOS ============ -->
  <section id="viewUsos" class="view">
    <div class="card wide" data-nodl="1">
      <div class="card-head">
        <div class="card-title"><span class="idx">U</span><h3>Diagnóstico — Usos Consuntivos Incrementais vs. Vazão Incremental Negativa</h3></div>
        <button class="dl-btn" id="btnCsvUsos" title="Baixar tabela em CSV">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>CSV
        </button>
      </div>
      <div class="badges" id="usosResumoRow" style="margin-top:6px;"></div>
      <p class="note">
        Para as usinas com meses de <b>Qinc</b> negativo, compara o déficit médio nesses meses com o consumo de água
        (usos consuntivos) atribuído ao trecho incremental entre montante e jusante (base ANA, Resolução 92/2021 v2).
        <b>Não é uma correção</b> — a vazão natural oficial já soma os usos consuntivos de volta em cada posto, então esse termo já está embutido no Qinc.
        "Cobertura" indica o tamanho do consumo frente ao déficit: é um indício de causa plausível (ex.: estimativa de usos desatualizada), não uma prova.
      </p>
      <div id="usosSemDados" class="empty-note" style="display:none;">
        Aba desativada — rode o gerador com <b>&nbsp;--ana-usos &lt;planilha da Resolução 92/2021&gt;&nbsp;</b> para habilitar.
      </div>
      <div id="usosComDados" class="tbl-wrap">
        <table class="tbl">
          <thead>
            <tr>
              <th>Usina</th>
              <th class="num">Meses com Qinc &lt; 0</th>
              <th class="num">Qinc médio nesses meses (m³/s)</th>
              <th class="num">Usos cons. incremental nesses meses (m³/s)</th>
              <th class="num">Cobertura do déficit</th>
            </tr>
          </thead>
          <tbody id="usosTbody"></tbody>
        </table>
      </div>
    </div>
  </section>

  <footer>
    <kbd>←</kbd> <kbd>→</kbd> navegam · <kbd>/</kbd> busca · <span id="footerCount">—</span> usinas · <span id="footerCascCount">—</span> cascatas
  </footer>

</div>

<script id="app-data" type="application/json">__APP_DATA__</script>

<script>
const APP = JSON.parse(document.getElementById('app-data').textContent);
const USINAS = APP.usinas;
const CASCATAS = APP.cascatas;
const RESUMO = APP.resumo || null;
let uCurrent = 0;
let cCurrent = 0;
let activeTab = 'geral';

Chart.defaults.color = '#9db3c4';
Chart.defaults.font.family = "'IBM Plex Mono', monospace";
Chart.defaults.font.size = 15;
Chart.defaults.borderColor = '#1c3348';

const COLORS = {cyan:'#4fd8e8', cyanDim:'#2b8ea3', amber:'#f2b84b', coral:'#f2765b', green:'#6fd6a0', grid:'#1c3348', textDim:'#5f7a90', white:'#eaf2f6'};
const PALETTE = [COLORS.cyan, COLORS.amber, COLORS.green, COLORS.coral, '#c792ea', '#7aa2f7', '#e0af68', '#89ddff'];
const CLASSE_INFO = {
  1:{txt:'Classe 1', cor:COLORS.coral},
  2:{txt:'Classe 2', cor:COLORS.amber},
  3:{txt:'Classe 3', cor:COLORS.cyan},
  0:{txt:'Sem negativos', cor:COLORS.green},
};
const fmt = (v, d=0) => v==null || Number.isNaN(v) ? '—' : Number(v).toLocaleString('pt-BR', {minimumFractionDigits:d, maximumFractionDigits:d});
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// linha tracejada em y = 0 (ativa com options.plugins.zeroLine.enabled)
Chart.register({
  id:'zeroLine',
  afterDatasetsDraw(chart, args, opts){
    if(!opts || !opts.enabled) return;
    const y = chart.scales.y;
    if(!y || y.min > 0 || y.max < 0) return;
    const py = y.getPixelForValue(0), a = chart.chartArea, ctx = chart.ctx;
    ctx.save();
    ctx.strokeStyle = opts.color || 'rgba(242,118,91,0.8)';
    ctx.lineWidth = 1; ctx.setLineDash([5,4]);
    ctx.beginPath(); ctx.moveTo(a.left, py); ctx.lineTo(a.right, py); ctx.stroke();
    ctx.restore();
  }
});
function comZero(opts){ opts.plugins = opts.plugins || {}; opts.plugins.zeroLine = {enabled:true}; return opts; }

function axisTitle(text){
  return { display:true, text, color:COLORS.textDim, font:{size:14, family:"'IBM Plex Mono', monospace", weight:600}, padding:{top:8} };
}
function baseScales({xTitle, yTitle, xTicks={}, yTicks={}}={}){
  return {
    x:{ grid:{color:COLORS.grid, drawTicks:false}, ticks:{maxRotation:0, autoSkip:true, maxTicksLimit:10, ...xTicks}, title: xTitle ? axisTitle(xTitle) : undefined },
    y:{ grid:{color:COLORS.grid, drawTicks:false}, ticks:{maxTicksLimit:6, ...yTicks}, title: yTitle ? axisTitle(yTitle) : undefined }
  };
}
const TOOLTIP = {backgroundColor:'#0f1e2b', borderColor:'#24405a', borderWidth:1, titleColor:'#eaf2f6', bodyColor:'#9db3c4', padding:10, cornerRadius:8};
const LEGEND_BOTTOM = {display:true, position:'bottom', labels:{padding:14, usePointStyle:true, pointStyleWidth:30, generateLabels:(c)=>legendaExport(c)}};
function baseOpts(overrides={}, axes={}){
  return Object.assign({
    responsive:true, maintainAspectRatio:false,
    interaction:{mode:'index', intersect:false},
    plugins:{ legend:{display:false}, tooltip:{...TOOLTIP} },
    scales: baseScales(axes)
  }, overrides);
}

let charts = {};
function destroyChart(key){ if(charts[key]){ charts[key].destroy(); charts[key]=null; } }
function wrapOf(id){ return document.querySelector(`.chart-wrap[data-for="${id}"]`); }
function resetCanvas(id){
  const wrap = wrapOf(id);
  wrap.innerHTML = `<canvas id="${id}"></canvas>`;
  return document.getElementById(id);
}

let modes = {natural:'anual', incremental:'anual', cascata:'natural', janela:'30', estacEscopo:'usinas', compVar:'natural'};

// ============================================================
// META / HEADER
// ============================================================
(function(){
  const m = APP.meta || {};
  const partes = [];
  if(m.deck) partes.push(`deck: ${m.deck}`);
  if(m.anoIni) partes.push(`${m.anoIni}–${m.anoFim}`);
  partes.push(`${USINAS.length} usinas`);
  partes.push('Qinc = Qi − ΣQm');
  document.getElementById('metaLinha').textContent = partes.join(' · ');
  document.querySelectorAll('.limC1').forEach(e => e.textContent = RESUMO ? RESUMO.limites.c1 : 30);
  document.querySelectorAll('.limC2').forEach(e => e.textContent = RESUMO ? RESUMO.limites.c2 : 2);
})();

// ============================================================
// TABS
// ============================================================
const VIEWS = {geral:'viewGeral', usinas:'viewUsinas', cascatas:'viewCascatas', estac:'viewEstac', saz:'viewSaz', perm:'viewPerm', evap:'viewEvap', corr:'viewCorr', comp:'viewComp', usos:'viewUsos'};
const TABS_POR_USINA = ['usinas','estac','saz','perm','evap','corr'];
document.querySelectorAll('#tabbar button').forEach(b => b.addEventListener('click', () => switchTab(b.dataset.tab)));
function switchTab(tab){
  activeTab = tab;
  document.querySelectorAll('#tabbar button').forEach(b => b.classList.toggle('active', b.dataset.tab===tab));
  Object.entries(VIEWS).forEach(([k,id]) => document.getElementById(id).classList.toggle('active', k===tab));
  renderAtivo();
}
function renderAtivo(){
  ({geral:renderGeralView, usinas:renderUsinaView, cascatas:renderCascataView, estac:renderEstacView,
    saz:renderSazView, perm:renderPermView, evap:renderEvapView, corr:renderCorrView, comp:renderCompView, usos:renderUsosConsuntivos})[activeTab]();
}

// ============================================================
// VISAO GERAL
// ============================================================
let geralBuilt = false;
let geralSort = {col:'pct', dir:-1};
let geralClasse = null;

const COLS_GERAL = [
  {k:'nome', t:'Usina'},
  {k:'cod', t:'Cód.', num:true},
  {k:'posto', t:'Posto', num:true},
  {k:'classe', t:'Classe'},
  {k:'neg', t:'Meses Qinc < 0', num:true},
  {k:'pct', t:'% meses', num:true},
  {k:'evap', t:'Evap. (m³/s)', num:true, evap:true},
  {k:'evQnat', t:'Evap/Qnat (%)', num:true, evap:true},
  {k:'nCorr', t:'Neg. após correção', num:true},
  {k:'tend', t:'Tendência Qnat'},
  {k:'quebra', t:'Quebra (Pettitt)', num:true},
];

function linhasGeral(){
  return USINAS.map((u,i) => {
    const inc = u.incremental, e = u.estacionariedade, ev = u.evaporacao;
    return {
      i, nome:u.nome, cod:u.codUsina, posto:u.codPosto, classe:u.classe,
      neg: inc ? inc.negativos : null, pct: inc ? inc.pctNegativos : null,
      evap: ev ? ev.evapMed : null, evQnat: ev ? ev.pctQnat : null,
      nCorr: u.correcao ? u.correcao.negDepois : null,
      tend: e ? e.mkTrendCorr : null,
      quebra: (e && e.pettittP < 0.05) ? e.pettittAno : null,
      qincNulo: u.qincNulo, especial: u.casoEspecial,
    };
  });
}

function pillClasse(c){
  if(c==null) return '<span style="color:var(--text-dim)">—</span>';
  const inf = CLASSE_INFO[c];
  return `<span class="pill" style="color:${inf.cor};border-color:${inf.cor}">${c===0 ? '0' : c}</span>`;
}
function txtTend(t){
  if(!t) return '<span style="color:var(--text-dim)">—</span>';
  if(t==='crescente') return `<span style="color:${COLORS.green}">▲ cresc.</span>`;
  if(t==='decrescente') return `<span style="color:${COLORS.coral}">▼ decr.</span>`;
  return '<span style="color:var(--text-dim)">n.s.</span>';
}

function geralFiltradas(){
  const txt = document.getElementById('geralFiltro').value.trim().toLowerCase();
  const soNeg = document.getElementById('geralSoNeg').checked;
  const soEsp = document.getElementById('geralSoEsp').checked;
  let L = linhasGeral().filter(r =>
    (!txt || r.nome.toLowerCase().includes(txt) || String(r.cod).includes(txt) || String(r.posto ?? '').includes(txt)) &&
    (!soNeg || (r.neg||0) > 0) && (!soEsp || r.especial) &&
    (geralClasse==null || r.classe===geralClasse));
  const {col, dir} = geralSort;
  L.sort((a,b) => {
    const va=a[col], vb=b[col];
    if(va==null && vb==null) return 0;
    if(va==null) return 1;
    if(vb==null) return -1;
    if(typeof va==='string') return va.localeCompare(vb, 'pt-BR')*dir;
    return (va-vb)*dir;
  });
  return L;
}

function renderGeralTabela(){
  const temEvap = !!APP.evaporacao;
  const cols = COLS_GERAL.filter(c => !c.evap || temEvap);
  document.getElementById('geralHead').innerHTML = cols.map(c => {
    const seta = geralSort.col===c.k ? (geralSort.dir>0 ? ' ▲' : ' ▼') : '';
    return `<th class="sortable ${c.num?'num':''}" data-col="${c.k}">${c.t}${seta}</th>`;
  }).join('');
  document.querySelectorAll('#geralHead th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.col;
    geralSort = {col:k, dir: geralSort.col===k ? -geralSort.dir : (k==='nome' ? 1 : -1)};
    renderGeralTabela();
  }));
  const L = geralFiltradas();
  document.getElementById('geralCount').textContent = `${L.length} de ${USINAS.length} usinas`;
  document.getElementById('geralTbody').innerHTML = L.map(r => {
    const flags = (r.qincNulo ? ` <span title="Qinc ≡ 0 (posto igual ao de montante)" style="color:${COLORS.amber}">∅</span>` : '')
                + (r.especial ? ` <span title="caso especial do projeto (derivação/transposição/complexo)" style="color:${COLORS.amber}">⚠</span>` : '');
    const cel = {
      nome: esc(r.nome) + flags, cod: r.cod, posto: r.posto ?? '—', classe: pillClasse(r.classe),
      neg: fmt(r.neg), pct: r.pct==null ? '—' : fmt(r.pct,2),
      evap: r.evap==null ? '—' : fmt(r.evap,2),
      evQnat: r.evQnat==null ? '—' : fmt(r.evQnat,2),
      nCorr: r.nCorr==null ? '—' : (r.nCorr>0 ? `<span style="color:${COLORS.coral}">${fmt(r.nCorr)}</span>` : fmt(r.nCorr)),
      tend: txtTend(r.tend), quebra: r.quebra ?? '—',
    };
    return `<tr class="clickrow" data-idx="${r.i}">` + cols.map(c => `<td class="${c.num?'num':''}">${cel[c.k]}</td>`).join('') + '</tr>';
  }).join('');
  document.querySelectorAll('#geralTbody tr').forEach(tr => tr.addEventListener('click', () => {
    uCurrent = parseInt(tr.dataset.idx); switchTab('usinas');
  }));
}

function renderGeralKpis(){
  const R = RESUMO;
  const el = document.getElementById('geralKpis');
  if(!R){ el.innerHTML = ''; return; }
  const k = (lbl, val, sub, cls='') => `<div class="kpi ${cls}"><div class="lbl">${lbl}</div><div class="val">${val}</div><div class="sub">${sub}</div></div>`;
  let h = k('Meses analisados', fmt(R.totalMeses), `${USINAS.length} usinas`, 'cyan');
  h += k('Qinc < 0', fmt(R.neg.orig), `${fmt(R.pct.orig,2)}% dos meses`, 'warn');
  if(APP.evaporacao){
    const E = APP.evaporacao;
    h += k('Evaporação líquida', `${fmt(E.totalMed,0)} m³/s`, `${fmt(E.volAnualKm3,1)} km³/ano${E.pctSIN!=null ? ` · ${fmt(E.pctSIN,2)}% da Qnat das bacias` : ''} · diagnóstico (não entra no Qinc)`, 'amber');
  }
  if(APP.correcao){
    h += k('Após correção de massa', fmt(APP.correcao.negDepois), `${fmt(APP.correcao.negDepois/R.totalMeses*100,2)}% · restantes em estruturais/anos não corrigíveis`, 'cyan');
  }
  const eu = APP.estacionariedadeResumo && APP.estacionariedadeResumo.usinas;
  if(eu && eu.nTotal){
    h += k('Tendência sig. (Qnat)', `${fmt(eu.pctSigCorr,1)}%`, `${eu.nSigCorr}/${eu.nTotal} postos de usinas`);
  }
  if(R.qincNulas && R.qincNulas.length){
    h += k('Qinc ≡ 0', R.qincNulas.length, R.qincNulas.join(', '));
  }
  el.innerHTML = h;

  const cl = document.getElementById('geralClasses');
  const lims = R.limites;
  const desc = {1:`> ${lims.c1}%`, 2:`${lims.c2}–${lims.c1}%`, 3:`< ${lims.c2}%`, 0:'nenhum mês'};
  cl.innerHTML = [1,2,3,0].map(c => `<div class="badge clickable ${geralClasse===c?'selected':''}" data-classe="${c}">
      <span class="dot" style="background:${CLASSE_INFO[c].cor}"></span>${CLASSE_INFO[c].txt} <span style="color:var(--text-dim)">(${desc[c]})</span>: <b>${R.classes[String(c)]}</b></div>`).join('')
    + (geralClasse!=null ? `<div class="badge clickable" data-classe="x"><span class="dot" style="background:${COLORS.textDim}"></span>limpar filtro</div>` : '');
  cl.querySelectorAll('[data-classe]').forEach(b => b.addEventListener('click', () => {
    const v = b.dataset.classe;
    geralClasse = (v==='x' || geralClasse===parseInt(v)) ? null : parseInt(v);
    renderGeralKpis(); renderGeralTabela();
  }));
}

function renderGeralCharts(){
  const R = RESUMO;
  if(!R) return;
  const MES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const ds = (serie, rot, cor) => ({label:rot, data:serie, backgroundColor:cor, borderWidth:0, barPercentage:0.8, categoryPercentage:0.75});
  const mk = (id, key, labels, xt) => {
    destroyChart(key);
    const sets = [ds(R[xt==='mes'?'porMes':'porDecada'].orig, 'Meses com Qinc < 0', 'rgba(242,118,91,0.75)')];
    charts[key] = new Chart(document.getElementById(id), {
      type:'bar', data:{labels, datasets:sets},
      options: baseOpts({plugins:{legend:{display:false}, tooltip:{...TOOLTIP}}},
        {xTitle: xt==='mes' ? 'Mês' : 'Década', yTitle:'Nº de meses', xTicks:{autoSkip:false, maxTicksLimit:14}})
    });
  };
  mk('chartGeralMes', 'geralMes', MES, 'mes');
  mk('chartGeralDec', 'geralDec', R.decadas.map(d => d+'s'), 'dec');
}

function renderGeralView(){
  if(!geralBuilt){
    geralBuilt = true;
    ['geralFiltro','geralSoNeg','geralSoEsp'].forEach(id => document.getElementById(id).addEventListener('input', renderGeralTabela));
    document.getElementById('btnCsvGeral').addEventListener('click', () => {
      const temEvap = !!APP.evaporacao;
      const cols = COLS_GERAL.filter(c => !c.evap || temEvap);
      exportarCSV('tabela_usinas.csv', cols.map(c => c.t), geralFiltradas().map(r => cols.map(c => {
        const v = r[c.k];
        return (typeof v === 'number' && !Number.isInteger(v)) ? v.toFixed(3).replace('.', ',') : v;
      })));
    });
    renderGeralCharts();
  }
  renderGeralKpis();
  renderGeralTabela();
}

// ============================================================
// USINAS VIEW
// ============================================================
function renderNatural(u){
  if(!u.natural){ wrapOf('chartNatural').innerHTML = '<div class="empty-note">Sem dados naturais para este posto</div>'; return; }
  const ctx = resetCanvas('chartNatural');
  const mode = modes.natural;
  destroyChart('natural');
  charts.natural = new Chart(ctx, {
    type:'line',
    data:{ labels: mode==='anual' ? u.natural.anos : u.natural.mensalLabels, datasets:[{
      label: mode==='anual' ? 'Média anual' : 'Mensal',
      data: mode==='anual' ? u.natural.media : u.natural.mensalValores,
      borderColor:COLORS.cyan, backgroundColor:'rgba(79,216,232,0.08)',
      borderWidth: mode==='anual' ? 1.6 : 0.9, pointRadius:0, pointHoverRadius:3, tension:0.1, fill:true
    }]},
    options: baseOpts({}, {xTitle: mode==='anual' ? 'Ano' : 'Mês/Ano', yTitle:'Qnat (m³/s)', xTicks: mode==='mensal' ? {maxTicksLimit:12} : {}})
  });
}

function renderIncremental(u){
  if(!u.incremental){ wrapOf('chartIncremental').innerHTML = '<div class="empty-note">Sem dados incrementais para este posto</div>'; return; }
  const ctx = resetCanvas('chartIncremental');
  const mode = modes.incremental;
  destroyChart('incremental');
  charts.incremental = new Chart(ctx, {
    type:'line',
    data:{ labels: mode==='anual' ? u.incremental.anos : u.incremental.mensalLabels, datasets:[{
      label: mode==='anual' ? 'Média anual' : 'Mensal',
      data: mode==='anual' ? u.incremental.media : u.incremental.mensalValores,
      borderWidth: mode==='anual' ? 1.6 : 0.9, pointRadius:0, pointHoverRadius:3, tension:0.1, fill:true,
      backgroundColor:'rgba(242,184,75,0.08)', borderColor:COLORS.amber,
      segment:{ borderColor: c => (c.p0.parsed.y<0 || c.p1.parsed.y<0) ? COLORS.coral : COLORS.amber }
    }]},
    options: comZero(baseOpts({}, {xTitle: mode==='anual' ? 'Ano' : 'Mês/Ano', yTitle:'Qinc (m³/s)', xTicks: mode==='mensal' ? {maxTicksLimit:12} : {}}))
  });
}

function renderEvap(u){
  if(!u.evaporacao){ wrapOf('chartEvap').innerHTML = '<div class="empty-note">Sem dados de evaporação para este posto</div>'; return; }
  const ctx = resetCanvas('chartEvap');
  const e = u.evaporacao;
  destroyChart('evap');
  charts.evap = new Chart(ctx, {
    type:'line',
    data:{ labels:e.meses, datasets:[
      {label:'Área Mín', data:e.min, borderColor:COLORS.cyanDim, borderWidth:1.4, pointRadius:2, tension:0.25, borderDash:[4,3]},
      {label:'Área Média', data:e.med, borderColor:COLORS.amber, borderWidth:2.2, pointRadius:2, tension:0.25},
      {label:'Área Máx', data:e.max, borderColor:COLORS.cyanDim, borderWidth:1.4, pointRadius:2, tension:0.25, borderDash:[6,2]},
    ]},
    options: baseOpts({plugins:{legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP}}}, {xTitle:'Mês', yTitle:'Evap (m³/s)', xTicks:{maxTicksLimit:12, autoSkip:false}})
  });
}

function renderBadges(u){
  const row = document.getElementById('badgesRow');
  let html = '';
  if(u.incremental){
    const pct = u.incremental.pctNegativos;
    const cls = pct>0 ? 'warn':'ok';
    html += `<div class="badge ${cls}"><span class="dot" style="background:${pct>0?COLORS.coral:COLORS.green}"></span>Qinc negativo: <b>${u.incremental.negativos}</b> de ${fmt(u.incremental.totalMeses)} meses <span style="color:var(--text-dim)">(${pct.toFixed(2)}%)</span></div>`;
  }
  if(u.qincNulo){
    html += `<div class="badge critical warn"><span class="dot" style="background:${COLORS.coral}"></span><b>Qinc ≡ 0</b> — posto igual ao de montante (incremental nula no deck)</div>`;
  }
  if(u.casoEspecial){
    html += `<div class="badge critical"><span class="dot" style="background:${COLORS.amber}"></span>Caso especial do projeto — derivação, transposição ou complexo hidrelétrico</div>`;
  }
  if(u.evaporacao){
    const e = u.evaporacao;
    html += `<div class="badge clickable" data-goto-evap="1" title="ver perfil mensal na aba Evaporação"><span class="dot" style="background:${COLORS.amber}"></span>Evaporação líquida: <b>${fmt(e.evapMed,2)} m³/s</b> <span style="color:var(--text-dim)">(${fmt(e.evapMin,2)}–${fmt(e.evapMax,2)}) · ${e.pctQnat!=null ? fmt(e.pctQnat,2)+'% da Qnat' : '—'} · nº ${e.rank} de ${APP.evaporacao ? APP.evaporacao.nUsinas : '—'}</span> →</div>`;
    html += `<div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Área Mín / Méd / Máx: <b>${fmt(e.areaMinKm2,1)} / ${fmt(e.areaMedKm2,1)} / ${fmt(e.areaMaxKm2,1)} km²</b> <span style="color:var(--text-dim)">· lâmina ${fmt(e.laminaAnualMm,0)} mm/ano</span></div>`;
    if(e.areaSuspeita){
      html += `<div class="badge warn critical"><span class="dot" style="background:${COLORS.coral}"></span>Área fora do plausível — <b>provável erro de dado</b> (PAC/cotas)</div>`;
    }
  }
  if(u.cascataIdx && u.cascataIdx.length){
    u.cascataIdx.forEach(idx=>{
      html += `<div class="badge clickable" data-goto-cascata="${idx}"><span class="dot" style="background:${COLORS.cyan}"></span>Cascata <b>#${String(idx).padStart(3,'0')}</b> →</div>`;
    });
  }
  row.innerHTML = html || `<div class="badge"><span class="dot" style="background:${COLORS.textDim}"></span>Sem indicadores disponíveis</div>`;
  row.querySelectorAll('[data-goto-evap]').forEach(el => el.addEventListener('click', () => {
    switchTab('evap'); document.getElementById('evapUsinaNav').scrollIntoView({block:'start'});
  }));
  row.querySelectorAll('[data-goto-cascata]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const found = CASCATAS.findIndex(c=>c.indice===parseInt(el.dataset.gotoCascata));
      if(found>=0){ cCurrent = found; switchTab('cascatas'); }
    });
  });
}

function atualizarNav(prefixo){
  document.getElementById(prefixo+'PosAtual').textContent = uCurrent+1;
  document.getElementById(prefixo+'PosTotal').textContent = USINAS.length;
  document.getElementById(prefixo+'ProgressFill').style.width = `${((uCurrent+1)/USINAS.length)*100}%`;
  document.getElementById(prefixo+'PrevBtn').disabled = uCurrent===0;
  document.getElementById(prefixo+'NextBtn').disabled = uCurrent===USINAS.length-1;
}
function identidade(prefixoIds, u){
  document.getElementById(prefixoIds.nome).textContent = u.nome;
  document.getElementById(prefixoIds.usina).textContent = `USINA ${u.codUsina}`;
  document.getElementById(prefixoIds.posto).textContent = u.codPosto!=null ? `POSTO ${u.codPosto}` : 'SEM POSTO';
}

function renderUsinaView(){
  const u = USINAS[uCurrent];
  identidade({nome:'usinaNome', usina:'tagUsina', posto:'tagPosto'}, u);
  const tc = document.getElementById('tagClasse');
  if(u.classe!=null){
    const inf = CLASSE_INFO[u.classe];
    tc.style.display=''; tc.textContent = inf.txt.toUpperCase(); tc.style.color = inf.cor; tc.style.borderColor = inf.cor;
  } else tc.style.display='none';
  const anosNat = u.natural ? `${u.natural.anos[0]}–${u.natural.anos[u.natural.anos.length-1]}` : '—';
  document.getElementById('usinaSub').textContent = `Série disponível (anos completos): ${anosNat}`;
  atualizarNav('u');
  renderBadges(u);
  renderNatural(u);
  renderIncremental(u);
}

function goToUsina(idx){
  if(idx<0 || idx>=USINAS.length) return;
  uCurrent = idx;
  if(!TABS_POR_USINA.includes(activeTab)) switchTab('usinas');
  else renderAtivo();
  fecharBusca();
}
['u','e','s','p','v','k'].forEach(p => {
  document.getElementById(p+'PrevBtn').addEventListener('click', ()=>goToUsina(uCurrent-1));
  document.getElementById(p+'NextBtn').addEventListener('click', ()=>goToUsina(uCurrent+1));
});

document.querySelectorAll('.toggle-group[data-chart="natural"], .toggle-group[data-chart="incremental"]').forEach(group=>{
  group.addEventListener('click', (e)=>{
    const btn = e.target.closest('button'); if(!btn) return;
    const key = group.dataset.chart;
    modes[key] = btn.dataset.mode;
    group.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
    if(key==='natural') renderNatural(USINAS[uCurrent]);
    if(key==='incremental') renderIncremental(USINAS[uCurrent]);
  });
});

// ============================================================
// CASCATAS VIEW
// ============================================================
function renderCascataView(){
  const c = CASCATAS[cCurrent];
  document.getElementById('cascataNomeH2').textContent = `Cascata #${String(c.indice).padStart(3,'0')}`;
  document.getElementById('cascataTagCount').textContent = `${c.usinas.length} usina${c.usinas.length>1?'s':''}`;
  document.getElementById('cascataSub').textContent = c.nomes.join(' → ');
  document.getElementById('cPosAtual').textContent = cCurrent+1;
  document.getElementById('cPosTotal').textContent = CASCATAS.length;
  document.getElementById('cProgressFill').style.width = `${((cCurrent+1)/CASCATAS.length)*100}%`;
  document.getElementById('cPrevBtn').disabled = cCurrent===0;
  document.getElementById('cNextBtn').disabled = cCurrent===CASCATAS.length-1;

  const mode = modes.cascata;
  const series = mode==='natural' ? c.seriesNatural : c.seriesIncremental;
  document.getElementById('cascataChartTitle').textContent = `Série em Cascata — ${mode==='natural' ? 'Natural' : 'Incremental'} (média anual)`;
  const wrap = document.querySelector('#viewCascatas .card.wide .chart-wrap');

  if(!series || series.length===0){
    wrap.innerHTML = `<div class="empty-note">Sem dados ${mode==='natural'?'naturais':'incrementais'} para as usinas desta cascata</div>`;
  } else {
    wrap.innerHTML = '<canvas id="chartCascata"></canvas>';
    destroyChart('cascata');
    // eixo = uniao dos anos (usinas podem ter series de tamanhos diferentes)
    const anos = [...new Set(series.flatMap(s => s.anos))].sort((a,b)=>a-b);
    const datasets = series.map((s,i)=>{
      const mapa = new Map(s.anos.map((a,j)=>[a, s.media[j]]));
      const cor = PALETTE[i % PALETTE.length];
      const base = {label: s.nome, data: anos.map(a => mapa.has(a) ? mapa.get(a) : null),
        borderColor: cor, backgroundColor: cor, borderWidth: 1.6, pointRadius:0, pointHoverRadius:3, tension:0.15, spanGaps:false};
      if(mode==='incremental'){
        base.segment = { borderColor: ctx => (ctx.p0.parsed.y<0 || ctx.p1.parsed.y<0) ? COLORS.coral : cor };
      }
      return base;
    });
    const opts = baseOpts({ plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP} } },
      {xTitle:'Ano', yTitle: mode==='natural' ? 'Qnat (m³/s)' : 'Qinc (m³/s)'});
    charts.cascata = new Chart(document.getElementById('chartCascata'), {
      type:'line', data:{ labels:anos, datasets }, options: mode==='incremental' ? comZero(opts) : opts
    });
  }

  const badgesEl = document.getElementById('cascataBadges');
  badgesEl.innerHTML = c.seriesIncremental.map(s=>{
    const pct = s.totalMeses ? (s.negativos/s.totalMeses*100) : 0;
    const cls = s.negativos>0 ? 'warn' : 'ok';
    const idx = USINAS.findIndex(u => u.codUsina===s.codUsina);
    return `<div class="badge clickable ${cls}" data-goto-usina="${idx}" title="abrir usina"><span class="dot" style="background:${s.negativos>0?COLORS.coral:COLORS.green}"></span>${esc(s.nome)}: <b>${s.negativos}</b> meses Qinc &lt; 0 <span style="color:var(--text-dim)">(${pct.toFixed(1)}%)</span></div>`;
  }).join('');
  badgesEl.querySelectorAll('[data-goto-usina]').forEach(el => el.addEventListener('click', () => {
    const i = parseInt(el.dataset.gotoUsina); if(i>=0){ uCurrent = i; switchTab('usinas'); }
  }));

  renderTopoView();
}

function goToCascata(idx){
  if(idx<0 || idx>=CASCATAS.length) return;
  cCurrent = idx;
  renderCascataView();
}
document.getElementById('cPrevBtn').addEventListener('click', ()=>goToCascata(cCurrent-1));
document.getElementById('cNextBtn').addEventListener('click', ()=>goToCascata(cCurrent+1));

document.querySelector('.toggle-group[data-chart="cascata"]').addEventListener('click', (e)=>{
  const btn = e.target.closest('button'); if(!btn) return;
  modes.cascata = btn.dataset.mode;
  e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
  renderCascataView();
});

// ============================================================
// TOPOLOGIA (dentro de Cascatas)
// ============================================================
let topoRendered = false;
function renderTopoView(){
  if(topoRendered) return;
  topoRendered = true;
  const t = APP.topologia;
  if(!t) return;
  document.getElementById('topoResumoRow').innerHTML = `
    <div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Total de usinas: <b>${t.totalUsinas}</b></div>
    <div class="badge warn"><span class="dot" style="background:${COLORS.amber}"></span>Em cascatas isoladas: <b>${t.totalIsoladas}</b> <span style="color:var(--text-dim)">(${(100*t.totalIsoladas/t.totalUsinas).toFixed(1)}%)</span></div>
    <div class="badge critical warn"><span class="dot" style="background:${COLORS.coral}"></span>Casos especiais entre elas: <b>${t.lista.filter(x=>x.casoEspecial).length}</b></div>
  `;
  const tbody = document.getElementById('topoTbody');
  tbody.innerHTML = t.lista.map(i => {
    const idx = USINAS.findIndex(u => u.codUsina===i.codUsina);
    return `<tr class="clickrow" data-idx="${idx}">
      <td style="${i.casoEspecial?'color:'+COLORS.amber+';font-weight:600;':''}">${esc(i.nome)}${i.casoEspecial?' ⚠':''}</td>
      <td class="num" style="color:var(--text-dim);">${i.codPosto ?? '—'}</td>
      <td class="num">${fmt(i.mediaAnual,1)}</td>
      <td style="color:var(--text-dim);">${i.casoEspecial ? 'ligação física conhecida, ausente no cadastro' : 'provável cabeceira real'}</td>
    </tr>`;}).join('');
  tbody.querySelectorAll('tr').forEach(tr => tr.addEventListener('click', () => {
    const i = parseInt(tr.dataset.idx); if(i>=0){ uCurrent = i; switchTab('usinas'); }
  }));
  document.getElementById('btnCsvTopo').addEventListener('click', () => {
    exportarCSV('topologia_isolada.csv', ['Usina','Posto','Vazao_Media_Anual_m3s','Caso_Especial'],
      t.lista.map(i => [i.nome, i.codPosto, i.mediaAnual, i.casoEspecial ? 'SIM' : 'NAO']));
  });
}

// ============================================================
// ESTACIONARIEDADE VIEW
// ============================================================
function resumoEstacAtual(){
  const R = APP.estacionariedadeResumo;
  if(!R) return null;
  return R.usinas ? R[modes.estacEscopo] : R;   // compat. com JSON antigo
}

function renderEstacResumo(){
  wireCsvResumoEstac();
  const r = resumoEstacAtual();
  const row = document.getElementById('estacResumoRow');
  if(!r){ row.innerHTML=''; return; }
  row.innerHTML = `
    <div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Postos analisados: <b>${r.nTotal}</b></div>
    <div class="badge warn"><span class="dot" style="background:${COLORS.coral}"></span>Tendência significativa (MK corrigido, p&lt;0,05): <b>${r.nSigCorr}</b> <span style="color:var(--text-dim)">(${r.pctSigCorr}% · ~${r.fpEsperados ?? '—'} esperados por acaso)</span></div>
    <div class="badge"><span class="dot" style="background:${COLORS.coral}"></span>Decrescente: <b>${r.nDecr}</b></div>
    <div class="badge"><span class="dot" style="background:${COLORS.green}"></span>Crescente: <b>${r.nCresc}</b></div>
    <div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Quebra significativa (Pettitt): <b>${r.nBreakSig}</b> <span style="color:var(--text-dim)">(${r.pctBreakSig}%${r.nBreakSemTend!=null ? ` · ${r.nBreakSemTend} sem tendência MK` : ''})</span></div>
    <div class="badge"><span class="dot" style="background:${COLORS.textDim}"></span>Autocorrelação lag-1 sig. (Yue-Wang): <b>${r.nR1Sig}</b> <span style="color:var(--text-dim)">(${r.pctR1Sig}%)</span></div>
  `;
  destroyChart('estacDecadas');
  charts.estacDecadas = new Chart(resetCanvas('chartEstacDecadas'), {
    type:'bar',
    data:{ labels:r.decadaLabels.map(d=>d+'s'), datasets:[{
      label:'Postos com quebra significativa', data:r.decadaValores,
      backgroundColor:'rgba(79,216,232,0.55)', borderWidth:0, barPercentage:0.7, categoryPercentage:0.7
    }]},
    options: baseOpts({}, {xTitle:'Década do ponto de quebra (Pettitt, p<0,05)', yTitle:'Nº de postos'})
  });
}
document.querySelector('.toggle-group[data-chart="estacEscopo"]').addEventListener('click', (e)=>{
  const btn = e.target.closest('button'); if(!btn) return;
  modes.estacEscopo = btn.dataset.mode;
  e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
  renderEstacResumo();
});

function renderEstac(u){
  if(!u.estacionariedade){ wrapOf('chartEstac').innerHTML = '<div class="empty-note">Sem dados de estacionariedade para este posto (série &lt; 30 anos)</div>'; return; }
  const e = u.estacionariedade;
  destroyChart('estac');
  charts.estac = new Chart(resetCanvas('chartEstac'), {
    type:'line',
    data:{ labels:e.anos, datasets:[
      {label:'Vazão natural anual', data:e.natural, borderColor:COLORS.cyan, backgroundColor:'rgba(79,216,232,0.06)', borderWidth:1.4, pointRadius:0, pointHoverRadius:3, tension:0.1, fill:true},
      {label:`Tendência de Sen (${e.senSlope>=0?'+':''}${e.senSlope} m³/s/ano)`, data:e.trendLine, borderColor:COLORS.amber, borderWidth:1.8, borderDash:[6,4], pointRadius:0, tension:0},
      {label:`Média pré-quebra (até ${e.pettittAno})`, data:e.preLine, borderColor:COLORS.green, borderWidth:2.4, pointRadius:0, spanGaps:false},
      {label:`Média pós-quebra (${e.pettittAno+1}+)`, data:e.posLine, borderColor:COLORS.coral, borderWidth:2.4, pointRadius:0, spanGaps:false},
    ]},
    options: baseOpts({ plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP} } }, {xTitle:'Ano', yTitle:'Qnat (m³/s)'})
  });
}

function renderEstacBadges(u){
  const row = document.getElementById('estacBadgesRow');
  if(!u.estacionariedade){ row.innerHTML = ''; return; }
  const e = u.estacionariedade;
  const clsTrend = e.mkTrendCorr==='crescente' ? 'ok' : (e.mkTrendCorr==='decrescente' ? 'warn' : '');
  const dotTrend = e.mkTrendCorr==='crescente' ? COLORS.green : (e.mkTrendCorr==='decrescente' ? COLORS.coral : COLORS.textDim);
  const media = e.natural.reduce((a,b)=>a+b,0)/e.natural.length;
  const rel = media ? e.senSlope/media*100 : 0;
  let html = `<div class="badge ${clsTrend}"><span class="dot" style="background:${dotTrend}"></span>Mann-Kendall (corrigido): <b>${e.mkTrendCorr}</b> <span style="color:var(--text-dim)">(p=${e.mkPCorr.toExponential(2)}; bruto p=${e.mkPBruto.toExponential(2)})</span></div>`;
  html += `<div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Declividade de Sen: <b>${e.senSlope} m³/s/ano</b> <span style="color:var(--text-dim)">(${rel>=0?'+':''}${rel.toFixed(2)}% da média/ano)</span></div>`;
  html += `<div class="badge"><span class="dot" style="background:${COLORS.textDim}"></span>Autocorrelação lag-1: <b>r1=${e.r1}</b> <span style="color:var(--text-dim)">(${e.r1Sig?'corrigida via pre-whitening':'não significativa'})</span></div>`;
  const clsBreak = e.pettittP<0.05 ? 'critical warn' : '';
  html += `<div class="badge ${clsBreak}"><span class="dot" style="background:${e.pettittP<0.05?COLORS.coral:COLORS.textDim}"></span>Pettitt — ponto de quebra: <b>${e.pettittAno}</b> <span style="color:var(--text-dim)">(p=${e.pettittP.toExponential(2)})</span></div>`;
  if(e.mediaPre!=null && e.mediaPos!=null){
    const delta = e.mediaPos - e.mediaPre;
    const pct = e.mediaPre ? (delta/e.mediaPre*100) : 0;
    html += `<div class="badge"><span class="dot" style="background:${delta>=0?COLORS.green:COLORS.coral}"></span>Média pré→pós: <b>${fmt(e.mediaPre,1)} → ${fmt(e.mediaPos,1)} m³/s</b> <span style="color:var(--text-dim)">(${pct>=0?'+':''}${pct.toFixed(1)}%)</span></div>`;
  }
  row.innerHTML = html;
}

function renderJanelaMovel(u){
  const tamanho = modes.janela;
  document.getElementById('janelaMovelTitulo').textContent = `Tendência em Janela Móvel (${tamanho} anos)`;
  const jm = (u.estacionariedade && u.estacionariedade.janelaMovel) ? u.estacionariedade.janelaMovel[tamanho] : null;
  if(!jm){ wrapOf('chartJanelaMovel').innerHTML = `<div class="empty-note">Série curta demais para janela móvel de ${tamanho} anos</div>`; return; }
  const corPonto = jm.tendencias.map(t => t==='crescente' ? COLORS.green : (t==='decrescente' ? COLORS.coral : COLORS.textDim));
  const raioPonto = jm.tendencias.map(t => t==='neutro' ? 2 : 4);
  destroyChart('janelaMovel');
  charts.janelaMovel = new Chart(resetCanvas('chartJanelaMovel'), {
    type:'line',
    data:{ labels:jm.anosFinal, datasets:[{
      label:`Declividade de Sen (janela de ${tamanho} anos)`, data:jm.senSlopes,
      borderColor:COLORS.textDim, borderWidth:1.2,
      pointBackgroundColor:corPonto, pointBorderColor:corPonto, pointRadius:raioPonto, pointHoverRadius:6,
      tension:0.2, fill:false,
    }]},
    options: comZero(baseOpts({ plugins:{ legend:{display:false},
      tooltip:{...TOOLTIP, callbacks:{ label: it => {
        const i=it.dataIndex; return [`Janela ${jm.anosFinal[i]-tamanho+1}–${jm.anosFinal[i]}`, `Sen: ${jm.senSlopes[i]} m³/s/ano`, `p=${jm.pValores[i]} (${jm.tendencias[i]})`];
      }}}
    }}, {xTitle:`Ano final da janela (${tamanho} anos)`, yTitle:'Sen (m³/s/ano)'}))
  });
}

document.querySelector('.toggle-group[data-chart="janela"]').addEventListener('click', (e)=>{
  const btn = e.target.closest('button'); if(!btn) return;
  modes.janela = btn.dataset.mode;
  e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
  renderJanelaMovel(USINAS[uCurrent]);
});

let estacResumoFeito = false;
function renderEstacView(){
  if(!estacResumoFeito){ estacResumoFeito = true; renderEstacResumo(); }
  const u = USINAS[uCurrent];
  identidade({nome:'estacUsinaNome', usina:'estacTagUsina', posto:'estacTagPosto'}, u);
  document.getElementById('estacUsinaSub').textContent = u.estacionariedade
    ? `Série testada: ${u.estacionariedade.anos[0]}–${u.estacionariedade.anos[u.estacionariedade.anos.length-1]} (Mann-Kendall + Pettitt, correção de autocorrelação Yue-Wang)`
    : 'Sem dados suficientes para o teste (< 30 anos)';
  atualizarNav('e');
  renderEstacBadges(u);
  renderEstac(u);
  renderJanelaMovel(u);
}

// ============================================================
// SAZONALIDADE VIEW
// ============================================================
const MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

function renderSazBadges(u){
  const rowSaz = document.getElementById('sazBadgesRow');
  const rowExt = document.getElementById('extremosBadgesRow');
  if(!u.sazonalidade){ rowSaz.innerHTML=''; rowExt.innerHTML=''; return; }
  const s = u.sazonalidade;
  rowSaz.innerHTML = `
    <div class="badge ok"><span class="dot" style="background:${COLORS.green}"></span>Mês mais cheio: <b>${s.mesMaisCheio}</b> <span style="color:var(--text-dim)">(${fmt(s.porMes[s.mesMaisCheio].media,1)} m³/s em média)</span></div>
    <div class="badge warn"><span class="dot" style="background:${COLORS.coral}"></span>Mês mais seco: <b>${s.mesMaisSeco}</b> <span style="color:var(--text-dim)">(${fmt(s.porMes[s.mesMaisSeco].media,1)} m³/s em média)</span></div>
    <div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Índice de sazonalidade: <b>${s.indice!=null ? (s.indice*100).toFixed(1)+'%' : '—'}</b> <span style="color:var(--text-dim)">(cheio − seco) / média</span></div>
  `;
  if(u.extremos){
    const e = u.extremos;
    let html = `<div class="badge"><span class="dot" style="background:${COLORS.cyanDim}"></span>Mín. mensal: <b>${fmt(e.minMensal.valor,1)} m³/s</b> <span style="color:var(--text-dim)">(${e.minMensal.quando})</span></div>`;
    html += `<div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Máx. mensal: <b>${fmt(e.maxMensal.valor,1)} m³/s</b> <span style="color:var(--text-dim)">(${e.maxMensal.quando})</span></div>`;
    if(e.minAnual) html += `<div class="badge"><span class="dot" style="background:${COLORS.coral}"></span>Ano mais seco: <b>${e.minAnual.ano}</b> <span style="color:var(--text-dim)">(${fmt(e.minAnual.valor,1)} m³/s)</span></div>`;
    if(e.maxAnual) html += `<div class="badge"><span class="dot" style="background:${COLORS.green}"></span>Ano mais cheio: <b>${e.maxAnual.ano}</b> <span style="color:var(--text-dim)">(${fmt(e.maxAnual.valor,1)} m³/s)</span></div>`;
    rowExt.innerHTML = html;
  } else rowExt.innerHTML = '';
}

function renderSaz(u){
  if(!u.sazonalidade){ wrapOf('chartSaz').innerHTML = '<div class="empty-note">Sem dados de sazonalidade para este posto</div>'; return; }
  const s = u.sazonalidade;
  const meses = MESES_PT.filter(m => s.porMes[m]);
  destroyChart('saz');
  charts.saz = new Chart(resetCanvas('chartSaz'), {
    type:'bar',
    data:{ labels:meses, datasets:[
      {type:'bar', label:'Faixa P10–P90', data:meses.map(m => [s.porMes[m].p10, s.porMes[m].p90]), backgroundColor:'rgba(79,216,232,0.28)', borderRadius:4, barPercentage:0.55, categoryPercentage:0.7, order:2},
      {type:'line', label:'Mediana', data:meses.map(m => s.porMes[m].mediana), borderColor:COLORS.amber, backgroundColor:COLORS.amber, borderWidth:0, pointRadius:5, pointStyle:'rectRot', showLine:false, order:1},
      {type:'line', label:'Média', data:meses.map(m => s.porMes[m].media), borderColor:COLORS.coral, backgroundColor:COLORS.coral, borderWidth:0, pointRadius:4, pointStyle:'circle', showLine:false, order:0},
    ]},
    options: baseOpts({ plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP} } }, {xTitle:'Mês', yTitle:'Qnat (m³/s)', xTicks:{maxTicksLimit:12, autoSkip:false}})
  });
}

function renderSazView(){
  const u = USINAS[uCurrent];
  identidade({nome:'sazUsinaNome', usina:'sazTagUsina', posto:'sazTagPosto'}, u);
  document.getElementById('sazUsinaSub').textContent = 'Distribuição da vazão natural por mês do calendário (todos os anos históricos agrupados)';
  atualizarNav('s');
  renderSazBadges(u);
  renderSaz(u);
}

// ============================================================
// PERMANÊNCIA VIEW
// ============================================================
function renderPermBadges(u){
  const row = document.getElementById('permBadgesRow');
  if(!u.permanencia){ row.innerHTML=''; return; }
  const p = u.permanencia;
  row.innerHTML = `
    <div class="badge"><span class="dot" style="background:${COLORS.green}"></span>Q10: <b>${fmt(p.q10,1)} m³/s</b></div>
    <div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Q50: <b>${fmt(p.q50,1)} m³/s</b></div>
    <div class="badge warn"><span class="dot" style="background:${COLORS.coral}"></span>Q95: <b>${fmt(p.q95,1)} m³/s</b></div>
    <div class="badge"><span class="dot" style="background:${COLORS.textDim}"></span>Q10/Q95: <b>${p.q95>0 ? fmt(p.q10/p.q95,1) : '—'}</b> <span style="color:var(--text-dim)">(variabilidade)</span></div>
  ` + (Math.min(...p.curvaGeral) <= 0 ? `<div class="badge warn"><span class="dot" style="background:${COLORS.coral}"></span>Q ≤ 0 em parte da curva — <b>omitido no eixo log</b></div>` : '');
}

function permPts(pontos, vals){
  return pontos.map((p,i) => ({x:p, y:(vals[i]!=null && vals[i]>0) ? vals[i] : null}));
}

function renderPerm(u, mesSelecionado){
  if(!u.permanencia){ wrapOf('chartPerm').innerHTML = '<div class="empty-note">Sem dados de permanência para este posto</div>'; return; }
  const p = u.permanencia;
  const sobre = mesSelecionado && p.curvaMensal[mesSelecionado];
  const datasets = [{
    label:'Curva geral', data:permPts(p.pontosGeral, p.curvaGeral),
    borderColor: sobre ? COLORS.textDim : COLORS.cyan, borderDash: sobre ? [4,4] : [],
    backgroundColor:'rgba(79,216,232,0.06)', borderWidth:1.8, pointRadius:2, fill:!sobre, tension:0.15, spanGaps:true
  }];
  if(sobre){
    datasets.push({label:`Curva — ${mesSelecionado}`, data:permPts(p.pontosMensal, p.curvaMensal[mesSelecionado]),
      borderColor:COLORS.amber, backgroundColor:'rgba(242,184,75,0.08)', borderWidth:2, pointRadius:3, fill:true, tension:0.15, spanGaps:true});
  }
  destroyChart('perm');
  const opts = baseOpts({ plugins:{ legend:LEGEND_BOTTOM,
      tooltip:{...TOOLTIP, callbacks:{ title: it => `${it[0].parsed.x}% do tempo` }} } },
    {xTitle:'% do tempo em que a vazão é igualada ou superada', yTitle:'Qnat (m³/s) — escala log'});
  opts.interaction = {mode:'nearest', axis:'x', intersect:false};
  opts.scales.x.type = 'linear';
  opts.scales.x.min = 0; opts.scales.x.max = 100;
  opts.scales.x.ticks = {...opts.scales.x.ticks, stepSize:10, maxTicksLimit:11, callback: v => v + '%'};
  opts.scales.y.type = 'logarithmic';
  charts.perm = new Chart(resetCanvas('chartPerm'), { type:'line', data:{ datasets }, options: opts });
}

function renderPermView(){
  const u = USINAS[uCurrent];
  identidade({nome:'permUsinaNome', usina:'permTagUsina', posto:'permTagPosto'}, u);
  document.getElementById('permUsinaSub').textContent = 'Curva de permanência da vazão natural mensal (Weibull) — geral, com opção de sobrepor um mês';
  atualizarNav('p');
  const sel = document.getElementById('permMesSelect');
  if(sel.dataset.built !== '1'){
    MESES_PT.forEach(m => sel.appendChild(new Option(m, m)));
    sel.dataset.built = '1';
    sel.addEventListener('change', () => renderPerm(USINAS[uCurrent], sel.value || null));
  }
  renderPermBadges(u);
  renderPerm(u, sel.value || null);
}

// ============================================================
// EVAPORACAO VIEW (perdas; nao entra no Qinc)
// ============================================================
const EVAP = APP.evaporacao || null;
modes.evapMetrica = 'm3s';
modes.evapRelX = 'pctQinc';
const escalaFaixa = (e, v, k) => (v==null ? null : (e.evapMed>0 ? v*e[k]/e.evapMed : v));
const EVAP_MET = {
  m3s:    {rot:'m³/s',      eixo:'Evaporação líquida média (m³/s)', v:e=>e.evapMed, lo:e=>e.evapMin, hi:e=>e.evapMax, d:2, un:' m³/s'},
  pctQnat:{rot:'% da Qnat', eixo:'Evaporação ÷ vazão natural média (%)', v:e=>e.pctQnat, lo:e=>escalaFaixa(e,e.pctQnat,'evapMin'), hi:e=>escalaFaixa(e,e.pctQnat,'evapMax'), d:2, un:'%'},
  pctQinc:{rot:'% da Qinc', eixo:'Evaporação ÷ vazão incremental média (%)', v:e=>e.pctQinc, lo:e=>escalaFaixa(e,e.pctQinc,'evapMin'), hi:e=>escalaFaixa(e,e.pctQinc,'evapMax'), d:1, un:'%'},
};
const COLS_EVAP = [
  {k:'rank', t:'Nº', num:true, d:0},
  {k:'nome', t:'Usina'},
  {k:'bacia', t:'Bacia até'},
  {k:'area', t:'Área méd. (km²)', num:true, d:1},
  {k:'lamina', t:'Lâmina (mm/ano)', num:true, d:0},
  {k:'evap', t:'Evap. (m³/s)', num:true, d:2},
  {k:'faixa', t:'Faixa mín–máx'},
  {k:'vol', t:'Volume (hm³/ano)', num:true, d:0},
  {k:'qnat', t:'Qnat (m³/s)', num:true, d:1},
  {k:'pctQnat', t:'Evap/Qnat (%)', num:true, d:2},
  {k:'qinc', t:'Qinc (m³/s)', num:true, d:1},
  {k:'pctQinc', t:'Evap/Qinc (%)', num:true, d:1},
  {k:'pctNeg', t:'% meses Qinc < 0', num:true, d:2},
  {k:'classe', t:'Classe'},
];
let evapSort = {col:'rank', dir:1};
let evapBuilt = false;

function linhasEvap(){
  return USINAS.map((u,i) => ({u,i,e:u.evaporacao})).filter(x => x.e).map(({u,i,e}) => ({
    i, rank:e.rank, nome:u.nome, bacia:e.terminalNome, area:e.areaMedKm2, lamina:e.laminaAnualMm,
    evap:e.evapMed, faixa:`${fmt(e.evapMin,2)}–${fmt(e.evapMax,2)}`, vol:e.volAnualHm3,
    qnat:e.qnatMedia, pctQnat:e.pctQnat, qinc:e.qincMedia, pctQinc:e.pctQinc,
    pctNeg: u.incremental ? u.incremental.pctNegativos : null, classe:u.classe,
    suspeita:e.areaSuspeita, especial:u.casoEspecial,
  }));
}
function evapFiltradas(){
  const txt = normTxt(document.getElementById('evapFiltro').value.trim());
  const L = linhasEvap().filter(r => !txt || normTxt(r.nome).includes(txt) || normTxt(r.bacia||'').includes(txt));
  const {col, dir} = evapSort;
  L.sort((a,b) => {
    const va=a[col], vb=b[col];
    if(va==null && vb==null) return 0;
    if(va==null) return 1;
    if(vb==null) return -1;
    if(typeof va==='string') return va.localeCompare(vb, 'pt-BR')*dir;
    return (va-vb)*dir;
  });
  return L;
}
function renderEvapTabela(){
  document.getElementById('evapHead').innerHTML = COLS_EVAP.map(c => {
    const seta = evapSort.col===c.k ? (evapSort.dir>0 ? ' ▲' : ' ▼') : '';
    return `<th class="sortable ${c.num?'num':''}" data-col="${c.k}">${c.t}${seta}</th>`;
  }).join('');
  document.querySelectorAll('#evapHead th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.col;
    evapSort = {col:k, dir: evapSort.col===k ? -evapSort.dir : (['rank','nome','bacia','faixa'].includes(k) ? 1 : -1)};
    renderEvapTabela();
  }));
  const L = evapFiltradas();
  document.getElementById('evapCount').textContent = `${L.length} usinas`;
  document.getElementById('evapTbody').innerHTML = L.map(r => {
    const cel = c => {
      const v = r[c.k];
      if(c.k==='classe') return pillClasse(v);
      if(c.k==='nome') return esc(v) + (r.especial ? ` <span style="color:${COLORS.amber}" title="caso especial do projeto">⚠</span>` : '')
                              + (r.suspeita ? ` <span style="color:${COLORS.coral}" title="área fora do plausível">!</span>` : '');
      if(!c.num) return esc(v ?? '—');
      return v==null ? '—' : fmt(v, c.d);
    };
    return `<tr class="clickrow" data-idx="${r.i}">` + COLS_EVAP.map(c => `<td class="${c.num?'num':''}">${cel(c)}</td>`).join('') + '</tr>';
  }).join('');
  document.querySelectorAll('#evapTbody tr').forEach(tr => tr.addEventListener('click', () => selecionarEvap(parseInt(tr.dataset.idx))));
}

function renderEvapKpis(){
  const el = document.getElementById('evapKpis');
  const k = (lbl, val, sub, cls='') => `<div class="kpi ${cls}"><div class="lbl">${lbl}</div><div class="val">${val}</div><div class="sub">${sub}</div></div>`;
  const com = USINAS.filter(u => u.evaporacao);
  const porM3s = [...com].sort((a,b) => a.evaporacao.rank - b.evaporacao.rank);
  let acum = 0, n80 = null;
  porM3s.forEach((u,j) => { acum += u.evaporacao.evapMed; if(n80==null && acum >= 0.8*EVAP.totalMed) n80 = j+1; });
  const porPct = com.filter(u => u.evaporacao.pctQnat!=null).sort((a,b) => b.evaporacao.pctQnat - a.evaporacao.pctQnat);
  const t1 = porM3s[0], p1 = porPct[0];
  const rel = EVAP.relacao && EVAP.relacao.pctQinc;
  let h = k('Evaporação líquida total', `${fmt(EVAP.totalMed,0)} m³/s`, `faixa ${fmt(EVAP.totalMin,0)}–${fmt(EVAP.totalMax,0)} m³/s · ${EVAP.nUsinas} usinas`, 'amber');
  h += k('Volume anual', `${fmt(EVAP.volAnualKm3,1)} km³`, EVAP.pctSIN!=null ? `${fmt(EVAP.pctSIN,2)}% da vazão natural das bacias` : '—', 'amber');
  if(t1) h += k('Maior perda absoluta', esc(t1.nome), `${fmt(t1.evaporacao.evapMed,1)} m³/s · ${fmt(t1.evaporacao.pctQnat,2)}% da Qnat`, 'cyan');
  if(p1) h += k('Maior perda relativa', esc(p1.nome), `${fmt(p1.evaporacao.pctQnat,2)}% da Qnat · ${fmt(p1.evaporacao.evapMed,1)} m³/s`, 'cyan');
  if(n80) h += k('Concentração', `${n80} usinas`, `respondem por 80% da evaporação total`);
  if(rel) h += k('Relação com Qinc < 0', `ρ = ${fmt(rel.rho,2)}`, `Spearman, Evap/Qinc × % meses Qinc<0 · p = ${rel.p.toExponential(1)} · n = ${rel.n}`);
  el.innerHTML = h;
}

function renderEvapRank(){
  const met = EVAP_MET[modes.evapMetrica];
  const n = parseInt(document.getElementById('evapTopN').value);
  const L = USINAS.map((u,i) => ({u,i,e:u.evaporacao})).filter(x => x.e && met.v(x.e)!=null)
    .sort((a,b) => met.v(b.e) - met.v(a.e)).slice(0, n);
  destroyChart('evapRank');
  const opts = baseOpts({ indexAxis:'y',
    plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP, callbacks:{
      label: it => {
        const x = L[it.dataIndex];
        if(it.datasetIndex===0) return `Faixa (Área Mín–Máx): ${fmt(met.lo(x.e),met.d)}–${fmt(met.hi(x.e),met.d)}${met.un}`;
        return `Área Méd: ${fmt(met.v(x.e),met.d)}${met.un}`;
      },
      afterBody: its => {
        const e = L[its[0].dataIndex].e;
        return [`Evap ${fmt(e.evapMed,2)} m³/s · ${fmt(e.volAnualHm3,0)} hm³/ano`, `Área ${fmt(e.areaMedKm2,1)} km² · lâmina ${fmt(e.laminaAnualMm,0)} mm/ano`,
                `Qnat ${fmt(e.qnatMedia,1)} · Qinc ${fmt(e.qincMedia,1)} m³/s`];
      }}}},
    onClick: (ev, els) => { if(els.length) selecionarEvap(L[els[0].index].i); }
  }, {xTitle: met.eixo, yTicks:{autoSkip:false, maxTicksLimit:100}});
  opts.interaction = {mode:'index', axis:'y', intersect:false};
  opts.scales.x.min = 0;
  charts.evapRank = new Chart(resetCanvas('chartEvapRank'), {
    type:'bar',
    data:{ labels:L.map(x => x.u.nome), datasets:[
      {label:'Faixa Área Mín–Máx', data:L.map(x => [met.lo(x.e), met.hi(x.e)]), backgroundColor:'rgba(242,184,75,0.35)', borderWidth:0, grouped:false, barPercentage:0.9, categoryPercentage:0.9},
      {label:'Área Méd (referência)', data:L.map(x => met.v(x.e)), backgroundColor:COLORS.amber, borderWidth:0, grouped:false, barPercentage:0.4, categoryPercentage:0.9},
    ]},
    options: opts
  });
  const extra = modes.evapMetrica==='pctQinc' ? ' Usinas com Qinc médio ≤ 0 ficam de fora. Um valor alto indica que a evaporação que o ONS devolveu à série é grande frente à incremental — a série dessa usina é sensível a erros na estimativa de evaporação.'
              : modes.evapMetrica==='pctQnat' ? ' Perda relativa: fração da vazão natural média do posto que evapora no reservatório da própria usina.' : '';
  document.getElementById('evapRankNota').innerHTML = `Barra cheia: área na cota média (referência); faixa translúcida: entre as áreas nas cotas mínima e máxima. Clique numa barra para ver o perfil mensal da usina (fim da página).${extra}`;
}

function renderEvapBacia(){
  const B = (EVAP.bacias || []).filter(b => b.pctQnat!=null).sort((a,b) => b.pctQnat - a.pctQnat).slice(0, 20);
  destroyChart('evapBacia');
  if(!B.length){ wrapOf('chartEvapBacia').innerHTML = '<div class="empty-note">Sem bacias com vazão natural na usina de jusante final</div>'; return; }
  const opts = baseOpts({ indexAxis:'y',
    plugins:{ legend:{display:false}, tooltip:{...TOOLTIP, callbacks:{
      label: it => `Perda: ${fmt(B[it.dataIndex].pctQnat,2)}% da Qnat`,
      afterBody: its => { const b = B[its[0].dataIndex];
        return [`Evap ${fmt(b.evapMed,1)} m³/s (${fmt(b.evapMin,1)}–${fmt(b.evapMax,1)}) · ${fmt(b.volAnualHm3,0)} hm³/ano`,
                `Qnat na jusante final ${fmt(b.qnatExutorio,0)} m³/s · ${b.nUsinas} usinas`, `Principais: ${b.principais.join(', ')}`]; }
    }}}
  }, {xTitle:'Evaporação da bacia ÷ Qnat na usina de jusante final (%)', yTicks:{autoSkip:false, maxTicksLimit:100}});
  opts.scales.x.min = 0;
  charts.evapBacia = new Chart(resetCanvas('chartEvapBacia'), {
    type:'bar',
    data:{ labels:B.map(b => `até ${b.nome} (${b.nUsinas})`), datasets:[
      {label:'Perda da bacia (%)', data:B.map(b => b.pctQnat), backgroundColor:COLORS.amber, borderWidth:0, barPercentage:0.7, categoryPercentage:0.85},
    ]},
    options: opts
  });
}

function renderEvapRel(){
  const chave = modes.evapRelX;
  const rot = chave==='pctQinc' ? 'Evap ÷ Qinc média (%) — escala log' : 'Evap ÷ Qnat média (%) — escala log';
  const pts = USINAS.map((u,i) => ({u,i,e:u.evaporacao}))
    .filter(x => x.e && !x.u.qincNulo && x.u.incremental && x.e[chave]!=null && x.e[chave]>0);
  const ds = [1,2,3,0].map(c => ({
    label: CLASSE_INFO[c].txt, backgroundColor: CLASSE_INFO[c].cor, borderColor: CLASSE_INFO[c].cor,
    pointRadius: c===0 ? 3 : 5, pointHoverRadius:7,
    data: pts.filter(x => x.u.classe===c).map(x => ({x:x.e[chave], y:x.u.incremental.pctNegativos, i:x.i, nome:x.u.nome})),
  })).filter(d => d.data.length);
  destroyChart('evapRel');
  const opts = baseOpts({ plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP, callbacks:{
      title: its => its[0].raw.nome,
      label: it => [`${chave==='pctQinc'?'Evap/Qinc':'Evap/Qnat'}: ${fmt(it.raw.x, 2)}%`, `Meses Qinc < 0: ${fmt(it.raw.y, 2)}%`] }}},
    onClick: (ev, els) => { if(els.length) selecionarEvap(charts.evapRel.data.datasets[els[0].datasetIndex].data[els[0].index].i); }
  }, {xTitle: rot, yTitle:'Meses com Qinc < 0 (%)'});
  opts.interaction = {mode:'nearest', intersect:true};
  opts.scales.x.type = 'logarithmic';
  opts.scales.x.ticks = {...opts.scales.x.ticks, maxTicksLimit:30, autoSkip:false, callback: v => {
    const l = Math.log10(v);
    if(Math.abs(l - Math.round(l)) > 1e-9) return '';
    return v >= 1 ? fmt(v, 0) : fmt(v, Math.round(-l));
  }};
  opts.scales.y.min = 0;
  charts.evapRel = new Chart(resetCanvas('chartEvapRel'), { type:'scatter', data:{datasets:ds}, options: opts });
  const r = EVAP.relacao && EVAP.relacao[chave];
  document.getElementById('evapRelNota').innerHTML = (r
      ? `Correlação de Spearman: <b>ρ = ${fmt(r.rho,2)}</b> (p = ${r.p.toExponential(1)}, n = ${r.n} usinas). `
      : '') + `Leitura <b>descritiva</b>: a evaporação já está embutida na vazão natural, então a relação só indica se as usinas com meses negativos são as que têm maior peso de evaporação na incremental — não é causa comprovada. `
      + `Ficam de fora as usinas com Qinc ≡ 0${chave==='pctQinc' ? ' e as com Qinc médio ≤ 0' : ''}.`;
}

function selecionarEvap(i){
  uCurrent = i;
  renderEvapUsina();
  document.getElementById('evapUsinaNav').scrollIntoView({behavior:'smooth', block:'start'});
}

function renderEvapUsina(){
  const u = USINAS[uCurrent];
  identidade({nome:'evapUsinaNome', usina:'evapTagUsina', posto:'evapTagPosto'}, u);
  atualizarNav('v');
  const e = u.evaporacao;
  document.getElementById('evapUsinaSub').textContent = e
    ? `Bacia até ${e.terminalNome} · perfil mensal da evaporação líquida do reservatório`
    : 'Sem dados de evaporação para esta usina';
  const row = document.getElementById('evapBadgesRow');
  if(!e){ row.innerHTML = ''; }
  else {
    let h = `<div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Evaporação: <b>${fmt(e.evapMed,2)} m³/s</b> <span style="color:var(--text-dim)">(${fmt(e.evapMin,2)}–${fmt(e.evapMax,2)}) · ${fmt(e.volAnualHm3,0)} hm³/ano · nº ${e.rank} de ${EVAP.nUsinas}</span></div>`;
    h += `<div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Evap/Qnat: <b>${e.pctQnat!=null ? fmt(e.pctQnat,2)+'%' : '—'}</b> <span style="color:var(--text-dim)">(Qnat ${fmt(e.qnatMedia,1)} m³/s)</span></div>`;
    h += `<div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Evap/Qinc: <b>${e.pctQinc!=null ? fmt(e.pctQinc,1)+'%' : '—'}</b> <span style="color:var(--text-dim)">(Qinc ${fmt(e.qincMedia,1)} m³/s${e.pctQinc==null ? ' · Qinc médio ≤ 0' : ''})</span></div>`;
    h += `<div class="badge"><span class="dot" style="background:${COLORS.textDim}"></span>Área Mín / Méd / Máx: <b>${fmt(e.areaMinKm2,1)} / ${fmt(e.areaMedKm2,1)} / ${fmt(e.areaMaxKm2,1)} km²</b> <span style="color:var(--text-dim)">· lâmina ${fmt(e.laminaAnualMm,0)} mm/ano</span></div>`;
    if(u.incremental) h += `<div class="badge ${u.incremental.negativos>0?'warn':'ok'}"><span class="dot" style="background:${u.incremental.negativos>0?COLORS.coral:COLORS.green}"></span>Meses com Qinc &lt; 0: <b>${u.incremental.negativos}</b> <span style="color:var(--text-dim)">(${fmt(u.incremental.pctNegativos,2)}%)</span></div>`;
    if(e.areaSuspeita) h += `<div class="badge warn critical"><span class="dot" style="background:${COLORS.coral}"></span>Área fora do plausível — <b>provável erro de dado</b> (PAC/cotas)</div>`;
    row.innerHTML = h;
  }
  renderEvap(u);
}

function renderEvapView(){
  if(!EVAP){
    document.getElementById('evapKpis').innerHTML = '<div class="empty-note" style="grid-column:1/-1;">Sem dados de evaporação — coloque o dados_hidroterm_completo.xlsx na pasta e rode o gerador de novo.</div>';
    return;
  }
  if(!evapBuilt){
    evapBuilt = true;
    renderEvapKpis();
    renderEvapBacia();
    renderEvapRel();
    renderEvapTabela();
    document.querySelector('.toggle-group[data-chart="evapMetrica"]').addEventListener('click', (e)=>{
      const btn = e.target.closest('button'); if(!btn) return;
      modes.evapMetrica = btn.dataset.mode;
      e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
      renderEvapRank();
    });
    document.querySelector('.toggle-group[data-chart="evapRelX"]').addEventListener('click', (e)=>{
      const btn = e.target.closest('button'); if(!btn) return;
      modes.evapRelX = btn.dataset.mode;
      e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
      renderEvapRel();
    });
    document.getElementById('evapTopN').addEventListener('change', renderEvapRank);
    renderEvapRank();
    document.getElementById('evapFiltro').addEventListener('input', renderEvapTabela);
    document.getElementById('btnCsvEvap').addEventListener('click', () => {
      exportarCSV('evaporacao_por_usina.csv', COLS_EVAP.map(c => c.t), evapFiltradas().map(r => COLS_EVAP.map(c => {
        const v = r[c.k];
        return (typeof v === 'number' && !Number.isInteger(v)) ? v.toFixed(3).replace('.', ',') : v;
      })));
    });
  }
  renderEvapUsina();
}

// ============================================================
// CORRECAO VIEW
// ============================================================
const STATUS_COR = {'corrigida':COLORS.green, 'corrigida parcial':COLORS.amber, 'estrutural':COLORS.coral,
                    'acompanha montante':COLORS.cyan, 'sem negativos':COLORS.textDim};
function pillStatus(s){
  if(!s) return '<span style="color:var(--text-dim)">—</span>';
  const c = STATUS_COR[s] || COLORS.textDim;
  return `<span class="pill" style="color:${c};border-color:${c}">${s}</span>`;
}
let corrBuilt = false;
function renderCorrResumo(){
  const C = APP.correcao;
  const el = document.getElementById('corrKpis');
  if(!C){ el.innerHTML = '<div class="empty-note">Correção não disponível neste arquivo.</div>'; return; }
  const k = (lbl, val, sub, cls='') => `<div class="kpi ${cls}"><div class="lbl">${lbl}</div><div class="val">${val}</div><div class="sub">${sub}</div></div>`;
  const st = C.status || {};
  const nEstr = USINAS.filter(u => u.correcao && u.correcao.status==='estrutural').reduce((s,u)=>s+u.correcao.negDepois,0);
  el.innerHTML =
    k('Qinc < 0 antes', fmt(C.negOrig), 'Qinc original', 'warn') +
    k('Qinc < 0 depois', fmt(C.negDepois), `${fmt(nEstr)} em usinas estruturais`, 'amber') +
    k('Usinas corrigidas', (st['corrigida']||0) + (st['corrigida parcial']||0), `${st['corrigida parcial']||0} parciais`, 'cyan') +
    k('Estruturais', st['estrutural']||0, `classes ${C.classesEstruturais.join(', ')}${C.excluiEspeciais ? ' + casos especiais' : ''} — não alteradas`) +
    k('Anos não corrigíveis', C.anosNaoCorrigiveis, 'Qinc anual ≤ 0') +
    k('Erro de volume anual', fmt(C.difVolAnualMax,3), 'm³/s (máx. por posto, arredondamento)');
  renderCorrTabela();
}
function renderCorrTabela(){
  const todas = document.getElementById('corrTodas').checked;
  const L = USINAS.map((u,i)=>({u,i})).filter(({u}) => u.correcao && (todas || u.correcao.status!=='sem negativos'))
    .sort((a,b) => b.u.correcao.negOrig - a.u.correcao.negOrig || b.u.correcao.negAntes - a.u.correcao.negAntes);
  document.getElementById('corrCount').textContent = `${L.length} usinas`;
  document.getElementById('corrTbody').innerHTML = L.map(({u,i}) => {
    const c = u.correcao;
    const vr = c.volRedistPct;
    const corVr = vr>20 ? COLORS.coral : (vr>10 ? COLORS.amber : 'inherit');
    const nc = c.anosNaoCorrigiveis.length;
    return `<tr class="clickrow" data-idx="${i}">
      <td>${esc(u.nome)}${u.casoEspecial ? ` <span style="color:${COLORS.amber}">⚠</span>` : ''}</td>
      <td>${pillClasse(u.classe)}</td><td>${pillStatus(c.status)}</td>
      <td class="num">${fmt(c.negOrig)}</td>
      <td class="num" style="color:${c.negDepois>0?COLORS.coral:COLORS.green}">${fmt(c.negDepois)}</td>
      <td class="num">${fmt(c.mesesCorrigidos)}</td><td class="num">${fmt(c.anosCorrigidos)}</td>
      <td class="num" title="${c.anosNaoCorrigiveis.join(', ')}">${nc || '—'}</td>
      <td class="num" style="color:${corVr}">${c.mesesCorrigidos ? fmt(vr,2) : '—'}</td>
      <td class="num">${c.mesesCorrigidos ? fmt(c.maxDelta) : '—'}</td>
    </tr>`;
  }).join('');
  document.querySelectorAll('#corrTbody tr').forEach(tr => tr.addEventListener('click', () => {
    uCurrent = parseInt(tr.dataset.idx); renderCorrUsina();
    document.querySelector('#viewCorr .nav-strip').scrollIntoView({behavior:'smooth', block:'start'});
  }));
}
function renderCorrUsina(){
  const u = USINAS[uCurrent];
  identidade({nome:'corrUsinaNome', usina:'corrTagUsina', posto:'corrTagPosto'}, u);
  atualizarNav('k');
  const c = u.correcao;
  document.getElementById('corrUsinaSub').textContent = c ? `Status: ${c.status}` : 'Sem correção para esta usina';
  const row = document.getElementById('corrBadgesRow');
  if(!c){ row.innerHTML = ''; }
  else {
    let h = `<div class="badge"><span class="dot" style="background:${STATUS_COR[c.status]||COLORS.textDim}"></span>Status: <b>${c.status}</b></div>`;
    h += `<div class="badge ${c.negDepois>0?'warn':'ok'}"><span class="dot" style="background:${c.negDepois>0?COLORS.coral:COLORS.green}"></span>Meses Qinc &lt; 0: <b>${c.negOrig} → ${c.negDepois}</b></div>`;
    if(c.mesesCorrigidos) h += `<div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>Volume redistribuído: <b>${fmt(c.volRedistPct,2)}%</b> <span style="color:var(--text-dim)">· máx |ΔQ| ${fmt(c.maxDelta)} m³/s · ${c.anosCorrigidos} anos</span></div>`;
    if(c.anosNaoCorrigiveis.length) h += `<div class="badge critical warn"><span class="dot" style="background:${COLORS.coral}"></span>Anos não corrigíveis (Qinc anual ≤ 0): <b>${c.anosNaoCorrigiveis.length}</b> <span style="color:var(--text-dim)">${c.anosNaoCorrigiveis.slice(0,12).join(', ')}${c.anosNaoCorrigiveis.length>12?'…':''}</span></div>`;
    if(c.status==='estrutural') h += `<div class="badge critical"><span class="dot" style="background:${COLORS.coral}"></span>Não alterada — <b>corrigir topologia/posto</b> (derivação, transposição ou montante cadastrada errada)</div>`;
    if(c.status==='acompanha montante') h += `<div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Qinc ≡ 0: a natural corrigida <b>acompanha a montante corrigida</b></div>`;
    row.innerHTML = h;
  }
  const wM = wrapOf('chartCorrMensal'), wP = wrapOf('chartCorrPerfil');
  destroyChart('corrMensal'); destroyChart('corrPerfil');
  if(!c || !c.mudou || !u.incremental){
    const msg = !c ? 'Sem dados' : (c.status==='estrutural' ? 'Usina estrutural — série não alterada pela correção'
              : 'Nenhum mês alterado pela correção nesta usina');
    wM.innerHTML = `<div class="empty-note">${msg}</div>`; wP.innerHTML = `<div class="empty-note">${msg}</div>`;
    return;
  }
  charts.corrMensal = new Chart(resetCanvas('chartCorrMensal'), {
    type:'line',
    data:{ labels:u.incremental.mensalLabels, datasets:[
      {label:'Qinc original', data:u.incremental.mensalValores, borderColor:COLORS.coral, borderWidth:1, pointRadius:0, tension:0},
      {label:'Qinc corrigido', data:c.mensalCorr, borderColor:COLORS.cyan, borderWidth:1.2, pointRadius:0, tension:0},
    ]},
    options: comZero(baseOpts({plugins:{legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP}}}, {xTitle:'Mês/Ano', yTitle:'Qinc (m³/s)', xTicks:{maxTicksLimit:12}}))
  });
  charts.corrPerfil = new Chart(resetCanvas('chartCorrPerfil'), {
    type:'bar',
    data:{ labels:MESES_PT, datasets:[
      {label:'Original', data:c.perfilOrig, backgroundColor:'rgba(242,118,91,0.75)', borderWidth:0, barPercentage:0.8, categoryPercentage:0.75},
      {label:'Corrigido', data:c.perfilCorr, backgroundColor:'rgba(79,216,232,0.55)', borderWidth:0, barPercentage:0.8, categoryPercentage:0.75},
    ]},
    options: comZero(baseOpts({plugins:{legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP}}}, {xTitle:'Mês', yTitle:'Qinc médio (m³/s)', xTicks:{autoSkip:false, maxTicksLimit:12}}))
  });
}
function renderCorrView(){
  if(!corrBuilt){
    corrBuilt = true;
    document.getElementById('corrTodas').addEventListener('change', renderCorrTabela);
    renderCorrResumo();
  }
  renderCorrUsina();
}

// ============================================================
// COMPARAR VIEW
// ============================================================
let compBuilt = false;
function renderCompView(){
  if(!compBuilt){
    compBuilt = true;
    const selA = document.getElementById('compSelectA');
    const selB = document.getElementById('compSelectB');
    USINAS.forEach((u,i) => {
      selA.appendChild(new Option(`${u.nome} (${u.codUsina})`, i));
      selB.appendChild(new Option(`${u.nome} (${u.codUsina})`, i));
    });
    selA.value = String(uCurrent);
    selB.selectedIndex = USINAS.length>1 ? (uCurrent+1) % USINAS.length : 0;
    selA.addEventListener('change', desenharComparacao);
    selB.addEventListener('change', desenharComparacao);
    document.getElementById('compMesmoEixo').addEventListener('change', desenharComparacao);
    document.querySelector('.toggle-group[data-chart="compVar"]').addEventListener('click', (e)=>{
      const btn = e.target.closest('button'); if(!btn) return;
      modes.compVar = btn.dataset.mode;
      e.currentTarget.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b===btn));
      desenharComparacao();
    });
  }
  desenharComparacao();
}

function desenharComparacao(){
  const a = USINAS[parseInt(document.getElementById('compSelectA').value)];
  const b = USINAS[parseInt(document.getElementById('compSelectB').value)];
  const v = modes.compVar;
  const rot = v==='natural' ? 'Qnat' : 'Qinc';
  document.getElementById('compTitulo').textContent = `Comparar duas usinas — Vazão ${v==='natural'?'Natural':'Incremental'} Anual`;
  const sa = a[v], sb = b[v];
  const mesmo = document.getElementById('compMesmoEixo').checked;
  document.getElementById('compBadgesRow').innerHTML = `
    <div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>${esc(a.nome)} <span style="color:var(--text-dim)">— posto ${a.codPosto ?? '—'}${mesmo?'':' — eixo esquerdo'}</span></div>
    <div class="badge"><span class="dot" style="background:${COLORS.amber}"></span>${esc(b.nome)} <span style="color:var(--text-dim)">— posto ${b.codPosto ?? '—'}${mesmo?'':' — eixo direito'}</span></div>
  `;
  const wrap = wrapOf('chartComp');
  if(!sa || !sb){ destroyChart('comp'); wrap.innerHTML = `<div class="empty-note">Uma das usinas não tem série ${v}</div>`; return; }
  const anos = [...new Set([...sa.anos, ...sb.anos])].sort((x,y)=>x-y);
  const serie = s => { const m = new Map(s.anos.map((x,j)=>[x, s.media[j]])); return anos.map(x => m.has(x) ? m.get(x) : null); };
  destroyChart('comp');
  const opts = baseOpts({ plugins:{ legend:LEGEND_BOTTOM, tooltip:{...TOOLTIP} } },
    {xTitle:'Ano', yTitle: mesmo ? `${rot} (m³/s)` : `${a.nome} — ${rot} (m³/s)`});
  if(!mesmo){
    opts.scales.y.ticks.color = COLORS.cyan;
    opts.scales.y.title.color = COLORS.cyan;
    opts.scales.y1 = { position:'right', grid:{drawOnChartArea:false}, ticks:{maxTicksLimit:6, color:COLORS.amber},
      title: {...axisTitle(`${b.nome} — ${rot} (m³/s)`), color:COLORS.amber} };
  }
  if(v==='incremental') comZero(opts);
  charts.comp = new Chart(resetCanvas('chartComp'), {
    type:'line',
    data:{ labels:anos, datasets:[
      {label:a.nome, data:serie(sa), borderColor:COLORS.cyan, backgroundColor:'rgba(79,216,232,0.06)', borderWidth:1.6, pointRadius:0, fill:false, tension:0.15, spanGaps:true, yAxisID:'y'},
      {label:b.nome, data:serie(sb), borderColor:COLORS.amber, backgroundColor:'rgba(242,184,75,0.06)', borderWidth:1.6, pointRadius:0, fill:false, tension:0.15, spanGaps:true, yAxisID: mesmo ? 'y' : 'y1'},
    ]},
    options: opts
  });
}

// ============================================================
// USOS CONSUNTIVOS VIEW
// ============================================================
let usosRendered = false;
function renderUsosConsuntivos(){
  if(usosRendered) return;
  usosRendered = true;
  const u = APP.usosConsuntivos;
  if(!u || !u.lista || !u.lista.length){
    document.getElementById('usosSemDados').style.display = 'flex';
    document.getElementById('usosComDados').style.display = 'none';
    return;
  }
  const top = u.lista[0];
  const n50 = u.lista.filter(i => i.coberturaPct!=null && i.coberturaPct>=50).length;
  document.getElementById('usosResumoRow').innerHTML = `
    <div class="badge"><span class="dot" style="background:${COLORS.cyan}"></span>Usinas cruzadas com a base ANA: <b>${u.lista.length}</b></div>
    <div class="badge warn"><span class="dot" style="background:${COLORS.amber}"></span>Cobertura ≥ 50%: <b>${n50}</b></div>
    <div class="badge critical warn"><span class="dot" style="background:${COLORS.coral}"></span>Maior cobertura: <b>${esc(top.nome)}</b> <span style="color:var(--text-dim)">(${top.coberturaPct}%)</span></div>
  `;
  const tbody = document.getElementById('usosTbody');
  tbody.innerHTML = u.lista.map(i => {
    const pct = i.coberturaPct;
    const cor = pct==null ? COLORS.textDim : (pct>=100 ? COLORS.coral : (pct>=50 ? COLORS.amber : COLORS.textDim));
    const idx = USINAS.findIndex(x => x.codUsina===i.codUsina);
    return `<tr class="clickrow" data-idx="${idx}" style="${pct!=null && pct>=50?'font-weight:600;':''}">
      <td>${esc(i.nome)}</td>
      <td class="num" style="color:var(--text-dim);">${i.nMesesNeg}</td>
      <td class="num" style="color:${COLORS.coral};">${fmt(i.qincMedioNeg,2)}</td>
      <td class="num">${fmt(i.usosMedioNeg,3)}</td>
      <td class="num" style="color:${cor};">${pct!=null ? fmt(pct,1)+'%' : '—'}</td>
    </tr>`;
  }).join('');
  tbody.querySelectorAll('tr').forEach(tr => tr.addEventListener('click', () => {
    const i = parseInt(tr.dataset.idx); if(i>=0){ uCurrent = i; switchTab('usinas'); }
  }));
  document.getElementById('btnCsvUsos').addEventListener('click', () => {
    exportarCSV('usos_consuntivos_diagnostico.csv',
      ['Usina','Meses_Qinc_Negativo','Qinc_Medio_Negativos_m3s','UsosCons_Incremental_m3s','Cobertura_Pct'],
      u.lista.map(i => [i.nome, i.nMesesNeg, i.qincMedioNeg, i.usosMedioNeg, i.coberturaPct]));
  });
}

// ============================================================
// CSV EXPORT (genérico, ; como separador p/ Excel pt-BR)
// ============================================================
function exportarCSV(filename, headers, rows){
  const e = v => {
    const s = String(v ?? '');
    return /[",;\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  };
  const linhas = [headers.map(e).join(';'), ...rows.map(r => r.map(e).join(';'))];
  const blob = new Blob(['\ufeff' + linhas.join('\n')], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function wireCsvResumoEstac(){
  const btn = document.getElementById('btnCsvResumoEstac');
  if(!btn || btn.dataset.wired) return;
  btn.dataset.wired = '1';
  btn.addEventListener('click', () => {
    const headers = ['CodUsina','CodPosto','Usina','r1','r1_significativa','Sen_slope','MK_p_bruto','MK_trend_bruto','MK_p_corrigido','MK_trend_corrigido','Pettitt_ano','Pettitt_p'];
    const rows = USINAS.filter(u=>u.estacionariedade).map(u => { const e=u.estacionariedade; return [
      u.codUsina, u.codPosto, u.nome, e.r1, e.r1Sig?'SIM':'NAO', e.senSlope, e.mkPBruto, e.mkTrendBruto, e.mkPCorr, e.mkTrendCorr, e.pettittAno, e.pettittP
    ];});
    exportarCSV('estacionariedade_usinas.csv', headers, rows);
  });
}

// ============================================================
// Keyboard nav
// ============================================================
document.addEventListener('keydown', (e)=>{
  const tag = document.activeElement && document.activeElement.tagName;
  if(e.key==='/' && tag!=='INPUT' && tag!=='SELECT'){ e.preventDefault(); searchInput.focus(); return; }
  if(tag==='INPUT' || tag==='SELECT') return;
  if(activeTab==='cascatas'){
    if(e.key==='ArrowLeft') goToCascata(cCurrent-1);
    if(e.key==='ArrowRight') goToCascata(cCurrent+1);
  } else if(TABS_POR_USINA.includes(activeTab)){
    if(e.key==='ArrowLeft') goToUsina(uCurrent-1);
    if(e.key==='ArrowRight') goToUsina(uCurrent+1);
  }
});

// ============================================================
// Search (usinas) — nome, codigo da usina ou posto; setas + Enter
// ============================================================
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');
let searchSel = 0;
function fecharBusca(){ searchInput.value=''; searchResults.classList.remove('show'); }
function normTxt(s){ return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); }
searchInput.addEventListener('input', ()=>{
  const q = normTxt(searchInput.value.trim());
  if(!q){ searchResults.classList.remove('show'); return; }
  const matches = USINAS.map((u,i)=>({u,i})).filter(({u}) =>
    normTxt(u.nome).includes(q) || String(u.codUsina)===q || String(u.codPosto ?? '')===q || String(u.codUsina).startsWith(q)).slice(0,12);
  searchSel = 0;
  searchResults.innerHTML = matches.length ? matches.map(({u,i},k)=>
    `<div class="item ${k===0?'sel':''}" data-idx="${i}"><span>${esc(u.nome)}</span><span class="code">USINA ${u.codUsina} · POSTO ${u.codPosto ?? '—'}</span></div>`
  ).join('') : '<div class="item">Nenhum resultado</div>';
  searchResults.classList.add('show');
});
searchInput.addEventListener('keydown', (e)=>{
  const itens = [...searchResults.querySelectorAll('.item[data-idx]')];
  if(e.key==='Escape'){ fecharBusca(); searchInput.blur(); return; }
  if(!itens.length) return;
  if(e.key==='ArrowDown' || e.key==='ArrowUp'){
    e.preventDefault();
    searchSel = (searchSel + (e.key==='ArrowDown'?1:-1) + itens.length) % itens.length;
    itens.forEach((it,k)=>it.classList.toggle('sel', k===searchSel));
  }
  if(e.key==='Enter'){ goToUsina(parseInt(itens[searchSel].dataset.idx)); searchInput.blur(); }
});
searchResults.addEventListener('click', (e)=>{
  const item = e.target.closest('.item[data-idx]');
  if(item) goToUsina(parseInt(item.dataset.idx));
});
document.addEventListener('click', (e)=>{ if(!e.target.closest('.search-box')) searchResults.classList.remove('show'); });

// ============================================================
// Download JPG (fundo branco, pronto p/ relatorio)
// ============================================================
function tituloCard(card){
  const t = card.querySelector('.card-title h3')?.textContent || 'Grafico';
  if(card.closest('#viewGeral')) return t;
  if(card.closest('#viewEvap') && !card.querySelector('#chartEvap')){
    if(card.querySelector('#chartEvapRank')) return `${t} (${EVAP_MET[modes.evapMetrica].rot}, Top ${document.getElementById('evapTopN').value})`;
    if(card.querySelector('#chartEvapRel')) return `${t} (${modes.evapRelX==='pctQinc' ? 'Evap/Qinc' : 'Evap/Qnat'})`;
    return t;
  }
  if(card.querySelector('#chartEstacDecadas')) return `${t} — ${modes.estacEscopo==='usinas' ? 'postos de usinas' : 'todos os postos do deck'}`;
  if(card.closest('#viewComp')){
    const a = USINAS[parseInt(document.getElementById('compSelectA').value)], b = USINAS[parseInt(document.getElementById('compSelectB').value)];
    return `${t} — ${a.nome} x ${b.nome}`;
  }
  if(card.closest('#viewCascatas')) return document.getElementById('cascataNomeH2').textContent + ' — ' + t;
  const u = USINAS[uCurrent];
  const sel = card.querySelector('#permMesSelect');
  const extra = sel && sel.value ? ` (sobreposto: ${sel.value})` : '';
  return `${u.nome} (usina ${u.codUsina}) — ${t}${extra}`;
}
function nomeArquivo(txt){
  return txt.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'') + '.jpg';
}
const TEMA_CLARO = {texto:'#1f2d3a', eixo:'#4a5b6b', grade:'#dde4ea', fundo:'#ffffff'};
const MAPA_CLARO = {
  '#eaf2f6':'#1f2d3a', '#4fd8e8':'#1591a8', '#f2b84b':'#c98a12', '#6fd6a0':'#2a9a5c', '#5f7a90':'#6b7c8c',
  'rgba(79,216,232,0.08)':'rgba(21,145,168,0.10)', 'rgba(79,216,232,0.06)':'rgba(21,145,168,0.08)',
  'rgba(79,216,232,0.28)':'rgba(21,145,168,0.25)', 'rgba(79,216,232,0.55)':'rgba(21,145,168,0.55)',
  'rgba(242,184,75,0.4)':'rgba(201,138,18,0.55)', 'rgba(242,184,75,0.45)':'rgba(201,138,18,0.55)', 'rgba(242,184,75,0.08)':'rgba(201,138,18,0.10)',
  'rgba(111,214,160,0.35)':'rgba(42,154,92,0.50)', 'rgba(242,184,75,0.35)':'rgba(201,138,18,0.35)', '#2b8ea3':'#1d6f80'
};
function legendaExport(chart){
  return Chart.defaults.plugins.legend.labels.generateLabels(chart).map(it => {
    const ds = chart.data.datasets[it.datasetIndex];
    const tipo = ds.type || chart.config.type;
    if(tipo === 'line' && ds.showLine !== false){
      it.pointStyle = 'line'; it.fillStyle = it.strokeStyle; it.lineWidth = 3;
    } else if(tipo === 'line' || tipo === 'scatter'){
      it.pointStyle = ds.pointStyle || 'circle';
      it.fillStyle = ds.backgroundColor; it.strokeStyle = ds.backgroundColor;
    } else {
      it.pointStyle = 'rect'; it.strokeStyle = it.fillStyle; it.lineWidth = 0;
    }
    return it;
  });
}
function aplicarTemaClaro(chart, claro){
  const o = chart.options;
  Object.values(o.scales || {}).forEach(sc => {
    if(claro){
      sc._bk = {t: sc.ticks.color, g: sc.grid.color, tt: sc.title ? sc.title.color : undefined};
      sc.ticks.color = TEMA_CLARO.texto; sc.grid.color = TEMA_CLARO.grade;
      if(sc.title) sc.title.color = TEMA_CLARO.eixo;
    } else if(sc._bk){
      sc.ticks.color = sc._bk.t; sc.grid.color = sc._bk.g;
      if(sc.title) sc.title.color = sc._bk.tt;
      delete sc._bk;
    }
  });
  const raw = chart.config.options;
  raw.plugins = raw.plugins || {};
  if(claro){
    const orig = raw.plugins.legend;
    chart._legBk = orig;
    raw.plugins.legend = {...(orig || {}), display:true, position:'bottom',
      labels:{...((orig && orig.labels) || {}), color:TEMA_CLARO.texto, usePointStyle:true,
              pointStyleWidth:30, padding:14, generateLabels:legendaExport}};
  } else if('_legBk' in chart){
    if(chart._legBk === undefined) delete raw.plugins.legend; else raw.plugins.legend = chart._legBk;
    delete chart._legBk;
  }
  chart.data.datasets.forEach(ds => {
    if(claro){
      ds._bk = {b: ds.borderColor, f: ds.backgroundColor};
      if(typeof ds.borderColor === 'string' && MAPA_CLARO[ds.borderColor]) ds.borderColor = MAPA_CLARO[ds.borderColor];
      if(typeof ds.backgroundColor === 'string' && MAPA_CLARO[ds.backgroundColor]) ds.backgroundColor = MAPA_CLARO[ds.backgroundColor];
    } else if(ds._bk){
      ds.borderColor = ds._bk.b; ds.backgroundColor = ds._bk.f; delete ds._bk;
    }
  });
  chart.update('none');
  chart.draw();
}
function baixarCard(card){
  const cvs = [...card.querySelectorAll('canvas')].filter(c => c.width && c.height);
  if(!cvs.length){ alert('Sem gráfico para baixar.'); return; }
  const escala = window.devicePixelRatio || 1;
  const pad = Math.round(20*escala), head = Math.round(46*escala);
  const w = Math.max(...cvs.map(c => c.width)) + 2*pad;
  const h = head + cvs.reduce((s,c) => s + c.height + pad, 0) + pad;
  const out = document.createElement('canvas'); out.width = w; out.height = h;
  const g = out.getContext('2d');
  g.fillStyle = TEMA_CLARO.fundo; g.fillRect(0, 0, w, h);
  const titulo = tituloCard(card);
  g.fillStyle = TEMA_CLARO.texto; g.textBaseline = 'middle';
  g.font = `600 ${Math.round(18*escala)}px Inter, 'Segoe UI', sans-serif`;
  g.fillText(titulo, pad, pad + head/2 - pad/2);
  let y = head + pad;
  cvs.forEach(c => {
    const ch = Chart.getChart(c);
    if(ch) aplicarTemaClaro(ch, true);
    g.drawImage(c, pad, y);
    if(ch) aplicarTemaClaro(ch, false);
    y += c.height + pad;
  });
  const a = document.createElement('a');
  a.href = out.toDataURL('image/jpeg', 0.95);
  a.download = nomeArquivo(titulo);
  document.body.appendChild(a); a.click(); a.remove();
}
const ICON_DL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></svg>';
document.querySelectorAll('.card').forEach(card => {
  if(card.dataset.nodl) return;
  const head = card.querySelector('.card-head'); if(!head) return;
  const wrap = document.createElement('div'); wrap.className = 'head-actions';
  head.querySelectorAll(':scope > .toggle-group, :scope > select, :scope > .dl-btn').forEach(el => wrap.appendChild(el));
  head.appendChild(wrap);
  const b = document.createElement('button'); b.className = 'dl-btn'; b.title = 'Baixar gráfico em JPG (fundo branco)';
  b.innerHTML = ICON_DL + 'JPG';
  b.addEventListener('click', () => baixarCard(card));
  wrap.appendChild(b);
});

document.getElementById('footerCount').textContent = USINAS.length;
document.getElementById('footerCascCount').textContent = CASCATAS.length;

renderGeralView();
</script>
</body>
</html>
'''


if __name__ == "__main__":
    main()
