"""
================================================================================
  CONCILIAÇÃO BANCÁRIA — Interface Streamlit
  Upload drag-and-drop → Prévia → Processamento → Gráficos → Download Excel
================================================================================
  Instalação:
      pip install streamlit pandas openpyxl plotly chardet xlsxwriter
  Execução:
      streamlit run conciliacao_app.py
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from decimal import Decimal, ROUND_HALF_UP, getcontext
from io import BytesIO
import re
from difflib import SequenceMatcher

getcontext().prec = 12

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================

st.set_page_config(
    page_title="Conciliação Bancária",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS customizado
st.markdown("""
<style>
    /* Header */
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 {
        margin: 0 0 0.3rem;
        font-size: 1.8rem;
        font-weight: 700;
    }
    .main-header p {
        margin: 0;
        opacity: 0.75;
        font-size: 0.95rem;
    }

    /* Metric cards */
    .metric-row {
        display: flex;
        gap: 1rem;
        margin: 1rem 0;
    }
    .metric-card {
        flex: 1;
        padding: 1.2rem 1.5rem;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        background: white;
    }
    .metric-card .label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        opacity: 0.6;
        margin-bottom: 0.3rem;
    }
    .metric-card .value {
        font-size: 1.5rem;
        font-weight: 700;
    }
    .metric-card .sub {
        font-size: 0.75rem;
        opacity: 0.5;
        margin-top: 0.2rem;
    }

    /* Status badge */
    .badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    .badge-green  { background: #dcfce7; color: #166534; }
    .badge-yellow { background: #fef9c3; color: #854d0e; }
    .badge-red    { background: #fee2e2; color: #991b1b; }
    .badge-blue   { background: #dbeafe; color: #1e40af; }

    /* Upload area */
    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 0.5rem;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #3b82f6;
    }

    /* Hide default header/footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.5rem 1.5rem;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# CONSTANTES E CONFIG
# ==============================================================================

CONFIG = {
    "tolerancia_dias": 3,
    "formatos_data": [
        "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
        "%d/%m/%y", "%Y/%m/%d", "%d.%m.%Y",
    ],
}

COLUMN_MAP = {
    "data": [
        "data", "date", "dt", "data_lancamento", "data_mov", "dt_mov",
        "data_movimentacao", "release_date", "fecha",
        "previsao_de_pagamento", "vencimento", "ultimo_pagamento",
    ],
    "descricao": [
        "descricao", "historico", "descricão", "histórico", "desc",
        "description", "memo", "obs", "lancamento", "detalhes",
        "transaction_type", "fornecedor_(nome_fantasia)",
    ],
    "valor": [
        "valor", "value", "vlr", "amount", "montante",
        "vl_lancamento", "valor_lancamento",
        "transaction_net_amount", "valor_da_conta", "valor_liquido",
    ],
    "id_doc": [
        "id", "documento", "doc", "id_documento", "num_doc", "numero",
        "nsu", "id_doc", "ref", "referencia", "numero_documento",
        "reference_id", "numero_do_documento",
    ],
}


# ==============================================================================
# FUNÇÕES DE NORMALIZAÇÃO
# ==============================================================================

def limpar_valor(valor_str) -> float:
    if pd.isna(valor_str) or str(valor_str).strip() == "":
        return 0.0
    s = str(valor_str).strip()
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo = True
        s = s[1:-1]
    if s.upper().endswith("D") or s.upper().endswith(" D"):
        negativo = True
        s = re.sub(r"\s*[dD]$", "", s)
    if s.startswith("-"):
        negativo = True
        s = s[1:]
    s = re.sub(r"[R$€£\s]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        val = float(s)
        return -val if negativo else val
    except ValueError:
        return 0.0


def parsear_data(data_str) -> pd.Timestamp:
    if pd.isna(data_str) or str(data_str).strip() == "":
        return pd.NaT
    s = str(data_str).strip()
    for fmt in CONFIG["formatos_data"]:
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(s, dayfirst=True)
    except Exception:
        return pd.NaT


def mapear_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mapeia colunas do arquivo para nomes padronizados.
    Blindado: detecta formato credits/debits, headers deslocados,
    gera id_doc sintético e descricao fallback quando ausentes.
    """
    # --- Normaliza nomes (minúsculo, sem acentos, underscores) ---
    df.columns = [
        re.sub(r"\s+", "_", col.strip().lower()
               .replace("á", "a").replace("ã", "a").replace("â", "a")
               .replace("é", "e").replace("ê", "e")
               .replace("í", "i").replace("ó", "o").replace("ô", "o")
               .replace("ú", "u").replace("ç", "c"))
        for col in df.columns
    ]

    # --- Remove colunas totalmente vazias ou "unnamed" vazias ---
    unnamed_vazias = [c for c in df.columns
                      if c.startswith("unnamed") and df[c].dropna().astype(str).str.strip().replace("", pd.NA).dropna().empty]
    df = df.drop(columns=unnamed_vazias, errors="ignore")
    df = df.dropna(axis=1, how="all")

    # --- Detecção especial: formato credits/debits (extrato bancário) ---
    cols_set = set(df.columns)
    tem_credits = any("credit" in c for c in cols_set)
    tem_debits = any("debit" in c for c in cols_set)

    if tem_credits and tem_debits and not any(
        c in cols_set for c in ["valor", "value", "amount", "transaction_net_amount"]
    ):
        col_cr = next(c for c in df.columns if "credit" in c)
        col_db = next(c for c in df.columns if "debit" in c)
        df[col_cr] = df[col_cr].apply(limpar_valor)
        df[col_db] = df[col_db].apply(limpar_valor)
        df["valor"] = df.apply(
            lambda r: r[col_cr] if r[col_cr] != 0 else -abs(r[col_db]), axis=1
        )
        df = df.drop(columns=[col_cr, col_db], errors="ignore")

    # --- Remove colunas de saldo/balance (não são lançamentos) ---
    df = df.loc[:, ~df.columns.str.contains("balance|saldo|partial")]

    # --- Mapeamento padrão ---
    mapeamento = {}
    ja_usado = set()
    for campo_padrao, sinonimos in COLUMN_MAP.items():
        for col in df.columns:
            if col in ja_usado:
                continue
            if col == campo_padrao or col in sinonimos or any(s in col for s in sinonimos):
                mapeamento[col] = campo_padrao
                ja_usado.add(col)
                break
    df = df.rename(columns=mapeamento)

    # --- Gera id_doc sintético se ausente ---
    if "id_doc" not in df.columns:
        df["id_doc"] = [f"AUTO_{i+1:06d}" for i in range(len(df))]

    # --- Gera descricao a partir de colunas alternativas se ausente ---
    if "descricao" not in df.columns:
        for fallback_col in df.columns:
            if any(fb in fallback_col for fb in [
                "fornecedor", "razao_social", "nome_fantasia",
                "observacao", "categoria", "historico",
            ]):
                df["descricao"] = df[fallback_col]
                break
        else:
            df["descricao"] = "SEM DESCRIÇÃO"

    # --- Validação final ---
    obrigatorias = {"data", "valor"}
    presentes = set(df.columns)
    ausentes = obrigatorias - presentes

    if ausentes:
        raise ValueError(
            f"⚠️ Arquivo não reconhecido.\n"
            f"Colunas faltando: {sorted(ausentes)}\n"
            f"Colunas encontradas: {sorted(df.columns.tolist())}\n"
            f"O arquivo deve conter pelo menos Data e Valor."
        )

    return df


def normalizar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").reset_index(drop=True)

    # Limpa valor (suporta strings e floats já convertidos)
    if df["valor"].dtype == object:
        df["valor"] = df["valor"].apply(limpar_valor)
    else:
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)

    df = df[df["valor"] != 0.0].reset_index(drop=True)
    df["data"] = df["data"].apply(parsear_data)
    df = df.dropna(subset=["data"]).reset_index(drop=True)
    df["id_doc"] = df["id_doc"].fillna("").astype(str).str.strip().str.upper()
    df["descricao"] = df["descricao"].fillna("").astype(str).str.strip()
    df["tipo"] = np.where(df["valor"] >= 0, "CRÉDITO", "DÉBITO")
    return df


def ler_upload(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        encodings = ["utf-8", "latin-1", "cp1252", "iso-8859-1"]
        seps = [";", ",", "\t", "|"]
        for enc in encodings:
            for sep in seps:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=sep, encoding=enc, dtype=str)
                    if len(df.columns) >= 3:
                        return df
                except Exception:
                    continue
        raise ValueError("Não foi possível ler o CSV. Verifique o formato.")

    elif name.endswith((".xlsx", ".xls")):
        engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"

        # --- Leitura bruta para detectar header real ---
        uploaded_file.seek(0)
        df_raw = pd.read_excel(uploaded_file, engine=engine, dtype=str, header=None)

        # Heurística: procura nas primeiras 10 linhas a que tem mais texto
        # (= provável header) e não é uma linha de resumo numérico
        melhor_row = 0
        melhor_score = 0
        for i in range(min(10, len(df_raw))):
            vals = df_raw.iloc[i].dropna().astype(str).str.strip()
            vals = vals[vals != ""]
            if len(vals) < 3:
                continue
            # Score = quantidade de valores não-numéricos (texto puro)
            score = sum(
                1 for v in vals
                if not v.replace(".", "").replace(",", "").replace("-", "").replace(" ", "").isdigit()
                and not v.startswith("unnamed")
                and len(v) > 1
            )
            if score > melhor_score:
                melhor_score = score
                melhor_row = i

        # Relê com header correto
        uploaded_file.seek(0)
        df = pd.read_excel(
            uploaded_file, engine=engine, dtype=str,
            skiprows=melhor_row, header=0,
        )

        # Remove linhas totalmente vazias
        df = df.dropna(how="all").reset_index(drop=True)

        # Remove colunas "Unnamed" totalmente vazias
        unnamed_vazias = [
            c for c in df.columns
            if str(c).startswith("Unnamed")
            and df[c].dropna().astype(str).str.strip().replace("", pd.NA).dropna().empty
        ]
        df = df.drop(columns=unnamed_vazias, errors="ignore")

        # Remove linhas que são cópia do header (texto repetido)
        if len(df) > 0:
            first_col = df.columns[0]
            df = df[df[first_col] != first_col].reset_index(drop=True)

        return df
    else:
        raise ValueError(f"Formato não suportado: {name}. Use .csv ou .xlsx")


# ==============================================================================
# MOTOR DE CONCILIAÇÃO
# ==============================================================================

def match_perfeito(extrato, controle):
    extrato, controle = extrato.copy(), controle.copy()

    # Chave usa valor absoluto para cruzar débitos do extrato com valores do controle
    extrato["_chave"] = (
        extrato["data"].dt.strftime("%Y%m%d") + "|" +
        extrato["valor"].abs().round(2).astype(str) + "|" +
        extrato["id_doc"]
    )
    controle["_chave"] = (
        controle["data"].dt.strftime("%Y%m%d") + "|" +
        controle["valor"].abs().round(2).astype(str) + "|" +
        controle["id_doc"]
    )
    conciliados = []
    ext_usado, ctrl_usado = set(), set()
    ctrl_por_chave = {}
    for idx, row in controle.iterrows():
        ctrl_por_chave.setdefault(row["_chave"], []).append(idx)

    for idx_e, row_e in extrato.iterrows():
        if idx_e in ext_usado:
            continue
        chave = row_e["_chave"]
        if chave in ctrl_por_chave:
            cands = [i for i in ctrl_por_chave[chave] if i not in ctrl_usado]
            if cands:
                idx_c = cands[0]
                conciliados.append({
                    "data_extrato": row_e["data"],
                    "descricao_extrato": row_e["descricao"],
                    "valor_extrato": row_e["valor"],
                    "id_extrato": row_e["id_doc"],
                    "data_controle": controle.loc[idx_c, "data"],
                    "descricao_controle": controle.loc[idx_c, "descricao"],
                    "valor_controle": controle.loc[idx_c, "valor"],
                    "id_controle": controle.loc[idx_c, "id_doc"],
                    "tipo_match": "PERFEITO",
                    "confianca": 100,
                    "diff_dias": 0,
                })
                ext_usado.add(idx_e)
                ctrl_usado.add(idx_c)

    ext_rest = extrato[~extrato.index.isin(ext_usado)].drop(columns=["_chave"])
    ctrl_rest = controle[~controle.index.isin(ctrl_usado)].drop(columns=["_chave"])
    return pd.DataFrame(conciliados), ext_rest, ctrl_rest


def match_aproximado(extrato, controle):
    tol = CONFIG["tolerancia_dias"]
    conciliados = []
    ext_usado, ctrl_usado = set(), set()

    # Indexa controle por valor absoluto para cruzar com débitos do extrato
    ctrl_por_valor = {}
    for idx, row in controle.iterrows():
        ctrl_por_valor.setdefault(round(abs(row["valor"]), 2), []).append(idx)

    for idx_e, row_e in extrato.iterrows():
        if idx_e in ext_usado:
            continue
        vk = round(abs(row_e["valor"]), 2)
        if vk not in ctrl_por_valor:
            continue
        cands = []
        for idx_c in ctrl_por_valor[vk]:
            if idx_c in ctrl_usado:
                continue
            dd = abs((row_e["data"] - controle.loc[idx_c, "data"]).days)
            if dd <= tol:
                cands.append((idx_c, dd))
        if cands:
            cands.sort(key=lambda x: x[1])
            idx_c, dd = cands[0]
            conf = max(50, 100 - dd * 15) if dd > 0 else 95
            conciliados.append({
                "data_extrato": row_e["data"],
                "descricao_extrato": row_e["descricao"],
                "valor_extrato": row_e["valor"],
                "id_extrato": row_e["id_doc"],
                "data_controle": controle.loc[idx_c, "data"],
                "descricao_controle": controle.loc[idx_c, "descricao"],
                "valor_controle": controle.loc[idx_c, "valor"],
                "id_controle": controle.loc[idx_c, "id_doc"],
                "tipo_match": "APROXIMADO",
                "confianca": conf,
                "diff_dias": dd,
            })
            ext_usado.add(idx_e)
            ctrl_usado.add(idx_c)

    ext_rest = extrato[~extrato.index.isin(ext_usado)]
    ctrl_rest = controle[~controle.index.isin(ctrl_usado)]
    return pd.DataFrame(conciliados), ext_rest, ctrl_rest


def executar_conciliacao(extrato, controle):
    c1, er1, cr1 = match_perfeito(extrato, controle)
    c2, er2, cr2 = match_aproximado(er1, cr1)

    if len(c1) > 0 and len(c2) > 0:
        conc = pd.concat([c1, c2], ignore_index=True)
    elif len(c1) > 0:
        conc = c1
    elif len(c2) > 0:
        conc = c2
    else:
        conc = pd.DataFrame()

    return {
        "conciliados": conc,
        "pend_extrato": er2,
        "pend_controle": cr2,
        "stats": {
            "total_extrato": len(extrato),
            "total_controle": len(controle),
            "match_perfeito": len(c1),
            "match_aproximado": len(c2),
            "pend_extrato": len(er2),
            "pend_controle": len(cr2),
        },
    }


def _extrair_nome(desc):
    """Extrai nome de pessoa/empresa da descrição do extrato."""
    desc_lower = desc.lower()
    for prefix in [
        "pix enviado ", "pix recebido ", "transferência pix enviada ",
        "transferência pix recebida ", "transferencia programada ",
        "transferência recebida ", "pagamento de conta ",
    ]:
        if desc_lower.startswith(prefix):
            return desc_lower[len(prefix):].strip()
    return desc_lower.strip()


def encontrar_matches_parciais(pend_ext, pend_ctrl, limite=30):
    """Encontra pares com nome parecido entre pendências do extrato e controle."""
    matches = []
    for _, re_ in pend_ext.iterrows():
        if re_["valor"] >= 0:
            continue
        nome_ext = _extrair_nome(re_["descricao"])
        for _, rc in pend_ctrl.iterrows():
            nome_ctrl = rc["descricao"].lower().strip()
            sim = SequenceMatcher(None, nome_ext, nome_ctrl).ratio()
            if sim > 0.55:
                matches.append({
                    "desc_ext": re_["descricao"],
                    "valor_ext": re_["valor"],
                    "data_ext": re_["data"],
                    "desc_ctrl": rc["descricao"],
                    "valor_ctrl": rc["valor"],
                    "data_ctrl": rc["data"],
                    "similaridade": round(sim * 100),
                    "diff_valor": round(abs(abs(re_["valor"]) - abs(rc["valor"])), 2),
                })
    matches.sort(key=lambda x: (-x["similaridade"], x["diff_valor"]))
    seen = set()
    dedup = []
    for m in matches:
        key = (m["desc_ext"][:30], m["desc_ctrl"][:30], m["valor_ext"], m["valor_ctrl"])
        if key not in seen:
            seen.add(key)
            dedup.append(m)
    return dedup[:limite]


def gerar_timeline(conciliados, pend_ext, pend_ctrl):
    """Gera dados de timeline dia a dia."""
    from collections import defaultdict
    tl = defaultdict(lambda: {"conc": 0, "pe": 0, "pc": 0, "v_conc": 0.0, "v_pend": 0.0})

    if len(conciliados) > 0:
        for _, r in conciliados.iterrows():
            d = pd.to_datetime(r["data_extrato"]).strftime("%d/%m")
            tl[d]["conc"] += 1
            tl[d]["v_conc"] += abs(r["valor_extrato"])

    if len(pend_ext) > 0:
        for _, r in pend_ext.iterrows():
            d = pd.to_datetime(r["data"]).strftime("%d/%m")
            tl[d]["pe"] += 1
            tl[d]["v_pend"] += abs(r["valor"])

    if len(pend_ctrl) > 0:
        for _, r in pend_ctrl.iterrows():
            d = pd.to_datetime(r["data"]).strftime("%d/%m")
            tl[d]["pc"] += 1
            tl[d]["v_pend"] += abs(r["valor"])

    return dict(sorted(tl.items()))


# ==============================================================================
# CÁLCULO FINANCEIRO
# ==============================================================================

def soma_decimal(series):
    total = Decimal("0.00")
    for v in series:
        total += Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return float(total)


def calcular_resumo(res):
    conc = res["conciliados"]
    r = {"cred": 0.0, "deb": 0.0, "saldo": 0.0, "pend_ext": 0.0, "pend_ctrl": 0.0}
    if len(conc) > 0:
        cr = conc[conc["valor_extrato"] > 0]["valor_extrato"]
        db = conc[conc["valor_extrato"] < 0]["valor_extrato"]
        r["cred"] = soma_decimal(cr) if len(cr) else 0.0
        r["deb"] = soma_decimal(db) if len(db) else 0.0
        r["saldo"] = round(r["cred"] + r["deb"], 2)
    if len(res["pend_extrato"]) > 0:
        r["pend_ext"] = soma_decimal(res["pend_extrato"]["valor"])
    if len(res["pend_controle"]) > 0:
        r["pend_ctrl"] = soma_decimal(res["pend_controle"]["valor"])
    return r


def formatar_brl(v):
    sinal = "-" if v < 0 else ""
    return f"{sinal}R$ {abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ==============================================================================
# GERAÇÃO DO EXCEL PARA DOWNLOAD
# ==============================================================================

def gerar_excel(res, resumo) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        # Aba Conciliados
        conc = res["conciliados"]
        if len(conc) > 0:
            dc = conc.copy()
            dc["data_extrato"] = pd.to_datetime(dc["data_extrato"]).dt.strftime("%d/%m/%Y")
            dc["data_controle"] = pd.to_datetime(dc["data_controle"]).dt.strftime("%d/%m/%Y")
            dc["confianca"] = dc["confianca"].astype(str) + "%"
            dc.columns = [
                "Data Extrato", "Desc. Extrato", "Valor Extrato", "ID Extrato",
                "Data Controle", "Desc. Controle", "Valor Controle", "ID Controle",
                "Tipo Match", "Confiança", "Diff Dias",
            ]
            dc.to_excel(w, sheet_name="Conciliados", index=False)
        else:
            pd.DataFrame({"Info": ["Nenhum item conciliado"]}).to_excel(
                w, sheet_name="Conciliados", index=False
            )

        # Aba Pendências Banco
        pe = res["pend_extrato"]
        if len(pe) > 0:
            dp = pe[["data", "descricao", "valor", "id_doc", "tipo"]].copy()
            dp["data"] = pd.to_datetime(dp["data"]).dt.strftime("%d/%m/%Y")
            dp.columns = ["Data", "Descrição", "Valor", "ID/Doc", "Tipo"]
            dp.to_excel(w, sheet_name="Pendências Banco", index=False)
        else:
            pd.DataFrame({"Info": ["Sem pendências"]}).to_excel(
                w, sheet_name="Pendências Banco", index=False
            )

        # Aba Pendências Controle
        pc = res["pend_controle"]
        if len(pc) > 0:
            dp2 = pc[["data", "descricao", "valor", "id_doc", "tipo"]].copy()
            dp2["data"] = pd.to_datetime(dp2["data"]).dt.strftime("%d/%m/%Y")
            dp2.columns = ["Data", "Descrição", "Valor", "ID/Doc", "Tipo"]
            dp2.to_excel(w, sheet_name="Pendências Controle", index=False)
        else:
            pd.DataFrame({"Info": ["Sem pendências"]}).to_excel(
                w, sheet_name="Pendências Controle", index=False
            )

        # Aba Resumo
        s = res["stats"]
        tc = s["match_perfeito"] + s["match_aproximado"]
        tt = max(s["total_extrato"], s["total_controle"])
        rows = [
            ["ESTATÍSTICAS", ""],
            ["Lançamentos no Extrato", s["total_extrato"]],
            ["Lançamentos no Controle", s["total_controle"]],
            ["Matches Perfeitos", s["match_perfeito"]],
            ["Matches Aproximados", s["match_aproximado"]],
            ["Total Conciliados", tc],
            ["Pendentes Extrato", s["pend_extrato"]],
            ["Pendentes Controle", s["pend_controle"]],
            ["Taxa de Conciliação", f"{(tc / tt * 100) if tt else 0:.1f}%"],
            ["", ""],
            ["FINANCEIRO", ""],
            ["Créditos Conciliados", formatar_brl(resumo["cred"])],
            ["Débitos Conciliados", formatar_brl(resumo["deb"])],
            ["Saldo Conciliado", formatar_brl(resumo["saldo"])],
            ["Pendências Extrato", formatar_brl(resumo["pend_ext"])],
            ["Pendências Controle", formatar_brl(resumo["pend_ctrl"])],
        ]
        pd.DataFrame(rows, columns=["Métrica", "Valor"]).to_excel(
            w, sheet_name="Resumo", index=False
        )

        # Auto-ajuste de largura em todas as abas
        for sn in w.sheets:
            ws = w.sheets[sn]
            ws.set_column("A:K", 20)

    return buf.getvalue()


# ==============================================================================
# GRÁFICOS PLOTLY
# ==============================================================================

def grafico_distribuicao(stats):
    labels = ["Match Perfeito", "Match Aproximado", "Pend. Extrato", "Pend. Controle"]
    values = [stats["match_perfeito"], stats["match_aproximado"], stats["pend_extrato"], stats["pend_controle"]]
    colors = ["#10b981", "#3b82f6", "#f59e0b", "#ef4444"]

    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.55,
        marker=dict(colors=colors, line=dict(color="white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>%{value} itens<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Distribuição dos Resultados", font=dict(size=16)),
        height=380,
        margin=dict(t=60, b=20, l=20, r=20),
        legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def grafico_financeiro(resumo):
    cats = ["Créditos\nConciliados", "Débitos\nConciliados", "Pend.\nExtrato", "Pend.\nControle"]
    vals = [resumo["cred"], abs(resumo["deb"]), resumo["pend_ext"], abs(resumo["pend_ctrl"])]
    colors = ["#10b981", "#ef4444", "#f59e0b", "#8b5cf6"]

    fig = go.Figure(go.Bar(
        x=cats, y=vals,
        marker=dict(color=colors, cornerradius=6),
        text=[formatar_brl(v) for v in [resumo["cred"], resumo["deb"], resumo["pend_ext"], resumo["pend_ctrl"]]],
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>%{text}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Resumo Financeiro", font=dict(size=16)),
        height=380,
        margin=dict(t=60, b=60, l=60, r=20),
        yaxis=dict(title="Valor (R$)", gridcolor="#f1f5f9"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
    )
    return fig


def grafico_timeline(conciliados):
    if len(conciliados) == 0:
        return None
    df = conciliados.copy()
    df["data_extrato"] = pd.to_datetime(df["data_extrato"])
    daily = df.groupby(df["data_extrato"].dt.date).agg(
        qtd=("valor_extrato", "count"),
        total=("valor_extrato", "sum"),
    ).reset_index()
    daily.columns = ["data", "qtd", "total"]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily["data"], y=daily["qtd"],
        mode="lines+markers",
        name="Qtd. Conciliados",
        line=dict(color="#3b82f6", width=2),
        marker=dict(size=6),
        hovertemplate="<b>%{x}</b><br>%{y} itens<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Conciliações por Dia", font=dict(size=16)),
        height=320,
        margin=dict(t=60, b=40, l=60, r=20),
        xaxis=dict(title="Data"),
        yaxis=dict(title="Quantidade", gridcolor="#f1f5f9"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def grafico_confianca(conciliados):
    if len(conciliados) == 0:
        return None
    df = conciliados.copy()
    bins = [0, 59, 79, 99, 100]
    labels_b = ["50-59%", "60-79%", "80-99%", "100%"]
    df["faixa"] = pd.cut(df["confianca"], bins=bins, labels=labels_b, include_lowest=True)
    counts = df["faixa"].value_counts().reindex(labels_b, fill_value=0)
    colors_b = ["#ef4444", "#f59e0b", "#3b82f6", "#10b981"]

    fig = go.Figure(go.Bar(
        x=counts.index, y=counts.values,
        marker=dict(color=colors_b, cornerradius=6),
        text=counts.values, textposition="outside",
        hovertemplate="<b>%{x}</b><br>%{y} itens<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Distribuição de Confiança", font=dict(size=16)),
        height=320,
        margin=dict(t=60, b=40, l=60, r=20),
        yaxis=dict(title="Quantidade", gridcolor="#f1f5f9"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# ==============================================================================
# APP PRINCIPAL
# ==============================================================================

def main():

    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🏦 Conciliação Bancária</h1>
        <p>Extrato Bancário × Controle Interno — Upload, processe e baixe o relatório</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar config ──
    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        tol = st.slider(
            "Tolerância de dias (Match Aproximado)",
            min_value=1, max_value=10, value=3,
            help="Janela em dias para considerar datas como 'próximas' no match aproximado.",
        )
        CONFIG["tolerancia_dias"] = tol

        st.markdown("---")
        st.markdown("### 📋 Colunas Obrigatórias")
        st.markdown("""
        Ambos os arquivos devem conter:
        - **Data** (data, dt, data_mov...)
        - **Descrição** (descricao, historico...)
        - **Valor** (valor, value, amount...)
        - **ID/Documento** (id, doc, nsu...)
        """)

        st.markdown("---")
        st.markdown("### 📖 Como funciona")
        st.markdown("""
        1. **Upload** dos dois arquivos
        2. **Prévia** dos dados carregados
        3. **Processar** a conciliação
        4. **Analisar** gráficos e tabelas
        5. **Download** do Excel final
        """)

    # ── Upload Section ──
    st.markdown("## 📁 Upload dos Arquivos")
    col_up1, col_up2 = st.columns(2)

    with col_up1:
        st.markdown("#### Extrato Bancário")
        file_extrato = st.file_uploader(
            "Arraste ou clique para enviar",
            type=["csv", "xlsx", "xls"],
            key="extrato",
            help="Arquivo do banco (.csv ou .xlsx)",
        )

    with col_up2:
        st.markdown("#### Controle Interno / ERP")
        file_controle = st.file_uploader(
            "Arraste ou clique para enviar",
            type=["csv", "xlsx", "xls"],
            key="controle",
            help="Arquivo do seu sistema (.csv ou .xlsx)",
        )

    # ── Prévia ──
    if file_extrato and file_controle:
        try:
            raw_ext = ler_upload(file_extrato)
            raw_ctrl = ler_upload(file_controle)
        except Exception as e:
            st.error(f"Erro ao ler arquivos: {e}")
            return

        st.markdown("---")
        st.markdown("## 👁️ Prévia dos Dados Brutos")

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown(f"**Extrato** — {len(raw_ext)} linhas, {len(raw_ext.columns)} colunas")
            st.dataframe(raw_ext.head(8), use_container_width=True, height=300)
        with col_p2:
            st.markdown(f"**Controle** — {len(raw_ctrl)} linhas, {len(raw_ctrl.columns)} colunas")
            st.dataframe(raw_ctrl.head(8), use_container_width=True, height=300)

        # ── Botão Processar ──
        st.markdown("---")
        col_btn = st.columns([1, 2, 1])
        with col_btn[1]:
            processar = st.button(
                "⚡ Processar Conciliação",
                use_container_width=True,
                type="primary",
            )

        if processar:
            with st.spinner("Normalizando dados..."):
                try:
                    df_ext = mapear_colunas(raw_ext.copy())
                    df_ext = normalizar_df(df_ext)
                    df_ctrl = mapear_colunas(raw_ctrl.copy())

                    # --- Filtro automático por conta corrente ---
                    col_conta = None
                    for c in df_ctrl.columns:
                        if "conta_corrente" in c or "conta" in c and "valor" not in c:
                            col_conta = c
                            break

                    if col_conta and df_ctrl[col_conta].nunique() > 1:
                        contas = df_ctrl[col_conta].dropna().unique().tolist()
                        nome_ext = file_extrato.name.lower()

                        # Tenta detectar qual conta bate com o nome do extrato
                        conta_filtro = None
                        for conta in contas:
                            palavras = [p for p in conta.lower().split() if len(p) > 3]
                            if any(p in nome_ext for p in palavras):
                                conta_filtro = conta
                                break

                        if conta_filtro:
                            total_antes = len(df_ctrl)
                            df_ctrl = df_ctrl[df_ctrl[col_conta] == conta_filtro]
                            st.info(
                                f"🏦 Filtro automático: **{conta_filtro}** "
                                f"({len(df_ctrl)} de {total_antes} registros)"
                            )
                        elif len(contas) > 1:
                            st.warning(
                                f"⚠️ O controle tem {len(contas)} contas correntes. "
                                f"Não foi possível filtrar automaticamente. "
                                f"Contas: {', '.join(str(c) for c in contas)}"
                            )

                    df_ctrl = normalizar_df(df_ctrl)
                except Exception as e:
                    st.error(f"Erro na normalização: {e}")
                    return

            with st.spinner("Executando conciliação..."):
                resultado = executar_conciliacao(df_ext, df_ctrl)
                resumo = calcular_resumo(resultado)

            st.session_state["resultado"] = resultado
            st.session_state["resumo"] = resumo

        # ── Resultados ──
        if "resultado" in st.session_state:
            resultado = st.session_state["resultado"]
            resumo = st.session_state["resumo"]
            stats = resultado["stats"]
            tc = stats["match_perfeito"] + stats["match_aproximado"]
            tt = max(stats["total_extrato"], stats["total_controle"])
            pct = (tc / tt * 100) if tt else 0
            val_conc = soma_decimal(resultado["conciliados"]["valor_extrato"].abs()) if len(resultado["conciliados"]) > 0 else 0
            val_pe = soma_decimal(resultado["pend_extrato"]["valor"].abs()) if len(resultado["pend_extrato"]) > 0 else 0
            val_pc = soma_decimal(resultado["pend_controle"]["valor"].abs()) if len(resultado["pend_controle"]) > 0 else 0

            st.markdown("---")
            st.markdown("## 📊 Resultados da Conciliação")

            # --- Métricas ---
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-card">
                    <div class="label">Taxa de Conciliação</div>
                    <div class="value" style="color: {'#10b981' if pct >= 80 else '#f59e0b' if pct >= 50 else '#ef4444'}">{pct:.1f}%</div>
                    <div class="sub">{tc} de {tt} itens</div>
                </div>
                <div class="metric-card">
                    <div class="label">Conciliados</div>
                    <div class="value" style="color: #10b981">{tc}</div>
                    <div class="sub">{formatar_brl(val_conc)}</div>
                </div>
                <div class="metric-card">
                    <div class="label">Pend. Extrato</div>
                    <div class="value" style="color: #f59e0b">{stats['pend_extrato']}</div>
                    <div class="sub">{formatar_brl(val_pe)}</div>
                </div>
                <div class="metric-card">
                    <div class="label">Pend. Controle</div>
                    <div class="value" style="color: #ef4444">{stats['pend_controle']}</div>
                    <div class="sub">{formatar_brl(val_pc)}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # --- Gráficos ---
            st.markdown("### 📈 Análise Visual")
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.plotly_chart(grafico_distribuicao(stats), use_container_width=True)
            with g_col2:
                st.plotly_chart(grafico_financeiro(resumo), use_container_width=True)

            # --- Timeline ---
            st.markdown("### 📅 Linha do Tempo — Dia a Dia")
            tl_data = gerar_timeline(
                resultado["conciliados"], resultado["pend_extrato"], resultado["pend_controle"]
            )
            if tl_data:
                tl_df = pd.DataFrame([
                    {"Data": k, "Conciliados": v["conc"], "Pend. Extrato": v["pe"], "Pend. Controle": v["pc"]}
                    for k, v in tl_data.items()
                ])
                fig_tl = go.Figure()
                fig_tl.add_trace(go.Bar(
                    x=tl_df["Data"], y=tl_df["Conciliados"], name="Conciliados",
                    marker_color="#10b981",
                ))
                fig_tl.add_trace(go.Bar(
                    x=tl_df["Data"], y=tl_df["Pend. Extrato"], name="Pend. Extrato",
                    marker_color="#f59e0b",
                ))
                fig_tl.add_trace(go.Bar(
                    x=tl_df["Data"], y=tl_df["Pend. Controle"], name="Pend. Controle",
                    marker_color="#ef4444",
                ))
                fig_tl.update_layout(
                    barmode="stack", height=320,
                    margin=dict(t=20, b=40, l=40, r=20),
                    legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center"),
                    xaxis=dict(title="Data", tickangle=-45),
                    yaxis=dict(title="Qtd.", gridcolor="#f1f5f9"),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                )
                st.plotly_chart(fig_tl, use_container_width=True)

            # --- Matches Parciais ---
            st.markdown("### 🔍 Matches Parciais — Possíveis Correspondências")
            st.caption("Nomes parecidos entre extrato e controle que não bateram por valor ou data. Podem ser o mesmo pagamento.")
            parciais = encontrar_matches_parciais(resultado["pend_extrato"], resultado["pend_controle"])

            if parciais:
                df_parciais = pd.DataFrame(parciais)
                df_parciais["data_ext"] = pd.to_datetime(df_parciais["data_ext"]).dt.strftime("%d/%m/%Y")
                df_parciais["data_ctrl"] = pd.to_datetime(df_parciais["data_ctrl"]).dt.strftime("%d/%m/%Y")

                display_p = df_parciais[[
                    "desc_ext", "valor_ext", "data_ext",
                    "desc_ctrl", "valor_ctrl", "data_ctrl",
                    "similaridade", "diff_valor",
                ]].copy()
                display_p["valor_ext"] = display_p["valor_ext"].apply(formatar_brl)
                display_p["valor_ctrl"] = display_p["valor_ctrl"].apply(formatar_brl)
                display_p["diff_valor"] = display_p["diff_valor"].apply(
                    lambda v: "✅ Valor OK" if v == 0 else formatar_brl(v)
                )
                display_p["similaridade"] = display_p["similaridade"].astype(str) + "%"
                display_p.columns = [
                    "Extrato", "Valor Ext.", "Data Ext.",
                    "Controle", "Valor Ctrl.", "Data Ctrl.",
                    "Similaridade", "Δ Valor",
                ]
                st.dataframe(display_p, use_container_width=True, height=400)
            else:
                st.info("Nenhum match parcial encontrado.")

            # --- Tabelas detalhadas ---
            st.markdown("### 📋 Detalhamento")
            tab1, tab2, tab3, tab4 = st.tabs([
                f"✅ Conciliados ({tc})",
                f"⚠️ Pend. Banco ({stats['pend_extrato']})",
                f"🔴 Pend. Controle ({stats['pend_controle']})",
                f"📈 Confiança",
            ])

            with tab1:
                conc = resultado["conciliados"]
                if len(conc) > 0:
                    display = conc.copy()
                    display["data_extrato"] = pd.to_datetime(display["data_extrato"]).dt.strftime("%d/%m/%Y")
                    display["data_controle"] = pd.to_datetime(display["data_controle"]).dt.strftime("%d/%m/%Y")
                    display["valor_extrato"] = display["valor_extrato"].apply(formatar_brl)
                    display["valor_controle"] = display["valor_controle"].apply(formatar_brl)
                    display["confianca"] = display["confianca"].astype(str) + "%"
                    display.columns = [
                        "Data Ext.", "Desc. Ext.", "Valor Ext.", "ID Ext.",
                        "Data Ctrl.", "Desc. Ctrl.", "Valor Ctrl.", "ID Ctrl.",
                        "Match", "Conf.", "Δ Dias",
                    ]
                    st.dataframe(display, use_container_width=True, height=400)
                else:
                    st.info("Nenhum item conciliado.")

            with tab2:
                pe = resultado["pend_extrato"]
                if len(pe) > 0:
                    dp = pe[["data", "descricao", "valor", "id_doc", "tipo"]].copy()
                    dp["data"] = pd.to_datetime(dp["data"]).dt.strftime("%d/%m/%Y")
                    dp["valor"] = dp["valor"].apply(formatar_brl)
                    dp.columns = ["Data", "Descrição", "Valor", "ID/Doc", "Tipo"]
                    st.dataframe(dp, use_container_width=True, height=400)
                else:
                    st.success("Nenhuma pendência no extrato bancário.")

            with tab3:
                pc = resultado["pend_controle"]
                if len(pc) > 0:
                    dp2 = pc[["data", "descricao", "valor", "id_doc", "tipo"]].copy()
                    dp2["data"] = pd.to_datetime(dp2["data"]).dt.strftime("%d/%m/%Y")
                    dp2["valor"] = dp2["valor"].apply(formatar_brl)
                    dp2.columns = ["Data", "Descrição", "Valor", "ID/Doc", "Tipo"]
                    st.dataframe(dp2, use_container_width=True, height=400)
                else:
                    st.success("Nenhuma pendência no controle interno.")

            with tab4:
                fig_cf = grafico_confianca(resultado["conciliados"])
                if fig_cf:
                    st.plotly_chart(fig_cf, use_container_width=True)
                else:
                    st.info("Sem dados de confiança.")

            # --- Download ---
            st.markdown("---")
            st.markdown("## 📥 Download do Relatório")
            excel_bytes = gerar_excel(resultado, resumo)

            col_dl = st.columns([1, 2, 1])
            with col_dl[1]:
                st.download_button(
                    label="⬇️  Baixar resultado_conciliacao.xlsx",
                    data=excel_bytes,
                    file_name="resultado_conciliacao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
                st.caption("Excel com abas: Conciliados, Pendências Banco, Pendências Controle e Resumo.")

    elif file_extrato or file_controle:
        st.info("📎 Envie ambos os arquivos para continuar.")


if __name__ == "__main__":
    main()
