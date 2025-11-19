import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import re
from datetime import datetime, date


# -------------------------------------------------------
# FUNÇÕES UTILITÁRIAS
# -------------------------------------------------------

def normalizar_nome(nome):
    if not isinstance(nome, str):
        return nome
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode()
    nome = nome.lower().strip()
    nome = re.sub(r"\s+", " ", nome)
    return nome


def formatar_data_coluna(dt):
    meses = {
        1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
        7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
    }
    dt = pd.to_datetime(dt)
    return f"{meses[dt.month]}/{str(dt.year)[2:]}"


DATE_PATTERNS = [
    r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$",
    r"^\d{1,2}[/-]\d{2,4}$",
    r"^[A-Za-zÀ-ÿ]{3,}[/-]\d{2,4}$",
    r"^\d{4}-\d{2}-\d{2}$",
]


def parece_data_str(s):
    s = s.strip()
    if " " in s:
        return False
    for pat in DATE_PATTERNS:
        if re.match(pat, s, flags=re.IGNORECASE):
            return True
    return False


def tentar_converter_para_data(valor):
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(valor)

    if isinstance(valor, (int, float)) and not pd.isna(valor):
        if 10000 <= valor <= 80000:
            try:
                return pd.to_datetime(int(valor), unit="D", origin="1899-12-30")
            except:
                return None

    if isinstance(valor, str):
        s = valor.strip()
        if s == "":
            return None
        if re.fullmatch(r"\d+", s):
            num = int(s)
            if 10000 <= num <= 80000:
                return pd.to_datetime(num, unit="D", origin="1899-12-30")
            return None
        if parece_data_str(s):
            try:
                return pd.to_datetime(s, dayfirst=True)
            except:
                try:
                    return pd.to_datetime(s)
                except:
                    return None
    return None


def mapear_fixas(df_cols):
    padrao = ["Regional", "Empreendimento", "Módulo", "Unidades", "Tipologia", "Fonte Curva"]
    norm = {normalizar_nome(c): c for c in df_cols}
    return {p: norm.get(normalizar_nome(p), None) for p in padrao}


# -------------------------------------------------------
# VP POR EMPREENDIMENTO
# -------------------------------------------------------

def calcular_vp_por_empreendimento(df, date_cols):

    resultado = {}

    for emp in df["Empreendimento"].unique():
        dfe = df[df["Empreendimento"] == emp]

        meses = dfe[date_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
        soma_up = meses.where(meses > 0).sum().sum()

        meses_com_up = (meses > 0).any(axis=0).sum()
        unidades_total = dfe["Unidades"].astype(float).sum()

        if meses_com_up > 0 and unidades_total > 0:
            resultado[emp] = soma_up / (unidades_total * meses_com_up)
        else:
            resultado[emp] = np.nan

    return df["Empreendimento"].map(resultado)


# -------------------------------------------------------
# VP POR ANO
# -------------------------------------------------------

def calcular_indicadores_ano(df, date_cols, datas, ano):

    anos = [dt.year for _, dt in datas]
    cols_ano = [date_cols[i] for i, a in enumerate(anos) if a == ano]

    if not cols_ano:
        return np.nan, 0.0

    meses = df[cols_ano].apply(pd.to_numeric, errors="coerce").fillna(0)
    unid_total = meses.sum().sum()

    soma_up_global = 0
    denom_global = 0

    for emp in df["Empreendimento"].unique():
        dfe = df[df["Empreendimento"] == emp]
        m_emp = dfe[cols_ano].apply(pd.to_numeric, errors="coerce").fillna(0)

        soma_up = m_emp.where(m_emp > 0).sum().sum()
        meses_com_up = (m_emp > 0).any(axis=0).sum()
        unidades_total = dfe["Unidades"].astype(float).sum()

        if meses_com_up > 0 and unidades_total > 0:
            soma_up_global += soma_up
            denom_global += unidades_total * meses_com_up

    if denom_global == 0:
        return np.nan, unid_total

    vp = soma_up_global / denom_global
    return vp, unid_total


def format_unidades(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# -------------------------------------------------------
# APP SIMPLIFICADO
# -------------------------------------------------------

def render():

    st.header("📊 Indicadores VP & UP — Forecast")

    up = st.file_uploader(
        "Carregue o arquivo Forecast (.xlsx, .xls, .xlsb)",
        type=["xlsx", "xls", "xlsb"]
    )
    if up is None:
        return

    try:
        # Leitura padrão que funciona para todos
        if up.name.lower().endswith(".xlsb"):
            df = pd.read_excel(up, engine="pyxlsb", sheet_name="Forecast", header=2)
        else:
            df = pd.read_excel(up, sheet_name="Forecast", header=2)

        df.columns = df.columns.map(str)

        fixas = mapear_fixas(df.columns)
        if not fixas["Empreendimento"]:
            st.error("A coluna 'Empreendimento' não foi encontrada.")
            return

        col_unid = df.columns[9]

        col_fixas = []
        for k in ["Regional", "Empreendimento", "Módulo"]:
            if fixas[k]:
                col_fixas.append(fixas[k])
        col_fixas.append(col_unid)
        for k in ["Tipologia", "Fonte Curva"]:
            if fixas[k]:
                col_fixas.append(fixas[k])

        ignorar = set(v for v in fixas.values() if v)
        ignorar.add(col_unid)

        datas = []
        for c in df.columns:
            if c in ignorar:
                continue
            dt = tentar_converter_para_data(c)
            if dt is not None:
                datas.append((c, pd.to_datetime(dt)))
        datas.sort(key=lambda x: x[1])

        cols_datas = [c for c, _ in datas]
        df_out = df[col_fixas + cols_datas].copy()

        df_out["Unidades"] = pd.to_numeric(df_out["Unidades"], errors="coerce").fillna(0)

        date_cols = df_out.columns[len(col_fixas):]

        df_out["VP"] = calcular_vp_por_empreendimento(df_out, date_cols)

        vp_2026, unid_2026 = calcular_indicadores_ano(df_out, date_cols, datas, 2026)
        vp_2027, unid_2027 = calcular_indicadores_ano(df_out, date_cols, datas, 2027)

        # ---------------------------
        # 🔥 EXIBIR RESULTADOS
        # ---------------------------
        st.subheader("Indicadores Carregados da Planilha")

        c1, c2 = st.columns(2)
        with c1:
            st.metric("VP 2026", f"{vp_2026*100:.2f}%".replace(".", ","))
            st.metric("Unidades 2026", format_unidades(unid_2026))

        with c2:
            st.metric("VP 2027", f"{vp_2027*100:.2f}%".replace(".", ","))
            st.metric("Unidades 2027", format_unidades(unid_2027))

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
