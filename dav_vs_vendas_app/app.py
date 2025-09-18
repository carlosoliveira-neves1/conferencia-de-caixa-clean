
import re
import io
from datetime import datetime, date
import pandas as pd
import streamlit as st

st.set_page_config(page_title="DAV x Vendas - Casa do Cigano", page_icon="📊", layout="wide")

# ---- THEME / STYLE ----
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.0rem; padding-bottom: 2rem;}
    .metric-card {background: #0b1220; padding: 14px 16px; border-radius: 16px; border: 1px solid #1f2937;}
    .metric-title {font-size: 0.80rem; color: #9CA3AF; margin-bottom: 6px;}
    .metric-value {font-size: 1.4rem; color: #FDE68A; font-weight: 700;}
    .sub {color:#9CA3AF; font-size: 12px;}
    .tag {display:inline-block; padding: 2px 8px; background:#111827; border:1px solid #1f2937; border-radius:12px; color:#e5e7eb; font-size:12px;}
    .footer-note {color:#9CA3AF; font-size: 12px; margin-top: 8px;}
    </style>
    """,
    unsafe_allow_html=True
)

# ---- HEADER ----
cols = st.columns([1,6])
with cols[0]:
    st.image("assets/logo.svg", width=140)
with cols[1]:
    st.title("Conferência de DAVs x Vendas Fechadas")
    st.caption("Cruzamento diário pelo número da venda • Verificação de valores • Resumo por vendedor • Dashboards")

# ---- SIDEBAR: UPLOADS ----
st.sidebar.header("Arquivos")
mov_file = st.sidebar.file_uploader("Movimento Diário (.xlsx)", type=["xlsx"], key="mov")
ven_file = st.sidebar.file_uploader("Minhas Vendas (.xlsx)", type=["xlsx"], key="ven")

# ---- Helpers ----
def read_first_sheet(xls_file):
    try:
        xl = pd.ExcelFile(xls_file)
        preferred = ["MovimentoDiario", "Movimento Diário", "Planilha1", "Sheet1"]
        use = xl.sheet_names[0]
        for p in preferred:
            if p in xl.sheet_names:
                use = p
                break
        return pd.read_excel(xl, sheet_name=use)
    except Exception as e:
        st.error(f"Erro ao ler planilha: {e}")
        return None

def extract_numero_mov(s):
    if pd.isna(s): return None
    m = re.search(r"(\d+)", str(s))
    return int(m.group(1)) if m else None

def extract_numero_vendas(s):
    if pd.isna(s): return None
    m = re.search(r"NF.*?(\d+)", str(s))
    return int(m.group(1)) if m else None

def parse_vendedor_line(txt):
    # Ex.: "Vendedor: 36 - ANA PAULA DOS SANTOS"
    if not isinstance(txt, str): return None, None, None
    if not txt.strip().lower().startswith("vendedor:"): return None, None, None
    m = re.match(r"Vendedor:\s*(\d+)\s*-\s*(.+)", txt.strip(), flags=re.IGNORECASE)
    if m:
        code = m.group(1)
        name = m.group(2).strip()
        return txt.strip(), code, name
    return txt.strip(), None, None

def normalize_mov_with_vendors(df):
    # Renomear campos básicos
    cols_map = {"Doc/Emp": "doc_emp","Valor do Documento": "valor_doc","Data": "data_mov","Série": "serie","Cliente": "cliente_mov"}
    df2 = df.copy()
    for k,v in cols_map.items():
        if k in df2.columns: df2.rename(columns={k:v}, inplace=True)

    # Encontrar linhas que são "Vendedor:" em QUALQUER coluna
    vend_marker_col = None
    vendor_raw = []
    for col in df2.columns:
        # Procura o primeiro valor que comece com Vendedor:
        mask = df2[col].astype(str).str.strip().str.lower().str.startswith("vendedor:")
        if mask.any():
            vend_marker_col = col
            vendor_raw = df2.loc[mask, col].astype(str).tolist()
            break

    # Construir coluna auxiliar 'vendor_header' com texto do vendedor apenas nas linhas que são cabeçalho
    df2["vendor_header"] = None
    if vend_marker_col:
        mask = df2[vend_marker_col].astype(str).str.strip().str.lower().str.startswith("vendedor:")
        df2.loc[mask, "vendor_header"] = df2.loc[mask, vend_marker_col].astype(str)
    else:
        # fallback: procura especificamente na coluna 'Data' (comum no exemplo)
        if "data_mov" in df2.columns:
            mask = df2["data_mov"].astype(str).str.strip().str.lower().str.startswith("vendedor:")
            df2.loc[mask, "vendor_header"] = df2.loc[mask, "data_mov"].astype(str)

    # Forward-fill para propagar o vendedor para as linhas de vendas até o próximo cabeçalho
    df2["vendor_header"] = df2["vendor_header"].ffill()

    # Extrai code e name
    parsed = df2["vendor_header"].apply(lambda x: parse_vendedor_line(x)[1:] if isinstance(x, str) else (None, None))
    df2["vendedor_codigo"] = [p[0] for p in parsed]
    df2["vendedor_nome"]   = [p[1] for p in parsed]

    # Número da venda
    df2["NumeroVenda"] = df2.get("doc_emp").apply(extract_numero_mov) if "doc_emp" in df2.columns else None

    return df2

def normalize_vendas(df):
    cols_map = {"Doc.": "doc","Emitido em": "emitido_em","Cliente": "cliente_ven","Valor": "valor_ven","Origem": "origem","Status": "status"}
    df2 = df.copy()
    for k,v in cols_map.items():
        if k in df2.columns: df2.rename(columns={k:v}, inplace=True)
    df2["NumeroVenda"] = df2.get("doc").apply(extract_numero_vendas) if "doc" in df2.columns else None
    if "emitido_em" in df2.columns:
        df2["emitido_em"] = pd.to_datetime(df2["emitido_em"], errors="coerce")
        df2["hora"] = df2["emitido_em"].dt.hour
        df2["data"] = df2["emitido_em"].dt.date
    return df2

def to_excel_bytes(sheets: dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for name, d in sheets.items():
            d.to_excel(writer, sheet_name=name, index=False)
    return out.getvalue()

# ---- PROCESSAMENTO ----
if mov_file and ven_file:
    df_mov = read_first_sheet(mov_file)
    df_ven = read_first_sheet(ven_file)

    if df_mov is not None and df_ven is not None:
        # Normalizações com vendedores
        mov_norm = normalize_mov_with_vendors(df_mov)
        ven_norm = normalize_vendas(df_ven)

        mov_ok = mov_norm.dropna(subset=["NumeroVenda"])
        ven_ok = ven_norm.dropna(subset=["NumeroVenda"])

        merged = ven_ok.merge(mov_ok, on="NumeroVenda", how="outer", suffixes=("_ven", "_mov"))

        total_mov = int(mov_ok["NumeroVenda"].nunique()) if not mov_ok.empty else 0
        total_ven = int(ven_ok["NumeroVenda"].nunique()) if not ven_ok.empty else 0

        faltando_em_mov = merged[merged["doc_emp"].isna()]
        faltando_em_ven = merged[merged["doc"].isna()]

        comp = merged.dropna(subset=["valor_ven", "valor_doc"]).copy()
        comp["Diferenca"] = comp["valor_ven"] - comp["valor_doc"]
        divergentes = comp[comp["Diferenca"].abs() > 0.01]

        total_valor_ven = float(ven_ok["valor_ven"].sum()) if "valor_ven" in ven_ok.columns else 0.0
        total_valor_mov = float(mov_ok["valor_doc"].sum()) if "valor_doc" in mov_ok.columns else 0.0
        ticket_medio_ven = (total_valor_ven / total_ven) if total_ven else 0.0
        ticket_medio_mov = (total_valor_mov / total_mov) if total_mov else 0.0
        taxa_div = (len(divergentes) / max(len(comp), 1)) * 100.0

        # KPIs
        st.subheader("Resumo do Dia")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Qtd Vendas (Vendas)</div><div class="metric-value">{total_ven}</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Qtd Vendas (Movimento)</div><div class="metric-value">{total_mov}</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Ticket Médio (Vendas)</div><div class="metric-value">{ticket_medio_ven:,.2f}</div><div class="sub">col. Valor</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Ticket Médio (Mov.)</div><div class="metric-value">{ticket_medio_mov:,.2f}</div><div class="sub">col. Valor do Documento</div></div>', unsafe_allow_html=True)
        with c5:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Divergências</div><div class="metric-value">{len(divergentes)}</div><div class="sub">{taxa_div:,.1f}%</div></div>', unsafe_allow_html=True)
        with c6:
            st.markdown(f'<div class="metric-card"><div class="metric-title">Diferença Total</div><div class="metric-value">{comp["Diferenca"].sum():,.2f}</div><div class="sub">Σ (Vendas - Mov.)</div></div>', unsafe_allow_html=True)

        # Lista de vendedores detectados (únicos) a partir do Movimento
        vendedores = mov_norm["vendor_header"].dropna().unique().tolist()
        if vendedores:
            st.markdown("**Vendedores detectados:** " + " • ".join(vendedores))

        # ---- DASHBOARDS / RELATÓRIOS ----
        tab_dash, tab_conf, tab_falta_mov, tab_falta_ven, tab_div, tab_vend, tab_export = st.tabs(
            ["📈 Dashboard Geral", "📋 Conferência", "🚫 Faltando no Movimento", "🟠 Faltando em Minhas Vendas", "⚠️ Divergências", "👤 Por Vendedor", "⬇️ Exportar"]
        )

        with tab_dash:
            st.markdown("#### Vendas por Hora (Minhas Vendas)")
            if "hora" in ven_ok.columns and not ven_ok.empty:
                por_hora = ven_ok.groupby("hora", dropna=False)["valor_ven"].sum().reset_index().sort_values("hora")
                por_hora["hora"] = por_hora["hora"].fillna(-1).astype(int)
                st.bar_chart(por_hora.set_index("hora"))
            else:
                st.info("Sem coluna de horário em 'Emitido em' para analisar por hora.")

            st.markdown("#### Top Clientes (Minhas Vendas)")
            if "cliente_ven" in ven_ok.columns:
                top_cli = ven_ok.groupby("cliente_ven")["valor_ven"].sum().reset_index().sort_values("valor_ven", ascending=False).head(10)
                st.bar_chart(top_cli.set_index("cliente_ven"))
            else:
                st.info("Coluna de cliente não encontrada em Minhas Vendas.")

            st.markdown("#### Origem da Venda (Minhas Vendas)")
            if "origem" in ven_ok.columns:
                por_origem = ven_ok.groupby("origem")["valor_ven"].sum().reset_index().sort_values("valor_ven", ascending=False)
                st.bar_chart(por_origem.set_index("origem"))
            else:
                st.info("Coluna 'Origem' não encontrada.")

            st.markdown("#### Formas de Pagamento (Movimento Diário)")
            fp_cols = [c for c in mov_ok.columns if c.lower() in ["dinheiro","cartão","cartao","pix","ch.vista","ch.prazo","crediário","crediario","convênio","convenio","outras moedas"]]
            if fp_cols:
                fp_df = pd.DataFrame()
                for c in fp_cols:
                    try:
                        vals = pd.to_numeric(mov_ok[c].astype(str).str.replace(",", ".", regex=False).str.replace("-", "0", regex=False), errors="coerce")
                    except Exception:
                        vals = pd.to_numeric(mov_ok[c], errors="coerce")
                    fp_df = pd.concat([fp_df, pd.DataFrame({"forma":[c], "total":[vals.fillna(0).sum()]})])
                fp_df = fp_df.groupby("forma")["total"].sum().reset_index().sort_values("total", ascending=False)
                st.bar_chart(fp_df.set_index("forma"))
            else:
                st.info("Não localizei colunas de formas de pagamento usuais.")

        with tab_conf:
            st.markdown("#### Tabela Consolidada")
            show_cols = ["NumeroVenda", "doc", "valor_ven", "doc_emp", "valor_doc", "vendedor_codigo", "vendedor_nome"]
            cols_exist = [c for c in show_cols if c in merged.columns]
            st.dataframe(merged[cols_exist].sort_values("NumeroVenda"), use_container_width=True)

        with tab_falta_mov:
            if faltando_em_mov.empty:
                st.success("Nenhum registro faltando no Movimento Diário. ✅")
            else:
                st.warning(f"{len(faltando_em_mov)} vendas presentes em Minhas Vendas não estão no Movimento Diário.")
                st.dataframe(faltando_em_mov[["NumeroVenda","doc","valor_ven"]], use_container_width=True)

        with tab_falta_ven:
            if faltando_em_ven.empty:
                st.success("Nenhum registro faltando em Minhas Vendas. ✅")
            else:
                st.warning(f"{len(faltando_em_ven)} vendas presentes no Movimento Diário não estão em Minhas Vendas.")
                st.dataframe(faltando_em_ven[["NumeroVenda","doc_emp","valor_doc","vendedor_codigo","vendedor_nome"]], use_container_width=True)

        with tab_div:
            if divergentes.empty:
                st.success("Todos os valores batem entre as planilhas. ✅")
            else:
                st.error(f"Encontradas {len(divergentes)} divergências de valor.")
                st.dataframe(divergentes[["NumeroVenda","valor_ven","valor_doc","Diferenca","vendedor_codigo","vendedor_nome"]], use_container_width=True)
                st.markdown(f"**Diferença total:** {comp['Diferenca'].sum():,.2f}")

        with tab_vend:
            st.markdown("#### Resumo por Vendedor (Movimento Diário)")
            if "vendedor_nome" in mov_ok.columns:
                vend_group = mov_ok.groupby(["vendedor_codigo","vendedor_nome"], dropna=False).agg(
                    qtd_vendas=("NumeroVenda","nunique"),
                    total_mov=("valor_doc","sum")
                ).reset_index()
                vend_group["ticket_medio_mov"] = vend_group["total_mov"] / vend_group["qtd_vendas"].replace(0, pd.NA)
                st.dataframe(vend_group.sort_values("total_mov", ascending=False), use_container_width=True)

                # Seletor para detalhar por vendedor
                nomes = vend_group["vendedor_nome"].astype(str).tolist()
                sel = st.selectbox("Selecionar vendedor para detalhe:", options=nomes)
                if sel:
                    cod_sel = vend_group[vend_group["vendedor_nome"] == sel]["vendedor_codigo"].astype(str).iloc[0]
                    mov_sel = mov_ok[(mov_ok["vendedor_nome"] == sel) & (mov_ok["vendedor_codigo"].astype(str) == str(cod_sel))]
                    st.markdown(f"**Vendas do vendedor:** {sel} (código {cod_sel})")
                    cols = ["NumeroVenda","doc_emp","valor_doc","cliente_mov","data_mov","serie"]
                    cols = [c for c in cols if c in mov_sel.columns]
                    st.dataframe(mov_sel[cols].sort_values("NumeroVenda"), use_container_width=True)
            else:
                st.info("Não foi possível identificar vendedores no Movimento Diário.")

        with tab_export:
            sheets = {
                "ConferenciaGeral": merged,
                "FaltandoNoMovimento": faltando_em_mov,
                "FaltandoEmMinhasVendas": faltando_em_ven,
                "Divergencias": divergentes
            }
            xbytes = None
            try:
                xbytes = to_excel_bytes(sheets)
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")
            if xbytes:
                st.download_button(
                    label="⬇️ Baixar Excel do Comparativo",
                    data=xbytes,
                    file_name=f"comparativo_dav_vendas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.markdown("<div class='footer-note'>Dica: a aba 'Por Vendedor' usa os cabeçalhos 'Vendedor:' do Movimento Diário e propaga para as vendas subsequentes até o próximo cabeçalho.</div>", unsafe_allow_html=True)

    else:
        st.info("Envie os dois arquivos para iniciar a conferência.")
else:
    st.info("Envie os dois arquivos na barra lateral para iniciar.")
