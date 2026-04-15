import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import difflib
import re
import datetime
import io
import os

from dotenv import load_dotenv
from databricks import sql as databricks_sql
from openpyxl.styles import PatternFill, Font as XLFont
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# =======================================================
# ==== Credenciais via .env ====
# =======================================================
load_dotenv()
HOST      = "dbc-9c96b0a5-aa83.cloud.databricks.com"
HTTP_PATH = "/sql/1.0/warehouses/66eb77f90d9b7746"
TOKEN     = os.getenv("DATABRICKS_TOKEN")

# =======================================================
# ==== Configuração da Página ====
# =======================================================
st.set_page_config(
    page_title="Comparador IMR vs Databricks",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Comparador de Entregas: IMR vs Databricks")
st.markdown("Selecione o período, faça o upload do IMR e gere o relatório de cruzamento (De-Para).")


# =======================================================
# ==== Constantes ====
# =======================================================
colunas_esperadas_db = [
    "numero_pedido", "data_de_corte", "data_prazo_entrega",
    "data_de_entrega", "armazem", "tipo_do_pedido",
    "destinatario", "programa_saude", "cidade_destinatario",
    "modal", "agendamento", "qualidade_de_entrega"
]

date_cols        = ["data_de_corte", "data_prazo_entrega", "data_de_entrega"]
cols_ignorar_hora = ["data_prazo_entrega", "data_de_entrega"]

COR_HEADER_AZUL = (0, 32, 96)
COR_DIVERGENCIA = (255, 199, 206)
COR_OK          = (198, 239, 206)
COR_AVISO       = (255, 235, 156)
COR_CINZA_CLARO = (242, 242, 242)

# =======================================================
# ==== Funções Utilitárias ====
# =======================================================
def limpar_nome_coluna(col):
    if not isinstance(col, str):
        return str(col)
    c = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('utf-8')
    c = re.sub(r'(?i)\bn[oº.]?\b', 'numero', c)
    c = re.sub(r'(?i)N[oº.]', 'numero', c)
    c = re.sub(r'[^a-zA-Z0-9]+', '_', c)
    return c.strip('_').lower()


def parse_dt_robust(series):
    s = pd.Series(series).copy()
    if pd.api.types.is_datetime64_any_dtype(s):
        return s.dt.tz_localize(None)
    s = s.astype(str).str.replace(r'[\r\n]', ' ', regex=True).str.strip()
    s = s.replace(['nan', 'NaT', 'None', ''], np.nan)
    is_num = s.str.isnumeric() | (
        s.str.replace('.', '', 1).str.isnumeric() & ~s.str.contains('[-/:]', na=False)
    )
    nums = pd.to_numeric(s[is_num], errors='coerce')
    s_num = pd.to_datetime(nums, unit='D', origin='1899-12-30', errors='coerce')
    s_txt = pd.to_datetime(s[~is_num], format='mixed', dayfirst=True, errors='coerce')
    return pd.concat([s_num, s_txt]).reindex(s.index)


def normalize_text(series):
    s = series.astype(str).replace(['nan', 'None', '<NA>', 'NaT'], '')
    s = s.str.replace(r'[\r\n]', '', regex=True).str.upper().str.strip()
    return s.apply(
        lambda x: unicodedata.normalize('NFKD', str(x)).encode('ASCII', 'ignore').decode('utf-8') if x else ''
    )


def compare_one_col(joined, col, key):
    nm_imr, nm_db = f"{col}_imr", f"{col}_db"
    val_imr = joined[nm_imr]
    val_db  = joined[nm_db] if nm_db in joined.columns else pd.Series([np.nan] * len(joined))

    if col in date_cols:
        match_final = (val_imr.isna() & val_db.isna()) | (
            val_imr.notna() & val_db.notna() & (val_imr == val_db)
        )
        fmt = "%d/%m/%Y" if col in cols_ignorar_hora else "%d/%m/%Y %H:%M"
        val_imr_str = val_imr.dt.strftime(fmt)
        val_db_str  = val_db.dt.strftime(fmt)
    else:
        t_imr, t_db = normalize_text(val_imr), normalize_text(val_db)
        match_final = (t_imr == t_db)
        val_imr_str = val_imr.astype(str)
        val_db_str  = val_db.astype(str)

    return pd.DataFrame({
        'Nº Pedido': joined[key],
        'coluna': col,
        'valor_imr': val_imr_str,
        'valor_db': val_db_str,
        'match_final': match_final
    })


def format_out(df):
    df_copy = df.copy()
    for col in df_copy.select_dtypes(include=['datetime64', 'datetimetz']).columns:
        fmt = "%d/%m/%Y" if col in cols_ignorar_hora else "%d/%m/%Y %H:%M:%S"
        df_copy[col] = df_copy[col].dt.strftime(fmt)
    return df_copy


def ordenar_por_pedido(df, col_name):
    if df.empty or col_name not in df.columns:
        return df
    df_copy = df.copy()
    df_copy['ordem_tmp'] = pd.to_numeric(df_copy[col_name], errors='coerce')
    df_copy = df_copy.sort_values(['ordem_tmp', col_name], ascending=[True, True])
    return df_copy.drop(columns=['ordem_tmp'])


# =======================================================
# ==== Extração Databricks ====
# =======================================================
@st.cache_data(show_spinner=False)
def extrair_databricks(ano, mes):
    conexao = databricks_sql.connect(
        server_hostname=HOST,
        http_path=HTTP_PATH,
        access_token=TOKEN
    )
    query = f"""
        SELECT *
        FROM vtclog.gold.cte_imr
        WHERE ano_imr_referencia = {ano}
          AND mes_imr_referencia = {mes}
    """
    df = pd.read_sql(query, conexao)
    conexao.close()
    return df


# =======================================================
# ==== Geração do Excel ====
# =======================================================
def gerar_excel(painel_match, stats_final, somente_no_db, somente_no_imr, df_base_out, df_ref_out):
    output = io.BytesIO()
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = XLFont(color="FFFFFF", bold=True)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        painel_match.to_excel(writer,   sheet_name='Painel_Match',        index=False)
        stats_final.to_excel(writer,    sheet_name='Stats_%_por_Coluna',   index=False)
        somente_no_db.to_excel(writer,  sheet_name='Somente_no_DB',        index=False)
        somente_no_imr.to_excel(writer, sheet_name='Somente_no_IMR',       index=False)
        df_base_out.to_excel(writer,    sheet_name='Base_IMR_Limpa',       index=False)
        df_ref_out.to_excel(writer,     sheet_name='Base_DB',              index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_length = 0
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[col[0].column_letter].width = max_length + 2

    output.seek(0)
    return output


# =======================================================
# ==== Helpers Word ====
# =======================================================
def set_cell_bg(cell, rgb_tuple):
    r, g, b = rgb_tuple
    hex_color = f"{r:02X}{g:02X}{b:02X}"
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.color.rgb = RGBColor(*COR_HEADER_AZUL)
    return p



def add_table_word(doc, df, max_rows=200, col_width_twips=None):
    """Adiciona tabela ocupando toda a largura da página."""
    if df.empty:
        doc.add_paragraph("Nenhum dado encontrado.")
        return

    df_show = df.head(max_rows).fillna('').astype(str)
    cols = list(df_show.columns)
    n_cols = len(cols)

    # Largura total útil em twips (página A4 retrato com margens 2.5cm cada lado)
    # 21cm - 5cm margens = 16cm = ~9072 twips
    PAGE_W = 9072
    w = col_width_twips if col_width_twips else PAGE_W // n_cols

    table = doc.add_table(rows=1 + len(df_show), cols=n_cols)
    table.style = 'Table Grid'

    # Força a tabela a ocupar largura total via XML
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(PAGE_W))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    # Cabeçalho
    for i, col_name in enumerate(cols):
        cell = table.rows[0].cells[i]
        cell.text = col_name
        cell.width = w
        set_cell_bg(cell, COR_HEADER_AZUL)
        run = cell.paragraphs[0].runs[0]
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.bold = True
        run.font.size = Pt(8)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Dados
    for row_idx, (_, row) in enumerate(df_show.iterrows()):
        bg = COR_CINZA_CLARO if row_idx % 2 == 0 else (255, 255, 255)
        for col_idx, val in enumerate(row):
            cell = table.rows[row_idx + 1].cells[col_idx]
            val_str = str(val)
            cell.text = val_str
            cell.width = w
            set_cell_bg(cell, bg)
            run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run(val_str)
            run.font.size = Pt(8)

    if len(df) > max_rows:
        doc.add_paragraph(f"⚠ Exibindo {max_rows} de {len(df)} linhas. Consulte o Excel para a tabela completa.")




# =======================================================
# ==== Geração do Word ====
# =======================================================
def gerar_word(stats_final, painel_match, somente_no_db, somente_no_imr,
               qtd_linhas_imr, qtd_linhas_db, total_pedidos_join,
               total_divergencias, pct_ok, comparacao):
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)

    hoje_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    titulo = doc.add_heading('Relatório de Cruzamento De-Para', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in titulo.runs:
        run.font.color.rgb = RGBColor(*COR_HEADER_AZUL)
        run.font.size = Pt(18)
    sub = doc.add_paragraph(f'IMR vs Databricks  |  Gerado em {hoje_str}')
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.color.rgb = RGBColor(89, 89, 89)
    sub.runs[0].font.size = Pt(11)
    doc.add_paragraph()

    add_heading(doc, '1. Resumo Executivo', level=1)
    resumo_data = [
        ("Total de linhas no IMR",       str(qtd_linhas_imr)),
        ("Total de linhas no Databricks", str(qtd_linhas_db)),
        ("Pedidos em comum (Inner Join)", str(total_pedidos_join)),
        ("Pedidos com divergência",       str(total_divergencias)),
        ("Pedidos somente no DB",         str(len(somente_no_db))),
        ("Pedidos somente no IMR",        str(len(somente_no_imr))),
        ("% de Conformidade",             f"{pct_ok:.1f}%"),
    ]
    t = doc.add_table(rows=len(resumo_data), cols=2)
    t.style = 'Table Grid'
    for i, (label, valor) in enumerate(resumo_data):
        bg = COR_CINZA_CLARO if i % 2 == 0 else (255, 255, 255)
        t.rows[i].cells[0].text = label
        set_cell_bg(t.rows[i].cells[0], bg)
        t.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
        t.rows[i].cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        t.rows[i].cells[0].width = Inches(4)
        t.rows[i].cells[1].text = valor
        if label == "% de Conformidade":
            val_cor = COR_OK if pct_ok >= 95 else (COR_AVISO if pct_ok >= 80 else COR_DIVERGENCIA)
            set_cell_bg(t.rows[i].cells[1], val_cor)
        elif label == "Pedidos com divergência" and total_divergencias > 0:
            set_cell_bg(t.rows[i].cells[1], COR_DIVERGENCIA)
        else:
            set_cell_bg(t.rows[i].cells[1], bg)
        t.rows[i].cells[1].paragraphs[0].runs[0].font.size = Pt(10)
        t.rows[i].cells[1].width = Inches(2)

    doc.add_paragraph()
    if pct_ok >= 95:
        msg = f"✅ Conformidade em {pct_ok:.1f}%, dentro do padrão esperado."
    elif pct_ok >= 80:
        msg = f"⚠ Conformidade em {pct_ok:.1f}%. Recomenda-se revisão das divergências."
    else:
        msg = f"❌ Conformidade em {pct_ok:.1f}%. Ação corretiva necessária."
    p = doc.add_paragraph(msg)
    p.runs[0].font.italic = True
    p.runs[0].font.size = Pt(10)
    doc.add_page_break()

    add_heading(doc, '2. Estatísticas por Coluna', level=1)
    doc.add_paragraph('Total de comparações por campo, com percentual de conformidade e divergência.')
    doc.add_paragraph()
    stats_display = stats_final.copy()
    for col in ['total_comparacoes', 'verdadeiro_n', 'falso_n']:
        if col in stats_display.columns:
            stats_display[col] = stats_display[col].apply(
                lambda x: str(int(x)) if pd.notna(x) else ''
            )
    add_table_word(doc, stats_display)
    doc.add_page_break()

    add_heading(doc, '3. Divergências Detalhadas por Campo', level=1)
    doc.add_paragraph('Pedidos onde os valores do IMR e do Databricks não coincidem, agrupados por campo.')
    if not comparacao.empty:
        colunas_div = comparacao[~comparacao['match_final']]['coluna'].unique()
        for col_div in sorted(colunas_div):
            df_col = comparacao[
                (comparacao['coluna'] == col_div) & (~comparacao['match_final'])
            ][['Nº Pedido', 'valor_imr', 'valor_db']].copy()
            df_col.columns = ['Nº Pedido', 'Valor IMR', 'Valor Databricks']
            add_heading(doc, f'Campo: {col_div}  ({len(df_col)} divergência(s))', level=2)
            add_table_word(doc, df_col, max_rows=100)
            doc.add_paragraph()
    else:
        doc.add_paragraph("Nenhuma divergência encontrada.")
    doc.add_page_break()

    add_heading(doc, '4. Painel de Match por Pedido', level=1)
    doc.add_paragraph('Pedidos com ao menos uma divergência e quais campos estão em conformidade.')
    doc.add_paragraph()
    if painel_match.empty:
        doc.add_paragraph("Nenhum pedido com divergência.")
    else:
        pm = painel_match.copy()
        for c in pm.columns:
            if pm[c].dtype == bool:
                pm[c] = pm[c].map({True: '✔', False: '✘'})
        add_table_word(doc, pm, max_rows=150)
    doc.add_page_break()

    add_heading(doc, '5. Pedidos Somente no Databricks', level=1)
    doc.add_paragraph('Pedidos no Databricks sem correspondência no IMR.')
    doc.add_paragraph()
    key = 'numero_pedido'
    db_simples = somente_no_db[[key]].rename(columns={key: 'Nº Pedido'}) if key in somente_no_db.columns else somente_no_db[['Nº Pedido']] if 'Nº Pedido' in somente_no_db.columns else somente_no_db.iloc[:, :1]
    add_table_word(doc, db_simples, max_rows=200)
    doc.add_page_break()

    add_heading(doc, '6. Pedidos Somente no IMR', level=1)
    doc.add_paragraph('Pedidos no IMR sem correspondência no Databricks.')
    doc.add_paragraph()
    imr_simples = somente_no_imr[[key]].rename(columns={key: 'Nº Pedido'}) if key in somente_no_imr.columns else somente_no_imr[['Nº Pedido']] if 'Nº Pedido' in somente_no_imr.columns else somente_no_imr.iloc[:, :1]
    add_table_word(doc, imr_simples, max_rows=200)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =======================================================
# ==== Interface ====
# =======================================================
if not TOKEN:
    st.error("⚠️ Token do Databricks não encontrado. Crie o arquivo `.env` com DATABRICKS_TOKEN=seu_token.")
    st.stop()

st.divider()
st.subheader("1. Selecione o período de referência")

col1, col2 = st.columns(2)
with col1:
    ano_input = st.selectbox("Ano", options=list(range(2025, 2029)), index=0)
with col2:
    mes_input = st.selectbox(
        "Mês",
        options=list(range(1, 13)),
        format_func=lambda m: ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                                "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"][m - 1]
    )

nomes_meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
               "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
nome_mes = nomes_meses[mes_input - 1]
st.success(f"Período selecionado: **{nome_mes} de {ano_input}**")

st.divider()
st.subheader("2. Upload do arquivo IMR")
file_imr = st.file_uploader("**📄 Base IMR** (Ex: IMR_SET.xlsx)", type=["xlsx", "xls"])

st.divider()

# Inicializa session_state
if "resultado" not in st.session_state:
    st.session_state.resultado = None

if file_imr:
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        processar = st.button("🔄 Buscar dados do Databricks e processar", type="primary", use_container_width=True)
    with col_btn2:
        if st.button("🗑️ Resetar", use_container_width=True):
            st.session_state.resultado = None
            st.rerun()

    if processar:
        with st.spinner(f"Conectando ao Databricks ({nome_mes} de {ano_input})..."):
            try:
                df_ref = extrair_databricks(ano_input, mes_input)
                st.success(f"✅ {len(df_ref):,} linhas extraídas do Databricks.")
            except Exception as e:
                st.error(f"Erro ao conectar ao Databricks: {e}")
                st.stop()

        with st.spinner("Processando os arquivos..."):
            df_base = pd.read_excel(file_imr)
            qtd_linhas_imr = len(df_base)
            qtd_linhas_db  = len(df_ref)

            mapa_colunas_imr = {}
            for col_original in df_base.columns:
                col_limpa = limpar_nome_coluna(col_original)
                matches = difflib.get_close_matches(col_limpa, colunas_esperadas_db, n=1, cutoff=0.6)
                if matches:
                    mapa_colunas_imr[col_original] = matches[0]
            df_base = df_base.rename(columns=mapa_colunas_imr)

            key = "numero_pedido"
            if key not in df_base.columns or key not in df_ref.columns:
                st.error(f"Coluna chave '{key}' não encontrada. Verifique os cabeçalhos.")
                st.stop()

            df_base[key] = df_base[key].astype(str).str.strip()
            df_ref[key]  = df_ref[key].astype(str).str.strip()

            for col in date_cols:
                if col in df_base.columns:
                    df_base[col] = parse_dt_robust(df_base[col])
                    df_base[col] = df_base[col].dt.floor('D') if col in cols_ignorar_hora else df_base[col].dt.floor('min')
                if col in df_ref.columns:
                    df_ref[col] = parse_dt_robust(df_ref[col])
                    df_ref[col] = df_ref[col].dt.floor('D') if col in cols_ignorar_hora else df_ref[col].dt.floor('min')

            joined = pd.merge(df_base, df_ref, on=key, how='inner', suffixes=('_imr', '_db'))
            cols_to_compare = [c for c in colunas_esperadas_db if c != key and c in df_base.columns]

            if len(joined) > 0:
                comparacao = pd.concat(
                    [compare_one_col(joined, col, key) for col in cols_to_compare],
                    ignore_index=True
                )
            else:
                comparacao = pd.DataFrame(columns=['Nº Pedido', 'coluna', 'valor_imr', 'valor_db', 'match_final'])

            stats_colunas = comparacao.groupby('coluna').agg(
                total_comparacoes=('match_final', 'count'),
                verdadeiro_n=('match_final', lambda x: x.sum()),
                falso_n=('match_final', lambda x: (~x).sum())
            ).reset_index()
            stats_colunas['pct_v_num'] = (stats_colunas['verdadeiro_n'] / stats_colunas['total_comparacoes']) * 100
            stats_colunas['pct_f_num'] = (stats_colunas['falso_n'] / stats_colunas['total_comparacoes']) * 100
            stats_colunas['pct_verdadeiro'] = stats_colunas['pct_v_num'].apply(lambda x: f"{x:,.2f}%".replace('.', ','))
            stats_colunas['pct_falso']      = stats_colunas['pct_f_num'].apply(lambda x: f"{x:,.2f}%".replace('.', ','))
            stats_colunas = stats_colunas[['coluna', 'total_comparacoes', 'verdadeiro_n', 'falso_n', 'pct_verdadeiro', 'pct_falso']]

            linhas_resumo = pd.DataFrame({
                'coluna': ["Qtd Linhas IMR Total", "Qtd Linhas DB Total"],
                'total_comparacoes': [qtd_linhas_imr, qtd_linhas_db],
                'verdadeiro_n': [np.nan, np.nan], 'falso_n': [np.nan, np.nan],
                'pct_verdadeiro': [np.nan, np.nan], 'pct_falso': [np.nan, np.nan]
            })
            stats_final = pd.concat([stats_colunas, linhas_resumo], ignore_index=True)

            if not comparacao.empty:
                painel_match = comparacao.pivot_table(
                    index='Nº Pedido', columns='coluna', values='match_final',
                    aggfunc='first', fill_value=False
                ).reset_index()
                colunas_teste = [c for c in painel_match.columns if c not in ['Nº Pedido', 'armazem']]
                mascara_divergencias = (painel_match[colunas_teste] == False).any(axis=1)
                painel_match = painel_match[mascara_divergencias]
                ordem_desejada = ['Nº Pedido'] + [c for c in df_base.columns if c in painel_match.columns and c != key]
                painel_match = painel_match[[c for c in ordem_desejada if c in painel_match.columns]]
            else:
                painel_match = pd.DataFrame()

            somente_no_db  = format_out(df_ref[~df_ref[key].isin(df_base[key])])
            somente_no_imr = format_out(df_base[~df_base[key].isin(df_ref[key])])
            painel_match   = ordenar_por_pedido(painel_match, 'Nº Pedido')
            somente_no_db  = ordenar_por_pedido(somente_no_db, key)
            somente_no_imr = ordenar_por_pedido(somente_no_imr, key)
            df_base_out    = ordenar_por_pedido(format_out(df_base), key)
            df_ref_out     = ordenar_por_pedido(format_out(df_ref), key)

            total_pedidos_join = len(joined)
            total_divergencias = len(painel_match) if not painel_match.empty else 0
            pct_ok = ((total_pedidos_join - total_divergencias) / total_pedidos_join * 100) if total_pedidos_join > 0 else 0

            # Salva tudo no session_state
            st.session_state.resultado = {
                "stats_final": stats_final,
                "painel_match": painel_match,
                "somente_no_db": somente_no_db,
                "somente_no_imr": somente_no_imr,
                "df_base_out": df_base_out,
                "df_ref_out": df_ref_out,
                "comparacao": comparacao,
                "mapa_colunas_imr": mapa_colunas_imr,
                "qtd_linhas_imr": qtd_linhas_imr,
                "qtd_linhas_db": qtd_linhas_db,
                "total_pedidos_join": total_pedidos_join,
                "total_divergencias": total_divergencias,
                "pct_ok": pct_ok,
                "ano_input": ano_input,
                "mes_input": mes_input,
            }

# Exibe resultados se existirem no session_state
if st.session_state.resultado:
    r = st.session_state.resultado
    st.success("Processamento concluído!")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Linhas IMR",         f"{r['qtd_linhas_imr']:,}")
    m2.metric("Linhas DB",          f"{r['qtd_linhas_db']:,}")
    m3.metric("Match (Inner Join)", f"{r['total_pedidos_join']:,}")
    m4.metric("Com Divergência",    f"{r['total_divergencias']:,}")
    m5.metric("% Conformidade",     f"{r['pct_ok']:.1f}%")

    st.divider()

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Stats por Coluna",
        "⚠️ Painel Match (Divergências)",
        "🔵 Somente no DB",
        "🟡 Somente no IMR"
    ])
    with tab1:
        st.dataframe(r['stats_final'], use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(r['painel_match'], use_container_width=True, hide_index=True) if not r['painel_match'].empty else st.info("Nenhuma divergência.")
    with tab3:
        st.dataframe(r['somente_no_db'], use_container_width=True, hide_index=True) if not r['somente_no_db'].empty else st.info("Nenhum pedido exclusivo no DB.")
    with tab4:
        st.dataframe(r['somente_no_imr'], use_container_width=True, hide_index=True) if not r['somente_no_imr'].empty else st.info("Nenhum pedido exclusivo no IMR.")

    if r['mapa_colunas_imr']:
        with st.expander("🔍 Colunas identificadas automaticamente (Fuzzy Match)"):
            for original, nova in r['mapa_colunas_imr'].items():
                st.write(f"• `{original}` → `{nova}`")

    st.divider()
    st.subheader("📥 Downloads")
    hoje = datetime.datetime.now().strftime("%Y-%m-%d")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        excel_bytes = gerar_excel(r['painel_match'], r['stats_final'], r['somente_no_db'], r['somente_no_imr'], r['df_base_out'], r['df_ref_out'])
        st.download_button(
            label="⬇️ Baixar relatório Excel",
            data=excel_bytes,
            file_name=f"({hoje}) Depara_entregas_{r['ano_input']}_{r['mes_input']:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_dl2:
        word_bytes = gerar_word(
            r['stats_final'], r['painel_match'], r['somente_no_db'], r['somente_no_imr'],
            r['qtd_linhas_imr'], r['qtd_linhas_db'], r['total_pedidos_join'],
            r['total_divergencias'], r['pct_ok'], r['comparacao']
        )
        st.download_button(
            label="⬇️ Baixar relatório Word",
            data=word_bytes,
            file_name=f"({hoje}) Relatorio_Depara_{r['ano_input']}_{r['mes_input']:02d}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary"
        )
else:
    st.info("⬆️ Faça o upload do arquivo IMR para continuar.")
