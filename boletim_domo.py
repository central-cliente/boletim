import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import textwrap
from fpdf import FPDF
from datetime import datetime

# ==========================================
# CONFIGURAÇÃO DE CAMINHOS
# ==========================================
CAMINHO_EXCEL = r"C:\Users\Sgs14\OneDrive - Energisa\Imagens\Central do Cliente\Boletins\domo de ferro\BASE CENTRAL DOMO DE FERRO - v2.xlsx"
CAMINHO_PDF = r"C:\Users\Sgs14\OneDrive - Energisa\Imagens\Central do Cliente\Boletins\domo de ferro\Boletim_Domo_de_Ferro_Automatico.pdf"
PASTA_TEMP = r"C:\Users\Sgs14\OneDrive - Energisa\Imagens\Central do Cliente\Boletins\domo de ferro\temp"
PASTA_FONTES = r"C:\Users\Sgs14\OneDrive - Energisa\Imagens\Central do Cliente\Boletins\domo de ferro\fontes"
os.makedirs(PASTA_TEMP, exist_ok=True)

# ==========================================
# CORES
# ==========================================
COR_LARANJA = "#F79646"      # barras laranja (padrão Excel)
COR_AZUL = "#4BACC6"         # barras azuis (padrão Excel)
COR_VERDE = "#A9C47F"
COR_CINZA = "#D9D9D9"
COR_FAIXA_LARANJA = "#F7941D"   # faixa lateral e linha do cabeçalho
COR_CAIXA_AZUL = "#0EA5E9"      # caixa azul do cabeçalho
COR_BOM = "#27AE60"
COR_RUIM = "#E74C3C"

# ==========================================
# TEXTOS DAS NOTAS (edite aqui toda semana)
# ==========================================
NOTA_EVO_SEMANAL = "Redução das ações dentro da expectativa devido as necessidades de reprogramação de datas, reflexo da atuação na contingência."
NOTA_EVO_EXPECTATIVA = "Aumento dos casos fora da expectativa, impactado principalmente, pela contingência."
NOTA_REPACTUACAO = "Em andamento contatos de repactuação de datas com os clientes afetados."

# ==========================================
# UTILIDADES
# ==========================================
DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

def data_extenso(d=None):
    d = d or datetime.now()
    return f"{DIAS_SEMANA[d.weekday()]}, {d.day} de {MESES[d.month - 1]} de {d.year}"

def _rgb(hexcor):
    h = hexcor.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

ABREV_ORIGEM = {"Assembleia Legislativa": "ALEMS"}   # nomes longos da Origem, como no Excel

def _rot_origem(txt):
    return ABREV_ORIGEM.get(str(txt).strip(), str(txt))

def _quebra(txt, largura=11):
    """Quebra o texto em linhas com <br> (rótulos do Plotly)."""
    return "<br>".join(textwrap.wrap(str(txt), largura)) or str(txt)

def _contagem(df, coluna, top, ascending, upper=False):
    serie = df[coluna].dropna()
    if upper:
        serie = serie.astype(str).str.strip().str.upper()
    return serie.value_counts().head(top).sort_values(ascending=ascending)

def variacao_pct(atual, anterior):
    return (atual - anterior) / anterior * 100 if anterior else 0.0

def evo_texto(v, bom_subir, unicode_ok=True):
    """Devolve (texto, cor_hex). Verde = melhorou, vermelho = piorou."""
    if round(v) == 0:
        return "—", "#7F7F7F"
    ok = (v > 0) == bom_subir
    cor = COR_BOM if ok else COR_RUIM
    if unicode_ok:
        return f"{'▲' if v > 0 else '▼'}{v:.0f}%", cor
    return f"{v:+.0f}%", cor

def evo_html(v, bom_subir):
    t, c = evo_texto(v, bom_subir)
    return f'<span style="color:{c};">{t}</span>'

# ==========================================
# LEITURA DOS DADOS
# ==========================================
@st.cache_data(ttl=5)
def carregar_resumo():
    try:
        df = pd.read_excel(CAMINHO_EXCEL, sheet_name='RESUMO', header=None)
        return {
            "total_demandas": int(df.iloc[4, 4]) if pd.notna(df.iloc[4, 4]) else 0,
            "backlog": int(df.iloc[4, 3]) if pd.notna(df.iloc[4, 3]) else 0,
            "em_andamento": int(df.iloc[4, 1]) if pd.notna(df.iloc[4, 1]) else 0,
            "concluida": int(df.iloc[4, 2]) if pd.notna(df.iloc[4, 2]) else 0,
            "dentro_expectativa_geral": int(df.iloc[12, 1]) if pd.notna(df.iloc[12, 1]) else 0,
            "total_tratativas": int(df.iloc[7, 10]) if pd.notna(df.iloc[7, 10]) else 0,
            "dentro_expectativa_tratativas": int(df.iloc[8, 7]) if pd.notna(df.iloc[8, 7]) else 0,
            "fora_expectativa_tratativas": int(df.iloc[8, 8]) if pd.notna(df.iloc[8, 8]) else 0,
            "risco_prazo": int(df.iloc[8, 9]) if pd.notna(df.iloc[8, 9]) else 0,
            "s1_em_andamento": int(df.iloc[134, 3]) if pd.notna(df.iloc[134, 3]) else 18,
            "s1_concluida": int(df.iloc[135, 3]) if pd.notna(df.iloc[135, 3]) else 135,
            "s1_backlog": int(df.iloc[136, 3]) if pd.notna(df.iloc[136, 3]) else 39,
            "s1_dentro_expectativa": float(df.iloc[137, 3]) if pd.notna(df.iloc[137, 3]) else 0.86,
            "s1_fora_expectativa": int(df.iloc[181, 1]) if pd.notna(df.iloc[181, 1]) else 6,
            "s1_dentro_exp": int(df.iloc[182, 1]) if pd.notna(df.iloc[182, 1]) else 36,
            "s1_risco": int(df.iloc[183, 1]) if pd.notna(df.iloc[183, 1]) else 9,
            "s1_total_aberto": int(df.iloc[184, 1]) if pd.notna(df.iloc[184, 1]) else 42,
        }
    except Exception as e:
        return {"erro": str(e)}

@st.cache_data(ttl=5)
def carregar_analitico():
    try:
        df = pd.read_excel(CAMINHO_EXCEL, sheet_name='Analitico')
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except:
        return pd.DataFrame()

# ==========================================
# DADOS DERIVADOS
# ==========================================
def calcular_aging(df):
    """Faixas de aging das demandas em aberto, na ordem natural (0-5 ... >50)."""
    if df.empty or 'SLA' not in df.columns or 'Status_Acao' not in df.columns:
        return None
    df_ab = df[df['Status_Acao'] != 'Finalizada'].copy()
    bins = [-1, 5, 15, 30, 50, 99999]   # -1 para incluir SLA = 0
    labels = ['0-5 dias', '6-15 dias', '16-30 dias', '30-50 dias', '>50 dias']
    df_ab['Faixa'] = pd.cut(df_ab['SLA'], bins=bins, labels=labels)
    contagem = df_ab['Faixa'].value_counts().reindex(labels).fillna(0).astype(int)
    total = contagem.sum()
    pct = (contagem / total * 100).round(0).astype(int) if total else contagem
    return pd.DataFrame({'Faixa': labels, 'Qtd': contagem.values, 'Pct': pct.values})

@st.cache_data(ttl=5)
def carregar_riscos():
    """Lê a tabela 'Riscos e Fora da Expectativa' da aba RESUMO (colunas H:K, até 'Total Geral').
    Devolve (tabela com Area/Fora/Risco/Total, [total_fora, total_risco, total_geral])."""
    vazio = pd.DataFrame(columns=['Area', 'Fora', 'Risco', 'Total'])
    try:
        df = pd.read_excel(CAMINHO_EXCEL, sheet_name='RESUMO', header=None)
        nomes = df.iloc[:, 7].astype(str).str.strip().str.lower()
        cab_i = df.iloc[:, 8].astype(str).str.lower()
        achou = df.index[nomes.isin(['área', 'area']) & cab_i.str.contains('fora')]
        if len(achou) == 0:
            return vazio, [0, 0, 0]
        linhas, total = [], None
        for i in range(achou[0] + 1, len(df)):
            nome = df.iloc[i, 7]
            if pd.isna(nome):
                break
            vals = [int(df.iloc[i, c]) if pd.notna(df.iloc[i, c]) else 0 for c in (8, 9, 10)]
            if str(nome).strip().lower() == 'total geral':
                total = vals
                break
            linhas.append([str(nome).strip()] + vals)
        tabela = pd.DataFrame(linhas, columns=['Area', 'Fora', 'Risco', 'Total'])
        if total is None:
            total = [int(tabela['Fora'].sum()), int(tabela['Risco'].sum()), int(tabela['Total'].sum())]
        return tabela, total
    except Exception:
        return vazio, [0, 0, 0]

# ==========================================
# GRÁFICOS PLOTLY (TELA)
# ==========================================
def grafico_barras_h(df, coluna, titulo, cor, top=5, upper=False, quebra=None, altura=250):
    if df.empty or coluna not in df.columns: return None
    contagem = _contagem(df, coluna, top, ascending=True, upper=upper)
    if contagem.empty: return None
    max_val = contagem.max()
    rotulos = [_quebra(i, quebra) if quebra else str(i) for i in contagem.index]
    fig = go.Figure(go.Bar(
        x=contagem.values, y=rotulos, orientation='h',
        marker_color=cor, text=contagem.values, textposition='outside',
        cliponaxis=False,
        textfont=dict(size=13, family="Arial Black", color="#1a1a1a")))
    fig.update_layout(
        title=dict(text=f"<b>{titulo}</b>", font=dict(size=11, color="#1a1a1a"), x=0.5, xanchor='center'),
        showlegend=False, height=altura, bargap=0.12,
        margin=dict(l=10, r=40, t=35, b=5),
        plot_bgcolor='white', paper_bgcolor='white',
        xaxis=dict(visible=False, range=[0, max_val * 1.2]),
        yaxis=dict(title=None, ticks='', showline=False, automargin=True,
                   tickfont=dict(size=10, color="#1a1a1a")))
    return fig

def grafico_barras_v(df, coluna, titulo, cor, top=5, upper=False):
    if df.empty or coluna not in df.columns: return None
    contagem = _contagem(df, coluna, top, ascending=False, upper=upper)
    if contagem.empty: return None
    max_val = contagem.max()
    fig = go.Figure(go.Bar(
        x=list(range(len(contagem))), y=contagem.values,
        marker_color=cor, text=contagem.values, textposition='outside',
        cliponaxis=False,
        textfont=dict(size=13, family="Arial Black", color="#1a1a1a")))
    fig.update_layout(
        title=dict(text=f"<b>{titulo}</b>", font=dict(size=11, color="#1a1a1a"), x=0.5, xanchor='center'),
        showlegend=False, height=250, bargap=0.12,
        margin=dict(l=5, r=5, t=35, b=5),
        plot_bgcolor='white', paper_bgcolor='white',
        xaxis=dict(tickmode='array', tickvals=list(range(len(contagem))),
                   ticktext=[_quebra(_rot_origem(i), 11) for i in contagem.index],
                   tickangle=0,   # força os nomes na horizontal (sem inclinar)
                   tickfont=dict(size=8, color="#1a1a1a"), ticks='', showline=False),
        yaxis=dict(visible=False, range=[0, max_val * 1.2]))
    return fig

def grafico_aging_plotly(df):
    df_plot = calcular_aging(df)
    if df_plot is None: return None
    max_val = max(df_plot['Qtd'].max(), 1)
    fig = go.Figure(go.Bar(
        x=df_plot['Qtd'], y=df_plot['Faixa'], orientation='h',
        marker_color=COR_AZUL, text=[f"{p}%" for p in df_plot['Pct']],
        textposition='outside', cliponaxis=False,
        textfont=dict(size=12, family="Arial Black", color="#1a1a1a")))
    fig.update_layout(
        showlegend=False, height=230, bargap=0.25,
        margin=dict(l=10, r=40, t=5, b=5),
        plot_bgcolor='white', paper_bgcolor='white',
        xaxis=dict(visible=False, range=[0, max_val * 1.2]),
        yaxis=dict(title=None, ticks='', showline=False, automargin=True,
                   autorange='reversed', tickfont=dict(size=11, color="#1a1a1a")))
    return fig

# ==========================================
# GRÁFICOS MATPLOTLIB (PDF)
# ==========================================
FIG_DETALHAMENTO = (3.0, 2.3)
FIG_IMPACTO = (3.2, 1.9)
FIG_AGING = (4.0, 1.6)

def _salvar(fig, nome):
    caminho = os.path.join(PASTA_TEMP, f"{nome}.png")
    fig.savefig(caminho, dpi=200, facecolor='white')   # sem bbox 'tight': mantém a proporção do figsize
    plt.close(fig)
    return caminho

def grafico_matplotlib_h(df, coluna, titulo, cor, nome, top=5, upper=False, quebra=None, figsize=FIG_DETALHAMENTO):
    if df.empty or coluna not in df.columns: return None
    contagem = _contagem(df, coluna, top, ascending=True, upper=upper)
    if contagem.empty: return None
    rotulos = [textwrap.fill(str(i), quebra) if quebra else str(i) for i in contagem.index]
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    bars = ax.barh(rotulos, contagem.values, color=cor, height=0.85)
    ax.set_title(titulo, fontsize=8, fontweight='bold', color='#1a1a1a', pad=8)
    max_val = max(contagem.values)
    for bar, val in zip(bars, contagem.values):
        ax.text(bar.get_width() + max_val * 0.03, bar.get_y() + bar.get_height() / 2,
                f'{int(val)}', va='center', ha='left', fontsize=10, fontweight='bold', color='#1a1a1a')
    for spine in ax.spines.values(): spine.set_visible(False)
    ax.tick_params(axis='y', labelsize=7, length=0)
    ax.tick_params(axis='x', bottom=False, labelbottom=False)
    ax.set_xlim(0, max_val * 1.22)
    plt.tight_layout(pad=0.4)
    return _salvar(fig, nome)

def grafico_matplotlib_v(df, coluna, titulo, cor, nome, top=5, upper=False, figsize=FIG_DETALHAMENTO):
    if df.empty or coluna not in df.columns: return None
    contagem = _contagem(df, coluna, top, ascending=False, upper=upper)
    if contagem.empty: return None
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    bars = ax.bar(range(len(contagem)), contagem.values, color=cor, width=0.85)
    ax.set_title(titulo, fontsize=8, fontweight='bold', color='#1a1a1a', pad=8)
    max_val = max(contagem.values)
    for bar, val in zip(bars, contagem.values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.03,
                f'{int(val)}', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#1a1a1a')
    ax.set_xticks(range(len(contagem)))
    ax.set_xticklabels([textwrap.fill(_rot_origem(i), 10) for i in contagem.index], fontsize=6.5, ha='center')
    for spine in ax.spines.values(): spine.set_visible(False)
    ax.tick_params(axis='x', length=0)
    ax.tick_params(axis='y', left=False, labelleft=False)
    ax.set_ylim(0, max_val * 1.25)
    plt.tight_layout(pad=0.4)
    return _salvar(fig, nome)

def aging_matplotlib(df):
    df_plot = calcular_aging(df)
    if df_plot is None: return None
    fig, ax = plt.subplots(figsize=FIG_AGING, dpi=200)
    faixas = list(df_plot['Faixa'])[::-1]      # barh desenha de baixo p/ cima: inverte para 0-5 ficar no topo
    qtds = list(df_plot['Qtd'])[::-1]
    pcts = list(df_plot['Pct'])[::-1]
    bars = ax.barh(faixas, qtds, color=COR_AZUL, height=0.6)
    max_val = max(max(qtds), 1)
    for bar, p in zip(bars, pcts):
        ax.text(bar.get_width() + max_val * 0.02, bar.get_y() + bar.get_height() / 2,
                f'{p}%', va='center', fontsize=8, fontweight='bold', color='#1a1a1a')
    for spine in ax.spines.values(): spine.set_visible(False)
    ax.tick_params(axis='y', labelsize=7, length=0)
    ax.tick_params(axis='x', bottom=False, labelbottom=False)
    ax.set_xlim(0, max_val * 1.15)
    plt.tight_layout(pad=0.4)
    return _salvar(fig, "aging")

# ==========================================
# ALERTAS
# ==========================================
def gerar_alertas(resumo, df_analitico):
    alertas = {}
    total_aberto = resumo['total_tratativas']
    fora = resumo['fora_expectativa_tratativas']
    pct_fora = (fora / total_aberto * 100) if total_aberto else 0
    if pct_fora > 15:
        alertas['ATENCAO'] = f"{pct_fora:.0f}% das demandas em tratativas estão FORA da expectativa do cliente ({fora} de {total_aberto}). RISCO DE ESCALONAMENTO INSTITUCIONAL."
    else:
        alertas['ATENCAO'] = f"Apenas {pct_fora:.0f}% das demandas estão fora da expectativa. Situação sob controle."
    if not df_analitico.empty and 'Area 2' in df_analitico.columns:
        em_and = df_analitico[df_analitico['Status_Acao'] == 'Em Andamento']
        if not em_and.empty:
            top_area = em_and['Area 2'].value_counts().idxmax()
            qtd_area = em_and['Area 2'].value_counts().max()
            pct_area = (qtd_area / len(em_and) * 100)
            alertas['CONCENTRACAO'] = f"{top_area} concentra {pct_area:.0f}% das demandas em andamento ({qtd_area} de {len(em_and)}), com evolução frente à semana anterior."
        else:
            alertas['CONCENTRACAO'] = "Sem demandas em andamento no momento."
    else:
        alertas['CONCENTRACAO'] = "Dados não disponíveis."
    alertas['EVOLUCAO'] = "Demandas concluídas com aging >50 dias apresentou avanço frente à semana anterior."
    return alertas

# ==========================================
# COMPONENTES HTML (TELA)
# ==========================================
SOMBRA = "box-shadow: 0 3px 6px rgba(0,0,0,0.15), 0 1px 3px rgba(0,0,0,0.10);"

def card_verde(titulo, valor, subtitulo=""):
    return f'''<div style="flex:1; background-color:{COR_VERDE}; border-radius:12px; padding:10px 6px; text-align:center; height:120px; display:flex; flex-direction:column; justify-content:center; {SOMBRA}">
        <div style="font-size:12px; font-weight:bold; color:#1a1a1a; line-height:1.2;">{titulo}</div>
        <div style="font-size:34px; font-weight:bold; color:#1a1a1a; line-height:1.1; margin:4px 0;">{valor}</div>
        <div style="font-size:12px; font-weight:bold; color:#1a1a1a;">{subtitulo}</div>
    </div>'''

def card_cinza(titulo, valor, subtitulo=""):
    return f'''<div style="flex:1; background-color:{COR_CINZA}; border-radius:12px; padding:10px 6px; text-align:center; height:95px; display:flex; flex-direction:column; justify-content:center; {SOMBRA}">
        <div style="font-size:11px; font-weight:bold; color:#1a1a1a; line-height:1.2;">{titulo}</div>
        <div style="font-size:26px; font-weight:bold; color:#1a1a1a; line-height:1.1; margin:4px 0;">{valor}</div>
        <div style="font-size:10px; font-weight:bold; color:#1a1a1a;">{subtitulo}</div>
    </div>'''

def termometro_geral_html(total, backlog, pct_back, andamento, pct_and, concluida, pct_concl, dentro, pct_dent):
    return ('<div style="display:flex; gap:8px; align-items:center; padding:4px;">'
            + card_verde("Total de<br>Demandas", total)
            + card_cinza("Backlog<br>(Não iniciada)", backlog, f"{pct_back:.0f}%")
            + card_cinza("Em<br>andamento", andamento, f"{pct_and:.0f}%")
            + card_cinza("Concluída", concluida, f"{pct_concl:.0f}%")
            + card_verde("Dentro da<br>Expectativa", dentro, f"{pct_dent:.0f}%")
            + '</div>')

def card_tratativas_lateral(total, dentro, pct_de, fora, pct_fe, risco, pct_rp):
    return f'''
    <div style="display:flex; gap:8px; min-height:120px; padding:4px;">
        <div style="background-color:{COR_CINZA}; border-radius:12px; padding:10px; text-align:center; width:24%; display:flex; flex-direction:column; justify-content:center; {SOMBRA}">
            <div style="font-size:11px; font-weight:bold; color:#1a1a1a; line-height:1.2;">Total de<br>Demandas</div>
            <div style="font-size:30px; font-weight:bold; color:#1a1a1a; margin-top:5px;">{total}</div>
        </div>
        <div style="width:76%; display:flex; flex-direction:column; gap:5px;">
            <div style="background-color:#D5F5D5; border-radius:8px; padding:6px 10px; display:flex; justify-content:space-between; align-items:center; flex:1; {SOMBRA}">
                <span style="font-size:11px; font-weight:bold; color:#1E8449;">● Dentro da Expectativa:</span>
                <span style="font-size:20px; font-weight:bold; color:#1a1a1a;">{dentro}</span>
                <span style="font-size:12px; font-weight:bold; color:#1E8449;">{pct_de:.0f}%</span>
            </div>
            <div style="background-color:#FADBD8; border-radius:8px; padding:6px 10px; display:flex; justify-content:space-between; align-items:center; flex:1; {SOMBRA}">
                <span style="font-size:11px; font-weight:bold; color:#C0392B;">● Fora da Expectativa:</span>
                <span style="font-size:20px; font-weight:bold; color:#C0392B;">{fora}</span>
                <span style="font-size:12px; font-weight:bold; color:#C0392B;">{pct_fe:.0f}%</span>
            </div>
            <div style="background-color:#FCF3CF; border-radius:8px; padding:6px 10px; display:flex; justify-content:space-between; align-items:center; flex:1; {SOMBRA}">
                <span style="font-size:11px; font-weight:bold; color:#B7950B;">● Risco de prazo:</span>
                <span style="font-size:20px; font-weight:bold; color:#1a1a1a;">{risco}</span>
                <span style="font-size:12px; font-weight:bold; color:#B7950B;">{pct_rp:.0f}%</span>
            </div>
        </div>
    </div>'''

def titulo_centralizado(txt, tam=13):
    return f'<div style="text-align:center; font-weight:bold; font-size:{tam}px; color:#1a1a1a; margin-bottom:4px;">{txt}</div>'

def titulo_esquerda(txt, tam=13):
    return f'<div style="font-weight:bold; font-size:{tam}px; color:#1a1a1a; margin:8px 0 4px 0;">{txt}</div>'

def tabela_html(cabecalho, linhas, classe="table-evo"):
    th = "".join(f"<th>{c}</th>" for c in cabecalho)
    trs = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in l) + "</tr>" for l in linhas)
    return f'<table class="{classe}"><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'

def nota_html(*paragrafos):
    corpo = "<br><br>".join(paragrafos)
    return f'<div style="border:1px dashed #888; padding:8px; font-size:11px; font-weight:bold; color:#333; border-radius:4px;"><b>Nota:</b> {corpo}</div>'

def riscos_html(tbl, total):
    linhas = ""
    for _, r in tbl.iterrows():
        linhas += f'<tr><td style="text-align:left;">{r["Area"]}</td><td>{r["Fora"]}</td><td>{r["Risco"]}</td><td>{r["Total"]}</td></tr>'
    tf, tr_, tt_ = total
    linhas += f'<tr class="total"><td style="text-align:left;">Total Geral</td><td>{tf}</td><td>{tr_}</td><td>{tt_}</td></tr>'
    return f'''<div style="border:1.5px solid {COR_FAIXA_LARANJA}; border-radius:10px; background:#FFF8EC; padding:8px 12px;">
        <div style="font-weight:bold; font-size:13px; color:#B9770E; margin-bottom:6px;">⚠️ Riscos e Fora da Expectativa</div>
        <table class="table-risco"><thead><tr><th style="text-align:left;">Área</th><th>Fora da Expectativa</th><th>Riscos</th><th>Total</th></tr></thead>
        <tbody>{linhas}</tbody></table></div>'''

def alerta_html(titulo, texto, bg, cor_titulo, cor_texto):
    return f'''<div style="background-color:{bg}; padding:8px 10px; border-radius:6px; min-height:100px; {SOMBRA}">
        <div style="font-weight:bold; color:{cor_titulo}; font-size:12px; margin-bottom:3px;">{titulo}</div>
        <div style="font-size:11px; color:{cor_texto}; font-weight:bold;">{texto}</div></div>'''

# ==========================================
# INTERFACE
# ==========================================
st.set_page_config(page_title="Boletim Domo de Ferro", layout="wide", page_icon="📊")
st.markdown("""<style>
.block-container {padding-top: 0.5rem; padding-bottom: 1rem; max-width: 1300px;}
h1, h2, h3 {margin: 0;}
.table-evo, .table-risco { width: 100%; border-collapse: collapse; font-size: 12px; }
.table-evo th, .table-risco th { padding: 5px; border-bottom: 2px solid #444; text-align: center; font-weight: bold; background: transparent; }
.table-evo td, .table-risco td { padding: 5px; border-bottom: 1px solid #DDD; text-align: center; font-weight: bold; }
.table-risco tr.total td { border-top: 2px solid #444; border-bottom: none; }
</style>""", unsafe_allow_html=True)

st.markdown(f'''
<div style="position:fixed; left:0; top:0; bottom:0; width:14px; background:{COR_FAIXA_LARANJA}; z-index:99;"></div>
<div style="border-top:4px solid #404040;">
  <div style="display:flex; align-items:stretch; margin-bottom:6px;">
    <div style="background:{COR_CAIXA_AZUL}; width:70px; min-height:48px;"></div>
    <div style="padding-left:14px; padding-top:4px;">
      <div style="font-size:26px; font-weight:bold; line-height:1.1;">Boletim Semanal</div>
      <div style="font-size:12px; color:#444;">{data_extenso()}</div>
    </div>
  </div>
  <div style="text-align:center; font-size:18px; letter-spacing:8px; padding-bottom:6px; border-bottom:3px solid {COR_FAIXA_LARANJA}; width:70%; margin:0 auto;">CENTRAL DO CLIENTE</div>
  <div style="font-size:46px; font-weight:bold; text-align:center; line-height:1.15; margin-bottom:8px;">Domo de Ferro</div>
</div>''', unsafe_allow_html=True)

resumo = carregar_resumo()
df_analitico = carregar_analitico()

if "erro" in resumo:
    st.error(f"Erro ao ler o Excel: {resumo['erro']}")
    st.stop()

# Demandas em aberto (tratativas) = tudo que não está Finalizado
if not df_analitico.empty and 'Status_Acao' in df_analitico.columns:
    df_aberto = df_analitico[df_analitico['Status_Acao'] != 'Finalizada'].copy()
else:
    df_aberto = pd.DataFrame()

total = resumo['total_demandas']
pct_back = (resumo['backlog'] / total * 100) if total else 0
pct_and = (resumo['em_andamento'] / total * 100) if total else 0
pct_concl = (resumo['concluida'] / total * 100) if total else 0
pct_dent = (resumo['dentro_expectativa_geral'] / total * 100) if total else 0
tt = resumo['total_tratativas']
pct_de = (resumo['dentro_expectativa_tratativas'] / tt * 100) if tt else 0
pct_fe = (resumo['fora_expectativa_tratativas'] / tt * 100) if tt else 0
pct_rp = (resumo['risco_prazo'] / tt * 100) if tt else 0

# --- Evolução semanal: (rótulo, S-1, S-Atual, variação %, "subir é bom?") ---
evo_dent = (resumo['dentro_expectativa_geral'] / total - resumo['s1_dentro_expectativa']) * 100 if total else 0
evo_semanal = [
    ("Em andamento", resumo['s1_em_andamento'], resumo['em_andamento'],
     variacao_pct(resumo['em_andamento'], resumo['s1_em_andamento']), False),
    ("Concluídas", resumo['s1_concluida'], resumo['concluida'],
     variacao_pct(resumo['concluida'], resumo['s1_concluida']), True),
    ("Backlog", resumo['s1_backlog'], resumo['backlog'],
     variacao_pct(resumo['backlog'], resumo['s1_backlog']), False),
    ("Dentro da expectativa", f"{resumo['s1_dentro_expectativa'] * 100:.0f}%", f"{pct_dent:.0f}%", evo_dent, True),
]

# --- Evolução expectativa ---
evo_expectativa = [
    ("Fora da Expectativa (atrasada)", resumo['s1_fora_expectativa'], resumo['fora_expectativa_tratativas'],
     variacao_pct(resumo['fora_expectativa_tratativas'], resumo['s1_fora_expectativa']), False),
    ("Dentro da Expectativa", resumo['s1_dentro_exp'], resumo['dentro_expectativa_tratativas'],
     variacao_pct(resumo['dentro_expectativa_tratativas'], resumo['s1_dentro_exp']), True),
    ("Risco de Prazo", resumo['s1_risco'], resumo['risco_prazo'],
     variacao_pct(resumo['risco_prazo'], resumo['s1_risco']), False),
    ("Total de demandas em aberto", resumo['s1_total_aberto'], resumo['total_tratativas'],
     variacao_pct(resumo['total_tratativas'], resumo['s1_total_aberto']), False),
]

tbl_riscos, tot_riscos = carregar_riscos()
alertas = gerar_alertas(resumo, df_analitico)

# --- TERMÔMETROS ---
col_geral, col_trat = st.columns([1.35, 1])

with col_geral:
    with st.container(border=True):
        st.markdown(titulo_centralizado("TERMÔMETRO GERAL DAS AÇÕES", 12), unsafe_allow_html=True)
        st.markdown(termometro_geral_html(total, resumo['backlog'], pct_back, resumo['em_andamento'], pct_and,
                                          resumo['concluida'], pct_concl, resumo['dentro_expectativa_geral'], pct_dent),
                    unsafe_allow_html=True)

with col_trat:
    with st.container(border=True):
        st.markdown(titulo_centralizado("TERMÔMETRO AÇÕES EM TRATATIVAS", 12), unsafe_allow_html=True)
        st.markdown(card_tratativas_lateral(tt, resumo['dentro_expectativa_tratativas'], pct_de,
                                            resumo['fora_expectativa_tratativas'], pct_fe,
                                            resumo['risco_prazo'], pct_rp), unsafe_allow_html=True)

# --- DETALHAMENTO DAS DEMANDAS ---
if not df_analitico.empty:
    with st.container(border=True):
        st.markdown(titulo_centralizado("DETALHAMENTO DAS DEMANDAS", 16), unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            fig = grafico_barras_h(df_analitico, 'Tipologia', 'TIPOLOGIA', COR_LARANJA, top=6)
            if fig: st.plotly_chart(fig, use_container_width=True, key="t1")
        with c2:
            fig = grafico_barras_v(df_analitico, 'Poder', 'ORIGEM', COR_AZUL, top=5)
            if fig: st.plotly_chart(fig, use_container_width=True, key="o1")
        with c3:
            fig = grafico_barras_h(df_analitico, 'Localidade', 'TOP 5 MUNICÍPIOS', COR_LARANJA, top=5)
            if fig: st.plotly_chart(fig, use_container_width=True, key="m1")
        with c4:
            fig = grafico_barras_h(df_analitico, 'Supervisão', 'TOP 5 SUPERVISÃO', COR_AZUL, top=5, upper=True)
            if fig: st.plotly_chart(fig, use_container_width=True, key="s1")

# --- EVOLUÇÃO SEMANAL ---
with st.container(border=True):
    c_tit, c_tab, c_nota = st.columns([1, 2.6, 1.8])
    with c_tit:
        st.markdown(titulo_esquerda("EVOLUÇÃO SEMANAL", 13), unsafe_allow_html=True)
    with c_tab:
        linhas = [[n, s1, at, evo_html(v, b)] for n, s1, at, v, b in evo_semanal]
        st.markdown(tabela_html(["Status", "S-1", "S-Atual", "Evolução"], linhas), unsafe_allow_html=True)
    with c_nota:
        st.markdown(nota_html(NOTA_EVO_SEMANAL, NOTA_REPACTUACAO), unsafe_allow_html=True)

# --- MAIORES IMPACTOS - DEMANDAS EM TRATATIVAS ---
with st.container(border=True):
    st.markdown(titulo_esquerda("MAIORES IMPACTO – DEMANDAS EM TRATATIVAS", 13), unsafe_allow_html=True)
    c0, c1, c2, c3 = st.columns([0.7, 2, 2, 2])
    with c0:
        st.markdown(f'''<div style="background-color:{COR_CINZA}; border-radius:10px; padding:12px 6px; text-align:center; margin-top:30px; {SOMBRA}">
            <div style="font-size:11px; font-weight:bold;">Total de Demandas</div>
            <div style="font-size:28px; font-weight:bold;">{tt}</div></div>''', unsafe_allow_html=True)
    with c1:
        fig = grafico_barras_h(df_aberto, 'Area 2', 'TOP 5 ÁREAS', COR_AZUL, top=5, quebra=24)
        if fig: st.plotly_chart(fig, use_container_width=True, key="ia")
    with c2:
        fig = grafico_barras_h(df_aberto, 'Tipologia', 'TOP 5 TIPOS DE DEMANDA', COR_LARANJA, top=5)
        if fig: st.plotly_chart(fig, use_container_width=True, key="it")
    with c3:
        fig = grafico_barras_h(df_aberto, 'Supervisão', 'TOP 5 SUPERVISÃO', COR_AZUL, top=5, upper=True)
        if fig: st.plotly_chart(fig, use_container_width=True, key="is")

# --- AGING + RISCOS ---
col_aging, col_riscos = st.columns([1, 1.2])

with col_aging:
    st.markdown('<div style="font-weight:bold; font-style:italic; font-size:14px; margin:8px 0 0 0;">AGING dos dias em aberto</div>', unsafe_allow_html=True)
    fig = grafico_aging_plotly(df_analitico)
    if fig: st.plotly_chart(fig, use_container_width=True, key="aging")

with col_riscos:
    st.markdown(riscos_html(tbl_riscos, tot_riscos), unsafe_allow_html=True)

# --- EVOLUÇÃO EXPECTATIVA ---
with st.container(border=True):
    c_tit, c_tab, c_nota = st.columns([1, 2.6, 1.8])
    with c_tit:
        st.markdown(titulo_esquerda("EVOLUÇÃO EXPECTATIVA", 13), unsafe_allow_html=True)
    with c_tab:
        linhas = [[n, s1, at, evo_html(v, b)] for n, s1, at, v, b in evo_expectativa]
        st.markdown(tabela_html(["Status", "S-1", "S-Atual", "Evolução"], linhas), unsafe_allow_html=True)
    with c_nota:
        st.markdown(nota_html(NOTA_EVO_EXPECTATIVA, NOTA_REPACTUACAO), unsafe_allow_html=True)

# --- ALERTAS EXECUTIVOS ---
with st.container(border=True):
    st.markdown(titulo_esquerda("ALERTAS EXECUTIVOS", 13), unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)
    with a1:
        st.markdown(alerta_html("ATENÇÃO", alertas["ATENCAO"], "#FADBD8", "#C0392B", "#7B241C"), unsafe_allow_html=True)
    with a2:
        st.markdown(alerta_html("CONCENTRAÇÃO", alertas["CONCENTRACAO"], "#FCF3CF", "#B9770E", "#7D6608"), unsafe_allow_html=True)
    with a3:
        st.markdown(alerta_html("EVOLUÇÃO", alertas["EVOLUCAO"], "#D5F5D5", "#1E8449", "#145A32"), unsafe_allow_html=True)

# ==========================================
# BOTÃO EXPORTAR PDF
# ==========================================
st.markdown("---")
if st.button("📄 Exportar para PDF", use_container_width=True, type="primary"):
    with st.spinner("Gerando PDF..."):
        try:
            # ---------- imagens dos gráficos ----------
            imgs = {}
            if not df_analitico.empty:
                imgs['tipologia'] = grafico_matplotlib_h(df_analitico, 'Tipologia', 'TIPOLOGIA', COR_LARANJA, 'tipologia', 6)
                imgs['origem'] = grafico_matplotlib_v(df_analitico, 'Poder', 'ORIGEM', COR_AZUL, 'origem', 5)
                imgs['municipios'] = grafico_matplotlib_h(df_analitico, 'Localidade', 'TOP 5 MUNICÍPIOS', COR_LARANJA, 'municipios', 5)
                imgs['supervisao'] = grafico_matplotlib_h(df_analitico, 'Supervisão', 'TOP 5 SUPERVISÃO', COR_AZUL, 'supervisao', 5, upper=True)
                imgs['aging'] = aging_matplotlib(df_analitico)
            if not df_aberto.empty:
                imgs['imp_areas'] = grafico_matplotlib_h(df_aberto, 'Area 2', 'TOP 5 ÁREAS', COR_AZUL, 'imp_areas', 5, quebra=24, figsize=FIG_IMPACTO)
                imgs['imp_tipos'] = grafico_matplotlib_h(df_aberto, 'Tipologia', 'TOP 5 TIPOS DE DEMANDA', COR_LARANJA, 'imp_tipos', 5, figsize=FIG_IMPACTO)
                imgs['imp_sup'] = grafico_matplotlib_h(df_aberto, 'Supervisão', 'TOP 5 SUPERVISÃO', COR_AZUL, 'imp_sup', 5, upper=True, figsize=FIG_IMPACTO)

            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_auto_page_break(auto=False)

            # ---------- fonte ----------
            FONTE = "Helvetica"
            UNICODE_OK = False
            try:
                f_reg = os.path.join(PASTA_FONTES, "DejaVuSans.ttf")
                f_bold = os.path.join(PASTA_FONTES, "DejaVuSans-Bold.ttf")
                if not os.path.exists(f_bold):
                    f_bold = f_reg   # sem o arquivo Bold, o "negrito" sai com o peso normal
                pdf.add_font("DejaVu", "", f_reg)
                pdf.add_font("DejaVu", "B", f_bold)
                FONTE = "DejaVu"
                UNICODE_OK = True
            except Exception:
                pass

            def limpar(t):
                return (str(t).replace("—", "-").replace("–", "-").replace("●", "")
                        .replace("🟢", "").replace("🔴", "").replace("🟡", "").replace("⚠️", "!"))

            # ---------- helpers de desenho ----------
            def caixa(x, y, w, h, fill=None, borda=(190, 190, 190), raio=2.5, lw=0.25):
                if fill: pdf.set_fill_color(*fill)
                if borda:
                    pdf.set_draw_color(*borda)
                    pdf.set_line_width(lw)
                estilo = ('DF' if borda else 'F') if fill else 'D'
                try:
                    pdf.rect(x, y, w, h, style=estilo, round_corners=True, corner_radius=raio)
                except TypeError:
                    pdf.rect(x, y, w, h, style=estilo)   # versões antigas do fpdf: sem cantos arredondados

            def caixa_tracejada(x, y, w, h):
                try: pdf.set_dash_pattern(dash=1, gap=1)
                except Exception: pass
                pdf.set_draw_color(120, 120, 120)
                pdf.set_line_width(0.2)
                pdf.rect(x, y, w, h)
                try: pdf.set_dash_pattern()
                except Exception: pass

            def texto(x, y, w, h, t, tam=6, negrito=True, alin="C", cor=(26, 26, 26)):
                pdf.set_xy(x, y)
                pdf.set_font(FONTE, "B" if negrito else "", tam)
                pdf.set_text_color(*cor)
                pdf.cell(w, h, limpar(t), align=alin)
                pdf.set_text_color(26, 26, 26)

            def paragrafo(x, y, w, h_linha, t, tam=5.5, cor=(51, 51, 51), negrito=False):
                pdf.set_xy(x, y)
                pdf.set_font(FONTE, "B" if negrito else "", tam)
                pdf.set_text_color(*cor)
                pdf.multi_cell(w, h_linha, limpar(t))
                pdf.set_text_color(26, 26, 26)

            def tabela(x, y, larguras, cab, linhas, rh=4.2, tam=6, aligns=None):
                """Tabela estilo Excel: cabeçalho com linha embaixo, linhas finas cinza."""
                pdf.set_font(FONTE, "B", tam)
                pdf.set_xy(x, y)
                for h, wc in zip(cab, larguras):
                    pdf.cell(wc, rh, limpar(h), align="C")
                pdf.set_draw_color(60, 60, 60)
                pdf.set_line_width(0.3)
                pdf.line(x, y + rh, x + sum(larguras), y + rh)
                for i, linha in enumerate(linhas):
                    yy = y + rh * (i + 1)
                    pdf.set_xy(x, yy)
                    for j, (v, wc) in enumerate(zip(linha, larguras)):
                        cor = (26, 26, 26)
                        if isinstance(v, tuple):
                            v, cor = v
                        pdf.set_text_color(*cor)
                        pdf.cell(wc, rh, limpar(v), align=(aligns[j] if aligns else "C"))
                    pdf.set_text_color(26, 26, 26)
                    pdf.set_draw_color(220, 220, 220)
                    pdf.set_line_width(0.15)
                    pdf.line(x, yy + rh, x + sum(larguras), yy + rh)
                return y + rh * (len(linhas) + 1)

            def evo_pdf(v, bom):
                t, c = evo_texto(v, bom, UNICODE_OK)
                return (t, _rgb(c))

            def card_pdf(x, y, w, h, cor, rotulos, valor, pct=None, tam_val=11):
                caixa(x, y, w, h, fill=cor, borda=None, raio=2)
                yy = y + 1.3
                for r in rotulos:
                    texto(x, yy, w, 2.4, r, 4.8)
                    yy += 2.3
                texto(x, yy + 0.2, w, 5, str(valor), tam_val)
                yy += 5.4
                if pct is not None:
                    texto(x, yy, w, 2.4, pct, 4.8)

            def moldura_pagina():
                pdf.set_fill_color(64, 64, 64)
                pdf.rect(0, 0, 210, 1.2, style='F')
                pdf.set_fill_color(*_rgb(COR_CAIXA_AZUL))
                pdf.rect(0, 1.2, 22, 13.5, style='F')
                pdf.set_fill_color(*_rgb(COR_FAIXA_LARANJA))
                pdf.rect(0, 14.7, 8, 282.3, style='F')

            VERDE, CINZA = _rgb(COR_VERDE), _rgb(COR_CINZA)
            ESCURO = (26, 26, 26)

            # ============ PÁGINA 1 ============
            moldura_pagina()

            # ---- cabeçalho ----
            texto(26, 3.5, 120, 6, "Boletim Semanal", 15, alin="L")
            texto(26, 10.5, 120, 4, data_extenso(), 6.5, negrito=False, alin="L")
            texto(14, 17, 188, 5, "C E N T R A L   D O   C L I E N T E", 9, negrito=False)
            pdf.set_draw_color(*_rgb(COR_FAIXA_LARANJA))
            pdf.set_line_width(0.6)
            pdf.line(45, 23.5, 180, 23.5)
            texto(14, 24.5, 188, 13, "Domo de Ferro", 30)

            # ---- termômetro geral ----
            y0 = 40
            caixa(14, y0, 108, 27)
            texto(14, y0 + 1.5, 108, 3, "TERMÔMETRO GERAL DAS AÇÕES", 6)
            card_pdf(17, y0 + 7, 20, 17, VERDE, ["Total de", "Demandas"], total, tam_val=13)
            card_pdf(39.5, y0 + 8.5, 17, 14, CINZA, ["Backlog", "(Não iniciada)"], resumo['backlog'], f"{pct_back:.0f}%", 10)
            card_pdf(59, y0 + 8.5, 17, 14, CINZA, ["Em", "andamento"], resumo['em_andamento'], f"{pct_and:.0f}%", 10)
            card_pdf(78.5, y0 + 8.5, 17, 14, CINZA, ["Concluída", " "], resumo['concluida'], f"{pct_concl:.0f}%", 10)
            card_pdf(98, y0 + 7, 20, 17, VERDE, ["Dentro da", "Expectativa"], resumo['dentro_expectativa_geral'], f"{pct_dent:.0f}%", 13)

            # ---- termômetro tratativas ----
            caixa(125, y0, 77, 27)
            texto(125, y0 + 1.5, 77, 3, "TERMÔMETRO AÇÕES EM TRATATIVAS", 6)
            card_pdf(128, y0 + 7.5, 17, 16, CINZA, ["Total de", "Demandas"], tt, tam_val=12)
            linhas_trat = [
                ("Dentro da Expectativa:", resumo['dentro_expectativa_tratativas'], pct_de, (213, 245, 213), (30, 132, 73)),
                ("Fora da Expectativa:", resumo['fora_expectativa_tratativas'], pct_fe, (250, 219, 216), (192, 57, 43)),
                ("Risco de prazo:", resumo['risco_prazo'], pct_rp, (252, 243, 207), (183, 149, 11)),
            ]
            for i, (rot, val, pc, bg, fg) in enumerate(linhas_trat):
                yl = y0 + 7 + i * 6.3
                caixa(148, yl, 51, 5.3, fill=bg, borda=None, raio=1.5)
                pdf.set_fill_color(*fg)
                pdf.ellipse(150, yl + 1.8, 1.7, 1.7, style='F')
                texto(153, yl, 30, 5.3, rot, 5.5, alin="L", cor=fg)
                texto(181, yl, 8, 5.3, str(val), 10, cor=ESCURO if i != 1 else fg)
                texto(189, yl, 9, 5.3, f"{pc:.0f}%", 6, cor=fg)

            # ---- detalhamento das demandas ----
            y0 = 69
            caixa(14, y0, 188, 43)
            texto(14, y0 + 1.5, 188, 4, "DETALHAMENTO DAS DEMANDAS", 7)
            for nome, xi in zip(['tipologia', 'origem', 'municipios', 'supervisao'], [16, 62.5, 109, 155.5]):
                if imgs.get(nome) and os.path.exists(imgs[nome]):
                    pdf.image(imgs[nome], x=xi, y=y0 + 6, w=45)

            # ---- evolução semanal ----
            y0 = 114
            caixa(14, y0, 188, 25)
            texto(17, y0 + 2, 34, 4, "EVOLUÇÃO SEMANAL", 6.5, alin="L")
            linhas = [[n, s1, at, evo_pdf(v, b)] for n, s1, at, v, b in evo_semanal]
            tabela(52, y0 + 2, [46, 14, 14, 16], ["Status", "S-1", "S-Atual", "Evolução"], linhas)
            caixa_tracejada(145, y0 + 2, 55, 21)
            paragrafo(147, y0 + 3, 51, 2.6, "Nota: " + NOTA_EVO_SEMANAL + "  " + NOTA_REPACTUACAO, 5)

            # ---- maiores impactos ----
            y0 = 141
            caixa(14, y0, 188, 39)
            texto(17, y0 + 1.5, 150, 4, "MAIORES IMPACTO - DEMANDAS EM TRATATIVAS", 6.5, alin="L")
            card_pdf(17, y0 + 12, 20, 17, CINZA, ["Total de", "Demandas"], tt, tam_val=12)
            for nome, xi in zip(['imp_areas', 'imp_tipos', 'imp_sup'], [40, 93, 146]):
                if imgs.get(nome) and os.path.exists(imgs[nome]):
                    pdf.image(imgs[nome], x=xi, y=y0 + 6, w=52)

            # ---- aging + riscos ----
            y0 = 182
            texto(14, y0, 88, 4, "AGING dos dias em aberto", 6.5, alin="L")
            if imgs.get('aging') and os.path.exists(imgs['aging']):
                pdf.image(imgs['aging'], x=14, y=y0 + 5, w=86)

            caixa(104, y0, 98, 41, fill=(255, 248, 236), borda=_rgb(COR_FAIXA_LARANJA), lw=0.5)
            texto(107, y0 + 1.5, 90, 4, "Riscos e Fora da Expectativa", 6.5, alin="L", cor=(185, 119, 14))
            linhas_r = [[str(r['Area'])[:34], int(r['Fora']), int(r['Risco']), int(r['Total'])]
                        for _, r in tbl_riscos.head(6).iterrows()]
            linhas_r.append(["Total Geral"] + [int(v) for v in tot_riscos])
            y_fim = tabela(107, y0 + 6.5, [40, 24, 14, 14], ["Área", "Fora da Expectativa", "Riscos", "Total"],
                           linhas_r, rh=4, tam=5.5, aligns=["L", "C", "C", "C"])
            pdf.set_draw_color(60, 60, 60)
            pdf.set_line_width(0.3)
            pdf.line(107, y_fim - 4, 107 + 92, y_fim - 4)

            # ---- evolução expectativa ----
            y0 = 225
            caixa(14, y0, 188, 25)
            texto(17, y0 + 2, 34, 4, "EVOLUÇÃO EXPECTATIVA", 6.5, alin="L")
            linhas = [[n, s1, at, evo_pdf(v, b)] for n, s1, at, v, b in evo_expectativa]
            tabela(52, y0 + 2, [46, 14, 14, 16], ["Status", "S-1", "S-Atual", "Evolução"], linhas)
            caixa_tracejada(145, y0 + 2, 55, 21)
            paragrafo(147, y0 + 3, 51, 2.6, "Nota: " + NOTA_EVO_EXPECTATIVA + "  " + NOTA_REPACTUACAO, 5)

            # ---- alertas ----
            y0 = 252
            caixa(14, y0, 188, 28)
            texto(17, y0 + 1.5, 100, 4, "ALERTAS EXECUTIVOS", 6.5, alin="L")
            blocos = [
                ("ATENÇÃO", alertas['ATENCAO'], (250, 219, 216), (192, 57, 43)),
                ("CONCENTRAÇÃO", alertas['CONCENTRACAO'], (252, 243, 207), (185, 119, 14)),
                ("EVOLUÇÃO", alertas['EVOLUCAO'], (213, 245, 213), (30, 132, 73)),
            ]
            for (tit, txt, bg, fg), xi in zip(blocos, [17, 78, 139]):
                caixa(xi, y0 + 6.5, 58, 19, fill=bg, borda=None, raio=2)
                texto(xi + 2, y0 + 7.5, 54, 3, tit, 6, alin="L", cor=fg)
                paragrafo(xi + 2, y0 + 11.5, 54, 2.6, txt, 5.5, cor=(51, 51, 51))

            # ============ PÁGINA 2 - TOP PRIORIDADES ============
            pdf.add_page()
            moldura_pagina()
            pdf.set_left_margin(12.5)
            pdf.set_font(FONTE, "B", 14)
            texto(0, 18, 210, 8, "Domo de Ferro", 14)
            texto(0, 27, 210, 6, "TOP - PRIORIDADES", 11)

            if not df_analitico.empty:
                df_prio = df_analitico[df_analitico['Status_Acao'] == 'Em Andamento'].copy()
                if 'Expectativa' in df_prio.columns:
                    df_prio = df_prio[df_prio['Expectativa'] == 'Fora']
                df_prio = df_prio.head(20)

                pdf.set_xy(12.5, 36)   # tabela de 185 mm centralizada na página (12,5 + 185 + 12,5 = 210)
                pdf.set_font(FONTE, "B", 7)
                pdf.set_fill_color(220, 220, 220)
                pdf.set_draw_color(150, 150, 150)
                pdf.set_line_width(0.2)
                for h, ww in zip(["ID", "Tipologia", "Origem", "Área", "Localidade", "Status"], [15, 28, 35, 40, 35, 32]):
                    pdf.cell(ww, 6, limpar(h), border=1, align="C", fill=True)
                pdf.ln()
                pdf.set_font(FONTE, "", 6.5)
                for _, row in df_prio.iterrows():
                    pdf.set_x(12.5)
                    pdf.cell(15, 5, limpar(str(row.get('ID', ''))[:6]), border=1, align="C")
                    pdf.cell(28, 5, limpar(str(row.get('Tipologia', ''))[:14]), border=1, align="L")
                    pdf.cell(35, 5, limpar(str(row.get('Poder', ''))[:18]), border=1, align="L")
                    pdf.cell(40, 5, limpar(str(row.get('Area 2', ''))[:22]), border=1, align="L")
                    pdf.cell(35, 5, limpar(str(row.get('Localidade', ''))[:18]), border=1, align="L")
                    pdf.cell(32, 5, limpar("Fora da Expectativa"), border=1, align="C")
                    pdf.ln()

            destino = CAMINHO_PDF
            try:
                pdf.output(destino)
            except PermissionError:
                # o PDF anterior está aberto (leitor de PDF / OneDrive): salva com outro nome
                destino = CAMINHO_PDF.replace(".pdf", f"_{datetime.now():%H%M%S}.pdf")
                pdf.output(destino)
                st.warning("O PDF anterior estava aberto em outro programa. Salvei com outro nome.")
            with open(destino, "rb") as f:
                st.session_state['pdf_bytes'] = f.read()
            st.session_state['pdf_nome'] = os.path.basename(destino)
            st.success(f"✅ PDF gerado! Salvo em: {destino}")
        except Exception as e:
            st.session_state.pop('pdf_bytes', None)
            st.error(f"Erro ao gerar PDF: {type(e).__name__}: {e}")
            st.exception(e)   # mostra o erro completo (linha e tipo) para facilitar o diagnóstico

# o botão de download fica fora do "if" para não sumir quando a página recarrega
if 'pdf_bytes' in st.session_state:
    st.download_button(label="⬇️ Baixar PDF", data=st.session_state['pdf_bytes'],
                       file_name=st.session_state.get('pdf_nome', "Boletim_Domo_de_Ferro.pdf"),
                       mime="application/pdf")