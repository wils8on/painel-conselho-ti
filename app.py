import io
import re
import unicodedata
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from matplotlib.patches import Circle, FancyBboxPatch
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


st.set_page_config(
    page_title="Painel Conselho TI",
    layout="wide"
)


def normalizar_coluna(nome):
    nome = str(nome)
    nome = unicodedata.normalize("NFKD", nome)
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    nome = nome.lower().strip()
    nome = re.sub(r"[^\w\s]", "", nome)
    nome = re.sub(r"\s+", "_", nome)
    return nome


def preparar_dados(df):
    df = df.copy()

    colunas_datas = [
        "data_de_criacao",
        "data_de_finalizacao",
        "data_de_primeira_resposta",
        "deadline",
        "inicializacao",
    ]

    for coluna in colunas_datas:
        if coluna in df.columns:
            df[coluna] = pd.to_datetime(df[coluna], errors="coerce", dayfirst=True)

    externo = df.get("qual_o_seu_setor_chamado_externo")
    interno = df.get("qual_o_seu_setor_chamado_interno")

    if externo is not None and interno is not None:
        df["setor_solicitante"] = externo.combine_first(interno)
    elif externo is not None:
        df["setor_solicitante"] = externo
    elif interno is not None:
        df["setor_solicitante"] = interno

    return df


def texto_finalizado(valor):
    if pd.isna(valor):
        return False

    texto = str(valor).lower().strip()

    termos = [
        "finalizado",
        "fechado",
        "encerrado",
        "resolvido",
        "concluido",
        "concluído",
        "cancelado",
    ]

    return any(t in texto for t in termos)


def chamado_finalizado(linha):
    if "data_de_finalizacao" in linha.index and pd.notna(linha["data_de_finalizacao"]):
        return True

    if "status" in linha.index and texto_finalizado(linha["status"]):
        return True

    if "ultima_situacao" in linha.index and texto_finalizado(linha["ultima_situacao"]):
        return True

    return False


def sim(valor):
    if pd.isna(valor):
        return False

    texto = str(valor).lower().strip()
    return texto in ["sim", "s", "yes", "true", "1", "cumprido"]


def calcular_backlog_na_data(df, data):
    criados = df[
        df["data_de_criacao"].notna()
        & (df["data_de_criacao"] <= data)
    ]

    finalizados = criados[
        criados["data_de_finalizacao"].notna()
        & (criados["data_de_finalizacao"] <= data)
    ]

    return len(criados) - len(finalizados)


def calcular_backlog_diario(df, data_inicio, data_fim):
    if "data_de_criacao" not in df.columns:
        return pd.DataFrame(columns=["data", "backlog"])

    df = df[df["data_de_criacao"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["data", "backlog"])

    datas = pd.date_range(
        start=data_inicio,
        end=data_fim,
        freq="D"
    )

    backlog = []

    for data in datas:
        backlog.append({
            "data": data,
            "backlog": calcular_backlog_na_data(df, data)
        })

    return pd.DataFrame(backlog)


def calcular_backlog_mensal(df, data_inicio, data_fim):
    if "data_de_criacao" not in df.columns:
        return pd.DataFrame(columns=["mes", "backlog"])

    df = df[df["data_de_criacao"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["mes", "backlog"])

    mes_inicio = pd.Timestamp(data_inicio).to_period("M")
    mes_fim = pd.Timestamp(data_fim).to_period("M")

    meses = pd.period_range(
        start=mes_inicio,
        end=mes_fim,
        freq="M"
    )

    backlog = []

    for mes in meses:
        ultimo_dia_mes = mes.to_timestamp(how="end")

        if ultimo_dia_mes > pd.Timestamp(data_fim):
            ultimo_dia_mes = pd.Timestamp(data_fim)

        backlog.append({
            "mes": str(mes),
            "backlog": calcular_backlog_na_data(df, ultimo_dia_mes)
        })

    return pd.DataFrame(backlog)


def aplicar_filtros_sidebar(df):
    df_filtrado = df.copy()

    filtros_texto = [
        ("departamento", "Departamento"),
        ("categoria", "Categoria"),
        ("atendente", "Atendente"),
        ("setor_solicitante", "Setor solicitante"),
        ("prioridade", "Prioridade"),
        ("ultima_situacao", "Última situação"),
    ]

    for coluna, label in filtros_texto:
        if coluna in df.columns:
            opcoes = sorted(
                df[coluna]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            selecionados = st.sidebar.multiselect(
                label,
                opcoes
            )

            if selecionados:
                df_filtrado = df_filtrado[
                    df_filtrado[coluna].astype(str).isin(selecionados)
                ]

    return df_filtrado


def filtrar_por_mes(df, periodo_mes):
    if "data_de_criacao" not in df.columns:
        return df.iloc[0:0].copy()

    return df[
        df["data_de_criacao"].notna()
        & (df["data_de_criacao"].dt.to_period("M") == periodo_mes)
    ].copy()


def calcular_indicadores_mes(df_mes, df_base, periodo_mes):
    total = len(df_mes)

    if total > 0:
        mask_finalizado = df_mes.apply(chamado_finalizado, axis=1)
        finalizados = mask_finalizado.sum()
        abertos = total - finalizados
    else:
        finalizados = 0
        abertos = 0

    reabertos = 0
    if "reaberto" in df_mes.columns and total > 0:
        reabertos = df_mes["reaberto"].apply(sim).sum()

    sla_deadline = None
    if "sla_de_deadline_cumprido" in df_mes.columns and total > 0:
        sla_deadline = df_mes["sla_de_deadline_cumprido"].apply(sim).mean() * 100

    sla_inicializacao = None
    if "sla_de_inicializacao_cumprido" in df_mes.columns and total > 0:
        sla_inicializacao = (
            df_mes["sla_de_inicializacao_cumprido"].apply(sim).mean() * 100
        )

    ultimo_dia_mes = periodo_mes.to_timestamp(how="end")
    backlog_fim = calcular_backlog_na_data(df_base, ultimo_dia_mes)

    return {
        "total": total,
        "finalizados": finalizados,
        "abertos": abertos,
        "reabertos": reabertos,
        "sla_deadline": sla_deadline,
        "sla_inicializacao": sla_inicializacao,
        "backlog_fim": backlog_fim,
    }


def calcular_dados_comparativo_mensal(df_base, data_fim):
    if data_fim is None or "data_de_criacao" not in df_base.columns:
        return None

    if not df_base["data_de_criacao"].notna().any():
        return None

    mes_atual_periodo = pd.Timestamp(data_fim).to_period("M")
    mes_anterior_periodo = mes_atual_periodo - 1

    df_mes_atual = filtrar_por_mes(df_base, mes_atual_periodo)
    df_mes_anterior = filtrar_por_mes(df_base, mes_anterior_periodo)

    ind_mes_atual = calcular_indicadores_mes(
        df_mes_atual,
        df_base,
        mes_atual_periodo,
    )
    ind_mes_anterior = calcular_indicadores_mes(
        df_mes_anterior,
        df_base,
        mes_anterior_periodo,
    )

    return {
        "mes_atual_periodo": mes_atual_periodo,
        "mes_anterior_periodo": mes_anterior_periodo,
        "label_mes_atual": formatar_periodo_mes(mes_atual_periodo),
        "label_mes_anterior": formatar_periodo_mes(mes_anterior_periodo),
        "ind_mes_atual": ind_mes_atual,
        "ind_mes_anterior": ind_mes_anterior,
        "var_chamados": variacao_absoluta(
            ind_mes_atual["total"],
            ind_mes_anterior["total"],
        ),
        "var_chamados_pct": variacao_percentual(
            ind_mes_atual["total"],
            ind_mes_anterior["total"],
        ),
        "var_backlog": variacao_absoluta(
            ind_mes_atual["backlog_fim"],
            ind_mes_anterior["backlog_fim"],
        ),
        "var_sla_deadline": variacao_absoluta(
            ind_mes_atual["sla_deadline"],
            ind_mes_anterior["sla_deadline"],
        ),
    }


def formatar_periodo_mes(periodo_mes):
    meses = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }
    return f"{meses[periodo_mes.month]} {periodo_mes.year}"


def variacao_absoluta(atual, anterior):
    if atual is None or anterior is None:
        return None
    return atual - anterior


def variacao_percentual(atual, anterior):
    if atual is None or anterior is None:
        return None
    if anterior == 0:
        return None
    return ((atual - anterior) / anterior) * 100


def formatar_valor_indicador(valor, sufixo=""):
    if valor is None:
        return "—"
    if isinstance(valor, float):
        if sufixo == "%":
            return f"{valor:.1f}%"
        return f"{valor:,.0f}"
    return f"{valor:,}{sufixo}"


def formatar_variacao(valor, sufixo=""):
    if valor is None:
        return "—"
    sinal = "+" if valor > 0 else ""
    if sufixo == "%":
        return f"{sinal}{valor:.1f} pp"
    if sufixo == "pct":
        return f"{sinal}{valor:.1f}%"
    return f"{sinal}{valor:,.0f}"


def montar_tabela_comparativa(ind_atual, ind_anterior, label_atual, label_anterior):
    linhas = [
        ("Total de chamados", "total", ""),
        ("Chamados finalizados", "finalizados", ""),
        ("Chamados em aberto", "abertos", ""),
        ("Chamados reabertos", "reabertos", ""),
        ("SLA Deadline cumprido", "sla_deadline", "%"),
        ("SLA Inicialização cumprido", "sla_inicializacao", "%"),
        ("Backlog no fim do mês", "backlog_fim", ""),
    ]

    dados = []

    for indicador, chave, sufixo in linhas:
        atual = ind_atual[chave]
        anterior = ind_anterior[chave]
        var_abs = variacao_absoluta(atual, anterior)

        if sufixo == "%":
            var_pct = var_abs
            var_pct_fmt = formatar_variacao(var_pct, "%")
        else:
            var_pct = variacao_percentual(atual, anterior)
            var_pct_fmt = formatar_variacao(var_pct, "pct")

        dados.append({
            "Indicador": indicador,
            "Mês Atual": formatar_valor_indicador(atual, sufixo),
            "Mês Anterior": formatar_valor_indicador(anterior, sufixo),
            "Variação Absoluta": formatar_variacao(var_abs, sufixo if sufixo == "%" else ""),
            "Variação %": var_pct_fmt,
        })

    return pd.DataFrame(dados)


def gerar_interpretacao_comparativo(
    ind_atual,
    ind_anterior,
    label_atual,
    label_anterior,
):
    if ind_atual["total"] == 0 and ind_anterior["total"] == 0:
        return (
            "Não há chamados criados nos meses comparados "
            "com os filtros selecionados."
        )

    var_chamados = variacao_absoluta(ind_atual["total"], ind_anterior["total"])
    var_backlog = variacao_absoluta(
        ind_atual["backlog_fim"],
        ind_anterior["backlog_fim"],
    )
    var_sla = variacao_absoluta(
        ind_atual["sla_deadline"],
        ind_anterior["sla_deadline"],
    )

    if var_chamados is None:
        texto_chamados = "Não foi possível comparar o volume de chamados."
    elif var_chamados > 0:
        texto_chamados = (
            f"Houve **aumento** de {var_chamados:,} chamados em relação a "
            f"{label_anterior}, passando de {ind_anterior['total']:,} para "
            f"{ind_atual['total']:,} em {label_atual}."
        )
    elif var_chamados < 0:
        texto_chamados = (
            f"Houve **redução** de {abs(var_chamados):,} chamados em relação a "
            f"{label_anterior}, passando de {ind_anterior['total']:,} para "
            f"{ind_atual['total']:,} em {label_atual}."
        )
    else:
        texto_chamados = (
            f"O volume de chamados permaneceu **estável** entre "
            f"{label_anterior} e {label_atual}."
        )

    if var_backlog is None:
        texto_backlog = "Não foi possível comparar o backlog."
    elif var_backlog > 0:
        texto_backlog = (
            f"O backlog ao fim do mês **aumentou** em {var_backlog:,} chamados "
            f"({ind_anterior['backlog_fim']:,} → {ind_atual['backlog_fim']:,})."
        )
    elif var_backlog < 0:
        texto_backlog = (
            f"O backlog ao fim do mês **reduziu** em {abs(var_backlog):,} "
            f"chamados ({ind_anterior['backlog_fim']:,} → "
            f"{ind_atual['backlog_fim']:,})."
        )
    else:
        texto_backlog = "O backlog ao fim do mês permaneceu **estável**."

    if var_sla is None:
        texto_sla = "Não há dados suficientes para avaliar a evolução do SLA."
    elif var_sla > 0:
        texto_sla = (
            f"O SLA de Deadline **melhorou** {var_sla:.1f} pontos percentuais "
            f"({ind_anterior['sla_deadline']:.1f}% → "
            f"{ind_atual['sla_deadline']:.1f}%)."
            if ind_anterior["sla_deadline"] is not None
            and ind_atual["sla_deadline"] is not None
            else "O SLA de Deadline apresentou **melhora** no período."
        )
    elif var_sla < 0:
        texto_sla = (
            f"O SLA de Deadline **piorou** {abs(var_sla):.1f} pontos "
            f"percentuais ({ind_anterior['sla_deadline']:.1f}% → "
            f"{ind_atual['sla_deadline']:.1f}%)."
            if ind_anterior["sla_deadline"] is not None
            and ind_atual["sla_deadline"] is not None
            else "O SLA de Deadline apresentou **piora** no período."
        )
    else:
        texto_sla = "O SLA de Deadline permaneceu **estável** entre os meses."

    atencao = []

    if var_chamados is not None and var_chamados > 0:
        atencao.append(
            "crescimento do volume de chamados, que pode pressionar a equipe"
        )

    if var_backlog is not None and var_backlog > 0:
        atencao.append(
            "aumento do backlog, indicando acúmulo de demandas pendentes"
        )

    if var_sla is not None and var_sla < 0:
        atencao.append(
            "queda no cumprimento do SLA de Deadline"
        )

    if ind_atual["reabertos"] > ind_anterior["reabertos"]:
        atencao.append(
            "elevação de chamados reabertos, possível indicador de retrabalho"
        )

    if ind_atual["abertos"] > ind_anterior["abertos"]:
        atencao.append(
            "maior volume de chamados em aberto no mês atual"
        )

    if atencao:
        texto_atencao = (
            "**Ponto de atenção:** " + "; ".join(atencao) + "."
        )
    else:
        texto_atencao = (
            "**Ponto de atenção:** nenhum indicador crítico se destacou; "
            "manter o acompanhamento da operação."
        )

    return f"""
Comparativo entre **{label_anterior}** (mês anterior) e **{label_atual}** (mês atual):

{texto_chamados}

{texto_backlog}

{texto_sla}

{texto_atencao}
"""


def aplicar_filtros(df):
    st.sidebar.header("Filtros")

    if "data_de_criacao" in df.columns and df["data_de_criacao"].notna().any():
        data_min = df["data_de_criacao"].min().date()
        data_max = df["data_de_criacao"].max().date()

        periodo = st.sidebar.date_input(
            "Período de criação",
            value=(data_min, data_max),
            min_value=data_min,
            max_value=data_max,
            format="DD/MM/YYYY"
        )

        if isinstance(periodo, tuple) and len(periodo) == 2:
            data_inicio, data_fim = periodo
        else:
            data_inicio = data_min
            data_fim = data_max
    else:
        data_inicio = None
        data_fim = None

    df_sidebar = aplicar_filtros_sidebar(df)
    df_filtrado = df_sidebar.copy()

    if "data_de_criacao" in df_filtrado.columns and data_inicio is not None:
        df_filtrado = df_filtrado[
            (df_filtrado["data_de_criacao"].dt.date >= data_inicio)
            & (df_filtrado["data_de_criacao"].dt.date <= data_fim)
        ]

    return df_filtrado, data_inicio, data_fim, df_sidebar


def grafico_barras(df, coluna, titulo):
    if coluna not in df.columns:
        return None

    dados = (
        df[coluna]
        .dropna()
        .astype(str)
        .str.strip()
    )

    dados = dados[dados != ""]

    if dados.empty:
        return None

    resumo = dados.value_counts().head(10).reset_index()
    resumo.columns = [coluna, "quantidade"]

    fig = px.bar(
        resumo,
        x="quantidade",
        y=coluna,
        orientation="h",
        text="quantidade",
        title=titulo
    )

    fig.update_layout(
        yaxis={"categoryorder": "total ascending"},
        height=400
    )

    return fig


def obter_maior_volume(df, coluna):
    if coluna not in df.columns or df.empty:
        return None, 0

    dados = (
        df[coluna]
        .dropna()
        .astype(str)
        .str.strip()
    )
    dados = dados[dados != ""]

    if dados.empty:
        return None, 0

    contagem = dados.value_counts()
    return contagem.index[0], int(contagem.iloc[0])


def gerar_insights_automaticos(
    total_chamados,
    reabertos,
    backlog_atual,
    sla_deadline,
    sla_inicializacao,
    df_filtrado,
    comparativo=None,
):
    insights = []
    dados_insuficientes = comparativo is None

    if total_chamados > 0:
        pct_reabertos = (reabertos / total_chamados) * 100

        if pct_reabertos > 10:
            insights.append({
                "tipo": "warning",
                "mensagem": (
                    f"Chamados reabertos acima de 10% do total: {reabertos:,} "
                    f"({pct_reabertos:.1f}%), indicador de possível retrabalho."
                ),
            })

    if sla_deadline is not None:
        if sla_deadline < 90:
            insights.append({
                "tipo": "warning",
                "mensagem": (
                    f"SLA de Deadline abaixo de 90% no período filtrado "
                    f"({sla_deadline:.1f}%)."
                ),
            })
        elif sla_deadline > 95:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"SLA de Deadline acima de 95% no período filtrado "
                    f"({sla_deadline:.1f}%)."
                ),
            })

    if sla_inicializacao is not None:
        if sla_inicializacao < 90:
            insights.append({
                "tipo": "warning",
                "mensagem": (
                    f"SLA de Inicialização abaixo de 90% no período filtrado "
                    f"({sla_inicializacao:.1f}%)."
                ),
            })
        elif sla_inicializacao > 95:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"SLA de Inicialização acima de 95% no período filtrado "
                    f"({sla_inicializacao:.1f}%)."
                ),
            })

    if comparativo is not None:
        label_atual = comparativo["label_mes_atual"]
        label_anterior = comparativo["label_mes_anterior"]
        ind_atual = comparativo["ind_mes_atual"]
        ind_anterior = comparativo["ind_mes_anterior"]
        var_chamados_pct = comparativo["var_chamados_pct"]
        var_backlog = comparativo["var_backlog"]
        var_sla_deadline = comparativo["var_sla_deadline"]

        if var_chamados_pct is not None and var_chamados_pct > 15:
            insights.append({
                "tipo": "warning",
                "mensagem": (
                    f"Crescimento de chamados acima de 15% em {label_atual}: "
                    f"{var_chamados_pct:.1f}% em relação a {label_anterior} "
                    f"({ind_anterior['total']:,} → {ind_atual['total']:,})."
                ),
            })
        elif var_chamados_pct is not None and var_chamados_pct < 0:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"Redução de chamados em {label_atual}: "
                    f"{var_chamados_pct:.1f}% em relação a {label_anterior} "
                    f"({ind_anterior['total']:,} → {ind_atual['total']:,})."
                ),
            })

        if var_backlog is not None and var_backlog > 0:
            insights.append({
                "tipo": "warning",
                "mensagem": (
                    f"Backlog em crescimento ao fim de {label_atual}: "
                    f"+{var_backlog:,} chamados em relação a {label_anterior} "
                    f"({ind_anterior['backlog_fim']:,} → "
                    f"{ind_atual['backlog_fim']:,}). "
                    f"Backlog atual no período: {backlog_atual:,}."
                ),
            })
        elif var_backlog is not None and var_backlog < 0:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"Backlog em redução ao fim de {label_atual}: "
                    f"{var_backlog:,} chamados em relação a {label_anterior} "
                    f"({ind_anterior['backlog_fim']:,} → "
                    f"{ind_atual['backlog_fim']:,}). "
                    f"Backlog atual no período: {backlog_atual:,}."
                ),
            })

        sla_atual = ind_atual["sla_deadline"]
        sla_anterior = ind_anterior["sla_deadline"]

        if sla_atual is not None and sla_atual < 90:
            mensagem_sla = (
                f"SLA de Deadline em {label_atual} abaixo de 90% "
                f"({sla_atual:.1f}%)."
            )

            if (
                var_sla_deadline is not None
                and var_sla_deadline < 0
                and sla_anterior is not None
            ):
                mensagem_sla += (
                    f" Houve redução de {abs(var_sla_deadline):.1f} pp em "
                    f"relação a {label_anterior} ({sla_anterior:.1f}%)."
                )

            insights.append({
                "tipo": "warning",
                "mensagem": mensagem_sla,
            })
        elif sla_atual is not None and sla_atual > 95:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"SLA de Deadline em {label_atual} acima de 95% "
                    f"({sla_atual:.1f}%)."
                ),
            })

        sla_ini_atual = ind_atual["sla_inicializacao"]
        sla_ini_anterior = ind_anterior["sla_inicializacao"]
        var_sla_ini = variacao_absoluta(sla_ini_atual, sla_ini_anterior)

        if sla_ini_atual is not None and sla_ini_atual < 90:
            mensagem_sla_ini = (
                f"SLA de Inicialização em {label_atual} abaixo de 90% "
                f"({sla_ini_atual:.1f}%)."
            )

            if (
                var_sla_ini is not None
                and var_sla_ini < 0
                and sla_ini_anterior is not None
            ):
                mensagem_sla_ini += (
                    f" Houve redução de {abs(var_sla_ini):.1f} pp em relação a "
                    f"{label_anterior} ({sla_ini_anterior:.1f}%)."
                )

            insights.append({
                "tipo": "warning",
                "mensagem": mensagem_sla_ini,
            })
        elif sla_ini_atual is not None and sla_ini_atual > 95:
            insights.append({
                "tipo": "success",
                "mensagem": (
                    f"SLA de Inicialização em {label_atual} acima de 95% "
                    f"({sla_ini_atual:.1f}%)."
                ),
            })

    if total_chamados > 0:
        setor, qtd_setor = obter_maior_volume(
            df_filtrado,
            "setor_solicitante",
        )

        if setor is not None:
            pct_setor = (qtd_setor / total_chamados) * 100
            insights.append({
                "tipo": "info",
                "mensagem": (
                    f"Setor com maior volume: **{setor}** "
                    f"({qtd_setor:,} chamados, {pct_setor:.1f}% do total)."
                ),
            })

        categoria, qtd_categoria = obter_maior_volume(
            df_filtrado,
            "categoria",
        )

        if categoria is not None:
            pct_categoria = (qtd_categoria / total_chamados) * 100
            insights.append({
                "tipo": "info",
                "mensagem": (
                    f"Categoria com maior volume: **{categoria}** "
                    f"({qtd_categoria:,} chamados, "
                    f"{pct_categoria:.1f}% do total)."
                ),
            })

        atendente, qtd_atendente = obter_maior_volume(
            df_filtrado,
            "atendente",
        )

        if atendente is not None:
            pct_atendente = (qtd_atendente / total_chamados) * 100
            insights.append({
                "tipo": "info",
                "mensagem": (
                    f"Atendente com maior volume: **{atendente}** "
                    f"({qtd_atendente:,} chamados, "
                    f"{pct_atendente:.1f}% do total)."
                ),
            })

    ordem = {"warning": 0, "success": 1, "info": 2}
    insights.sort(key=lambda item: ordem[item["tipo"]])

    return insights, dados_insuficientes


def exibir_insights_automaticos(insights, dados_insuficientes=False):
    st.subheader("Insights Automáticos")

    if dados_insuficientes:
        st.info(
            "Não há dados suficientes para gerar todos os insights automáticos."
        )

    for insight in insights:
        if insight["tipo"] == "warning":
            st.warning(insight["mensagem"])
        elif insight["tipo"] == "success":
            st.success(insight["mensagem"])
        else:
            st.info(insight["mensagem"])

def calcular_aging(df):
    if "data_de_criacao" not in df.columns or "data_de_finalizacao" not in df.columns:
        return pd.DataFrame()

    hoje = pd.Timestamp.today().normalize()

    df_abertos = df[
        df["data_de_finalizacao"].isna()
    ].copy()

    if df_abertos.empty:
        return pd.DataFrame()

    df_abertos["idade_dias"] = (
        hoje - df_abertos["data_de_criacao"]
    ).dt.days

    def faixa_aging(dias):
        if dias <= 3:
            return "0-3 dias"
        elif dias <= 7:
            return "4-7 dias"
        elif dias <= 15:
            return "8-15 dias"
        elif dias <= 30:
            return "16-30 dias"
        else:
            return "+30 dias"

    df_abertos["faixa_aging"] = df_abertos["idade_dias"].apply(faixa_aging)

    return df_abertos
def gerar_resumo_executivo(
    total_chamados,
    finalizados,
    abertos,
    backlog_atual,
    reabertos,
    sla_deadline,
    sla_inicializacao,
    df_filtrado
):
    if total_chamados == 0:
        return "Não há chamados no período e filtros selecionados."

    percentual_finalizados = (finalizados / total_chamados) * 100
    percentual_abertos = (abertos / total_chamados) * 100

    departamento_top = "não identificado"

    if (
        "departamento" in df_filtrado.columns
        and not df_filtrado["departamento"].dropna().empty
    ):
        departamento_top = (
            df_filtrado["departamento"]
            .value_counts()
            .idxmax()
        )

    categoria_top = "não identificada"

    if (
        "categoria" in df_filtrado.columns
        and not df_filtrado["categoria"].dropna().empty
    ):
        categoria_top = (
            df_filtrado["categoria"]
            .value_counts()
            .idxmax()
        )

    setor_top = "não identificado"

    if (
        "setor_solicitante" in df_filtrado.columns
        and not df_filtrado["setor_solicitante"].dropna().empty
    ):
        setor_top = (
            df_filtrado["setor_solicitante"]
            .value_counts()
            .idxmax()
        )

    sla_deadline_txt = (
        f"{sla_deadline:.1f}%"
        if sla_deadline is not None
        else "não disponível"
    )

    sla_inicializacao_txt = (
        f"{sla_inicializacao:.1f}%"
        if sla_inicializacao is not None
        else "não disponível"
    )

    resumo = f"""
No período analisado, foram registrados **{total_chamados:,} chamados**, dos quais **{finalizados:,} foram finalizados** ({percentual_finalizados:.1f}%) e **{abertos:,} permanecem em aberto** ({percentual_abertos:.1f}%).

O backlog atual é de **{backlog_atual:,} chamados**, representando o volume de demandas ainda pendentes de conclusão.

Foram identificados **{reabertos:,} chamados reabertos**, indicador que deve ser acompanhado por representar possível retrabalho ou necessidade de revisão na qualidade das tratativas.

Em relação aos acordos de nível de serviço, o SLA de Deadline apresenta índice de **{sla_deadline_txt}**, enquanto o SLA de Inicialização apresenta **{sla_inicializacao_txt}**.

A maior concentração de chamados está no departamento **{departamento_top}**, com destaque para a categoria **{categoria_top}** e para o setor solicitante **{setor_top}**.

Esses pontos devem ser observados como possíveis focos de demanda operacional, necessidade de automação, padronização de processos ou reforço de orientação aos usuários.

De forma geral, os indicadores permitem acompanhar o comportamento da operação de TI, avaliar capacidade de atendimento e apoiar decisões sobre priorização, melhoria contínua e alocação de recursos.
"""

    return resumo


COR_FUNDO_SLIDE = RGBColor(248, 249, 250)
COR_TITULO_SLIDE = RGBColor(33, 37, 41)
COR_TEXTO_SLIDE = RGBColor(73, 80, 87)
COR_CABECALHO_TABELA = RGBColor(52, 58, 64)
COR_PRIMARIA = RGBColor(15, 52, 96)
COR_SECUNDARIA = RGBColor(0, 123, 255)
COR_CARD_FUNDO = RGBColor(255, 255, 255)
COR_CARD_BORDA = RGBColor(206, 212, 218)
COR_SUBTITULO = RGBColor(108, 117, 125)
COR_RODAPE = RGBColor(134, 142, 150)


def limpar_markdown(texto):
    if not texto:
        return ""

    texto = re.sub(r"\*\*([^*]+)\*\*", r"\1", texto)
    texto = re.sub(r"\*([^*]+)\*", r"\1", texto)
    return texto.strip()


def aplicar_fundo_slide(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = COR_FUNDO_SLIDE


def definir_texto_run(run, tamanho=14, negrito=False, cor=COR_TEXTO_SLIDE):
    font = run.font
    font.size = Pt(tamanho)
    font.bold = negrito
    font.color.rgb = cor
    font.name = "Calibri"


def _estilo_grafico_executivo(ax, titulo):
    ax.set_title(titulo, fontsize=11, color="#0F3460", pad=10, fontweight="bold")
    ax.tick_params(colors="#495057", labelsize=8)
    ax.grid(True, alpha=0.2, color="#CED4DA")
    for spine in ax.spines.values():
        spine.set_color("#CED4DA")


def _salvar_figura_matplotlib(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor="#F8F9FA")
    plt.close(fig)
    buf.seek(0)
    return buf


def criar_logotipo_ti_generico():
    fig, ax = plt.subplots(figsize=(2.2, 2.2))
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    monitor = FancyBboxPatch(
        (2.2, 2.5),
        5.6,
        4.2,
        boxstyle="round,pad=0.15,rounding_size=0.3",
        facecolor="#0F3460",
        edgecolor="none",
    )
    ax.add_patch(monitor)
    ax.text(5, 4.6, "TI", ha="center", va="center", fontsize=20, color="white", fontweight="bold")

    for pos_x, pos_y in [(1.2, 8.2), (8.8, 8.2), (5, 9.2)]:
        ax.add_patch(Circle((pos_x, pos_y), 0.55, facecolor="#007BFF", edgecolor="none"))

    ax.plot([5, 1.2], [9.2, 8.2], color="#007BFF", linewidth=1.5, alpha=0.6)
    ax.plot([5, 8.8], [9.2, 8.2], color="#007BFF", linewidth=1.5, alpha=0.6)
    ax.plot([5, 5], [9.2, 6.7], color="#007BFF", linewidth=1.5, alpha=0.6)

    return _salvar_figura_matplotlib(fig)


def adicionar_faixa_cabecalho(slide, titulo, subtitulo=None):
    faixa = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(13.333),
        Inches(1.05),
    )
    faixa.fill.solid()
    faixa.fill.fore_color.rgb = COR_PRIMARIA
    faixa.line.fill.background()

    adicionar_caixa_texto(
        slide,
        titulo,
        Inches(0.55),
        Inches(0.18),
        Inches(10.5),
        Inches(0.5),
        tamanho=24,
        negrito=True,
        cor=RGBColor(255, 255, 255),
    )

    if subtitulo:
        adicionar_caixa_texto(
            slide,
            subtitulo,
            Inches(0.55),
            Inches(0.62),
            Inches(11),
            Inches(0.35),
            tamanho=11,
            cor=RGBColor(200, 220, 240),
        )


def adicionar_rodape_executivo(slide, numero_slide, total_slides):
    adicionar_caixa_texto(
        slide,
        f"Painel Conselho TI  |  Slide {numero_slide} de {total_slides}",
        Inches(0.55),
        Inches(7.05),
        Inches(12),
        Inches(0.3),
        tamanho=9,
        cor=COR_RODAPE,
    )


def adicionar_card_kpi(slide, rotulo, valor, left, top, width, height):
    cartao = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
    )
    cartao.fill.solid()
    cartao.fill.fore_color.rgb = COR_CARD_FUNDO
    cartao.line.color.rgb = COR_CARD_BORDA
    cartao.adjustments[0] = 0.08

    barra = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left,
        top,
        width,
        Inches(0.1),
    )
    barra.fill.solid()
    barra.fill.fore_color.rgb = COR_SECUNDARIA
    barra.line.fill.background()

    tf_rotulo = slide.shapes.add_textbox(
        left + Inches(0.15),
        top + Inches(0.22),
        width - Inches(0.3),
        Inches(0.45),
    ).text_frame
    p_rotulo = tf_rotulo.paragraphs[0]
    p_rotulo.text = rotulo
    definir_texto_run(p_rotulo.runs[0], tamanho=11, cor=COR_SUBTITULO)

    tf_valor = slide.shapes.add_textbox(
        left + Inches(0.15),
        top + Inches(0.62),
        width - Inches(0.3),
        height - Inches(0.7),
    ).text_frame
    p_valor = tf_valor.paragraphs[0]
    p_valor.text = str(valor)
    definir_texto_run(
        p_valor.runs[0],
        tamanho=22,
        negrito=True,
        cor=COR_PRIMARIA,
    )


def adicionar_caixa_texto(
    slide,
    texto,
    left,
    top,
    width,
    height,
    tamanho=14,
    negrito=False,
    cor=COR_TEXTO_SLIDE,
):
    caixa = slide.shapes.add_textbox(left, top, width, height)
    tf = caixa.text_frame
    tf.word_wrap = True

    linhas = texto.split("\n")

    for indice, linha in enumerate(linhas):
        paragrafo = tf.paragraphs[0] if indice == 0 else tf.add_paragraph()
        paragrafo.text = linha
        paragrafo.space_after = Pt(6)
        paragrafo.alignment = PP_ALIGN.LEFT

        if paragrafo.runs:
            run = paragrafo.runs[0]
        else:
            run = paragrafo.add_run()

        definir_texto_run(run, tamanho=tamanho, negrito=negrito, cor=cor)

    return caixa


def criar_grafico_backlog_diario(backlog_diario):
    fig, ax = plt.subplots(figsize=(6.2, 3.0))
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#F8F9FA")

    if backlog_diario.empty:
        ax.text(
            0.5,
            0.5,
            "Sem dados de backlog diário",
            ha="center",
            va="center",
            fontsize=11,
            color="#495057",
        )
        ax.axis("off")
    else:
        ax.fill_between(
            backlog_diario["data"],
            backlog_diario["backlog"],
            alpha=0.15,
            color="#007BFF",
        )
        ax.plot(
            backlog_diario["data"],
            backlog_diario["backlog"],
            color="#0F3460",
            linewidth=2.2,
            marker="o",
            markersize=2.5,
        )
        _estilo_grafico_executivo(ax, "Backlog Diário")

    return _salvar_figura_matplotlib(fig)


def criar_grafico_backlog_mensal(backlog_mensal):
    fig, ax = plt.subplots(figsize=(6.2, 3.0))
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#F8F9FA")

    if backlog_mensal.empty:
        ax.text(
            0.5,
            0.5,
            "Sem dados de backlog mensal",
            ha="center",
            va="center",
            fontsize=11,
            color="#495057",
        )
        ax.axis("off")
    else:
        barras = ax.bar(
            backlog_mensal["mes"],
            backlog_mensal["backlog"],
            color="#0F3460",
            edgecolor="#007BFF",
            linewidth=0.6,
        )
        for barra in barras:
            altura = barra.get_height()
            ax.text(
                barra.get_x() + barra.get_width() / 2,
                altura,
                f"{int(altura):,}",
                ha="center",
                va="bottom",
                fontsize=7,
                color="#495057",
            )
        _estilo_grafico_executivo(ax, "Backlog Mensal")
        ax.tick_params(axis="x", rotation=35)

    return _salvar_figura_matplotlib(fig)


def criar_grafico_barras_operacao(df, coluna, titulo, limite=8):
    fig, ax = plt.subplots(figsize=(6.2, 3.2))
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#F8F9FA")

    if coluna not in df.columns or df.empty:
        ax.text(
            0.5,
            0.5,
            f"Sem dados para {titulo.lower()}",
            ha="center",
            va="center",
            fontsize=11,
            color="#495057",
        )
        ax.axis("off")
    else:
        dados = (
            df[coluna]
            .dropna()
            .astype(str)
            .str.strip()
        )
        dados = dados[dados != ""]

        if dados.empty:
            ax.text(
                0.5,
                0.5,
                f"Sem dados para {titulo.lower()}",
                ha="center",
                va="center",
                fontsize=11,
                color="#495057",
            )
            ax.axis("off")
        else:
            resumo = dados.value_counts().head(limite).sort_values()
            cores = plt.cm.Blues(np.linspace(0.45, 0.85, len(resumo)))
            ax.barh(resumo.index, resumo.values, color=cores)
            _estilo_grafico_executivo(ax, titulo)
            ax.tick_params(axis="y", labelsize=7)

            for indice, valor in enumerate(resumo.values):
                ax.text(
                    valor,
                    indice,
                    f" {valor:,}",
                    va="center",
                    fontsize=7,
                    color="#495057",
                )

    return _salvar_figura_matplotlib(fig)


def adicionar_tabela_dataframe(slide, df_tabela, left, top, width, height):
    if df_tabela is None or df_tabela.empty:
        adicionar_caixa_texto(
            slide,
            "Sem dados disponíveis para exibição.",
            left,
            top,
            width,
            height,
            tamanho=12,
        )
        return

    linhas = len(df_tabela) + 1
    colunas = len(df_tabela.columns)
    tabela = slide.shapes.add_table(linhas, colunas, left, top, width, height).table

    for col_idx, nome_coluna in enumerate(df_tabela.columns):
        celula = tabela.cell(0, col_idx)
        celula.text = str(nome_coluna)
        for paragrafo in celula.text_frame.paragraphs:
            for run in paragrafo.runs:
                definir_texto_run(
                    run,
                    tamanho=10,
                    negrito=True,
                    cor=COR_CABECALHO_TABELA,
                )

    for row_idx, registro in enumerate(df_tabela.itertuples(index=False), start=1):
        for col_idx, valor in enumerate(registro):
            celula = tabela.cell(row_idx, col_idx)
            celula.text = str(valor)
            for paragrafo in celula.text_frame.paragraphs:
                for run in paragrafo.runs:
                    definir_texto_run(run, tamanho=9)


def formatar_periodo_analisado(data_inicio, data_fim):
    if data_inicio is None or data_fim is None:
        return "Período não informado"

    return (
        f"{pd.Timestamp(data_inicio).strftime('%d/%m/%Y')} a "
        f"{pd.Timestamp(data_fim).strftime('%d/%m/%Y')}"
    )


def formatar_insights_pptx(insights, dados_insuficientes):
    linhas = []

    if dados_insuficientes:
        linhas.append(
            "• Não há dados suficientes para gerar todos os insights automáticos."
        )

    prefixos = {
        "warning": "Atenção",
        "success": "Positivo",
        "info": "Informação",
    }

    for insight in insights:
        tipo = insight.get("tipo", "info")
        prefixo = prefixos.get(tipo, "Informação")
        mensagem = limpar_markdown(insight.get("mensagem", ""))
        linhas.append(f"• [{prefixo}] {mensagem}")

    if not linhas:
        return "Nenhum insight automático identificado para o período analisado."

    return "\n".join(linhas)


def gerar_recomendacoes_executivas(
    total_chamados,
    reabertos,
    backlog_atual,
    sla_deadline,
    sla_inicializacao,
    insights,
    comparativo_mensal,
    df_filtrado,
):
    recomendacoes = []

    if total_chamados == 0:
        return [
            "Carregar e validar a base de chamados para habilitar "
            "recomendações orientadas ao Conselho.",
        ]

    if sla_deadline is not None and sla_deadline < 90:
        recomendacoes.append(
            "Implementar plano de ação para elevar o cumprimento do SLA "
            "de Deadline, com revisão de prazos e gargalos operacionais."
        )

    if sla_inicializacao is not None and sla_inicializacao < 90:
        recomendacoes.append(
            "Reforçar o tempo de primeira resposta com escalonamento "
            "automático e monitoramento diário do SLA de Inicialização."
        )

    if total_chamados > 0 and (reabertos / total_chamados) * 100 > 10:
        recomendacoes.append(
            "Investigar causas de reabertura e padronizar critérios de "
            "encerramento para reduzir retrabalho."
        )

    if comparativo_mensal is not None:
        var_chamados_pct = comparativo_mensal.get("var_chamados_pct")
        var_backlog = comparativo_mensal.get("var_backlog")

        if var_chamados_pct is not None and var_chamados_pct > 15:
            recomendacoes.append(
                "Avaliar capacidade da equipe e priorização de demandas "
                "diante do crescimento acima de 15% no volume de chamados."
            )

        if var_backlog is not None and var_backlog > 0:
            recomendacoes.append(
                "Definir metas de redução de backlog com acompanhamento "
                "semanal até estabilização do estoque de chamados."
            )

    for insight in insights:
        if insight.get("tipo") != "warning":
            continue

        mensagem = limpar_markdown(insight.get("mensagem", ""))

        if "reaberto" in mensagem.lower() and not any(
            "reabertura" in r.lower() for r in recomendacoes
        ):
            recomendacoes.append(
                "Tratar reaberturas como indicador de qualidade e revisar "
                "o fluxo de validação antes do fechamento."
            )
        elif "backlog" in mensagem.lower() and not any(
            "backlog" in r.lower() for r in recomendacoes
        ):
            recomendacoes.append(
                "Priorizar chamados mais antigos do backlog e estabelecer "
                "rituais de triagem com as áreas demandantes."
            )
        elif "sla" in mensagem.lower() and len(recomendacoes) < 6:
            recomendacoes.append(
                "Realizar comitê mensal de SLAs com indicadores por "
                "departamento e plano corretivo."
            )

    categoria, _ = obter_maior_volume(df_filtrado, "categoria")
    if categoria is not None:
        recomendacoes.append(
            f"Concentrar esforços de melhoria na categoria '{categoria}', "
            "com foco em automação, FAQ e padronização de atendimento."
        )

    if backlog_atual > 0 and len(recomendacoes) < 6:
        recomendacoes.append(
            f"Manter governança do backlog atual ({backlog_atual:,} chamados) "
            "com metas de aging e reporte executivo ao Conselho."
        )

    padrao = [
        "Consolidar visão mensal dos indicadores para apoiar decisões "
        "de investimento, priorização e alocação de recursos em TI.",
        "Promover alinhamento com as áreas de maior demanda para reduzir "
        "volume recorrente e elevar satisfação dos usuários.",
    ]

    for item in padrao:
        if len(recomendacoes) >= 6:
            break
        if item not in recomendacoes:
            recomendacoes.append(item)

    return recomendacoes[:6]


def _novo_slide_executivo(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    aplicar_fundo_slide(slide)
    return slide


def gerar_apresentacao_executiva(
    data_inicio,
    data_fim,
    total_chamados,
    finalizados,
    abertos,
    backlog_atual,
    reabertos,
    sla_deadline,
    sla_inicializacao,
    backlog_diario,
    backlog_mensal,
    comparativo_mensal,
    insights,
    dados_insuficientes,
    resumo_executivo,
    df_filtrado,
    metricas_mensais,
    performance_sla,
    backlog_executivo,
):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total_slides = 11

    periodo_txt = formatar_periodo_analisado(data_inicio, data_fim)
    gerado_em = datetime.now().strftime("%d/%m/%Y %H:%M")

    sla_deadline_txt = (
        f"{sla_deadline:.1f}%"
        if sla_deadline is not None
        else "—"
    )
    sla_inicializacao_txt = (
        f"{sla_inicializacao:.1f}%"
        if sla_inicializacao is not None
        else "—"
    )

    slide_capa = _novo_slide_executivo(prs)
    faixa_capa = slide_capa.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(13.333),
        Inches(7.5),
    )
    faixa_capa.fill.solid()
    faixa_capa.fill.fore_color.rgb = COR_PRIMARIA
    faixa_capa.line.fill.background()

    slide_capa.shapes.add_picture(
        criar_logotipo_ti_generico(),
        Inches(0.9),
        Inches(1.6),
        width=Inches(1.8),
    )

    adicionar_caixa_texto(
        slide_capa,
        "Painel Conselho TI",
        Inches(3.0),
        Inches(1.9),
        Inches(9.5),
        Inches(0.9),
        tamanho=38,
        negrito=True,
        cor=RGBColor(255, 255, 255),
    )
    adicionar_caixa_texto(
        slide_capa,
        "Apresentação Executiva para o Conselho",
        Inches(3.0),
        Inches(2.75),
        Inches(9.5),
        Inches(0.5),
        tamanho=18,
        cor=RGBColor(200, 220, 240),
    )
    adicionar_caixa_texto(
        slide_capa,
        f"Período analisado: {periodo_txt}\nGerado em: {gerado_em}",
        Inches(3.0),
        Inches(3.6),
        Inches(9),
        Inches(1.2),
        tamanho=15,
        cor=RGBColor(255, 255, 255),
    )
    adicionar_rodape_executivo(slide_capa, 1, total_slides)

    slide_kpis = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_kpis,
        "KPIs Principais",
        "Indicadores consolidados do período filtrado",
    )
    adicionar_rodape_executivo(slide_kpis, 2, total_slides)

    cards = [
        ("Total de chamados", f"{total_chamados:,}"),
        ("Finalizados", f"{finalizados:,}"),
        ("Em aberto", f"{abertos:,}"),
        ("Backlog atual", f"{backlog_atual:,}"),
        ("Reabertos", f"{reabertos:,}"),
        ("SLA Deadline", sla_deadline_txt),
        ("SLA Inicialização", sla_inicializacao_txt),
    ]

    card_w = Inches(2.95)
    card_h = Inches(1.55)
    margem_x = Inches(0.55)
    margem_y = Inches(1.35)
    gap_x = Inches(0.25)
    gap_y = Inches(0.25)

    for indice, (rotulo, valor) in enumerate(cards):
        linha = indice // 4
        coluna = indice % 4
        left = margem_x + coluna * (card_w + gap_x)
        top = margem_y + linha * (card_h + gap_y)
        if indice == 6:
            left = margem_x + 1 * (card_w + gap_x)
            top = margem_y + 2 * (card_h + gap_y)
        adicionar_card_kpi(slide_kpis, rotulo, valor, left, top, card_w, card_h)

    slide_backlog = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_backlog,
        "Evolução do Backlog",
        "Série diária e visão mensal do estoque de chamados",
    )
    adicionar_rodape_executivo(slide_backlog, 3, total_slides)

    slide_backlog.shapes.add_picture(
        criar_grafico_backlog_diario(backlog_diario),
        Inches(0.45),
        Inches(1.25),
        width=Inches(6.15),
    )
    slide_backlog.shapes.add_picture(
        criar_grafico_backlog_mensal(backlog_mensal),
        Inches(6.75),
        Inches(1.25),
        width=Inches(6.15),
    )

    slide_operacao = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_operacao,
        "Visão da Operação",
        "Distribuição de chamados por dimensões operacionais",
    )
    adicionar_rodape_executivo(slide_operacao, 4, total_slides)

    graficos_operacao = [
        (df_filtrado, "departamento", "Chamados por departamento"),
        (df_filtrado, "categoria", "Chamados por categoria"),
        (df_filtrado, "atendente", "Chamados por atendente"),
        (df_filtrado, "setor_solicitante", "Chamados por setor solicitante"),
    ]
    posicoes = [
        (Inches(0.45), Inches(1.25)),
        (Inches(6.75), Inches(1.25)),
        (Inches(0.45), Inches(4.05)),
        (Inches(6.75), Inches(4.05)),
    ]

    for (dados, coluna, titulo), (left, top) in zip(graficos_operacao, posicoes):
        slide_operacao.shapes.add_picture(
            criar_grafico_barras_operacao(dados, coluna, titulo),
            left,
            top,
            width=Inches(6.15),
        )

    slide_comparativo = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(slide_comparativo, "Comparativo Mensal")
    adicionar_rodape_executivo(slide_comparativo, 5, total_slides)

    if comparativo_mensal is None:
        adicionar_caixa_texto(
            slide_comparativo,
            "Não há dados suficientes para o comparativo mensal.",
            Inches(0.8),
            Inches(1.8),
            Inches(11.5),
            Inches(1.0),
            tamanho=16,
        )
    else:
        tabela_comparativo = montar_tabela_comparativa(
            comparativo_mensal["ind_mes_atual"],
            comparativo_mensal["ind_mes_anterior"],
            comparativo_mensal["label_mes_atual"],
            comparativo_mensal["label_mes_anterior"],
        )
        tabela_comparativo = tabela_comparativo.rename(
            columns={
                "Mês Atual": comparativo_mensal["label_mes_atual"],
                "Mês Anterior": comparativo_mensal["label_mes_anterior"],
            }
        )
        adicionar_caixa_texto(
            slide_comparativo,
            (
                f"Mês atual: {comparativo_mensal['label_mes_atual']} · "
                f"Mês anterior: {comparativo_mensal['label_mes_anterior']}"
            ),
            Inches(0.55),
            Inches(1.15),
            Inches(12),
            Inches(0.35),
            tamanho=11,
            cor=COR_SUBTITULO,
        )
        adicionar_tabela_dataframe(
            slide_comparativo,
            tabela_comparativo,
            Inches(0.4),
            Inches(1.55),
            Inches(12.5),
            Inches(5.2),
        )

    slide_insights = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_insights,
        "Insights Automáticos",
        "Pontos de atenção identificados para o Conselho",
    )
    adicionar_rodape_executivo(slide_insights, 6, total_slides)

    if dados_insuficientes:
        adicionar_caixa_texto(
            slide_insights,
            "Não há dados suficientes para gerar todos os insights automáticos.",
            Inches(0.7),
            Inches(1.25),
            Inches(12),
            Inches(0.5),
            tamanho=12,
            cor=COR_SUBTITULO,
        )

    texto_insights = formatar_insights_pptx(insights, False)
    adicionar_caixa_texto(
        slide_insights,
        texto_insights,
        Inches(0.7),
        Inches(1.7 if dados_insuficientes else 1.25),
        Inches(12),
        Inches(5.5),
        tamanho=12,
    )

    slide_resumo = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_resumo,
        "Resumo Executivo",
        "Síntese para apresentação ao Conselho",
    )
    adicionar_rodape_executivo(slide_resumo, 7, total_slides)
    adicionar_caixa_texto(
        slide_resumo,
        limpar_markdown(resumo_executivo),
        Inches(0.7),
        Inches(1.2),
        Inches(12),
        Inches(5.8),
        tamanho=12,
    )

    recomendacoes = gerar_recomendacoes_executivas(
        total_chamados,
        reabertos,
        backlog_atual,
        sla_deadline,
        sla_inicializacao,
        insights,
        comparativo_mensal,
        df_filtrado,
    )
    texto_recomendacoes = "\n".join(
        f"{indice + 1}. {texto}"
        for indice, texto in enumerate(recomendacoes)
    )

    slide_visao = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_visao,
        "Visão Executiva",
        "Volume mensal de chamados, urgências e tendência"
    )
    adicionar_rodape_executivo(slide_visao, 8, total_slides)

    if metricas_mensais is not None and not metricas_mensais.empty:
        tabela_visao = metricas_mensais[
            ["mes", "chamados_abertos", "chamados_urgentes"]
        ].copy()

        adicionar_tabela_dataframe(
            slide_visao,
            tabela_visao,
            Inches(0.7),
            Inches(1.35),
            Inches(12),
            Inches(4.8),
        )
    else:
        adicionar_caixa_texto(
            slide_visao,
            "Sem dados disponíveis para visão executiva.",
            Inches(0.8),
            Inches(1.8),
            Inches(11),
            Inches(1),
            tamanho=16,
        )

    slide_performance = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_performance,
        "Performance e SLA",
        "Eficiência mensal e cumprimento de prazo"
    )
    adicionar_rodape_executivo(slide_performance, 9, total_slides)

    if performance_sla is not None and not performance_sla.empty:
        tabela_performance = performance_sla[
            ["mes", "total_chamados", "finalizados", "eficiencia", "sla_deadline"]
        ].copy()

        tabela_performance["eficiencia"] = tabela_performance["eficiencia"].map(lambda x: f"{x:.2f}%")
        tabela_performance["sla_deadline"] = tabela_performance["sla_deadline"].map(lambda x: f"{x:.2f}%")

        adicionar_tabela_dataframe(
            slide_performance,
            tabela_performance,
            Inches(0.7),
            Inches(1.35),
            Inches(12),
            Inches(4.8),
        )
    else:
        adicionar_caixa_texto(
            slide_performance,
            "Sem dados disponíveis para performance e SLA.",
            Inches(0.8),
            Inches(1.8),
            Inches(11),
            Inches(1),
            tamanho=16,
        )

    slide_backlog_exec = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_backlog_exec,
        "Backlog Mensal",
        "Backlog inicial, acúmulo atual e baixas do estoque herdado"
    )
    adicionar_rodape_executivo(slide_backlog_exec, 10, total_slides)

    if backlog_executivo:
        cards_backlog = [
            ("Total abertos", f"{backlog_executivo['total_abertos']:,}"),
            ("Backlog inicial", f"{backlog_executivo['backlog_inicial']:,}"),
            ("Base operacional", f"{backlog_executivo['base_operacional']:,}"),
            ("Fechados", f"{backlog_executivo['finalizados']:,}"),
            ("Em atendimento", f"{backlog_executivo['em_atendimento']:,}"),
            ("Em pausa", f"{backlog_executivo['em_pausa']:,}"),
        ]

        for indice, (rotulo, valor) in enumerate(cards_backlog):
            linha = indice // 3
            coluna = indice % 3
            adicionar_card_kpi(
                slide_backlog_exec,
                rotulo,
                valor,
                Inches(0.7) + coluna * Inches(4.1),
                Inches(1.35) + linha * Inches(1.75),
                Inches(3.7),
                Inches(1.35),
            )

    if (
        metricas_mensais is not None
        and not metricas_mensais.empty
        and "backlog_herdado_finalizado" in metricas_mensais.columns
    ):
        tabela_backlog = metricas_mensais[
            ["mes", "backlog_inicial", "backlog_herdado_finalizado", "backlog_final"]
        ].copy()

        adicionar_tabela_dataframe(
            slide_backlog_exec,
            tabela_backlog,
            Inches(0.7),
            Inches(4.7),
            Inches(12),
            Inches(1.6),
        )
    slide_recomendacoes = _novo_slide_executivo(prs)
    adicionar_faixa_cabecalho(
        slide_recomendacoes,
        "Recomendações ao Conselho",
        "Próximos passos sugeridos com base nos indicadores",
    )
    adicionar_rodape_executivo(slide_recomendacoes, 11, total_slides)
    adicionar_caixa_texto(
        slide_recomendacoes,
        texto_recomendacoes,
        Inches(0.75),
        Inches(1.25),
        Inches(12),
        Inches(5.8),
        tamanho=14,
    )

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def calcular_metricas_mensais_operacionais(df, data_inicio, data_fim):
    if "data_de_criacao" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df = df[df["data_de_criacao"].notna()]

    if df.empty:
        return pd.DataFrame()

    mes_inicio = pd.Timestamp(data_inicio).to_period("M")
    mes_fim = pd.Timestamp(data_fim).to_period("M")

    meses = pd.period_range(
        start=mes_inicio,
        end=mes_fim,
        freq="M"
    )

    linhas = []

    for mes in meses:
        inicio_mes = mes.to_timestamp(how="start")
        fim_mes = mes.to_timestamp(how="end")

        chamados_abertos_mes = df[
            df["data_de_criacao"].dt.to_period("M") == mes
        ]

        chamados_fechados_mes = df[
            df["data_de_finalizacao"].notna()
            & (df["data_de_finalizacao"].dt.to_period("M") == mes)
        ]

        backlog_inicial = df[
            (df["data_de_criacao"] < inicio_mes)
            & (
                df["data_de_finalizacao"].isna()
                | (df["data_de_finalizacao"] >= inicio_mes)
            )
        ]

        backlog_final = df[
            (df["data_de_criacao"] <= fim_mes)
            & (
                df["data_de_finalizacao"].isna()
                | (df["data_de_finalizacao"] > fim_mes)
            )
        ]

        backlog_herdado_finalizado = df[
            (df["data_de_criacao"] < inicio_mes)
            & df["data_de_finalizacao"].notna()
            & (df["data_de_finalizacao"] >= inicio_mes)
            & (df["data_de_finalizacao"] <= fim_mes)
        ]

        urgentes = 0
        if "prioridade" in df.columns:
            urgentes = chamados_abertos_mes[
                chamados_abertos_mes["prioridade"]
                .astype(str)
                .str.lower()
                .str.contains("urgent", na=False)
            ].shape[0]

        linhas.append({
            "mes": str(mes),
            "chamados_abertos": len(chamados_abertos_mes),
            "chamados_fechados": len(chamados_fechados_mes),
            "chamados_urgentes": urgentes,
            "backlog_inicial": len(backlog_inicial),
            "abertos_mais_backlog_inicial": len(chamados_abertos_mes) + len(backlog_inicial),
            "backlog_final": len(backlog_final),
            "backlog_herdado_finalizado": len(backlog_herdado_finalizado),
            "backlog_herdado_remanescente": max(
                len(backlog_inicial) - len(backlog_herdado_finalizado),
                0
            ),
        })

    return pd.DataFrame(linhas)
def grafico_volume_executivo(metricas_mensais):
    if metricas_mensais.empty:
        return None

    fig = px.bar(
        metricas_mensais,
        x="mes",
        y=["chamados_abertos", "chamados_urgentes"],
        barmode="group",
        text_auto=True,
        title="Chamados por mês"
    )

    fig.add_scatter(
        x=metricas_mensais["mes"],
        y=metricas_mensais["chamados_abertos"].rolling(2, min_periods=1).mean(),
        mode="lines+markers",
        name="Tendência"
    )

    fig.update_layout(
        height=480,
        xaxis_title="Mês",
        yaxis_title="Quantidade",
        legend_title="Indicadores"
    )

    return fig    
def calcular_crescimento_ano_anterior(df_sidebar, data_inicio, data_fim):
    if data_inicio is None or data_fim is None:
        return None

    if "data_de_criacao" not in df_sidebar.columns:
        return None

    inicio_atual = pd.Timestamp(data_inicio).date()
    fim_atual = pd.Timestamp(data_fim).date()

    inicio_anterior = pd.Timestamp(data_inicio) - pd.DateOffset(years=1)
    fim_anterior = pd.Timestamp(data_fim) - pd.DateOffset(years=1)

    inicio_anterior = inicio_anterior.date()
    fim_anterior = fim_anterior.date()

    periodo_atual = df_sidebar[
        df_sidebar["data_de_criacao"].notna()
        & (df_sidebar["data_de_criacao"].dt.date >= inicio_atual)
        & (df_sidebar["data_de_criacao"].dt.date <= fim_atual)
    ]

    periodo_anterior = df_sidebar[
        df_sidebar["data_de_criacao"].notna()
        & (df_sidebar["data_de_criacao"].dt.date >= inicio_anterior)
        & (df_sidebar["data_de_criacao"].dt.date <= fim_anterior)
    ]

    atual = len(periodo_atual)
    anterior = len(periodo_anterior)

    if anterior == 0:
        return {
            "atual": atual,
            "anterior": anterior,
            "variacao_pct": None
        }

    variacao_pct = ((atual - anterior) / anterior) * 100

    return {
        "atual": atual,
        "anterior": anterior,
        "variacao_pct": variacao_pct
    }
def calcular_performance_sla_mensal(df, data_inicio, data_fim):
    if "data_de_criacao" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df = df[df["data_de_criacao"].notna()]

    if df.empty:
        return pd.DataFrame()

    data_inicio_ref = pd.Timestamp(data_inicio)
    data_fim_ref = pd.Timestamp(data_fim)

    mes_inicio = data_inicio_ref.to_period("M")
    mes_fim = data_fim_ref.to_period("M")

    meses = pd.period_range(
        start=mes_inicio,
        end=mes_fim,
        freq="M"
    )

    if len(meses) == 0:
        return pd.DataFrame()

    linhas = []

    for mes in meses:
        chamados_mes = df[
            df["data_de_criacao"].dt.to_period("M") == mes
        ]

        total_mes = len(chamados_mes)

        finalizados_mes = 0
        if "data_de_finalizacao" in chamados_mes.columns:
            finalizados_mes = chamados_mes[
                chamados_mes["data_de_finalizacao"].notna()
                & (
                    chamados_mes["data_de_finalizacao"].dt.to_period("M")
                    == chamados_mes["data_de_criacao"].dt.to_period("M")
                )
            ].shape[0]

        eficiencia = (
            (finalizados_mes / total_mes) * 100
            if total_mes > 0
            else 0
        )

        sla_mes = 0

        if (
            "sla_de_deadline_cumprido" in chamados_mes.columns
            and "data_de_finalizacao" in chamados_mes.columns
            and total_mes > 0
        ):
            qtd_sla_ok = chamados_mes[
                chamados_mes["data_de_finalizacao"].notna()
                & chamados_mes["sla_de_deadline_cumprido"]
                    .astype(str)
                    .str.lower()
                    .str.strip()
                    .isin(["sim", "s"])
            ].shape[0]

            sla_mes = (qtd_sla_ok / total_mes) * 100

        fora_mes_abertura = 0
        if "data_de_finalizacao" in chamados_mes.columns:
            fora_mes_abertura = chamados_mes[
                chamados_mes["data_de_finalizacao"].notna()
                & (
                    chamados_mes["data_de_finalizacao"].dt.to_period("M")
                    != chamados_mes["data_de_criacao"].dt.to_period("M")
                )
            ].shape[0]

        pct_fora_mes = (
            (fora_mes_abertura / total_mes) * 100
            if total_mes > 0
            else 0
        )

        linhas.append({
            "mes": str(mes),
            "total_chamados": total_mes,
            "finalizados": finalizados_mes,
            "eficiencia": eficiencia,
            "sla_deadline": sla_mes,
            "fora_mes_abertura": fora_mes_abertura,
            "pct_fora_mes_abertura": pct_fora_mes,
        })

    return pd.DataFrame(linhas)

def calcular_backlog_executivo(df, data_inicio, data_fim):
    if "data_de_criacao" not in df.columns:
        return {}

    df = df.copy()
    df = df[df["data_de_criacao"].notna()]

    inicio = pd.Timestamp(data_inicio)
    fim = pd.Timestamp(data_fim)

    backlog_inicial_df = df[
        (df["data_de_criacao"] < inicio)
        & (
            df["data_de_finalizacao"].isna()
            | (df["data_de_finalizacao"] >= inicio)
        )
    ]

    chamados_periodo = df[
        (df["data_de_criacao"].dt.date >= inicio.date())
        & (df["data_de_criacao"].dt.date <= fim.date())
    ]

    finalizados_periodo = df[
        df["data_de_finalizacao"].notna()
        & (df["data_de_finalizacao"].dt.date >= inicio.date())
        & (df["data_de_finalizacao"].dt.date <= fim.date())
    ]

    em_atendimento = df[
        (df["data_de_criacao"] <= fim)
        & (
            df["data_de_finalizacao"].isna()
            | (df["data_de_finalizacao"] > fim)
        )
    ]

    em_pausa = pd.DataFrame()

    if "ultima_situacao" in em_atendimento.columns:
        em_pausa = em_atendimento[
            em_atendimento["ultima_situacao"]
            .astype(str)
            .str.lower()
            .str.contains("pausa|espera|aguard", na=False)
        ]

    total = len(chamados_periodo)
    backlog_inicial = len(backlog_inicial_df)
    base_operacional = total + backlog_inicial

    finalizados = len(finalizados_periodo)
    
    em_pausa = pd.DataFrame()

    if "ultima_situacao" in em_atendimento.columns:
        em_pausa = em_atendimento[
            em_atendimento["ultima_situacao"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("em pausa")
        ]

    pausa = len(em_pausa)
    atendimento = len(em_atendimento) - pausa
    

    return {
        "backlog_inicial": backlog_inicial,
        "total_abertos": total,
        "base_operacional": base_operacional,
        "finalizados": finalizados,
        "em_atendimento": atendimento,
        "em_pausa": pausa,
        "pct_finalizados": (finalizados / base_operacional * 100) if base_operacional else 0,
        "pct_atendimento": (atendimento / base_operacional * 100) if base_operacional else 0,
        "pct_pausa": (pausa / base_operacional * 100) if base_operacional else 0,
    }
st.title("Painel Conselho TI")
st.caption("Análise mensal de chamados para apresentação ao Conselho")

uploaded_file = st.file_uploader(
    "Faça upload da planilha de chamados",
    type=["xlsx"]
)

if uploaded_file:
    try:
        df = pd.read_excel(
            uploaded_file,
            header=5,
            engine="openpyxl"
        )

        df.columns = [normalizar_coluna(col) for col in df.columns]
        df = preparar_dados(df)

        df_filtrado, data_inicio, data_fim, df_sidebar = aplicar_filtros(df)

        total_chamados = len(df_filtrado)

        finalizados = df_filtrado.apply(
            chamado_finalizado,
            axis=1
        ).sum()

        abertos = total_chamados - finalizados

        backlog_atual = 0
        if "data_de_finalizacao" in df_sidebar.columns:
            backlog_atual = df_sidebar[
                df_sidebar["data_de_finalizacao"].isna()
            ].shape[0]

        reabertos = 0
        if "reaberto" in df_filtrado.columns:
            reabertos = df_filtrado["reaberto"].apply(sim).sum()

        sla_deadline = None
        if "sla_de_deadline_cumprido" in df_filtrado.columns and len(df_filtrado) > 0:
            sla_deadline = df_filtrado["sla_de_deadline_cumprido"].apply(sim).mean() * 100

        sla_inicializacao = None
        if "sla_de_inicializacao_cumprido" in df_filtrado.columns and len(df_filtrado) > 0:
            sla_inicializacao = df_filtrado["sla_de_inicializacao_cumprido"].apply(sim).mean() * 100

        if data_inicio is None or data_fim is None:
            data_inicio = df["data_de_criacao"].min().date()
            data_fim = df["data_de_criacao"].max().date()

        backlog_diario = calcular_backlog_diario(
            df_sidebar,
            data_inicio,
            data_fim
        )

        backlog_mensal = calcular_backlog_mensal(
            df_sidebar,
            data_inicio,
            data_fim
        )
        metricas_mensais = calcular_metricas_mensais_operacionais(
            df_sidebar,
            data_inicio,
            data_fim
        )
        performance_sla = calcular_performance_sla_mensal(
            df_sidebar,
            data_inicio,
            data_fim
        )
        backlog_executivo = calcular_backlog_executivo(
            df_sidebar,
            data_inicio,
            data_fim
        )
        st.caption(
            f"Auditoria Performance: período usado = {data_inicio} até {data_fim}"
        )
        crescimento_ano_anterior = calcular_crescimento_ano_anterior(
            df_sidebar,
            data_inicio,
            data_fim
        )
        aging_df = calcular_aging(df_sidebar)
        comparativo_mensal = calcular_dados_comparativo_mensal(
            df_sidebar,
            data_fim,
        )

        resumo_executivo = gerar_resumo_executivo(
            total_chamados,
            finalizados,
            abertos,
            backlog_atual,
            reabertos,
            sla_deadline,
            sla_inicializacao,
            df_filtrado,
        )

        insights_automaticos, dados_insuficientes_insights = (
            gerar_insights_automaticos(
                total_chamados,
                reabertos,
                backlog_atual,
                sla_deadline,
                sla_inicializacao,
                df_filtrado,
                comparativo_mensal,
            )
        )

        st.success("Planilha carregada com sucesso!")

        col_export, _ = st.columns([1, 3])

        with col_export:
            if st.button(
                "Gerar Apresentação Executiva",
                key="btn_gerar_apresentacao",
            ):
                with st.spinner("Gerando apresentação executiva..."):
                    try:
                        pptx_bytes = gerar_apresentacao_executiva(
                            data_inicio,
                            data_fim,
                            total_chamados,
                            finalizados,
                            abertos,
                            backlog_atual,
                            reabertos,
                            sla_deadline,
                            sla_inicializacao,
                            backlog_diario,
                            backlog_mensal,
                            comparativo_mensal,
                            insights_automaticos,
                            dados_insuficientes_insights,
                            resumo_executivo,
                            df_filtrado,
                            metricas_mensais,
                            performance_sla,
                            backlog_executivo,
                        )
                        st.session_state["pptx_executivo"] = pptx_bytes
                        st.session_state["pptx_executivo_nome"] = (
                            f"Painel_Conselho_TI_"
                            f"{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
                        )
                    except Exception as erro_pptx:
                        st.session_state.pop("pptx_executivo", None)
                        st.error(
                            f"Erro ao gerar apresentação: {erro_pptx}"
                        )

        if st.session_state.get("pptx_executivo"):
            st.download_button(
                label="Baixar Apresentação Executiva (.pptx)",
                data=st.session_state["pptx_executivo"],
                file_name=st.session_state.get(
                    "pptx_executivo_nome",
                    "Painel_Conselho_TI.pptx",
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                key="download_apresentacao_executiva",
            )

        tab_visao, tab_performance, tab_backlog_exec, tab_dashboard, tab_comparativo, tab_backlog,  tab_aging, tab_operacao, tab_resumo, tab_dados = st.tabs(
            [
                "Visão Executiva",
                "Performance e SLA",
                "Backlog Mensal",
                "Dashboard",
                "Comparativo Mensal",
                "Backlog",
                "Aging",
                "Operação",
                "Resumo Executivo",
                "Dados"
            ]
        )
        with tab_visao:
            st.caption("Métricas")
            st.title("Suporte Interno - Volume de Chamados")

            st.divider()

            coluna_grafico, coluna_cards = st.columns([2.2, 1])

            with coluna_grafico:
                fig_executivo = grafico_volume_executivo(metricas_mensais)

                if fig_executivo:
                    st.plotly_chart(
                        fig_executivo,
                        use_container_width=True,
                        key="grafico_visao_executiva"
                    )
                else:
                    st.info("Não há dados suficientes para montar a visão executiva.")

            with coluna_cards:
                st.metric(
                    "Total do período",
                    f"{total_chamados:,}",
                    help="Total de chamados registrados no período filtrado"
                )

                st.metric(
                    "Finalizados",
                    f"{finalizados:,}"
                )

                st.metric(
                    "Em aberto",
                    f"{abertos:,}"
                )

                st.metric(
                    "Backlog atual",
                    f"{backlog_atual:,}"
                )

            st.divider()

            st.subheader("Status dos Chamados")

            percentual_finalizados = (finalizados / total_chamados * 100) if total_chamados else 0
            percentual_abertos = (abertos / total_chamados * 100) if total_chamados else 0

            s1, s2, s3 = st.columns(3)

            s1.metric("Finalizados", f"{percentual_finalizados:.2f}%")
            s2.metric("Em andamento", f"{percentual_abertos:.2f}%")
            s3.metric("Em pausa / espera", "0.00%")

            st.divider()

            st.subheader("Leitura Executiva")

            c1, c2, c3, c4 = st.columns(4)

            variacao_chamados_txt = "Sem comparação"

            if (
                crescimento_ano_anterior is not None
                and crescimento_ano_anterior["variacao_pct"] is not None
            ):
                variacao_chamados_txt = (
                    f"{crescimento_ano_anterior['variacao_pct']:+.2f}% "
                    "vs mesmo período do ano anterior"
            )
            if crescimento_ano_anterior is not None:
                st.caption(
                    f"Base crescimento: "
                    f"{crescimento_ano_anterior['atual']} chamados no período atual "
                    f"vs {crescimento_ano_anterior['anterior']} no mesmo período anterior."
                )

            c1.metric(
                "Crescimento",
                variacao_chamados_txt
            )

            status_sla = "Acompanhar"
            if sla_deadline is not None:
                if sla_deadline >= 95:
                    status_sla = "Excelente"
                elif sla_deadline >= 90:
                    status_sla = "Estável"
                else:
                    status_sla = "Atenção"

            c2.metric(
                "SLA",
                status_sla
            )

            media_urgentes = 0
            if not metricas_mensais.empty:
                media_urgentes = metricas_mensais["chamados_urgentes"].mean()

            c3.metric(
                "Urgências",
                f"Média {media_urgentes:.1f}/mês"
            )

            maturidade = "Em evolução"
            if backlog_atual <= 10 and sla_deadline is not None and sla_deadline >= 90:
                maturidade = "Equilibrado"
            elif backlog_atual > 30:
                maturidade = "Atenção"

            c4.metric(
                "Maturidade",
                maturidade
            )

        with tab_performance:
            st.caption("Performance")
            st.title("Suporte Interno - Performance e SLA")

            if performance_sla.empty:
                st.info("Não há dados suficientes para calcular performance e SLA.")
            else:
                resolvidos_media = performance_sla["eficiencia"].mean()
                sla_media = performance_sla["sla_deadline"].mean()
                fora_mes_media = performance_sla["pct_fora_mes_abertura"].mean()

                k1, k2, k3 = st.columns(3)

                k1.metric(
                    "Resolvidos no mês",
                    f"{resolvidos_media:.2f}%"
                )

                k2.metric(
                    "Cumprimento de prazo",
                    f"{sla_media:.2f}%"
                )

                k3.metric(
                    "Fora do mês de abertura",
                    f"{fora_mes_media:.2f}%"
                )

                st.divider()

                g1, g2 = st.columns(2)

                fig_eficiencia = px.line(
                    performance_sla,
                    x="mes",
                    y="eficiencia",
                    markers=True,
                    text="eficiencia",
                    title="Eficiência mensal"
                )

                fig_eficiencia.update_traces(
                    texttemplate="%{text:.2f}%",
                    textposition="top center"
                )

                fig_eficiencia.update_layout(
                    yaxis_title="Percentual",
                    xaxis_title="Mês"
                )

                g1.plotly_chart(
                    fig_eficiencia,
                    use_container_width=True,
                    key="grafico_eficiencia_mensal"
                )

                fig_sla = px.bar(
                    performance_sla,
                    x="mes",
                    y="sla_deadline",
                    text="sla_deadline",
                    title="Cumprimento de SLA"
                )

                fig_sla.update_traces(
                    texttemplate="%{text:.2f}%",
                    textposition="outside"
                )

                fig_sla.update_layout(
                    yaxis_title="Percentual",
                    xaxis_title="Mês"
                )

                g2.plotly_chart(
                    fig_sla,
                    use_container_width=True,
                    key="grafico_sla_mensal"
                )

                st.divider()

                c1, c2 = st.columns(2)

                melhor_mes = performance_sla.sort_values(
                    "eficiencia",
                    ascending=False
                ).iloc[0]

                c1.info(
                    f"⭐ Pico de eficiência em **{melhor_mes['mes']}**, "
                    f"com {melhor_mes['eficiencia']:.2f}% dos chamados resolvidos."
                )

                if sla_media >= 95:
                    c2.success(
                        f"✅ Cumprimento médio de SLA acima de 95% "
                        f"({sla_media:.2f}%), indicando estabilidade operacional."
                    )
                elif sla_media >= 90:
                    c2.info(
                        f"📌 Cumprimento médio de SLA em nível estável "
                        f"({sla_media:.2f}%)."
                    )
                else:
                    c2.warning(
                        f"⚠️ Cumprimento médio de SLA abaixo de 90% "
                        f"({sla_media:.2f}%). Requer atenção."
                    )

                st.subheader("Tabela de performance mensal")

                st.dataframe(
                    performance_sla,
                    use_container_width=True,
                    hide_index=True
                )
        with tab_backlog_exec:
            st.caption("Operações")
            st.title("Backlog Mensal - Chamados")

            if not backlog_executivo:
                st.info("Não há dados suficientes para calcular o backlog mensal executivo.")
            else:
                c1, c2, c3 = st.columns(3)

                c1.metric(
                    "Total de chamados abertos",
                    f"{backlog_executivo['total_abertos']:,}"
                )

                c2.metric(
                    "Backlog inicial",
                    f"{backlog_executivo['backlog_inicial']:,}",
                    help="Chamados herdados do período anterior ainda em aberto no início do período."
                )

                c3.metric(
                    "Acúmulo atual",
                    f"{backlog_atual:,}"
                )

                st.divider()

                s1, s2, s3 = st.columns(3)

                s1.metric(
                    "Fechados",
                    f"{backlog_executivo['finalizados']:,}",
                    f"{backlog_executivo['pct_finalizados']:.2f}%"
                )

                s2.metric(
                    "Em atendimento",
                    f"{backlog_executivo['em_atendimento']:,}",
                    f"{backlog_executivo['pct_atendimento']:.2f}%"
                )

                s3.metric(
                    "Em pausa",
                    f"{backlog_executivo['em_pausa']:,}",
                    f"{backlog_executivo['pct_pausa']:.2f}%"
                )

                st.divider()

                st.subheader("Backlogs finalizados por mês")

                if (
                    not metricas_mensais.empty
                    and "backlog_herdado_finalizado" in metricas_mensais.columns
                ):
                    fig_backlogs_finalizados = px.bar(
                        metricas_mensais,
                        x="mes",
                        y="backlog_herdado_finalizado",
                        text="backlog_herdado_finalizado",
                        title="Backlogs finalizados por mês"
                    )

                    fig_backlogs_finalizados.update_traces(
                        textposition="outside"
                    )

                    fig_backlogs_finalizados.update_layout(
                        xaxis_title="Mês",
                        yaxis_title="Quantidade"
                    )

                    st.plotly_chart(
                        fig_backlogs_finalizados,
                        use_container_width=True,
                        key="grafico_backlogs_finalizados_exec"
                    )
                else:
                    st.info("Sem dados para backlogs finalizados por mês.")

                st.divider()

                st.warning(
                    "Backlog representa a quantidade de chamados que ainda não foram finalizados, "
                    "ou seja, demandas pendentes que estão em atendimento ou aguardando resolução. "
                    "Esse indicador ajuda a medir o volume de trabalho acumulado e o nível de controle da operação."
                )

                st.subheader("Distribuição mensal do backlog")

                if not backlog_mensal.empty:
                    st.line_chart(
                        backlog_mensal.set_index("mes")["backlog"]
                    )
                else:
                    st.info("Sem dados para distribuição mensal do backlog.")
        with tab_dashboard:
            st.subheader("Indicadores principais")

            col1, col2, col3, col4 = st.columns(4)

            col1.metric("Total de chamados", f"{total_chamados:,}")
            col2.metric("Finalizados", f"{finalizados:,}")
            col3.metric("Em aberto", f"{abertos:,}")
            col4.metric("Backlog atual", f"{backlog_atual:,}")

            col5, col6, col7 = st.columns(3)

            col5.metric("Reabertos", f"{reabertos:,}")

            if sla_deadline is not None:
                col6.metric("SLA Deadline cumprido", f"{sla_deadline:.1f}%")
            else:
                col6.metric("SLA Deadline cumprido", "—")

            if sla_inicializacao is not None:
                col7.metric("SLA Inicialização cumprido", f"{sla_inicializacao:.1f}%")
            else:
                col7.metric("SLA Inicialização cumprido", "—")

            exibir_insights_automaticos(
                insights_automaticos,
                dados_insuficientes_insights,
            )

            st.divider()

            st.subheader("Visão rápida da operação")

            g1, g2 = st.columns(2)

            fig_departamento = grafico_barras(
                df_filtrado,
                "departamento",
                "Chamados por departamento"
            )

            if fig_departamento:
                g1.plotly_chart(
                    fig_departamento,
                    use_container_width=True,
                    key="dashboard_departamento"
                )
            else:
                g1.info("Sem dados para departamento.")

            fig_categoria = grafico_barras(
                df_filtrado,
                "categoria",
                "Chamados por categoria"
            )

            if fig_categoria:
                g2.plotly_chart(
                    fig_categoria,
                    use_container_width=True,
                    key="dashboard_categoria"
                )
            else:
                g2.info("Sem dados para categoria.")

        with tab_comparativo:
            st.subheader("Comparativo Mensal")

            if "data_de_criacao" not in df.columns:
                st.warning(
                    "A coluna data_de_criacao não está disponível na planilha."
                )
            elif comparativo_mensal is None:
                st.info(
                    "Não há dados suficientes para o comparativo mensal."
                )
            else:
                label_mes_atual = comparativo_mensal["label_mes_atual"]
                label_mes_anterior = comparativo_mensal["label_mes_anterior"]
                ind_mes_atual = comparativo_mensal["ind_mes_atual"]
                ind_mes_anterior = comparativo_mensal["ind_mes_anterior"]
                var_chamados = comparativo_mensal["var_chamados"]
                var_chamados_pct = comparativo_mensal["var_chamados_pct"]
                var_backlog = comparativo_mensal["var_backlog"]
                var_backlog_pct = variacao_percentual(
                    ind_mes_atual["backlog_fim"],
                    ind_mes_anterior["backlog_fim"],
                )
                var_sla = comparativo_mensal["var_sla_deadline"]

                st.caption(
                    f"Mês atual: {label_mes_atual} · "
                    f"Mês anterior: {label_mes_anterior}"
                )

                var_chamados_txt = "—"
                if var_chamados is not None:
                    var_chamados_txt = f"{var_chamados:+,}"
                    if var_chamados_pct is not None:
                        var_chamados_txt += f" ({var_chamados_pct:+.1f}%)"

                var_backlog_txt = "—"
                if var_backlog is not None:
                    var_backlog_txt = f"{var_backlog:+,}"
                    if var_backlog_pct is not None:
                        var_backlog_txt += f" ({var_backlog_pct:+.1f}%)"

                sla_atual_txt = (
                    f"{ind_mes_atual['sla_deadline']:.1f}%"
                    if ind_mes_atual["sla_deadline"] is not None
                    else "—"
                )

                var_sla_txt = (
                    f"{var_sla:+.1f} pp"
                    if var_sla is not None
                    else "—"
                )

                c1, c2, c3, c4, c5, c6 = st.columns(6)

                c1.metric(
                    f"Total de chamados ({label_mes_atual})",
                    f"{ind_mes_atual['total']:,}",
                )
                c2.metric("Variação de chamados", var_chamados_txt)
                c3.metric(
                    f"Backlog fim do mês ({label_mes_atual})",
                    f"{ind_mes_atual['backlog_fim']:,}",
                )
                c4.metric("Variação de backlog", var_backlog_txt)
                c5.metric(
                    f"SLA Deadline ({label_mes_atual})",
                    sla_atual_txt,
                )
                c6.metric("Variação do SLA Deadline", var_sla_txt)

                st.divider()

                tabela_comparativa = montar_tabela_comparativa(
                    ind_mes_atual,
                    ind_mes_anterior,
                    label_mes_atual,
                    label_mes_anterior,
                )

                tabela_comparativa = tabela_comparativa.rename(
                    columns={
                        "Mês Atual": label_mes_atual,
                        "Mês Anterior": label_mes_anterior,
                    }
                )

                st.dataframe(
                    tabela_comparativa,
                    use_container_width=True,
                    hide_index=True,
                    key="comparativo_tabela",
                )

                st.divider()

                st.subheader("Interpretação automática")

                interpretacao = gerar_interpretacao_comparativo(
                    ind_mes_atual,
                    ind_mes_anterior,
                    label_mes_atual,
                    label_mes_anterior,
                )

                st.markdown(interpretacao)

        with tab_backlog:
            st.subheader("Backlog Diário")

            if backlog_diario.empty:
                st.info("Não há dados suficientes para calcular o backlog diário.")
            else:
                st.line_chart(
                    backlog_diario.set_index("data")["backlog"]
                )

            st.subheader("Backlog Mensal")

            if backlog_mensal.empty:
                st.info("Não há dados suficientes para calcular o backlog mensal.")
            else:
                st.bar_chart(
                    backlog_mensal.set_index("mes")["backlog"]
                )

            st.subheader("Indicadores de Backlog")

            b1, b2, b3 = st.columns(3)

            if not backlog_diario.empty:
                maior_backlog = backlog_diario["backlog"].max()
                media_backlog = backlog_diario["backlog"].mean()
            else:
                maior_backlog = 0
                media_backlog = 0

            b1.metric("Backlog atual", f"{backlog_atual:,}")
            b2.metric("Maior backlog diário", f"{maior_backlog:,.0f}")
            b3.metric("Backlog médio diário", f"{media_backlog:,.0f}")
            st.divider()

            st.subheader("Análises Mensais Operacionais")

            if metricas_mensais.empty:
                st.info("Não há dados suficientes para gerar as análises mensais.")
            else:
                fig_abertos_fechados = px.bar(
                    metricas_mensais,
                    x="mes",
                    y=["chamados_abertos", "chamados_fechados"],
                    barmode="group",
                    title="Chamados abertos x fechados por mês"
                )

                st.plotly_chart(
                    fig_abertos_fechados,
                    use_container_width=True,
                    key="grafico_abertos_fechados_mes"
                )

                fig_urgentes = px.bar(
                    metricas_mensais,
                    x="mes",
                    y="chamados_urgentes",
                    text="chamados_urgentes",
                    title="Chamados urgentes por mês"
                )

                st.plotly_chart(
                    fig_urgentes,
                    use_container_width=True,
                    key="grafico_urgentes_mes"
                )

                fig_abertos_backlog = px.bar(
                    metricas_mensais,
                    x="mes",
                    y=["chamados_abertos", "backlog_inicial"],
                    barmode="stack",
                    title="Chamados abertos + backlog inicial por mês"
                )

                st.plotly_chart(
                    fig_abertos_backlog,
                    use_container_width=True,
                    key="grafico_abertos_backlog_inicial"
                )

                fig_backlog_finalizados = px.bar(
    metricas_mensais,
    x="mes",
    y=[
        "backlog_final",
        "backlog_herdado_finalizado",
        "backlog_herdado_remanescente"
    ],
    barmode="group",
    title="Backlog final x backlog herdado finalizado x redução líquida"
)

                st.plotly_chart(
                    fig_backlog_finalizados,
                    use_container_width=True,
                    key="grafico_backlog_finalizados"
                )

                st.subheader("Tabela mensal consolidada")

                st.dataframe(
                    metricas_mensais,
                    use_container_width=True,
                    hide_index=True
                )
        with tab_aging:
            st.subheader("Análise de Aging dos Chamados em Aberto")

            if aging_df.empty:
                st.info("Não há chamados em aberto para análise de aging.")
            else:
                acima_7 = (aging_df["idade_dias"] > 7).sum()
                acima_15 = (aging_df["idade_dias"] > 15).sum()
                acima_30 = (aging_df["idade_dias"] > 30).sum()
                mais_antigo = aging_df["idade_dias"].max()
                media_aging = aging_df["idade_dias"].mean()

                a1, a2, a3, a4, a5 = st.columns(5)

                a1.metric("> 7 dias", f"{acima_7:,}")
                a2.metric("> 15 dias", f"{acima_15:,}")
                a3.metric("> 30 dias", f"{acima_30:,}")
                a4.metric("Mais antigo", f"{mais_antigo:.0f} dias")
                a5.metric("Média de idade", f"{media_aging:.1f} dias")

                st.divider()

                resumo_aging = (
                    aging_df["faixa_aging"]
                    .value_counts()
                    .reindex(
                        ["0-3 dias", "4-7 dias", "8-15 dias", "16-30 dias", "+30 dias"],
                        fill_value=0
                    )
                    .reset_index()
                )

                resumo_aging.columns = ["Faixa", "Quantidade"]

                fig_aging = px.bar(
                    resumo_aging,
                    x="Faixa",
                    y="Quantidade",
                    text="Quantidade",
                    title="Distribuição dos chamados em aberto por faixa de aging"
                )

                st.plotly_chart(
                    fig_aging,
                    use_container_width=True,
                    key="grafico_aging"
                )

                st.divider()

                st.subheader("Top 20 chamados mais antigos")

                colunas_top = []

                for coluna in [
                    "protocolo",
                    "assunto",
                    "idade_dias",
                    "data_de_criacao",
                    "departamento",
                    "categoria",
                    "atendente",
                    "setor_solicitante",
                    "ultima_situacao",
                ]:
                    if coluna in aging_df.columns:
                        colunas_top.append(coluna)

                top_aging = (
                    aging_df
                    .sort_values("idade_dias", ascending=False)
                    [colunas_top]
                    .head(20)
                )

                st.dataframe(
                    top_aging,
                    use_container_width=True,
                    height=420
                )

                st.divider()

                st.subheader("Insights de Aging")

                if acima_30 > 0:
                    st.warning(f"Existem {acima_30:,} chamados com mais de 30 dias em aberto.")

                if acima_15 > 0:
                    st.warning(f"Existem {acima_15:,} chamados com mais de 15 dias em aberto.")

                if acima_7 == 0:
                    st.success("Não há chamados acima de 7 dias em aberto.")

                if "setor_solicitante" in aging_df.columns and not aging_df["setor_solicitante"].dropna().empty:
                    setor_critico = aging_df["setor_solicitante"].value_counts().idxmax()
                    qtd_setor = aging_df["setor_solicitante"].value_counts().max()

                    st.info(
                        f"Setor com maior concentração de chamados em aberto: {setor_critico} "
                        f"({qtd_setor:,} chamados)."
                    )

                if "atendente" in aging_df.columns and not aging_df["atendente"].dropna().empty:
                    aging_atendente = (
                        aging_df
                        .groupby("atendente")["idade_dias"]
                        .mean()
                        .sort_values(ascending=False)
                    )

                    atendente_maior_aging = aging_atendente.index[0]
                    media_atendente = aging_atendente.iloc[0]

                    st.info(
                        f"Atendente com maior média de aging: {atendente_maior_aging} "
                        f"({media_atendente:.1f} dias em média)."
                    )
        with tab_operacao:
            st.subheader("Análise por dimensão")

            g1, g2 = st.columns(2)

            fig_departamento = grafico_barras(
                df_filtrado,
                "departamento",
                "Chamados por departamento"
            )

            if fig_departamento:
                g1.plotly_chart(
                    fig_departamento,
                    use_container_width=True,
                    key="operacao_departamento"
                )
            else:
                g1.info("Sem dados para departamento.")

            fig_categoria = grafico_barras(
                df_filtrado,
                "categoria",
                "Chamados por categoria"
            )

            if fig_categoria:
                g2.plotly_chart(
                    fig_categoria,
                    use_container_width=True,
                    key="operacao_categoria"
                )
            else:
                g2.info("Sem dados para categoria.")

            g3, g4 = st.columns(2)

            fig_atendente = grafico_barras(
                df_filtrado,
                "atendente",
                "Chamados por atendente"
            )

            if fig_atendente:
                g3.plotly_chart(
                    fig_atendente,
                    use_container_width=True,
                    key="operacao_atendente"
                )
            else:
                g3.info("Sem dados para atendente.")

            fig_setor = grafico_barras(
                df_filtrado,
                "setor_solicitante",
                "Chamados por setor solicitante"
            )

            if fig_setor:
                g4.plotly_chart(
                    fig_setor,
                    use_container_width=True,
                    key="operacao_setor"
                )
            else:
                g4.info("Sem dados para setor solicitante.")

            fig_prioridade = grafico_barras(
                df_filtrado,
                "prioridade",
                "Chamados por prioridade"
            )

            if fig_prioridade:
                st.plotly_chart(
                    fig_prioridade,
                    use_container_width=True,
                    key="operacao_prioridade"
                )
            else:
                st.info("Sem dados para prioridade.")

        with tab_resumo:
            st.subheader("Resumo Executivo")
            st.markdown(resumo_executivo)

        with tab_dados:
            st.subheader("Prévia dos dados filtrados")
            st.dataframe(
                df_filtrado,
                use_container_width=True,
                height=500
            )

    except Exception as e:
        st.error(f"Erro ao carregar planilha: {e}")
else:
    st.info("Carregue uma planilha .xlsx para iniciar a análise.")