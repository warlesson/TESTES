import flet as ft
import pandas as pd
import os
from matplotlib import pyplot as plt

ARQUIVO_EXCEL = "dados_fiis.xlsx"

def carregar_dados():
    if os.path.exists(ARQUIVO_EXCEL):
        return pd.read_excel(ARQUIVO_EXCEL, sheet_name=None)
    else:
        return {
            "Aportes": pd.DataFrame(columns=["FII", "Nº Cotas", "Valor Cota"]),
            "Proventos": pd.DataFrame(columns=["FII", "Provento"])
        }

def salvar_dados(dados):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl", mode="w") as writer:
        for nome_aba, df in dados.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)

def gerar_resumo():
    dados = carregar_dados()
    aportes = dados["Aportes"]
    proventos = dados["Proventos"]

    if aportes.empty or proventos.empty:
        return pd.DataFrame(columns=["FII", "Nº Cotas", "Provento", "RENDIMENTO MÊS"])

    df = pd.merge(aportes, proventos, on="FII", how="inner")
    df["RENDIMENTO MÊS"] = df["Nº Cotas"] * df["Provento"]
    return df[["FII", "Nº Cotas", "Provento", "RENDIMENTO MÊS"]]

def gerar_grafico(df):
    if df.empty:
        return None
    fig, ax = plt.subplots()
    agrupado = df.groupby("FII")["Nº Cotas"].sum()
    ax.pie(agrupado, labels=agrupado.index, autopct="%1.1f%%", startangle=90)
    ax.set_title("Proporção de Cotas por FII")
    caminho = "grafico_pizza.png"
    plt.savefig(caminho)
    return caminho

def main(page: ft.Page):
    page.title = "Controle de FIIs"
    page.scroll = ft.ScrollMode.ALWAYS

    dados = carregar_dados()

    # Tabela de Aportes
    def tabela_aportes():
        return ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII")),
                ft.DataColumn(ft.Text("Nº Cotas")),
                ft.DataColumn(ft.Text("Valor Cota")),
            ],
            rows=[
                ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["FII"]))),
                    ft.DataCell(ft.Text(str(row["Nº Cotas"]))),
                    ft.DataCell(ft.Text(str(row["Valor Cota"]))),
                ]) for i, row in dados["Aportes"].iterrows()
            ]
        )

    # Tabela de Proventos
    def tabela_proventos():
        return ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII")),
                ft.DataColumn(ft.Text("Provento")),
            ],
            rows=[
                ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["FII"]))),
                    ft.DataCell(ft.Text(str(row["Provento"]))),
                ]) for i, row in dados["Proventos"].iterrows()
            ]
        )

    # Tabela de Resumo
    def tabela_resumo():
        resumo = gerar_resumo()
        return ft.Column([
            ft.DataTable(
                columns=[
                    ft.DataColumn(ft.Text("FII")),
                    ft.DataColumn(ft.Text("Nº Cotas")),
                    ft.DataColumn(ft.Text("Provento")),
                    ft.DataColumn(ft.Text("RENDIMENTO MÊS")),
                ],
                rows=[
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(str(row["FII"]))),
                        ft.DataCell(ft.Text(str(row["Nº Cotas"]))),
                        ft.DataCell(ft.Text(str(row["Provento"]))),
                        ft.DataCell(ft.Text(str(row["RENDIMENTO MÊS"]))),
                    ]) for i, row in resumo.iterrows()
                ]
            ),
            ft.Image(src=gerar_grafico(resumo), width=400, height=400)
            if not resumo.empty else ft.Text("Sem dados para gerar gráfico.")
        ])

    # Tabs
    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="Aportes",
                content=tabela_aportes()
            ),
            ft.Tab(
                text="Proventos",
                content=tabela_proventos()
            ),
            ft.Tab(
                text="Resumo",
                content=tabela_resumo()
            ),
        ]
    )

    page.add(tabs)

ft.app(target=main)
