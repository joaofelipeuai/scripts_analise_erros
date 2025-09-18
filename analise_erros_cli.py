#!/usr/bin/env python3
import argparse
from pathlib import Path
import pandas as pd

def gerar_dashboard(src_path: Path, out_path: Path, sheet_name: str | int = 0):
    print(f"- Carregando dados de: {src_path}")
    df = pd.read_excel(src_path, sheet_name=sheet_name)
    print(f"- Dados carregados: {len(df)} registros encontrados")

    col_msg = next((c for c in df.columns if c.lower() == "message"), None)
    col_inst = next((c for c in df.columns if c.lower() in ["instant", "instant_utc", "date", "data"]), None)
    col_cat  = next((c for c in df.columns if c.lower() in ["categoria", "category"]), None)
    col_module  = next((c for c in df.columns if c.lower() in ["module_name", "module", "application_name", "espace_name"]), None)

    if col_msg is None:
        raise ValueError("Coluna 'MESSAGE' não encontrada.")
    if col_inst is None:
        df["_instant_falso"] = pd.Timestamp.today().normalize()
        col_inst = "_instant_falso"
    if col_cat is None:
        df["_categoria_aux"] = "Não Categorizado"
        col_cat = "_categoria_aux"
    if col_module is None:
        df["_origem_aux"] = "Origem Desconhecida"
        col_module = "_origem_aux"

    print("- Processando dados e identificando colunas...")
    df[col_inst] = pd.to_datetime(df[col_inst], errors="coerce")

    total_erros = len(df)

    print("- Gerando resumo de mensagens de erro...")
    top_msgs = (
        df.groupby(col_msg, dropna=False)
          .size().reset_index(name="Quantidade")
          .sort_values("Quantidade", ascending=False)
    )
    top_msgs["%"] = (top_msgs["Quantidade"] / total_erros * 100).round(2)

    print("- Analisando categorias...")
    cat_counts = (
        df.groupby(col_cat, dropna=False)
          .size().reset_index(name="Quantidade")
          .sort_values("Quantidade", ascending=False)
    )
    cat_counts["%"] = (cat_counts["Quantidade"] / total_erros * 100).round(2)

    print("- Processando dados por origem/módulo...")
    origem_counts = (
        df.groupby(col_module, dropna=False)
          .size().reset_index(name="Quantidade")
          .sort_values("Quantidade", ascending=False)
    )

    print("- Gerando análise temporal...")
    if df[col_inst].notna().any():
        weekly = (
            df.assign(AnoSemana=df[col_inst].dt.strftime("%G-W%V"))
              .groupby("AnoSemana").size()
              .reset_index(name="Quantidade")
              .sort_values("AnoSemana")
        )
    else:
        weekly = pd.DataFrame({"AnoSemana": [], "Quantidade": []})

    print("- Criando arquivo Excel...")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        print("  Salvando dados base...")
        df.to_excel(writer, index=False, sheet_name="Base")
        print("  Salvando resumo de erros...")
        top_msgs.to_excel(writer, index=False, sheet_name="Resumo_Erros")
        print("  Salvando resumo de categorias...")
        cat_counts.to_excel(writer, index=False, sheet_name="Resumo_Categoria")
        print("  Salvando resumo de origens...")
        origem_counts.to_excel(writer, index=False, sheet_name="Resumo_Origem")
        print("  Salvando resumo semanal...")
        weekly.to_excel(writer, index=False, sheet_name="Resumo_Semana")

        print("- Criando dashboard e gráficos...")
        wb  = writer.book
        ws_dash = wb.add_worksheet("Dashboard")

        h1 = wb.add_format({"bold": True, "font_size": 16})
        kpi_label = wb.add_format({"bold": True, "font_size": 12, "align": "left"})
        kpi_val   = wb.add_format({"bold": True, "font_size": 14, "align": "left"})
        small = wb.add_format({"font_size": 9})

        min_dt = df[col_inst].min()
        max_dt = df[col_inst].max()
        ws_dash.write("A1", "Dashboard de Erros", h1)
        periodo_txt = f"Período: {min_dt.strftime('%d/%m/%Y')} a {max_dt.strftime('%d/%m/%Y')}" if pd.notna(min_dt) and pd.notna(max_dt) else "Período: não disponível"
        ws_dash.write("A2", periodo_txt, small)

        ws_dash.write("A4", "Total de ocorrências", kpi_label)
        ws_dash.write("A5", len(df), kpi_val)

        ws_dash.write("C4", "Mensagens distintas", kpi_label)
        ws_dash.write("C5", top_msgs.shape[0], kpi_val)

        ws_err = writer.sheets["Resumo_Erros"]
        ws_err.set_column(0, 0, 70)
        ws_err.set_column(1, 2, 14)

        top_n = min(10, len(top_msgs))
        err_first_row = 1
        err_last_row  = top_n
        chart_err = wb.add_chart({"type": "column"})
        chart_err.add_series({
            "name":       "Quantidade",
            "categories": ["Resumo_Erros", err_first_row+1, 0, err_last_row, 0],
            "values":     ["Resumo_Erros", err_first_row+1, 1, err_last_row, 1],
            "data_labels": {"value": True}
        })
        chart_err.set_title({"name": "Top Erros (Top 10)"})
        chart_err.set_x_axis({"name": "Mensagem"})
        chart_err.set_y_axis({"name": "Ocorrências"})
        chart_err.set_legend({"none": True})
        ws_dash.insert_chart("A8", chart_err, {"x_scale": 1.35, "y_scale": 1.2})
        print("  Gráfico de top erros criado")

        ws_cat = writer.sheets["Resumo_Categoria"]
        ws_cat.set_column(0, 2, 22)
        cat_n = len(cat_counts)
        if cat_n >= 1:
            chart_cat = wb.add_chart({"type": "pie"})
            chart_cat.add_series({
                "name": "Distribuição por Categoria",
                "categories": ["Resumo_Categoria", 1, 0, cat_n, 0],
                "values": ["Resumo_Categoria", 1, 1, cat_n, 1],
                "data_labels": {"percentage": True}
            })
            chart_cat.set_title({"name": "Distribuição por Categoria"})
            ws_dash.insert_chart("H8", chart_cat, {"x_scale": 1.0, "y_scale": 1.2})
            print("  Gráfico de categorias criado")

        ws_week = writer.sheets["Resumo_Semana"]
        ws_week.set_column(0, 1, 18)
        week_n = len(weekly)
        if week_n >= 2:
            chart_week = wb.add_chart({"type": "line"})
            chart_week.add_series({
                "name": "Ocorrências por Semana",
                "categories": ["Resumo_Semana", 1, 0, week_n, 0],
                "values": ["Resumo_Semana", 1, 1, week_n, 1],
                "data_labels": {"value": False}
            })
            chart_week.set_title({"name": "Ocorrências por Semana (ISO)"})
            chart_week.set_x_axis({"name": "Ano-Semana"})
            chart_week.set_y_axis({"name": "Ocorrências"})
            ws_dash.insert_chart("A25", chart_week, {"x_scale": 1.35, "y_scale": 1.2})
            print("  Gráfico temporal criado")

        ws_origem = writer.sheets["Resumo_Origem"]
        ws_origem.set_column(0, 1, 30)
        origem_n = min(10, len(origem_counts))
        if origem_n >= 1:
            chart_orig = wb.add_chart({"type": "bar"})
            chart_orig.add_series({
                "name": "Ocorrências por Origem (Top 10)",
                "categories": ["Resumo_Origem", 1, 0, origem_n, 0],
                "values": ["Resumo_Origem", 1, 1, origem_n, 1],
                "data_labels": {"value": True}
            })
            chart_orig.set_title({"name": "Ocorrências por Origem (Top 10)"})
            chart_orig.set_x_axis({"name": "Origem"})
            chart_orig.set_y_axis({"name": "Ocorrências"})
            chart_orig.set_legend({"none": True})
            ws_dash.insert_chart("H25", chart_orig, {"x_scale": 1.0, "y_scale": 1.2})
            print("  Gráfico de origens criado")

    print("- Processamento concluído!")
    print(f"- Arquivo gerado em: {out_path}")

def main():
    parser = argparse.ArgumentParser(description="Gera dashboard de erros a partir de uma planilha.")
    parser.add_argument("-i", "--input", required=True, help="Caminho do Excel de entrada (mesma estrutura).")
    parser.add_argument("-o", "--output", help="Caminho do Excel de saída (default: mesmo diretório com nome Dashboard_Erros.xlsx).")
    parser.add_argument("-s", "--sheet", default="0", help="Aba a ler (nome da aba ou índice; default: 0).")
    args = parser.parse_args()

    print("- Iniciando análise de erros...")

    src = Path(args.input).expanduser().resolve()
    if not src.exists():
        raise SystemExit(f"Arquivo de entrada não encontrado: {src}")

    # Interpretar sheet como int se for número
    try:
        sheet = int(args.sheet)
    except ValueError:
        sheet = args.sheet

    if args.output:
        out = Path(args.output).expanduser().resolve()
    else:
        out = src.with_name("Dashboard_Erros.xlsx")

    out.parent.mkdir(parents=True, exist_ok=True)

    gerar_dashboard(src, out, sheet_name=sheet)

if __name__ == "__main__":
    main()
