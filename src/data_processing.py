from __future__ import annotations

from typing import Tuple
import unicodedata
import warnings

import pandas as pd

from .models import SheetTable


def _normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _normalizar_sem_acento(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch)).lower().strip()


def _nome_aba_seguro(nome: str, usados: set[str]) -> str:
    inval = ['\\', '/', '*', '[', ']', ':', '?']
    base = nome
    for ch in inval:
        base = base.replace(ch, "_")
    base = (base.strip() or "SEM_UF")[:31]

    candidato = base
    idx = 1
    while candidato in usados:
        sufixo = f"_{idx}"
        candidato = f"{base[:31-len(sufixo)]}{sufixo}"
        idx += 1

    usados.add(candidato)
    return candidato


def _display_name_from_sheet_name(sheet_name: str) -> str:
    nome = sheet_name
    if nome.upper().startswith("P1_") or nome.upper().startswith("P3_"):
        nome = nome[3:]
    return nome.replace("_", " ").strip() or sheet_name


def _read_excel_base(arquivo_excel: str, aba_origem: str) -> pd.DataFrame:
    # Alguns arquivos possuem células marcadas como data com serial inválido;
    # isso é ruído de origem e não deve quebrar nem poluir a execução.
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=r"Cell .* is marked as a date but the serial value .* is outside the limits for dates.*",
            category=UserWarning,
        )
        return pd.read_excel(arquivo_excel, sheet_name=aba_origem)


def _calcular_indicadores_df(
    df_base: pd.DataFrame,
    coluna_pergunta: str,
    valor_alvo: str,
) -> pd.DataFrame:
    df = df_base.copy()

    if coluna_pergunta not in df.columns:
        raise ValueError(f'A coluna "{coluna_pergunta}" nao foi encontrada na planilha BASE.')
    if "ESPECIALISTA" not in df.columns:
        raise ValueError('A coluna "ESPECIALISTA" nao foi encontrada na planilha BASE.')

    df["ESPECIALISTA"] = df["ESPECIALISTA"].map(_normalizar_texto).replace("", "SEM_ESPECIALISTA")
    df[coluna_pergunta] = df[coluna_pergunta].map(_normalizar_texto)

    total_cirurgias = df.groupby("ESPECIALISTA", dropna=False).size().rename("TOTAL_CIRURGIAS")

    alvo_norm = _normalizar_sem_acento(_normalizar_texto(valor_alvo))
    serie_norm = df[coluna_pergunta].map(_normalizar_sem_acento)
    total_alvo = (
        df[serie_norm == alvo_norm]
        .groupby("ESPECIALISTA", dropna=False)
        .size()
        .rename("TOTAL_SIM")
    )

    resultado = (
        total_cirurgias.to_frame()
        .join(total_alvo, how="left")
        .fillna({"TOTAL_SIM": 0})
        .reset_index()
    )
    resultado["TOTAL_SIM"] = resultado["TOTAL_SIM"].astype(int)
    resultado["PERCENTUAL_SIM"] = (resultado["TOTAL_SIM"] / resultado["TOTAL_CIRURGIAS"] * 100).round(1)

    total_geral_alvo = int(resultado["TOTAL_SIM"].sum())
    if total_geral_alvo == 0:
        resultado["REPRESENTATIVIDADE"] = 0.0
    else:
        resultado["REPRESENTATIVIDADE"] = (resultado["TOTAL_SIM"] / total_geral_alvo * 100).round(1)

    resultado = resultado.sort_values("TOTAL_CIRURGIAS", ascending=False).reset_index(drop=True)

    total_cirurgias_geral = int(resultado["TOTAL_CIRURGIAS"].sum())
    total_alvo_geral = int(resultado["TOTAL_SIM"].sum())
    percentual_alvo_geral = (
        round(total_alvo_geral / total_cirurgias_geral * 100, 1) if total_cirurgias_geral > 0 else 0.0
    )

    linha_total = pd.DataFrame(
        [
            {
                "ESPECIALISTA": "TOTAL",
                "TOTAL_CIRURGIAS": total_cirurgias_geral,
                "TOTAL_SIM": total_alvo_geral,
                "PERCENTUAL_SIM": percentual_alvo_geral,
                "REPRESENTATIVIDADE": 100.0 if total_alvo_geral > 0 else 0.0,
            }
        ]
    )
    resultado = pd.concat([resultado, linha_total], ignore_index=True)
    return resultado


def _df_para_sheet_table(nome_aba: str, df: pd.DataFrame, rotulo_benef: str = "SIM") -> SheetTable:
    headers = [
        "ESPECIALISTA",
        "TOTAL DE CIRURGIAS REALIZADAS",
        f'Nº BENEF "{rotulo_benef}"',
        "PROPORCIONALIDADE",
        "REPRESENTATIVIDADE",
    ]
    rows: list[list[object]] = []
    for _, row in df.iterrows():
        rows.append(
            [
                row["ESPECIALISTA"],
                int(row["TOTAL_CIRURGIAS"]),
                int(row["TOTAL_SIM"]),
                float(row["PERCENTUAL_SIM"]),
                float(row["REPRESENTATIVIDADE"]),
            ]
        )

    return SheetTable(
        name=nome_aba,
        display_name=_display_name_from_sheet_name(nome_aba),
        headers=headers,
        rows=rows,
    )


def contar_respostas_sim_nao(
    arquivo_excel: str,
    aba_origem: str = "BASE",
    tipo_filtro: str | None = None,
    colunas: tuple[str, ...] = ("P1", "P3"),
) -> dict[str, dict[str, int]]:
    df_base = _read_excel_base(arquivo_excel, aba_origem)

    if tipo_filtro is not None:
        if "TIPO" not in df_base.columns:
            raise ValueError('A coluna "TIPO" nao foi encontrada na planilha BASE.')
        tipo_norm = df_base["TIPO"].map(_normalizar_texto).map(_normalizar_sem_acento)
        filtro_norm = _normalizar_sem_acento(tipo_filtro)
        df_base = df_base[tipo_norm == filtro_norm].copy()

    saida: dict[str, dict[str, int]] = {}
    for coluna in colunas:
        if coluna not in df_base.columns:
            raise ValueError(f'A coluna "{coluna}" nao foi encontrada na planilha BASE.')

        serie = df_base[coluna].map(_normalizar_texto).map(_normalizar_sem_acento)
        qtd_sim = int((serie == "sim").sum())
        qtd_nao = int((serie == "nao").sum())
        saida[coluna] = {"Sim": qtd_sim, "Não": qtd_nao}

    return saida


def calcular_taxas_resposta(
    arquivo_excel: str,
    aba_origem: str = "BASE",
    tipo_filtro: str | None = None,
) -> dict[str, int]:
    df_base = _read_excel_base(arquivo_excel, aba_origem)

    if tipo_filtro is not None:
        if "TIPO" not in df_base.columns:
            raise ValueError('A coluna "TIPO" nao foi encontrada na planilha BASE.')
        tipo_norm = df_base["TIPO"].map(_normalizar_texto).map(_normalizar_sem_acento)
        filtro_norm = _normalizar_sem_acento(tipo_filtro)
        df_base = df_base[tipo_norm == filtro_norm].copy()

    if "P1" not in df_base.columns:
        raise ValueError('A coluna "P1" nao foi encontrada na planilha BASE.')
    status_col = next((c for c in df_base.columns if str(c).strip().upper() == "STATUS"), None)
    if status_col is None:
        raise ValueError('A coluna "Status" nao foi encontrada na planilha BASE.')

    p1_norm = df_base["P1"].map(_normalizar_texto).map(_normalizar_sem_acento)
    status_norm = df_base[status_col].map(_normalizar_texto).map(_normalizar_sem_acento)

    resposta = int(((p1_norm == "sim") | (p1_norm == "nao")).sum())
    nqa = int((status_norm == "nao quis").sum())
    nao_contato = int(((p1_norm == "") & (status_norm != "nao quis")).sum())

    return {
        "Resposta": resposta,
        "NQA": nqa,
        "Não conseguimos contato": nao_contato,
    }


def montar_tabelas_por_uf(
    arquivo_excel: str,
    aba_origem: str = "BASE",
    tipo_filtro: str | None = None,
) -> Tuple[list[str], dict[str, SheetTable]]:
    df_base = _read_excel_base(arquivo_excel, aba_origem)

    if "UF" not in df_base.columns:
        raise ValueError('A coluna "UF" nao foi encontrada na planilha BASE.')

    if tipo_filtro is not None:
        if "TIPO" not in df_base.columns:
            raise ValueError('A coluna "TIPO" nao foi encontrada na planilha BASE.')
        tipo_norm = df_base["TIPO"].map(_normalizar_texto).map(_normalizar_sem_acento)
        filtro_norm = _normalizar_sem_acento(tipo_filtro)
        df_base = df_base[tipo_norm == filtro_norm].copy()

    df_base["UF"] = df_base["UF"].map(_normalizar_texto).replace("", "SEM_UF")

    usados: set[str] = set()
    ordered_names: list[str] = []
    tables_by_sheet: dict[str, SheetTable] = {}

    blocos = [
        ("P1", "P1", "Sim", "SIM"),
        ("P3", "P3", "Não", "NÃO"),
    ]
    for prefixo, coluna_pergunta, valor_alvo, rotulo_benef in blocos:
        nome_geral = _nome_aba_seguro(f"{prefixo}_GERAL", usados)
        df_geral = _calcular_indicadores_df(df_base, coluna_pergunta=coluna_pergunta, valor_alvo=valor_alvo)
        tables_by_sheet[nome_geral] = _df_para_sheet_table(nome_geral, df_geral, rotulo_benef=rotulo_benef)
        ordered_names.append(nome_geral)

        ordem_ufs = (
            df_base.groupby("UF", dropna=False)
            .size()
            .sort_values(ascending=False)
            .index
            .tolist()
        )
        for uf in ordem_ufs:
            df_uf = df_base[df_base["UF"] == uf]
            df_tabela = _calcular_indicadores_df(df_uf, coluna_pergunta=coluna_pergunta, valor_alvo=valor_alvo)
            nome_aba = _nome_aba_seguro(f"{prefixo}_{uf}", usados)
            tables_by_sheet[nome_aba] = _df_para_sheet_table(nome_aba, df_tabela, rotulo_benef=rotulo_benef)
            ordered_names.append(nome_aba)

    return ordered_names, tables_by_sheet
