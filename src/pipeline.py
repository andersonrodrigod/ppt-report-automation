from __future__ import annotations

from pathlib import Path

from .data_processing import contar_respostas_sim_nao, montar_tabelas_por_uf
from .ppt_renderer import gerar_ppt


def executar(
    arquivo_entrada: str,
    arquivo_saida: str,
    assets_dir: str,
    aba_origem: str = "BASE",
    tipo_filtro: str | None = None,
    layout_mode: str = "paired",
    header_layout_gap_px: int = 24,
) -> Path:
    ordered_names, tables_by_sheet = montar_tabelas_por_uf(
        arquivo_excel=arquivo_entrada,
        aba_origem=aba_origem,
        tipo_filtro=tipo_filtro,
    )
    contagens_sim_nao = contar_respostas_sim_nao(
        arquivo_excel=arquivo_entrada,
        aba_origem=aba_origem,
        tipo_filtro=tipo_filtro,
        colunas=("P1", "P3"),
    )

    saida = Path(arquivo_saida)
    saida.parent.mkdir(parents=True, exist_ok=True)

    gerar_ppt(
        ordered_names=ordered_names,
        tables_by_sheet=tables_by_sheet,
        arquivo_saida=str(saida),
        assets_dir=assets_dir,
        layout_mode=layout_mode,
        contagens_sim_nao=contagens_sim_nao,
        header_layout_gap_px=header_layout_gap_px,
    )
    return saida
