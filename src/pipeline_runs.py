from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .pipeline import executar


def _montar_saida_run(base_dir: Path, run_nome: str, ts: datetime, arquivo_nome: str = "indicadores_apresentacao.pptx") -> Path:
    timestamp = ts.strftime("%Y%m%d_%H%M%S")
    return base_dir / "data" / "output" / f"{run_nome}_{timestamp}" / arquivo_nome


def executar_run_cirurgia(
    arquivo_entrada: str,
    assets_dir: str,
    base_dir: Path,
    aba_origem: str = "BASE",
    header_layout_gap_px: int = 24,
    arquivo_saida: str | None = None,
    ts: datetime | None = None,
) -> Path:
    run_ts = ts or datetime.now()
    saida = Path(arquivo_saida) if arquivo_saida else _montar_saida_run(base_dir, "cirurgia_run", run_ts)
    return executar(
        arquivo_entrada=arquivo_entrada,
        arquivo_saida=str(saida),
        assets_dir=assets_dir,
        aba_origem=aba_origem,
        tipo_filtro=None,
        layout_mode="paired",
        header_layout_gap_px=header_layout_gap_px,
    )


def executar_run_video(
    arquivo_entrada: str,
    assets_dir: str,
    base_dir: Path,
    aba_origem: str = "BASE",
    header_layout_gap_px: int = 24,
    arquivo_saida: str | None = None,
    ts: datetime | None = None,
) -> Path:
    run_ts = ts or datetime.now()
    saida = Path(arquivo_saida) if arquivo_saida else _montar_saida_run(base_dir, "video_run", run_ts)
    return executar(
        arquivo_entrada=arquivo_entrada,
        arquivo_saida=str(saida),
        assets_dir=assets_dir,
        aba_origem=aba_origem,
        tipo_filtro="VIDEO ABDOMINAL",
        layout_mode="grid4",
        header_layout_gap_px=header_layout_gap_px,
    )
