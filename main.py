from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from src.pipeline_runs import executar_run_cirurgia, executar_run_video


def _default_entrada(base_dir: Path) -> Path:
    candidatos = sorted((base_dir / "data" / "input").glob("*.xlsx"))
    if not candidatos:
        raise FileNotFoundError("Nenhum XLSX encontrado em data/input.")
    return candidatos[0]


def main() -> None:
    base_dir = Path(__file__).resolve().parent

    try:
        default_in = _default_entrada(base_dir)
    except FileNotFoundError:
        default_in = base_dir / "data" / "input" / "COMPLICACAO.xlsx"

    parser = argparse.ArgumentParser(description="Gera PPT direto da base sem excel intermediario.")
    parser.add_argument("--entrada", default=str(default_in), help="Arquivo XLSX de entrada.")
    parser.add_argument("--saida", default=None, help="Arquivo PPTX de saida (somente quando --modo for cirurgia ou video).")
    parser.add_argument("--aba", default="BASE", help="Aba de origem.")
    parser.add_argument("--modo", default="ambos", choices=["ambos", "cirurgia", "video"], help="Define quais apresentacoes gerar.")
    parser.add_argument("--assets", default=str(base_dir / "assets"), help="Pasta de assets (logo.jpg e inicio.jpg).")
    parser.add_argument("--header-gap", type=int, default=24, help="Espacamento em px entre cabecalho e layouts.")
    args = parser.parse_args()

    run_ts = datetime.now()

    if args.modo == "cirurgia":
        saida = executar_run_cirurgia(
            arquivo_entrada=args.entrada,
            assets_dir=args.assets,
            base_dir=base_dir,
            aba_origem=args.aba,
            header_layout_gap_px=args.header_gap,
            arquivo_saida=args.saida,
            ts=run_ts,
        )
        print(f"Arquivo gerado (cirurgia): {saida}")
        return

    if args.modo == "video":
        saida = executar_run_video(
            arquivo_entrada=args.entrada,
            assets_dir=args.assets,
            base_dir=base_dir,
            aba_origem=args.aba,
            header_layout_gap_px=args.header_gap,
            arquivo_saida=args.saida,
            ts=run_ts,
        )
        print(f"Arquivo gerado (video): {saida}")
        return

    saida_cirurgia = executar_run_cirurgia(
        arquivo_entrada=args.entrada,
        assets_dir=args.assets,
        base_dir=base_dir,
        aba_origem=args.aba,
        header_layout_gap_px=args.header_gap,
        ts=run_ts,
    )
    saida_video = executar_run_video(
        arquivo_entrada=args.entrada,
        assets_dir=args.assets,
        base_dir=base_dir,
        aba_origem=args.aba,
        header_layout_gap_px=args.header_gap,
        ts=run_ts,
    )
    print(f"Arquivo gerado (cirurgia): {saida_cirurgia}")
    print(f"Arquivo gerado (video): {saida_video}")


if __name__ == "__main__":
    main()
