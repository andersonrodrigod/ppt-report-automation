from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from src.pipeline import executar


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

    default_out = (
        base_dir
        / "data"
        / "output"
        / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        / "indicadores_apresentacao.pptx"
    )

    parser = argparse.ArgumentParser(description="Gera PPT direto da base sem excel intermediario.")
    parser.add_argument("--entrada", default=str(default_in), help="Arquivo XLSX de entrada.")
    parser.add_argument("--saida", default=str(default_out), help="Arquivo PPTX de saida.")
    parser.add_argument("--aba", default="BASE", help="Aba de origem.")
    parser.add_argument("--tipo", default=None, help='Filtro na coluna TIPO (ex.: "VIDEO ABDOMINAL").')
    parser.add_argument("--layout", default="paired", choices=["paired", "grid4"], help="Layout dos slides.")
    parser.add_argument("--assets", default=str(base_dir / "assets"), help="Pasta de assets (logo.jpg e inicio.jpg).")
    args = parser.parse_args()

    saida = executar(
        arquivo_entrada=args.entrada,
        arquivo_saida=args.saida,
        assets_dir=args.assets,
        aba_origem=args.aba,
        tipo_filtro=args.tipo,
        layout_mode=args.layout,
    )
    print(f"Arquivo gerado: {saida}")


if __name__ == "__main__":
    main()
