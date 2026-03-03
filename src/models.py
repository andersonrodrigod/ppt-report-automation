from __future__ import annotations

from dataclasses import dataclass


@dataclass
class SheetTable:
    name: str
    display_name: str
    headers: list[str]
    rows: list[list[object]]
