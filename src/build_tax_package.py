from __future__ import annotations

from datetime import datetime
from pathlib import Path

from mercari_converter import convert_sold_csv_to_excel


def build_tax_accountant_package(input_csv: str | Path, output_dir: str | Path) -> Path:
    """Create the accountant-facing Excel workbook from a Mercari sold CSV."""
    output_path = Path(output_dir) / f"税理士提出用_メルカリ会計資料_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return convert_sold_csv_to_excel(input_csv, output_path)
