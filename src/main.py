from __future__ import annotations

from pathlib import Path

from build_tax_package import build_tax_accountant_package


ROOT = Path(__file__).resolve().parents[1]
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"


def find_input_csv() -> Path:
    candidates = sorted(INPUT_DIR.glob("mercari_sold_*.csv")) or sorted(INPUT_DIR.glob("*.csv"))
    if not candidates:
        raise FileNotFoundError("inputフォルダにメルカリsold CSVが見つかりません。")
    return candidates[0]


def main() -> None:
    input_csv = find_input_csv()
    output_path = build_tax_accountant_package(input_csv, OUTPUT_DIR)
    print(f"税理士提出用Excelを出力しました: {output_path}")


if __name__ == "__main__":
    main()
