import argparse
import sys
from pathlib import Path

import pandas as pd


def load_reports(folder: Path) -> pd.DataFrame:

    files = sorted(folder.glob("*.xlsx"))
    if not files:
        print(f"В папке {folder} нет .xlsx файлов.")
        sys.exit(1)

    frames = []
    for file in files:
        try:
            df = pd.read_excel(file)
        except Exception as exc:
            print(f"[!] Пропускаю {file.name}: не удалось прочитать ({exc})")
            continue

        df["Файл-источник"] = file.name
        frames.append(df)
        print(f"Прочитан {file.name}: {len(df)} строк")

    if not frames:
        print("Не удалось прочитать ни одного файла.")
        sys.exit(1)

    return pd.concat(frames, ignore_index=True)

def main() -> None:
    parser = argparse.ArgumentParser(description="Объединение Excel-отчётов")
    parser.add_argument("--input", default="./reports", help="папка с .xlsx файлами")
    parser.add_argument("--output", default="итог.xlsx", help="имя итогового файла")
    args = parser.parse_args()

    data = load_reports(Path(args.input))
    data.to_excel(args.output, index=False)

    print(f"\nГотово: {len(data)} строк объединены в «{args.output}»")


if __name__ == "__main__":
    main()
