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


def build_summary(data: pd.DataFrame) -> pd.DataFrame:

    summary = (
        data.groupby("Менеджер", as_index=False)
        .agg({"Количество": "sum", "Сумма": "sum"})
        .sort_values("Сумма", ascending=False)
    )

    total = pd.DataFrame(
        [{
            "Менеджер": "ИТОГО",
            "Количество": summary["Количество"].sum(),
            "Сумма": summary["Сумма"].sum(),
        }]
    )
    return pd.concat([summary, total], ignore_index=True)


def autofit_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:

    worksheet = writer.sheets[sheet_name]
    for i, column in enumerate(df.columns, start=1):
        max_len = max(df[column].astype(str).map(len).max(), len(str(column)))
        letter = worksheet.cell(row=1, column=i).column_letter
        worksheet.column_dimensions[letter].width = min(max_len + 3, 40)


def main() -> None:
    parser = argparse.ArgumentParser(description="Объединение Excel-отчётов")
    parser.add_argument("--input", default="./reports", help="папка с .xlsx файлами")
    parser.add_argument("--output", default="итог.xlsx", help="имя итогового файла")
    args = parser.parse_args()

    data = load_reports(Path(args.input))

    required = {"Менеджер", "Товар", "Количество", "Сумма"}
    missing = required - set(data.columns)
    if missing:
        print(f"В данных не хватает колонок: {', '.join(missing)}")
        sys.exit(1)

    summary = build_summary(data)

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="Все данные", index=False)
        summary.to_excel(writer, sheet_name="Сводка", index=False)
        autofit_columns(writer, "Все данные", data)
        autofit_columns(writer, "Сводка", summary)

    print(
        f"\nГотово: {len(data)} строк из исходных файлов объединены в «{args.output}»\n"
        f"Листы: «Все данные» и «Сводка» (продажи по менеджерам + итог)."
    )


if __name__ == "__main__":
    main()
