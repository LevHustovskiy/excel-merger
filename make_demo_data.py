import random
from pathlib import Path

import pandas as pd

MANAGERS = ["Иванов", "Петрова", "Сидоров", "Кузнецова"]
PRODUCTS = ["Ноутбук", "Монитор", "Клавиатура", "Мышь", "Наушники"]
PRICES = {"Ноутбук": 65000, "Монитор": 18000, "Клавиатура": 3500, "Мышь": 1500, "Наушники": 4200}


def make_report(rows: int) -> pd.DataFrame:
    data = []
    for _ in range(rows):
        product = random.choice(PRODUCTS)
        qty = random.randint(1, 10)
        data.append(
            {
                "Менеджер": random.choice(MANAGERS),
                "Товар": product,
                "Количество": qty,
                "Сумма": qty * PRICES[product],
            }
        )
    return pd.DataFrame(data)


def main() -> None:
    folder = Path("./reports")
    folder.mkdir(exist_ok=True)

    for month in ["январь", "февраль", "март"]:
        df = make_report(rows=random.randint(15, 25))
        path = folder / f"продажи_{month}.xlsx"
        df.to_excel(path, index=False)
        print(f"Создан {path} ({len(df)} строк)")

    print("\nТеперь запустите: python merge_reports.py")


if __name__ == "__main__":
    main()
