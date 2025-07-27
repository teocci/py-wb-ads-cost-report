# WB Ads & Orders Report Builder

This small utility merges two Wildberries-related Excel reports and produces a single, tidy Excel file with
per‑SKU advertisement costs, orders, and **CPO** (cost per order). It also writes a compact summary block to the
right side of the sheet.

> Script file: `build_wb_ads_report.py`

## What it does

1. **Reads** the *ads-cost* file (downloaded from MP Manager), which contains:
   - `Товар` (product name)
   - `Артикул товара` (article/SKU)
   - `Сумма затрат на рекламу` (ads spend for the day)

2. **Reads** the *supplier-goods* file (downloaded from Wildberries supplier portal), which contains rows per warehouse.
   The script **aggregates by article** and keeps:
   - `Артикул WB` → article key
   - `шт.` → summed quantity (**Заказано, шт.**)
   - `Сумма заказов минус комиссия WB, руб.` → summed revenue (**Заказано, руб.**)

3. **Left-joins** the two datasets by article (string match).

4. **Calculates** `Затраты на 1 заказ` (CPO) = `Сумма затрат на рекламу` / `Заказано, шт.`
   (left blank when quantity is 0).

5. **Writes** the output Excel file with formatted numbers and a summary block:
   - *Финансовые показатели на основе отчетов ВБ*
   - *Итого заказов, руб.* = sum of `Заказано, руб.`
   - *Итого затрат на рекламу, руб.* = sum of `Сумма затрат на рекламу`


## Input files

- **Ads-cost file** (e.g., `Аналитика по товарам от 27.07.2025.xlsx`):
  - Required columns: `Товар`, `Артикул товара`, `Сумма затрат на рекламу`

- **Supplier-goods file** (e.g., `supplier-goods-3925272-2025-06-21-2025-06-21-*.xlsx`):
  - Required columns (some headers may be a few rows below due to a preamble; the script auto-detects):
    - `Артикул WB`, `шт.`, `Сумма заказов минус комиссия WB, руб.`

> All article IDs are normalized to **string** during processing to avoid number/text mismatches.


## Output

A single Excel workbook (default name: `wb_ads_report.xlsx`) with one sheet `Отчет` and the columns:

1. `Наименование`
2. `Артикул`
3. `Сумма затрат на рекламу`
4. `Заказано, руб.`
5. `Заказано, шт.`
6. `Затраты на 1 заказ`

The sheet includes number formatting and a summary block on the right.

## Installation

```bash
python -V  # Python 3.10+ recommended (3.11 tested)
pip install -r requirements.txt
```

## Usage

```bash
python build_wb_ads_report_old.py   --ads "Аналитика по товарам от 27.07.2025.xlsx"   --goods "supplier-goods-3925272-2025-06-21-2025-06-21-khzvlnmig.xlsx"   --out "wb_ads_report_2025-07-27.xlsx"
```

Optional arguments:
- `--ads-sheet`   (sheet index or name in the ads file; default 0)
- `--goods-sheet` (sheet index or name in the supplier-goods file; default 0)

### Windows tips
- Wrap paths with spaces in quotes.
- Ensure file encoding is preserved; Excel files are binary so encoding of filenames is usually the only concern.

## Column mapping

| Source file     | Input column (RU)                                  | Output column (RU)           |
|-----------------|----------------------------------------------------|------------------------------|
| ads-cost        | `Товар`                                            | `Наименование`               |
| ads-cost        | `Артикул товара`                                   | `Артикул`                    |
| ads-cost        | `Сумма затрат на рекламу`                          | `Сумма затрат на рекламу`    |
| supplier-goods  | `Артикул WB`                                       | `Артикул` (join key)         |
| supplier-goods  | `шт.` (summed)                                     | `Заказано, шт.`              |
| supplier-goods  | `Сумма заказов минус комиссия WB, руб.` (summed)   | `Заказано, руб.`             |
| computed        | —                                                  | `Затраты на 1 заказ`         |

## Troubleshooting

- **Missing column error**: The script validates the presence of required columns and shows which ones are missing.
  If your WB report contains a multi-line preamble, the script tries several header rows automatically.
- **Zero or blank CPO**: `Затраты на 1 заказ` is left blank if `Заказано, шт.` = 0.
- **Mismatched keys**: Make sure the same article IDs exist in both files. The join is by string article IDs.
- **Multiple rows per article in supplier-goods**: This is expected (per warehouse). The script groups and sums by `Артикул WB`.

## Development notes

- Implemented with `pandas`, `openpyxl`, and `XlsxWriter`.
- The code is in a single file for simplicity and can be embedded into larger pipelines easily.

## License

Use internally at your own discretion. No warranty.
