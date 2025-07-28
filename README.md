# WB Ads & Orders Report Builder

This utility merges two Wildberries-related Excel reports and produces a single Excel file with
per‑SKU advertisement costs, orders, and **CPO** (cost per order). It also writes a compact summary block
to the right side of the sheet.

> Script: `build_wb_ads_report.py`

## File layout & naming (automated mode)

Put your files under the `data/` directory using these names:

- **Ads cost**: `ads-cost-<supplier_id>-<YYYY-MM-DD>.xlsx`
- **Supplier goods**: `supplier-goods-<supplier_id>-<YYYY-MM-DD>.xlsx`

The report will be written to the `reports/` directory with the name:

- `wb-ads-report-<supplier_id>-<YYYY-MM-DD>.xlsx`

> The script also tolerates additional suffixes after the date (e.g. exports with hash codes):
> it will fallback to a glob pattern like `supplier-goods-<supplier>-<date>*`. The exact name
> above is recommended for consistency.

## What it does

1. Reads the ads-cost file and keeps: `Товар`, `Артикул товара`, `Сумма затрат на рекламу`.
2. Reads the supplier-goods file, aggregates by `Артикул WB`, and keeps summed:
   - `шт.` → **Заказано, шт.**
   - `Сумма заказов минус комиссия WB, руб.` → **Заказано, руб.**
3. Left‑joins datasets by article and calculates **Затраты на 1 заказ** = cost / qty (blank if qty = 0).
4. Writes a formatted Excel with a summary block:
   - *Финансовые показатели на основе отчетов ВБ*
   - *Итого заказов, руб.* (sum of **Заказано, руб.**)
   - *Итого затрат на рекламу, руб.* (sum of **Сумма затрат на рекламу**)

## Installation

```bash
python -V  # Python 3.10+ (3.11 tested)
pip install -r requirements.txt
```

## Usage

### Automated (recommended)
Organize files:

```
data/
  ads-cost-3925272-2025-06-21.xlsx
  supplier-goods-3925272-2025-06-21.xlsx
reports/   # created automatically
```

Run:

```bash
python build_daily_ads_report.py --supplier-id 3925272 --date 2025-06-21
```

Optional:
```bash
python build_daily_ads_report.py --supplier-id 3925272 --date 2025-06-21 --data-dir "D:/path/to/data" --reports-dir "D:/path/to/reports" --ads-sheet 0 --goods-sheet 0
```

### Manual (explicit paths)
```bash
python build_daily_ads_report.py   --ads "D:/files/Аналитика по товарам от 27.07.2025.xlsx"   --goods "D:/files/supplier-goods-3925272-2025-06-21.xlsx"   --out "D:/files/wb-ads-report-3925272-2025-06-21.xlsx"
```

## Output columns (sheet `Отчет`)

1. `Наименование`
2. `Артикул`
3. `Сумма затрат на рекламу`
4. `Заказано, руб.`
5. `Заказано, шт.`
6. `Затраты на 1 заказ`

## Troubleshooting

- **File not found**: the script looks first for the exact name, then for a glob with extra suffixes;
  ensure the `supplier_id` and `date` are correct.
- **Missing columns**: the script validates inputs and shows what’s missing.
- **Headers shifted by a preamble**: the script tries several header rows automatically.
- **CPO blank**: quantity was zero; this is expected.

## License

Use internally at your own discretion. No warranty.
