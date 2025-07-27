import argparse
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# -----------------------------
# Readers
# -----------------------------

def read_ads_cost(path: str | Path, sheet_name=0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    df.columns = df.columns.astype(str).str.strip()

    required = ['Товар', 'Артикул товара', 'Сумма затрат на рекламу']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f'Missing columns in {Path(path).name}: {missing}')

    out = df[required].rename(columns={
        'Товар': 'Наименование',
        'Артикул товара': 'Артикул',
        'Сумма затрат на рекламу': 'Сумма затрат на рекламу',
    }).copy()

    out['Артикул'] = out['Артикул'].astype(str).str.strip()
    out['Сумма затрат на рекламу'] = pd.to_numeric(out['Сумма затрат на рекламу'], errors='coerce').fillna(0.0)
    return out


def read_supplier_goods(path: str | Path, sheet_name=0) -> pd.DataFrame:
    # Try several header positions because WB often inserts a preamble row
    df = None
    for header in [0, 1, 2, 3, 4, 5]:
        try:
            tmp = pd.read_excel(path, sheet_name=sheet_name, header=header, engine='openpyxl')
            tmp.columns = tmp.columns.astype(str).str.strip()
            if {'Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.'}.issubset(set(tmp.columns)):
                df = tmp
                break
        except Exception:
            continue

    if df is None:
        raise KeyError(
            f'Cannot find required headers in {Path(path).name}: '
            f'"Артикул WB", "шт.", "Сумма заказов минус комиссия WB, руб."'
        )

    df = df[['Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.']].copy()
    df['Артикул WB'] = df['Артикул WB'].astype(str).str.strip()
    df['шт.'] = pd.to_numeric(df['шт.'], errors='coerce').fillna(0)
    df['Сумма заказов минус комиссия WB, руб.'] = pd.to_numeric(
        df['Сумма заказов минус комиссия WB, руб.'], errors='coerce'
    ).fillna(0.0)

    grouped = (
        df.groupby('Артикул WB', as_index=False)
        .agg({'шт.': 'sum', 'Сумма заказов минус комиссия WB, руб.': 'sum'})
        .rename(columns={
            'Артикул WB': 'Артикул',
            'шт.': 'Заказано, шт.',
            'Сумма заказов минус комиссия WB, руб.': 'Заказано, руб.',
        })
    )
    grouped['Заказано, шт.'] = grouped['Заказано, шт.'].astype(int)
    return grouped


# -----------------------------
# Report builder
# -----------------------------

def build_report(ads_cost_path: str | Path,
                 supplier_goods_path: str | Path,
                 out_path: str | Path,
                 ads_sheet=0,
                 goods_sheet=0) -> dict:
    ads = read_ads_cost(ads_cost_path, sheet_name=ads_sheet)
    goods = read_supplier_goods(supplier_goods_path, sheet_name=goods_sheet)

    # Left join to keep all rows from ads-cost
    report = ads.merge(goods, on='Артикул', how='left')

    # Ensure numeric
    if 'Заказано, шт.' not in report.columns:
        report['Заказано, шт.'] = 0
    if 'Заказано, руб.' not in report.columns:
        report['Заказано, руб.'] = 0.0
    report['Заказано, шт.'] = pd.to_numeric(report['Заказано, шт.'], errors='coerce').fillna(0).astype(int)
    report['Заказано, руб.'] = pd.to_numeric(report['Заказано, руб.'], errors='coerce').fillna(0.0)
    report['Сумма затрат на рекламу'] = pd.to_numeric(report['Сумма затрат на рекламу'], errors='coerce').fillna(0.0)

    # CPO
    qty = pd.to_numeric(report['Заказано, шт.'], errors='coerce')
    cost = pd.to_numeric(report['Сумма затрат на рекламу'], errors='coerce')
    report['Затраты на 1 заказ'] = (cost / qty.replace(0, np.nan)).round(2)

    # Arrange columns
    report = report[
        ['Наименование', 'Артикул', 'Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.', 'Затраты на 1 заказ']]
    report = report.sort_values(by=['Наименование', 'Артикул'], kind='stable')

    # Totals
    total_orders_rub = float(report['Заказано, руб.'].sum())
    total_ads = float(report['Сумма затрат на рекламу'].sum())

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Write Excel
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        report.to_excel(writer, sheet_name='Отчет', index=False, startrow=0, startcol=0)

        wb = writer.book
        ws = writer.sheets['Отчет']

        money_fmt = wb.add_format({'num_format': '#,##0.00'})
        qty_fmt = wb.add_format({'num_format': '0'})
        header_fmt = wb.add_format({'bold': True})

        # Auto width (sample up to 200 rows to size columns)
        for idx, col in enumerate(report.columns):
            max_len = max(len(str(col)), *(len(str(v)) for v in report[col].astype(str).values[:200]))
            ws.set_column(idx, idx, min(max_len + 2, 60))

        # Number formats
        money_cols = [2, 3, 5]  # Сумма затрат на рекламу, Заказано, руб., Затраты на 1 заказ
        qty_col = 4
        for c in money_cols:
            ws.set_column(c, c, None, money_fmt)
        ws.set_column(qty_col, qty_col, None, qty_fmt)

        # Summary block
        start_col = report.shape[1] + 2
        ws.write(0, start_col, 'Финансовые показатели на основе отчетов ВБ', header_fmt)
        ws.write(2, start_col, 'Итого заказов, руб.')
        ws.write(2, start_col + 1, total_orders_rub, money_fmt)
        ws.write(3, start_col, 'Итого затрат на рекламу, руб.')
        ws.write(3, start_col + 1, total_ads, money_fmt)

        ws.freeze_panes(1, 0)

    return {
        'rows': len(report),
        'total_orders_rub': total_orders_rub,
        'total_ads': total_ads,
        'out_path': str(out_path),
    }


# -----------------------------
# File discovery helpers
# -----------------------------

def validate_date(date_str: str) -> str:
    """Validate YYYY-MM-DD and return normalized string."""
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return dt.strftime('%Y-%m-%d')
    except ValueError as e:
        raise argparse.ArgumentTypeError('Date must be in YYYY-MM-DD format') from e


def resolve_input_path(data_dir: Path, prefix: str, supplier_id: str, date_str: str) -> Path:
    """
    Build expected path: data/<prefix>-<supplier_id>-<date>.xlsx
    If not present, fall back to: data/<prefix>-<supplier_id>-<date>*.xlsx (first by mtime desc).
    """
    exact = data_dir / f'{prefix}-{supplier_id}-{date_str}.xlsx'
    if exact.exists():
        return exact

    # Fallback glob to tolerate extra suffixes
    candidates = sorted(data_dir.glob(f'{prefix}-{supplier_id}-{date_str}*.xlsx'), key=lambda p: p.stat().st_mtime,
                        reverse=True)
    if candidates:
        return candidates[0]
    raise FileNotFoundError(
        f'File not found: {exact} (and no matches for pattern {prefix}-{supplier_id}-{date_str}*.xlsx)')


def default_paths(data_dir: str | Path, supplier_id: str, date_str: str) -> tuple[Path, Path]:
    data_dir = Path(data_dir)
    ads_path = resolve_input_path(data_dir, 'ads-cost', supplier_id, date_str)
    goods_path = resolve_input_path(data_dir, 'supplier-goods', supplier_id, date_str)
    return ads_path, goods_path


def default_out_path(reports_dir: str | Path, supplier_id: str, date_str: str) -> Path:
    reports_dir = Path(reports_dir)
    reports_dir.mkdir(parents=True, exist_ok=True)
    return reports_dir / f'wb-ads-report-{supplier_id}-{date_str}.xlsx'


# -----------------------------
# CLI
# -----------------------------

def main():
    parser = argparse.ArgumentParser(description='Merge WB supplier goods with ads cost and compute CPO.')
    # Mode A: explicit paths
    parser.add_argument('--ads', help='Explicit path to ads-cost Excel file')
    parser.add_argument('--goods', help='Explicit path to supplier-goods Excel file')
    parser.add_argument('--out', help='Explicit output Excel path')

    # Mode B: directory automation
    parser.add_argument('--supplier-id', help='Supplier ID, e.g. 3925272')
    parser.add_argument('--date', type=validate_date, help='Date in YYYY-MM-DD, e.g. 2025-06-21')
    parser.add_argument('--data-dir', default='data', help='Directory with input files (default: ./data)')
    parser.add_argument('--reports-dir', default='reports', help='Directory for outputs (default: ./reports)')

    # Optional sheet names/indices
    parser.add_argument('--ads-sheet', default=0, help='Sheet for ads-cost file (default first)')
    parser.add_argument('--goods-sheet', default=0, help='Sheet for supplier-goods file (default first)')

    args = parser.parse_args()

    # Determine input paths
    if args.ads and args.goods:
        ads_path = Path(args.ads)
        goods_path = Path(args.goods)
        # Derive supplier/date for default output naming if not provided
        supplier_id = args.supplier_id or 'unknown'
        date_str = args.date or datetime.today().strftime('%Y-%m-%d')
        out_path = Path(args.out) if args.out else default_out_path(args.reports_dir, supplier_id, date_str)
    else:
        if not (args.supplier_id and args.date):
            parser.error(
                'Provide either both --ads and --goods, or use --supplier-id and --date with --data-dir/--reports-dir.')
        supplier_id = args.supplier_id
        date_str = args.date
        ads_path, goods_path = default_paths(args.data_dir, supplier_id, date_str)
        out_path = Path(args.out) if args.out else default_out_path(args.reports_dir, supplier_id, date_str)

    result = build_report(ads_path, goods_path, out_path, ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet)
    print(f'Saved: {result["out_path"]} (rows: {result["rows"]})')
    print(f'Итого заказов, руб.: {result["total_orders_rub"]:.2f}')
    print(f'Итого затрат на рекламу, руб.: {result["total_ads"]:.2f}')


if __name__ == "__main__":
    main()
