import argparse
import os

import numpy as np
import pandas as pd


def read_ads_cost(path: str, sheet_name=0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    df.columns = df.columns.astype(str).str.strip()

    # Expected column names
    required = ['Товар', 'Артикул товара', 'Сумма затрат на рекламу']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f'В файле {os.path.basename(path)} отсутствуют столбцы: {missing}')

    out = df[required].rename(columns={
        'Товар': 'Наименование',
        'Артикул товара': 'Артикул',
        'Сумма затрат на рекламу': 'Сумма затрат на рекламу',
    }).copy()

    out['Артикул'] = out['Артикул'].astype(str).str.strip()
    out['Сумма затрат на рекламу'] = pd.to_numeric(out['Сумма затрат на рекламу'], errors='coerce').fillna(0.0)
    return out


def _try_read_with_header(path, header_row):
    try:
        df = pd.read_excel(path, header=header_row, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception:
        return None


def read_supplier_goods(path: str, sheet_name=0) -> pd.DataFrame:
    # Try several header positions because WB often inserts a preamble row
    candidates = [0, 1, 2, 3, 4]
    df = None
    for h in candidates:
        df = pd.read_excel(path, sheet_name=sheet_name, header=h, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()
        if {'Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.'}.issubset(set(df.columns)):
            break
        df = None
    if df is None:
        raise KeyError(
            f'Не удалось найти заголовки "Артикул WB", "шт.", "Сумма заказов минус комиссия WB, руб." '
            f'в файле {os.path.basename(path)}. Проверьте формат.'
        )

    # Keep only needed columns and aggregate by article
    df = df[['Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.']].copy()
    df['Артикул WB'] = df['Артикул WB'].astype(str).str.strip()
    df['шт.'] = pd.to_numeric(df['шт.'], errors='coerce').fillna(0)
    df['Сумма заказов минус комиссия WB, руб.'] = pd.to_numeric(df['Сумма заказов минус комиссия WB, руб.'],
                                                                errors='coerce').fillna(0.0)

    grouped = (
        df.groupby('Артикул WB', as_index=False)
        .agg({'шт.': 'sum', 'Сумма заказов минус комиссия WB, руб.': 'sum'})
        .rename(columns={
            'Артикул WB': 'Артикул',
            'шт.': 'Заказано, шт.',
            'Сумма заказов минус комиссия WB, руб.': 'Заказано, руб.',
        })
    )
    # Ensure integer for qty but keep float for money
    grouped['Заказано, шт.'] = grouped['Заказано, шт.'].astype(int)
    return grouped


def build_report(ads_cost_path: str, supplier_goods_path: str, out_path: str, ads_sheet=0, goods_sheet=0):
    ads = read_ads_cost(ads_cost_path, sheet_name=ads_sheet)
    goods = read_supplier_goods(supplier_goods_path, sheet_name=goods_sheet)

    # Left join to keep all rows from ads-cost
    report = ads.merge(goods, on='Артикул', how='left')

    # Fill empties for metrics
    if 'Заказано, шт.' not in report.columns:
        report['Заказано, шт.'] = 0
    if 'Заказано, руб.' not in report.columns:
        report['Заказано, руб.'] = 0.0

    report['Заказано, шт.'] = report['Заказано, шт.'].fillna(0).astype(int)
    report['Заказано, руб.'] = pd.to_numeric(report['Заказано, руб.'], errors='coerce').fillna(0.0)

    # CPO: cost per order; blank if qty == 0
    # qty = report['Заказано, шт.'].replace(0, pd.NA)
    # report['Затраты на 1 заказ'] = (report['Сумма затрат на рекламу'] / qty).round(2)
    qty = pd.to_numeric(report["Заказано, шт."], errors="coerce")
    cost = pd.to_numeric(report["Сумма затрат на рекламу"], errors="coerce")
    report["Затраты на 1 заказ"] = (cost / qty.replace(0, np.nan)).round(2)
    # Keep NaN as blank in Excel

    # Sort by name for readability
    report = report[
        ['Наименование', 'Артикул', 'Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.', 'Затраты на 1 заказ']]
    report = report.sort_values(by=['Наименование', 'Артикул'], kind='stable')

    # Totals
    total_orders_rub = float(report['Заказано, руб.'].sum())
    total_ads = float(report['Сумма затрат на рекламу'].sum())

    # Write to Excel with a summary block on the right
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        report.to_excel(writer, sheet_name='Отчет', index=False, startrow=0, startcol=0)

        wb = writer.book
        ws = writer.sheets['Отчет']

        money_fmt = wb.add_format({'num_format': '#,##0.00'})
        qty_fmt = wb.add_format({'num_format': '0'})
        header_fmt = wb.add_format({'bold': True})

        # Auto width
        for idx, col in enumerate(report.columns):
            max_len = max(len(str(col)), *(len(str(v)) for v in report[col].astype(str).values[:200]))
            ws.set_column(idx, idx, min(max_len + 2, 60))

        # Apply number formats
        money_cols = [2, 3, 5]  # Сумма затрат на рекламу, Заказано, руб., Затраты на 1 заказ
        qty_col = 4  # Заказано, шт.

        # Determine last row
        last_row = len(report) + 1  # +1 because header at row 0

        for c in money_cols:
            ws.set_column(c, c, None, money_fmt)
        ws.set_column(qty_col, qty_col, None, qty_fmt)

        # Summary block to the right (one blank column gap)
        start_col = report.shape[1] + 2  # e.g., if 6 cols => start at col index 8 (column I)
        ws.write(0, start_col, 'Финансовые показатели на основе отчетов ВБ', header_fmt)
        ws.write(2, start_col, 'Итого заказов, руб.')
        ws.write(2, start_col + 1, total_orders_rub, money_fmt)
        ws.write(3, start_col, 'Итого затрат на рекламу, руб.')
        ws.write(3, start_col + 1, total_ads, money_fmt)

        # Freeze first row
        ws.freeze_panes(1, 0)

    return {
        'rows': len(report),
        'total_orders_rub': total_orders_rub,
        'total_ads': total_ads,
        'out_path': out_path,
    }


def main():
    parser = argparse.ArgumentParser(description='Слить данные затрат на рекламу с отчетом WB и посчитать CPO.')
    parser.add_argument('--ads', required=True, help='Путь к файлу "Аналитика по товарам ...xlsx"')
    parser.add_argument('--goods', required=True, help='Путь к файлу "supplier-goods-....xlsx"')
    parser.add_argument('--out', default='wb_ads_report.xlsx', help='Куда сохранить итоговый отчет')
    parser.add_argument('--ads-sheet', default=0, help='Лист для файла затрат (по умолчанию первый)')
    parser.add_argument('--goods-sheet', default=0, help='Лист для файла WB (по умолчанию первый)')
    args = parser.parse_args()

    result = build_report(args.ads, args.goods, args.out, ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet)
    print(f'Готово: {result["out_path"]} (строк: {result["rows"]})')
    print(f'Итого заказов, руб.: {result["total_orders_rub"]:.2f}')
    print(f'Итого затрат на рекламу, руб.: {result["total_ads"]:.2f}')


if __name__ == '__main__':
    main()
