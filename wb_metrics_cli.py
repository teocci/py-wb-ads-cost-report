import argparse
from pathlib import Path
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# -----------------------------
# Helpers to read raw inputs
# -----------------------------

def read_ads_cost(path: str | Path, sheet_name: int = 0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    df.columns = df.columns.astype(str).str.strip()
    req = ['Товар', 'Артикул товара', 'Сумма затрат']
    miss = [c for c in req if c not in df.columns]
    if miss:
        raise KeyError(f'В файле {Path(path).name} не найдены столбцы: {miss}')
    out = df[req].rename(columns={
        'Товар': 'Наименование',
        'Артикул товара': 'Артикул',
        'Сумма затрат': 'Сумма затрат на рекламу'
    }).copy()
    out['Артикул'] = out['Артикул'].astype(str).str.strip()
    out['Сумма затрат на рекламу'] = pd.to_numeric(
        out['Сумма затрат на рекламу'], errors='coerce'
    ).fillna(0.0)
    return out


def read_supplier_goods(path: str | Path, sheet_name: int = 0) -> pd.DataFrame:
    df = None
    for header in [0, 1, 2, 3, 4, 5]:
        try:
            tmp = pd.read_excel(path, sheet_name=sheet_name,
                                header=header, engine='openpyxl')
            tmp.columns = tmp.columns.astype(str).str.strip()
            if {'Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.'}.issubset(tmp.columns):
                df = tmp
                break
        except Exception:
            continue
    if df is None:
        raise KeyError(f'В файле {Path(path).name} не найдены нужные заголовки')
    df = df[['Артикул WB', 'шт.', 'Сумма заказов минус комиссия WB, руб.']].copy()
    df['Артикул WB'] = df['Артикул WB'].astype(str).str.strip()
    df['шт.'] = pd.to_numeric(df['шт.'], errors='coerce').fillna(0).astype(int)
    df['Сумма заказов минус комиссия WB, руб.'] = pd.to_numeric(
        df['Сумма заказов минус комиссия WB, руб.'],
        errors='coerce'
    ).fillna(0.0)

    grouped = (
        df.groupby('Артикул WB', as_index=False)
          .agg({'шт.': 'sum', 'Сумма заказов минус комиссия WB, руб.': 'sum'})
          .rename(columns={
              'Артикул WB': 'Артикул',
              'шт.': 'Заказано, шт.',
              'Сумма заказов минус комиссия WB, руб.': 'Заказано, руб.'
          })
    )
    return grouped


def parse_date(s: str) -> datetime:
    return datetime.strptime(s, '%Y-%m-%d')


def daterange(start: datetime, end: datetime):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def find_file(data_dir: Path, prefix: str, supplier_id: str, date_str: str) -> Path | None:
    exact = data_dir / f'{prefix}-{supplier_id}-{date_str}.xlsx'
    if exact.exists():
        return exact
    candidates = sorted(
        data_dir.glob(f'{prefix}-{supplier_id}-{date_str}*.xlsx'),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    return candidates[0] if candidates else None


# -----------------------------
# Build daily metrics
# -----------------------------

def build_daily_metrics(
    data_dir: Path,
    supplier_id: str,
    start_date: str,
    end_date: str,
    ads_sheet: int = 0,
    goods_sheet: int = 0
) -> pd.DataFrame:
    start = parse_date(start_date)
    end = parse_date(end_date)
    records: list[pd.DataFrame] = []

    for day in daterange(start, end):
        ds = day.strftime('%Y-%m-%d')
        ads_path = find_file(data_dir, 'mp-ads-cost', supplier_id, ds)
        goods_path = find_file(data_dir, 'wb-supplier-goods', supplier_id, ds)

        if not ads_path and not goods_path:
            continue

        ads_df = (
            read_ads_cost(ads_path, sheet_name=ads_sheet)
            if ads_path else
            pd.DataFrame(columns=['Наименование', 'Артикул', 'Сумма затрат на рекламу'])
        )
        goods_df = (
            read_supplier_goods(goods_path, sheet_name=goods_sheet)
            if goods_path else
            pd.DataFrame(columns=['Артикул', 'Заказано, шт.', 'Заказано, руб.'])
        )

        df = ads_df.merge(goods_df, on='Артикул', how='outer')
        df['Дата'] = ds
        df['Поставщик'] = supplier_id

        df['Наименование'] = df['Наименование'].fillna('')
        df['Сумма затрат на рекламу'] = pd.to_numeric(
            df['Сумма затрат на рекламу'], errors='coerce'
        ).fillna(0.0)
        df['Заказано, шт.'] = pd.to_numeric(
            df['Заказано, шт.'], errors='coerce'
        ).fillna(0).astype(int)
        df['Заказано, руб.'] = pd.to_numeric(
            df['Заказано, руб.'], errors='coerce'
        ).fillna(0.0)

        df['Затраты на 1 заказ'] = (
            df['Сумма затрат на рекламу'] /
            df['Заказано, шт.'].replace(0, np.nan)
        ).round(2)
        df['ROAS'] = (
            df['Заказано, руб.'] /
            df['Сумма затрат на рекламу'].replace(0, np.nan)
        ).round(2)
        df['ACOS'] = (
            df['Сумма затрат на рекламу'] /
            df['Заказано, руб.'].replace(0, np.nan)
        ).round(4)

        records.append(df)

    if not records:
        return pd.DataFrame(columns=[
            'Дата', 'Поставщик', 'Артикул', 'Наименование',
            'Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.',
            'Затраты на 1 заказ', 'ROAS', 'ACOS'
        ])

    daily = pd.concat(records, ignore_index=True)
    daily['Дата'] = pd.to_datetime(daily['Дата'])
    return daily


# -----------------------------
# Export & Charts
# -----------------------------

def export_excel(daily: pd.DataFrame, out_xlsx: Path) -> None:
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    weekly = (
        daily
        .set_index('Дата')
        .groupby('Артикул')[['Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.']]
        .resample('W-MON').sum()
        .reset_index()
    )
    monthly = (
        daily
        .set_index('Дата')
        .groupby('Артикул')[['Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.']]
        .resample('MS').sum()
        .reset_index()
    )
    portfolio = daily.agg({
        'Сумма затрат на рекламу': 'sum',
        'Заказано, руб.': 'sum',
        'Заказано, шт.': 'sum'
    }).to_frame().T
    portfolio['Затраты на 1 заказ'] = (
        portfolio['Сумма затрат на рекламу'] /
        portfolio['Заказано, шт.'].replace(0, np.nan)
    ).round(2)
    portfolio['ROAS'] = (
        portfolio['Заказано, руб.'] /
        portfolio['Сумма затрат на рекламу'].replace(0, np.nan)
    ).round(2)
    portfolio['ACOS'] = (
        portfolio['Сумма затрат на рекламу'] /
        portfolio['Заказано, руб.'].replace(0, np.nan)
    ).round(4)

    by_sku = daily.groupby('Артикул').agg({
        'Сумма затрат на рекламу': 'sum',
        'Заказано, руб.': 'sum',
        'Заказано, шт.': 'sum'
    }).reset_index()
    by_sku['Затраты на 1 заказ'] = (
        by_sku['Сумма затрат на рекламу'] /
        by_sku['Заказано, шт.'].replace(0, np.nan)
    ).round(2)
    by_sku['ROAS'] = (
        by_sku['Заказано, руб.'] /
        by_sku['Сумма затрат на рекламу'].replace(0, np.nan)
    ).round(2)
    by_sku['ACOS'] = (
        by_sku['Сумма затрат на рекламу'] /
        by_sku['Заказано, руб.'].replace(0, np.nan)
    ).round(4)

    top_spend = by_sku.sort_values('Сумма затрат на рекламу', ascending=False).head(20)
    top_roas = by_sku.sort_values('ROAS', ascending=False).head(20)
    best_acos = by_sku.sort_values('ACOS', ascending=True).head(20)

    with pd.ExcelWriter(out_xlsx, engine='xlsxwriter') as writer:
        daily_out = daily.copy()
        daily_out['Дата'] = daily_out['Дата'].dt.strftime('%Y-%m-%d')
        daily_out.to_excel(writer, sheet_name='Daily', index=False)
        weekly.to_excel(writer, sheet_name='Weekly_SKU', index=False)
        monthly.to_excel(writer, sheet_name='Monthly_SKU', index=False)
        portfolio.to_excel(writer, sheet_name='Portfolio', index=False)
        by_sku.to_excel(writer, sheet_name='By_SKU', index=False)
        top_spend.to_excel(writer, sheet_name='Top_Spend', index=False)
        top_roas.to_excel(writer, sheet_name='Top_ROAS', index=False)
        best_acos.to_excel(writer, sheet_name='Best_ACOS', index=False)


def make_charts(daily: pd.DataFrame, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    port = daily.groupby('Дата', as_index=False).agg({
        'Сумма затрат на рекламу': 'sum',
        'Заказано, руб.': 'sum',
        'Заказано, шт.': 'sum'
    })
    port['ROAS'] = port['Заказано, руб.'] / port['Сумма затрат на рекламу'].replace(0, np.nan)
    port['ACOS'] = port['Сумма затрат на рекламу'] / port['Заказано, руб.'].replace(0, np.nan)
    port['CPO'] = port['Сумма затрат на рекламу'] / port['Заказано, шт.'].replace(0, np.nan)

    rolling = port[['Сумма затрат на рекламу', 'Заказано, руб.', 'ROAS', 'ACOS', 'CPO']] \
        .rolling(7, min_periods=1).mean()
    rolling['Дата'] = port['Дата']

    # Charts code unchanged...
    fig = plt.figure()
    plt.plot(port['Дата'], port['Сумма затрат на рекламу'], label='Расходы (дн.)')
    plt.plot(port['Дата'], port['Заказано, руб.'], label='Выручка (дн.)')
    plt.plot(rolling['Дата'], rolling['Сумма затрат на рекламу'], label='Расходы 7дн ср')
    plt.plot(rolling['Дата'], rolling['Заказано, руб.'], label='Выручка 7дн ср')
    plt.legend()
    plt.title('Ежедневные расходы vs выручка')
    plt.xlabel('Дата')
    plt.ylabel('Сумма, руб.')
    fig.autofmt_xdate()
    fig.savefig(out_dir / 'daily_spend_vs_revenue.png', bbox_inches='tight')
    plt.close(fig)
    # ... other charts omitted for brevity


# -----------------------------
# CLI
# -----------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description='Сбор и анализ KPI (CPO, ROAS, ACOS)'
    )
    sub = parser.add_subparsers(dest='cmd', required=True)

    pa = sub.add_parser('analyze', help='Построить метрики и графики за диапазон дат')
    pa.add_argument('--supplier-id', required=True, help='Код поставщика, напр. 3925272')
    pa.add_argument('--start', required=True, help='Дата начала YYYY-MM-DD')
    pa.add_argument('--end', required=True, help='Дата конца YYYY-MM-DD')
    pa.add_argument('--data-dir', default='data', help='Папка с исходными файлами')
    pa.add_argument('--reports-dir', default='reports', help='Папка для результатов')
    pa.add_argument('--ads-sheet', default=0, type=int, help='Лист в файле затрат')
    pa.add_argument('--goods-sheet', default=0, type=int, help='Лист в файле продаж')

    pp = sub.add_parser('performance', help='Формирует сводные показатели по SKU')
    pp.add_argument('--supplier-id', required=True, help='Код поставщика')
    pp.add_argument('--start', required=True, help='Дата начала YYYY-MM-DD')
    pp.add_argument('--end', required=True, help='Дата конца YYYY-MM-DD')
    pp.add_argument('--data-dir', default='data', help='Папка с исходниками')
    pp.add_argument('--reports-dir', default='reports', help='Папка для результатов')
    pp.add_argument('--ads-sheet', default=0, type=int, help='Лист затрат')
    pp.add_argument('--goods-sheet', default=0, type=int, help='Лист продаж')

    args = parser.parse_args()

    if args.cmd == 'analyze':
        data_dir = Path(args.data_dir)
        reports_dir = Path(args.reports_dir)
        charts_dir = reports_dir / 'charts'

        daily = build_daily_metrics(
            data_dir, args.supplier_id, args.start, args.end,
            ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet
        )
        if daily.empty:
            print('Нет данных за указанный период.')
            return
        out_xlsx = reports_dir / f'kpi-summary-{args.supplier_id}-{args.start}_to_{args.end}.xlsx'
        export_excel(daily, out_xlsx)
        make_charts(daily, charts_dir)
        print(f'Excel сохранён: {out_xlsx}')

    elif args.cmd == 'performance':
        data_dir = Path(args.data_dir)
        reports_dir = Path(args.reports_dir)

        daily = build_daily_metrics(
            data_dir, args.supplier_id, args.start, args.end,
            ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet
        )
        if daily.empty:
            print('Нет данных за указанный период.')
            return

        by_sku = daily.groupby('Артикул', as_index=False).agg({
            'Сумма затрат на рекламу': 'sum',
            'Заказано, руб.': 'sum',
            'Заказано, шт.': 'sum'
        })
        by_sku['avg_cpo'] = (
            by_sku['Сумма затрат на рекламу'] /
            by_sku['Заказано, шт.'].replace(0, np.nan)
        ).round(2)
        by_sku['avg_roas'] = (
            by_sku['Заказано, руб.'] /
            by_sku['Сумма затрат на рекламу'].replace(0, np.nan)
        ).round(2)
        by_sku['avg_acos'] = (
            by_sku['Сумма затрат на рекламу'] /
            by_sku['Заказано, руб.'].replace(0, np.nan)
        ).round(4)

        stats = daily.groupby('Артикул').agg(
            std_cpo=('Затраты на 1 заказ','std'),
            std_roas=('ROAS','std'),
            max_daily_spend=('Сумма затрат на рекламу','max'),
            min_daily_spend=('Сумма затрат на рекламу','min')
        )
        perf = by_sku.set_index('Артикул').join(stats).reset_index()

        perf['roas_pctile'] = perf['avg_roas'].rank(pct=True)
        perf['cpo_pctile'] = perf['avg_cpo'].rank(pct=True)
        perf['perf_score'] = (
            (perf['roas_pctile'] + (1 - perf['cpo_pctile'])) / 2
        ).round(4)

        med_roas = perf['avg_roas'].median()
        med_spend = perf['Сумма затрат на рекламу'].median()
        perf['dist'] = (
            (perf['avg_roas'] - med_roas).abs()/med_roas +
            (perf['Сумма затрат на рекламу'] - med_spend).abs()/med_spend
        )
        perf['balanced'] = False
        if not perf.empty:
            perf.loc[perf['dist'].idxmin(), 'balanced'] = True

        out_perf = reports_dir / f'performance-summary-{args.supplier_id}-{args.start}_to_{args.end}.xlsx'
        perf.to_excel(out_perf, index=False)
        print(f'Performance summary saved to: {out_perf}')


if __name__ == '__main__':
    main()
