import argparse
from datetime import datetime, timedelta
from pathlib import Path
import zipfile

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# -----------------------------
# Readers
# -----------------------------

def read_ads_cost(path: str | Path, sheet_name=0) -> pd.DataFrame:
    """Read and clean ads cost data from Excel file."""
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
    df.columns = df.columns.astype(str).str.strip()

    required = ['Товар', 'Артикул товара', 'Сумма затрат']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f'Missing columns in {Path(path).name}: {missing}')

    out = df[required].rename(columns={
        'Товар': 'Наименование',
        'Артикул товара': 'Артикул',
        'Сумма затрат': 'Сумма затрат на рекламу',
    }).copy()

    out['Артикул'] = out['Артикул'].astype(str).str.strip()
    out['Сумма затрат на рекламу'] = pd.to_numeric(out['Сумма затрат на рекламу'], errors='coerce').fillna(0.0)
    return out


def read_supplier_goods(path: str | Path, sheet_name=0) -> pd.DataFrame:
    """Read and clean supplier goods data from Excel file."""
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
# File discovery helpers
# -----------------------------

def parse_date(s: str) -> datetime:
    """Parse date string in YYYY-MM-DD format."""
    return datetime.strptime(s, '%Y-%m-%d')


def daterange(start: datetime, end: datetime):
    """Generate date range from start to end inclusive."""
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def find_file(data_dir: Path, prefix: str, supplier_id: str, date_str: str) -> Path | None:
    """Find file with exact name or pattern match."""
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
# Daily report builder
# -----------------------------

def build_daily_report(ads_cost_path: str | Path,
                      supplier_goods_path: str | Path,
                      out_path: str | Path,
                      ads_sheet=0,
                      goods_sheet=0) -> dict:
    """Build single daily report combining ads cost and supplier goods data."""
    ads = read_ads_cost(ads_cost_path, sheet_name=ads_sheet)
    goods = read_supplier_goods(supplier_goods_path, sheet_name=goods_sheet)

    # Left join to keep only products that exist in ads data (no garbage data)
    report = ads.merge(goods, on='Артикул', how='left')

    # Ensure numeric columns
    if 'Заказано, шт.' not in report.columns:
        report['Заказано, шт.'] = 0
    if 'Заказано, руб.' not in report.columns:
        report['Заказано, руб.'] = 0.0
    
    report['Заказано, шт.'] = pd.to_numeric(report['Заказано, шт.'], errors='coerce').fillna(0).astype(int)
    report['Заказано, руб.'] = pd.to_numeric(report['Заказано, руб.'], errors='coerce').fillna(0.0)
    report['Сумма затрат на рекламу'] = pd.to_numeric(report['Сумма затрат на рекламу'], errors='coerce').fillna(0.0)

    # Calculate metrics
    qty = pd.to_numeric(report['Заказано, шт.'], errors='coerce')
    cost = pd.to_numeric(report['Сумма затрат на рекламу'], errors='coerce')
    revenue = pd.to_numeric(report['Заказано, руб.'], errors='coerce')
    
    report['Затраты на 1 заказ'] = (cost / qty.replace(0, np.nan)).round(2)
    report['ROAS'] = (revenue / cost.replace(0, np.nan)).round(2)
    report['ACOS'] = (cost / revenue.replace(0, np.nan)).round(4)

    # Arrange columns
    report = report[
        ['Наименование', 'Артикул', 'Сумма затрат на рекламу', 'Заказано, руб.', 
         'Заказано, шт.', 'Затраты на 1 заказ', 'ROAS', 'ACOS']
    ]
    report = report.sort_values(by=['Наименование', 'Артикул'], kind='stable')

    # Calculate totals
    total_orders_rub = float(report['Заказано, руб.'].sum())
    total_ads = float(report['Сумма затрат на рекламу'].sum())
    total_orders_qty = int(report['Заказано, шт.'].sum())

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Write Excel with formatting
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

        # Number formats
        money_cols = [2, 3, 5, 6]  # Ads cost, Revenue, CPO, ROAS
        percent_col = 7  # ACOS
        qty_col = 4
        
        for c in money_cols:
            ws.set_column(c, c, None, money_fmt)
        ws.set_column(qty_col, qty_col, None, qty_fmt)
        ws.set_column(percent_col, percent_col, None, wb.add_format({'num_format': '0.00%'}))

        # Summary block
        start_col = report.shape[1] + 2
        ws.write(0, start_col, 'Финансовые показатели на основе отчетов ВБ', header_fmt)
        ws.write(2, start_col, 'Итого заказов, руб.')
        ws.write(2, start_col + 1, total_orders_rub, money_fmt)
        ws.write(3, start_col, 'Итого затрат на рекламу, руб.')
        ws.write(3, start_col + 1, total_ads, money_fmt)
        ws.write(4, start_col, 'Итого заказов, шт.')
        ws.write(4, start_col + 1, total_orders_qty, qty_fmt)

        ws.freeze_panes(1, 0)

    return {
        'rows': len(report),
        'total_orders_rub': total_orders_rub,
        'total_ads': total_ads,
        'total_orders_qty': total_orders_qty,
        'out_path': str(out_path),
    }


# -----------------------------
# Multi-day metrics builder
# -----------------------------

def build_daily_metrics(
    data_dir: Path,
    supplier_id: str,
    start_date: str,
    end_date: str,
    ads_sheet: int = 0,
    goods_sheet: int = 0
) -> pd.DataFrame:
    """Build daily metrics for date range, keeping only products with ads data."""
    start = parse_date(start_date)
    end = parse_date(end_date)
    records: list[pd.DataFrame] = []

    for day in daterange(start, end):
        ds = day.strftime('%Y-%m-%d')
        ads_path = find_file(data_dir, 'mp-ads-cost', supplier_id, ds)
        goods_path = find_file(data_dir, 'wb-supplier-goods', supplier_id, ds)

        # Skip if no ads data (main requirement)
        if not ads_path:
            continue

        ads_df = read_ads_cost(ads_path, sheet_name=ads_sheet)
        goods_df = (
            read_supplier_goods(goods_path, sheet_name=goods_sheet)
            if goods_path else
            pd.DataFrame(columns=['Артикул', 'Заказано, шт.', 'Заказано, руб.'])
        )

        # Left join to keep only products from ads data
        df = ads_df.merge(goods_df, on='Артикул', how='left')
        df['Дата'] = ds
        df['Поставщик'] = supplier_id

        # Fill missing values
        df['Сумма затрат на рекламу'] = pd.to_numeric(
            df['Сумма затрат на рекламу'], errors='coerce'
        ).fillna(0.0)
        df['Заказано, шт.'] = pd.to_numeric(
            df['Заказано, шт.'], errors='coerce'
        ).fillna(0).astype(int)
        df['Заказано, руб.'] = pd.to_numeric(
            df['Заказано, руб.'], errors='coerce'
        ).fillna(0.0)

        # Calculate metrics
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
# Export functions
# -----------------------------

def export_kpi_summary(daily: pd.DataFrame, out_xlsx: Path) -> None:
    """Export KPI summary with multiple aggregation levels."""
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    # Weekly aggregation
    weekly = (
        daily
        .set_index('Дата')
        .groupby('Артикул')[['Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.']]
        .resample('W-MON').sum()
        .reset_index()
    )
    
    # Monthly aggregation  
    monthly = (
        daily
        .set_index('Дата')
        .groupby('Артикул')[['Сумма затрат на рекламу', 'Заказано, руб.', 'Заказано, шт.']]
        .resample('MS').sum()
        .reset_index()
    )
    
    # Portfolio totals
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

    # By SKU aggregation
    by_sku = daily.groupby('Артикул').agg({
        'Наименование': 'first',
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

    # Top performers
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


def export_performance_summary(daily: pd.DataFrame, out_xlsx: Path) -> None:
    """Export performance analysis with statistical measures."""
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    # Aggregate by SKU
    by_sku = daily.groupby('Артикул', as_index=False).agg({
        'Наименование': 'first',
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

    # Calculate volatility statistics
    stats = daily.groupby('Артикул').agg(
        std_cpo=('Затраты на 1 заказ', 'std'),
        std_roas=('ROAS', 'std'),
        max_daily_spend=('Сумма затрат на рекламу', 'max'),
        min_daily_spend=('Сумма затрат на рекламу', 'min'),
        days_with_data=('Дата', 'count')
    )
    perf = by_sku.set_index('Артикул').join(stats).reset_index()

    # Calculate percentiles and performance scores
    perf['roas_pctile'] = perf['avg_roas'].rank(pct=True)
    perf['cpo_pctile'] = perf['avg_cpo'].rank(pct=True, ascending=False)  # Lower CPO is better
    perf['perf_score'] = (
        (perf['roas_pctile'] + perf['cpo_pctile']) / 2
    ).round(4)

    # Find balanced performers
    med_roas = perf['avg_roas'].median()
    med_spend = perf['Сумма затрат на рекламу'].median()
    perf['dist_from_median'] = (
        (perf['avg_roas'] - med_roas).abs() / max(med_roas, 0.01) +
        (perf['Сумма затрат на рекламу'] - med_spend).abs() / max(med_spend, 0.01)
    )
    perf['is_balanced'] = False
    if not perf.empty and not perf['dist_from_median'].isna().all():
        perf.loc[perf['dist_from_median'].idxmin(), 'is_balanced'] = True

    # Sort by performance score
    perf = perf.sort_values('perf_score', ascending=False)

    perf.to_excel(out_xlsx, index=False)


def make_charts(daily: pd.DataFrame, out_dir: Path, supplier_id: str, start_date: str, end_date: str) -> None:
    """Generate performance charts with date range in filename."""
    out_dir.mkdir(parents=True, exist_ok=True)
    
    # Date range suffix for filenames
    date_suffix = f'{supplier_id}-{start_date}-{end_date}'
    
    # Portfolio daily aggregation
    port = daily.groupby('Дата', as_index=False).agg({
        'Сумма затрат на рекламу': 'sum',
        'Заказано, руб.': 'sum',
        'Заказано, шт.': 'sum'
    })
    port['ROAS'] = port['Заказано, руб.'] / port['Сумма затрат на рекламу'].replace(0, np.nan)
    port['ACOS'] = port['Сумма затрат на рекламу'] / port['Заказано, руб.'].replace(0, np.nan)
    port['CPO'] = port['Сумма затрат на рекламу'] / port['Заказано, шт.'].replace(0, np.nan)

    # 7-day rolling averages
    rolling = port[['Сумма затрат на рекламу', 'Заказано, руб.', 'ROAS', 'ACOS', 'CPO']] \
        .rolling(7, min_periods=1).mean()
    rolling['Дата'] = port['Дата']

    # Chart 1: Daily spend vs revenue
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(port['Дата'], port['Сумма затрат на рекламу'], label='Расходы (дн.)', alpha=0.7)
    ax.plot(port['Дата'], port['Заказано, руб.'], label='Выручка (дн.)', alpha=0.7)
    ax.plot(rolling['Дата'], rolling['Сумма затрат на рекламу'], label='Расходы 7дн ср', linewidth=2)
    ax.plot(rolling['Дата'], rolling['Заказано, руб.'], label='Выручка 7дн ср', linewidth=2)
    ax.set_title('Ежедневные расходы vs выручка')
    ax.set_xlabel('Дата')
    ax.set_ylabel('Сумма, руб.')
    ax.legend()
    fig.autofmt_xdate()
    fig.savefig(out_dir / f'daily_spend_vs_revenue-{date_suffix}.png', bbox_inches='tight', dpi=150)
    plt.close(fig)

    # Chart 2: ROAS trends
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(port['Дата'], port['ROAS'], label='ROAS (дн.)', alpha=0.7)
    ax.plot(rolling['Дата'], rolling['ROAS'], label='ROAS 7дн ср', linewidth=2)
    ax.axhline(y=1.0, color='red', linestyle='--', alpha=0.5, label='Точка безубыточности')
    ax.set_title('Динамика ROAS (Return on Ad Spend)')
    ax.set_xlabel('Дата')
    ax.set_ylabel('ROAS')
    ax.legend()
    fig.autofmt_xdate()
    fig.savefig(out_dir / f'roas_trends-{date_suffix}.png', bbox_inches='tight', dpi=150)
    plt.close(fig)

    # Chart 3: ACOS trends
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(port['Дата'], port['ACOS'] * 100, label='ACOS (дн.) %', alpha=0.7)
    ax.plot(rolling['Дата'], rolling['ACOS'] * 100, label='ACOS 7дн ср %', linewidth=2)
    ax.set_title('Динамика ACOS (Advertising Cost of Sales)')
    ax.set_xlabel('Дата')
    ax.set_ylabel('ACOS %')
    ax.legend()
    fig.autofmt_xdate()
    fig.savefig(out_dir / f'acos_trends-{date_suffix}.png', bbox_inches='tight', dpi=150)
    plt.close(fig)

    # Chart 4: CPO trends
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(port['Дата'], port['CPO'], label='CPO (дн.)', alpha=0.7)
    ax.plot(rolling['Дата'], rolling['CPO'], label='CPO 7дн ср', linewidth=2)
    ax.set_title('Динамика CPO (Cost Per Order)')
    ax.set_xlabel('Дата')
    ax.set_ylabel('CPO, руб.')
    ax.legend()
    fig.autofmt_xdate()
    fig.savefig(out_dir / f'cpo_trends-{date_suffix}.png', bbox_inches='tight', dpi=150)
    plt.close(fig)


# -----------------------------
# Zip functionality
# -----------------------------

def create_zip_package(generated_files: list[Path], charts_dir: Path, supplier_id: str, start_date: str, end_date: str, zips_dir: Path) -> Path:
    """Create zip package with all generated files from this run."""
    zips_dir.mkdir(parents=True, exist_ok=True)
    
    zip_filename = f'wb-analytics-{supplier_id}-{start_date}-{end_date}.zip'
    zip_path = zips_dir / zip_filename
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add generated report files
        for file_path in generated_files:
            if file_path.exists():
                zipf.write(file_path, file_path.name)
        
        # Add chart files for this specific run
        date_suffix = f'{supplier_id}-{start_date}-{end_date}'
        chart_patterns = [
            f'daily_spend_vs_revenue-{date_suffix}.png',
            f'roas_trends-{date_suffix}.png', 
            f'acos_trends-{date_suffix}.png',
            f'cpo_trends-{date_suffix}.png'
        ]
        
        charts_added = 0
        for chart_name in chart_patterns:
            chart_path = charts_dir / chart_name
            if chart_path.exists():
                zipf.write(chart_path, f'charts/{chart_name}')
                charts_added += 1
        
        # Add a summary file
        summary_content = f'''WB Analytics Report Package
Supplier ID: {supplier_id}
Date Range: {start_date} to {end_date}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Files included:
- Daily reports: {len([f for f in generated_files if 'daily-ads-report' in f.name])} files
- KPI summary: {'Yes' if any('kpi-summary' in f.name for f in generated_files) else 'No'}
- Performance summary: {'Yes' if any('performance-summary' in f.name for f in generated_files) else 'No'}  
- Charts: {charts_added} files
'''
        zipf.writestr('README.txt', summary_content)
    
    return zip_path


# -----------------------------
# Command functions
# -----------------------------

def cmd_daily(args):
    """Generate daily reports for each day in range."""
    data_dir = Path(args.data_dir)
    reports_dir = Path(args.reports_dir)
    start = parse_date(args.start)
    end = parse_date(args.end)
    
    generated_files = []
    
    for day in daterange(start, end):
        date_str = day.strftime('%Y-%m-%d')
        ads_path = find_file(data_dir, 'mp-ads-cost', args.supplier_id, date_str)
        goods_path = find_file(data_dir, 'wb-supplier-goods', args.supplier_id, date_str)
        
        if not ads_path:
            print(f'Пропуск {date_str}: Не найдены данные о рекламе')
            continue
            
        out_path = reports_dir / f'daily-ads-report-{args.supplier_id}-{date_str}.xlsx'
        
        try:
            result = build_daily_report(
                ads_path, 
                goods_path or data_dir / 'dummy.xlsx',  # Use dummy if no goods data
                out_path,
                ads_sheet=args.ads_sheet,
                goods_sheet=args.goods_sheet
            )
            generated_files.append(out_path)
            print(f'Создан: {result["out_path"]} (строк: {result["rows"]})')
        except Exception as e:
            print(f'Ошибка при создании {date_str}: {e}')
    
    print(f'\nСоздано ежедневных отчетов: {len(generated_files)} файлов')
    
    # Create zip if requested
    if args.zip and generated_files:
        zip_path = create_zip_package(
            generated_files, 
            reports_dir / 'charts', 
            args.supplier_id, 
            args.start, 
            args.end, 
            Path(args.zips_dir)
        )
        print(f'Создан архив: {zip_path}')


def cmd_analyze(args):
    """Generate KPI summary and charts."""
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

    out_xlsx = reports_dir / f'kpi-summary-{args.supplier_id}-{args.start}-{args.end}.xlsx'
    export_kpi_summary(daily, out_xlsx)
    make_charts(daily, charts_dir, args.supplier_id, args.start, args.end)
    
    print(f'Сводка KPI сохранена: {out_xlsx}')
    print(f'Графики сохранены в: {charts_dir}')
    
    # Create zip if requested
    if args.zip:
        generated_files = [out_xlsx]
        zip_path = create_zip_package(
            generated_files, 
            charts_dir, 
            args.supplier_id, 
            args.start, 
            args.end, 
            Path(args.zips_dir)
        )
        print(f'Создан архив: {zip_path}')


def cmd_performance(args):
    """Generate performance summary."""
    data_dir = Path(args.data_dir)
    reports_dir = Path(args.reports_dir)

    daily = build_daily_metrics(
        data_dir, args.supplier_id, args.start, args.end,
        ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet
    )
    
    if daily.empty:
        print('Нет данных за указанный период.')
        return

    out_xlsx = reports_dir / f'performance-summary-{args.supplier_id}-{args.start}-{args.end}.xlsx'
    export_performance_summary(daily, out_xlsx)
    
    print(f'Сводка эффективности сохранена: {out_xlsx}')
    
    # Create zip if requested
    if args.zip:
        generated_files = [out_xlsx]
        zip_path = create_zip_package(
            generated_files, 
            reports_dir / 'charts', 
            args.supplier_id, 
            args.start, 
            args.end, 
            Path(args.zips_dir)
        )
        print(f'Создан архив: {zip_path}')


def cmd_all(args):
    """Generate all report types."""
    print('=== Создание ежедневных отчетов ===')
    data_dir = Path(args.data_dir)
    reports_dir = Path(args.reports_dir)
    charts_dir = reports_dir / 'charts'
    start = parse_date(args.start)
    end = parse_date(args.end)
    
    all_generated_files = []
    
    # Generate daily reports
    for day in daterange(start, end):
        date_str = day.strftime('%Y-%m-%d')
        ads_path = find_file(data_dir, 'mp-ads-cost', args.supplier_id, date_str)
        goods_path = find_file(data_dir, 'wb-supplier-goods', args.supplier_id, date_str)
        
        if not ads_path:
            print(f'Пропуск {date_str}: Не найдены данные о рекламе')
            continue
            
        out_path = reports_dir / f'daily-ads-report-{args.supplier_id}-{date_str}.xlsx'
        
        try:
            result = build_daily_report(
                ads_path, 
                goods_path or data_dir / 'dummy.xlsx',
                out_path,
                ads_sheet=args.ads_sheet,
                goods_sheet=args.goods_sheet
            )
            all_generated_files.append(out_path)
            print(f'Создан: {result["out_path"]} (строк: {result["rows"]})')
        except Exception as e:
            print(f'Ошибка при создании {date_str}: {e}')
    
    print(f'\nСоздано ежедневных отчетов: {len([f for f in all_generated_files if "daily-ads-report" in f.name])} файлов')
    
    # Generate KPI summary and charts
    print('\n=== Создание сводки KPI и графиков ===')
    daily = build_daily_metrics(
        data_dir, args.supplier_id, args.start, args.end,
        ads_sheet=args.ads_sheet, goods_sheet=args.goods_sheet
    )
    
    if not daily.empty:
        kpi_xlsx = reports_dir / f'kpi-summary-{args.supplier_id}-{args.start}-{args.end}.xlsx'
        export_kpi_summary(daily, kpi_xlsx)
        make_charts(daily, charts_dir, args.supplier_id, args.start, args.end)
        all_generated_files.append(kpi_xlsx)
        
        print(f'Сводка KPI сохранена: {kpi_xlsx}')
        print(f'Графики сохранены в: {charts_dir}')
    
    # Generate performance summary
    print('\n=== Создание сводки эффективности ===')
    if not daily.empty:
        perf_xlsx = reports_dir / f'performance-summary-{args.supplier_id}-{args.start}-{args.end}.xlsx'
        export_performance_summary(daily, perf_xlsx)
        all_generated_files.append(perf_xlsx)
        
        print(f'Сводка эффективности сохранена: {perf_xlsx}')
    
    # Create zip if requested
    if args.zip and all_generated_files:
        zip_path = create_zip_package(
            all_generated_files, 
            charts_dir, 
            args.supplier_id, 
            args.start, 
            args.end, 
            Path(args.zips_dir)
        )
        print(f'\nСоздан архив: {zip_path}')


# -----------------------------
# CLI
# -----------------------------

def validate_date(date_str: str) -> str:
    """Validate YYYY-MM-DD format."""
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return dt.strftime('%Y-%m-%d')
    except ValueError as e:
        raise argparse.ArgumentTypeError('Дата должна быть в формате YYYY-MM-DD') from e


def main():
    parser = argparse.ArgumentParser(description='Единая аналитика ВБ - генерация ежедневных отчетов, сводки KPI и анализ эффективности')
    subparsers = parser.add_subparsers(dest='command', required=True, help='Доступные команды')

    # Common arguments for all commands
    common_args = argparse.ArgumentParser(add_help=False)
    common_args.add_argument('--supplier-id', required=True, help='Код поставщика, например 25169')
    common_args.add_argument('--start', type=validate_date, required=True, help='Дата начала YYYY-MM-DD')
    common_args.add_argument('--end', type=validate_date, required=True, help='Дата конца YYYY-MM-DD')
    common_args.add_argument('--data-dir', default='data', help='Папка с исходными файлами (по умолчанию: ./data)')
    common_args.add_argument('--reports-dir', default='reports', help='Папка для отчетов (по умолчанию: ./reports)')
    common_args.add_argument('--zips-dir', default='zips', help='Папка для архивов (по умолчанию: ./zips)')
    common_args.add_argument('--ads-sheet', default=0, type=int, help='Номер листа в файле затрат на рекламу (по умолчанию: 0)')
    common_args.add_argument('--goods-sheet', default=0, type=int, help='Номер листа в файле товаров поставщика (по умолчанию: 0)')
    common_args.add_argument('--zip', '-z', action='store_true', help='Создать ZIP архив с результатами')

    # Daily command
    parser_daily = subparsers.add_parser('daily', parents=[common_args], help='Создать ежедневные отчеты для каждого дня')
    parser_daily.set_defaults(func=cmd_daily)

    # Analyze command  
    parser_analyze = subparsers.add_parser('analyze', parents=[common_args], help='Создать сводку KPI и графики')
    parser_analyze.set_defaults(func=cmd_analyze)

    # Performance command
    parser_perf = subparsers.add_parser('performance', parents=[common_args], help='Создать анализ эффективности')
    parser_perf.set_defaults(func=cmd_performance)

    # All command
    parser_all = subparsers.add_parser('all', parents=[common_args], help='Создать все типы отчетов')
    parser_all.set_defaults(func=cmd_all)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
