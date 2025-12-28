import os
import sys
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Patch
from matplotlib.colors import to_rgba
from matplotlib.ticker import FuncFormatter, MultipleLocator


def setup_chinese_font():
    # Try user-specified font path first (file path or family name).
    font_path = os.getenv('CHINESE_FONT_PATH')
    if font_path:
        font_file = Path(font_path).expanduser()
        if font_file.exists():
            try:
                font_manager.fontManager.addfont(str(font_file))
                font_name = font_manager.FontProperties(fname=str(font_file)).get_name()
                plt.rcParams['font.sans-serif'] = [font_name]
                return font_name
            except Exception as exc:
                print(f'加载自定义字体失败: {exc}')
        else:
            available = {font.name for font in font_manager.fontManager.ttflist}
            if font_path in available:
                plt.rcParams['font.sans-serif'] = [font_path]
                return font_path

    # Try a local bundled font file in the project directory.
    local_candidates = [
        Path.cwd() / 'NotoSansCJK-Regular.otf',
        Path(__file__).resolve().parent / 'NotoSansCJK-Regular.otf',
    ]
    seen = set()
    for font_file in local_candidates:
        if font_file in seen:
            continue
        seen.add(font_file)
        if not font_file.exists():
            continue
        try:
            font_manager.fontManager.addfont(str(font_file))
            font_name = font_manager.FontProperties(fname=str(font_file)).get_name()
            plt.rcParams['font.sans-serif'] = [font_name]
            return font_name
        except Exception as exc:
            print(f'加载本地字体失败: {exc}')

    # Try known Chinese font families on the system.
    candidates = [
        'Noto Sans CJK SC',
        'Noto Sans SC',
        'Source Han Sans SC',
        'WenQuanYi Zen Hei',
        'Microsoft YaHei',
        'PingFang SC',
        'Hiragino Sans GB',
        'STHeiti',
        'SimHei',
        'Arial Unicode MS',
    ]
    available = {font.name for font in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            plt.rcParams['font.sans-serif'] = [name]
            return name

    print('未检测到可用中文字体。')
    print('请安装中文字体（如 fonts-noto-cjk 或 fonts-wqy-zenhei），或设置 CHINESE_FONT_PATH。')
    return None


def find_data_file(file_path=None, search_dir=Path.cwd()):
    if file_path:
        path = Path(file_path).expanduser()
        if path.is_dir():
            search_dir = path
            path = None
        elif path.exists():
            return path

    csv_candidates = sorted(
        search_dir.glob('*.csv'),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    xlsx_candidates = sorted(
        search_dir.glob('*.xlsx'),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    xls_candidates = sorted(
        search_dir.glob('*.xls'),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if not csv_candidates and not xlsx_candidates and not xls_candidates:
        raise FileNotFoundError('未找到 .csv 或 .xlsx/.xls 数据文件')

    if csv_candidates:
        if len(csv_candidates) > 1 or xlsx_candidates or xls_candidates:
            print('检测到多个数据文件，已优先选择最新 CSV 文件。')
        return csv_candidates[0]

    excel_candidates = xlsx_candidates + xls_candidates
    if len(excel_candidates) > 1:
        print('检测到多个 Excel 数据文件，已默认选择最新文件。')
    return excel_candidates[0]


def load_data(file_path):
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == '.csv':
        return pd.read_csv(path, skiprows=2)
    if suffix in ('.xlsx', '.xls'):
        try:
            return pd.read_excel(path, skiprows=2)
        except ImportError as exc:
            if suffix == '.xlsx' and 'openpyxl' in str(exc).lower():
                raise ImportError(
                    '检测到 Excel 文件，但未安装 openpyxl。'
                    '请运行: pip install openpyxl 或 conda install openpyxl。'
                ) from exc
            if suffix == '.xls' and 'xlrd' in str(exc).lower():
                raise ImportError(
                    '检测到 Excel 文件，但未安装 xlrd。'
                    '请运行: pip install xlrd 或 conda install xlrd。'
                ) from exc
            raise
    raise ValueError('仅支持 .csv 或 .xlsx/.xls 文件')


def preprocess_data(df):
    if '交易地点' not in df.columns:
        return df.copy()
    return df[~df['交易地点'].str.contains('楼|天猫|学生卡成本', na=False)].copy()


def create_bar_chart(location_spending, output_dir):
    fig, ax = plt.subplots(figsize=(12, 6))
    location_spending.plot(kind='bar', ax=ax, color='skyblue', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各食堂消费总额统计（柱形图）'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('交易地点', fontsize=12)
    ax.set_ylabel('总金额 (元)', fontsize=12)
    ax.tick_params(axis='x', rotation=45)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    safe_title = chart_title.replace('/', '_')
    fig.savefig(output_dir / f'{safe_title}.png')


def create_pie_chart(location_spending, output_dir, min_percent=3):
    total = location_spending.sum()
    if total > 0:
        shares = location_spending / total * 100
        major = location_spending[shares >= min_percent]
        minor = location_spending[shares < min_percent]
        if not minor.empty:
            data = pd.concat([major, pd.Series({'其他': minor.sum()})])
        else:
            data = location_spending.copy()
    else:
        data = location_spending.copy()

    labels = list(data.index)
    idx = 0

    def format_pct(pct):
        nonlocal idx
        label = labels[idx] if idx < len(labels) else ''
        idx += 1
        if label == '其他':
            return f'{pct:.1f}%'
        return f'{pct:.1f}%' if pct >= 3 else ''

    fig, ax = plt.subplots(figsize=(9, 9))
    wedges, _, _ = ax.pie(
        data.values,
        autopct=format_pct,
        startangle=90,
        counterclock=False,
        pctdistance=0.7,
    )
    chart_title = '各食堂消费总额统计（扇形图）'
    ax.set_title(chart_title, fontsize=15)
    ax.axis('equal')
    ax.legend(
        wedges,
        data.index,
        title='食堂',
        loc='center left',
        bbox_to_anchor=(1.02, 0.5),
    )
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_monthly_daily_avg_chart(spending_df, output_dir):
    df = spending_df.copy()
    df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce')
    df = df.dropna(subset=['交易时间'])
    if df.empty:
        print('没有可用于统计的交易时间数据，跳过各月份日均食堂消费图表生成。')
        return

    daily_totals = (
        df.groupby(df['交易时间'].dt.date)['交易金额（元）']
        .sum()
        .reset_index()
    )
    daily_totals['年月'] = pd.to_datetime(daily_totals['交易时间']).dt.to_period('M')
    monthly_avg = (
        daily_totals.groupby('年月')['交易金额（元）']
        .mean()
        .sort_index()
    )

    fig, ax = plt.subplots(figsize=(12, 6))
    monthly_avg.plot(kind='bar', ax=ax, color='mediumseagreen', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各月份日均食堂消费'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('年月', fontsize=12)
    ax.set_ylabel('日均消费金额 (元)', fontsize=12)
    ax.tick_params(axis='x', rotation=0)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def prepare_meal_period_data(spending_df):
    df = spending_df.copy()
    df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce')
    df = df.dropna(subset=['交易时间'])
    if df.empty:
        return None, None, '没有可用于统计的交易时间数据'

    minutes = df['交易时间'].dt.hour * 60 + df['交易时间'].dt.minute
    bins = [6 * 60 + 30, 10 * 60, 15 * 60, 19 * 60 + 30, 22 * 60 + 30]
    labels = ['早饭', '午饭', '晚饭', '夜宵']
    df['餐次'] = pd.cut(minutes, bins=bins, labels=labels, right=False)
    df = df.dropna(subset=['餐次'])
    if df.empty:
        return None, labels, '未找到 6:30-22:30 的消费记录'

    return df, labels, None


def create_meal_period_chart(spending_df, output_dir):
    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各餐消费图表生成。')
        return

    meal_spending = (
        df.groupby('餐次')['交易金额（元）']
        .sum()
        .reindex(labels, fill_value=0)
    )

    fig, ax = plt.subplots(figsize=(8, 6))
    meal_spending.plot(kind='bar', ax=ax, color='cornflowerblue', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各餐消费总额'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('餐次', fontsize=12)
    ax.set_ylabel('消费总额 (元)', fontsize=12)
    ax.tick_params(axis='x', rotation=0)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_meal_count_chart(spending_df, output_dir):
    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各餐次数图表生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    meal_counts = (
        df.drop_duplicates(subset=['日期', '餐次'])
        .groupby('餐次')
        .size()
        .reindex(labels, fill_value=0)
    )

    fig, ax = plt.subplots(figsize=(8, 6))
    meal_counts.plot(kind='bar', ax=ax, color='lightsalmon', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        y = height if height >= 0 else height
        va = 'bottom' if height >= 0 else 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.0f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各餐次数统计'
    ax.set_title(chart_title, fontsize=15)
    ax.set_ylabel('用餐次数', fontsize=12)
    ax.set_xlabel('')
    ax.tick_params(axis='x', rotation=0)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_meal_avg_chart(spending_df, output_dir):
    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各餐顿均消费图表生成。')
        return

    meal_spending = (
        df.groupby('餐次')['交易金额（元）']
        .sum()
        .reindex(labels, fill_value=0)
    )
    df['日期'] = df['交易时间'].dt.date
    meal_counts = (
        df.drop_duplicates(subset=['日期', '餐次'])
        .groupby('餐次')
        .size()
        .reindex(labels, fill_value=0)
    )
    meal_avg = meal_spending.divide(meal_counts.where(meal_counts != 0)).fillna(0)

    fig, ax = plt.subplots(figsize=(8, 6))
    meal_avg.plot(kind='bar', ax=ax, color='mediumslateblue', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各餐顿均消费'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('餐次', fontsize=12)
    ax.set_ylabel('顿均消费 (元)', fontsize=12)
    ax.tick_params(axis='x', rotation=0)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_canteen_avg_per_meal_chart(spending_df, output_dir):
    df, _, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各食堂顿均消费图表生成。')
        return
    if '交易地点' not in df.columns:
        print('缺少交易地点字段，跳过各食堂顿均消费图表生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    total_spending = df.groupby('交易地点')['交易金额（元）'].sum()
    meal_counts = (
        df.drop_duplicates(subset=['日期', '餐次', '交易地点'])
        .groupby('交易地点')
        .size()
    )
    avg_spending = total_spending.divide(meal_counts.where(meal_counts != 0)).fillna(0)
    avg_spending = avg_spending.sort_values(ascending=False)
    if avg_spending.empty:
        print('没有可用于统计的食堂数据，跳过各食堂顿均消费图表生成。')
        return

    fig, ax = plt.subplots(figsize=(12, 6))
    avg_spending.plot(kind='bar', ax=ax, color='plum', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各食堂顿均消费'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('交易地点', fontsize=12)
    ax.set_ylabel('顿均消费 (元)', fontsize=12)
    ax.tick_params(axis='x', rotation=45)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_yearly_spending_chart(all_df, output_dir, year=2025):
    df = all_df.copy()
    df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce')
    df = df.dropna(subset=['交易时间'])
    df = df[df['交易时间'].dt.year == year]
    if df.empty:
        print(f'没有 {year} 年的交易记录，跳过消费/充值统计图表生成。')
        return

    spending_events = ['持卡人消费', '离线码在线消费']
    spending_total = df[df['交易事件'].isin(spending_events)]['交易金额（元）'].sum()

    df['月份'] = df['交易时间'].dt.month
    monthly = df[df['交易事件'].isin(spending_events)].groupby('月份')['交易金额（元）'].sum()
    monthly = monthly.reindex(range(1, 13), fill_value=0)

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(
        [f'{month}月' for month in monthly.index],
        monthly.values,
        color='cornflowerblue',
        edgecolor='black',
    )
    for bar in bars:
        height = bar.get_height()
        y = height if height >= 0 else height
        va = 'bottom' if height >= 0 else 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.2f}',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = f'{year}年食堂消费统计'
    ax.set_title(f'{year}年共消费 {spending_total:.2f} 元\n{chart_title}', fontsize=15)
    ax.set_xlabel('月份', fontsize=12)
    ax.set_ylabel('每月消费 (元)', fontsize=12)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_yearly_extremes_table(spending_df, output_dir, year=2025):
    df = spending_df.copy()
    df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce')
    df = df.dropna(subset=['交易时间'])
    df = df[df['交易时间'].dt.year == year]
    if df.empty:
        print(f'没有 {year} 年的刷卡消费记录，跳过“之最”表格生成。')
        return

    df, _, err = prepare_meal_period_data(df)
    if err:
        print(f'{err}，跳过“之最”表格生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    df_sorted = df.sort_values('交易时间')
    first_info = (
        df_sorted.groupby(['日期', '餐次'], as_index=False)
        .first()[['日期', '餐次', '交易时间', '交易地点']]
        .rename(columns={'交易时间': 'first_time', '交易地点': 'first_location'})
    )
    total_amount = (
        df.groupby(['日期', '餐次'], as_index=False)['交易金额（元）']
        .sum()
        .rename(columns={'交易金额（元）': 'total_amount'})
    )
    meal_summary = first_info.merge(total_amount, on=['日期', '餐次'], how='left')
    meal_summary['first_location'] = meal_summary['first_location'].fillna('未知')
    meal_summary['clock_minutes'] = (
        meal_summary['first_time'].dt.hour * 60
        + meal_summary['first_time'].dt.minute
        + meal_summary['first_time'].dt.second / 60
    )

    def pick_first_row(data):
        if data.empty:
            return None
        return data.sort_values('first_time').iloc[0]

    def pick_last_row(data):
        if data.empty:
            return None
        return data.sort_values('first_time', ascending=False).iloc[0]

    def pick_max_amount(data):
        if data.empty:
            return None
        return data.loc[data['total_amount'].idxmax()]

    first_meal = pick_first_row(meal_summary)
    breakfast_meals = meal_summary[meal_summary['餐次'] == '早饭']
    night_meals = meal_summary[meal_summary['餐次'] == '夜宵']
    if not breakfast_meals.empty:
        earliest_breakfast = breakfast_meals.sort_values(['clock_minutes', 'first_time']).iloc[0]
    else:
        earliest_breakfast = None
    if not night_meals.empty:
        latest_night = night_meals.sort_values(['clock_minutes', 'first_time'], ascending=[False, False]).iloc[0]
    else:
        latest_night = None
    most_expensive = pick_max_amount(meal_summary)

    def format_row(row):
        if row is None or pd.isna(row['first_time']):
            return ['无记录', '无记录', '无记录']
        time_str = row['first_time'].strftime('%Y-%m-%d %H:%M:%S')
        return [
            str(row['first_location']),
            f"{row['total_amount']:.2f}",
            time_str,
        ]

    rows = [
        ['刷的第一顿食堂', *format_row(first_meal)],
        ['吃的最早的一顿早饭', *format_row(earliest_breakfast)],
        ['吃的最晚的一顿夜宵', *format_row(latest_night)],
        ['吃的最贵的一顿饭', *format_row(most_expensive)],
    ]
    columns = ['项目', '食堂', '金额(元)', '时间']

    fig, ax = plt.subplots(figsize=(10, 3.8))
    ax.axis('off')
    table = ax.table(
        cellText=rows,
        colLabels=columns,
        cellLoc='center',
        colLoc='center',
        loc='center',
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.6)

    title = f'{year}年刷卡消费之最'
    ax.set_title(title, fontsize=15, pad=10)
    fig.tight_layout()
    fig.savefig(output_dir / f'{title}.png')


def create_meal_first_last_time_chart(spending_df, output_dir):
    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各餐第一笔与最后一笔消费时间图表生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    per_meal_day = (
        df.groupby(['日期', '餐次'])['交易时间']
        .agg(first='min', last='max')
        .reset_index()
    )
    if per_meal_day.empty:
        print('没有可用于统计的餐次数据，跳过各餐第一笔与最后一笔消费时间图表生成。')
        return

    per_meal_day['first_minute'] = (
        per_meal_day['first'].dt.hour * 60
        + per_meal_day['first'].dt.minute
        + per_meal_day['first'].dt.second / 60
    )
    per_meal_day['last_minute'] = (
        per_meal_day['last'].dt.hour * 60
        + per_meal_day['last'].dt.minute
        + per_meal_day['last'].dt.second / 60
    )

    summary = (
        per_meal_day.groupby('餐次')
        .agg(
            first_minute=('first_minute', 'min'),
            last_minute=('last_minute', 'max'),
            avg_first_minute=('first_minute', 'mean'),
        )
        .reindex(labels)
    )

    def minutes_to_label(value):
        rounded = int(round(value))
        hours = max(0, rounded) // 60
        minutes = max(0, rounded) % 60
        return f'{hours:02d}:{minutes:02d}'

    fig, ax = plt.subplots(figsize=(8, 6))
    x = list(range(len(labels)))
    first_color = 'seagreen'
    last_color = 'tomato'
    avg_color = 'royalblue'

    label_fontsize = 11
    label_box = dict(boxstyle='round,pad=0.2', facecolor='white', edgecolor='none', alpha=0.75)

    for i, meal in enumerate(labels):
        if meal not in summary.index or pd.isna(summary.loc[meal, 'first_minute']):
            continue
        first_val = summary.loc[meal, 'first_minute']
        last_val = summary.loc[meal, 'last_minute']
        ax.plot([i, i], [first_val, last_val], color='gray', linewidth=2, zorder=1)
        avg_val = summary.loc[meal, 'avg_first_minute']
        ax.scatter(i, first_val, color=first_color, s=60, zorder=3)
        ax.scatter(i, last_val, color=last_color, s=60, zorder=3)
        if pd.notna(avg_val):
            ax.scatter(i, avg_val, color=avg_color, s=70, marker='D', zorder=4)
        ax.annotate(
            minutes_to_label(first_val),
            xy=(i, first_val),
            xytext=(0, -10),
            textcoords='offset points',
            ha='center',
            va='top',
            fontsize=label_fontsize,
            bbox=label_box,
        )
        ax.annotate(
            minutes_to_label(last_val),
            xy=(i, last_val),
            xytext=(0, 10),
            textcoords='offset points',
            ha='center',
            va='bottom',
            fontsize=label_fontsize,
            bbox=label_box,
        )
        if pd.notna(avg_val):
            ax.annotate(
                f'均 {minutes_to_label(avg_val)}',
                xy=(i, avg_val),
                xytext=(10, 0),
                textcoords='offset points',
                ha='left',
                va='center',
                fontsize=label_fontsize,
                bbox=label_box,
            )

    chart_title = '各餐第一笔与最后一笔消费时间'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('餐次', fontsize=12)
    ax.set_ylabel('时间', fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    ax.scatter([], [], color=first_color, label='第一笔消费')
    ax.scatter([], [], color=last_color, label='最后一笔消费')
    ax.scatter([], [], color=avg_color, marker='D', label='平均开始时间')
    ax.legend(loc='upper left')
    ax.yaxis.set_major_locator(MultipleLocator(60))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: minutes_to_label(v)))

    min_val = summary['first_minute'].min()
    max_val = summary['last_minute'].max()
    if pd.notna(min_val) and pd.notna(max_val):
        lower = 6 * 60
        upper = 22 * 60 + 30
        if lower < upper:
            ax.set_ylim(lower, upper)

    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')


def create_meal_attendance_chart(spending_df, output_dir):
    df_all = spending_df.copy()
    df_all['交易时间'] = pd.to_datetime(df_all['交易时间'], errors='coerce')
    df_all = df_all.dropna(subset=['交易时间'])
    if df_all.empty:
        print('没有可用于统计的交易时间数据，跳过各餐出勤率图表生成。')
        return

    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过各餐出勤率图表生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    min_date = df_all['交易时间'].dt.date.min()
    max_date = df_all['交易时间'].dt.date.max()
    total_days = (max_date - min_date).days + 1
    if total_days == 0:
        print('未找到有效消费日期，跳过各餐出勤率图表生成。')
        return

    meal_days = (
        df.drop_duplicates(subset=['日期', '餐次'])
        .groupby('餐次')
        .size()
        .reindex(labels, fill_value=0)
    )
    attendance = (meal_days / total_days * 100).fillna(0)

    fig, ax = plt.subplots(figsize=(8, 6))
    attendance.plot(kind='bar', ax=ax, color='seagreen', edgecolor='black')

    for bar in ax.patches:
        height = bar.get_height()
        if height >= 0:
            y = height
            va = 'bottom'
        else:
            y = height
            va = 'top'
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f'{height:.1f}%',
            ha='center',
            va=va,
            fontsize=10,
        )

    chart_title = '各餐出勤率（用餐次数/统计周期）'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('餐次', fontsize=12)
    ax.set_ylabel('出勤率 (%)', fontsize=12)
    ax.tick_params(axis='x', rotation=0)
    ax.set_ylim(0, max(100, attendance.max() * 1.15))
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    fig.tight_layout()
    fig.savefig(output_dir / '各餐出勤率.png')


def create_meal_canteen_chart(spending_df, output_dir):
    df, labels, err = prepare_meal_period_data(spending_df)
    if err:
        print(f'{err}，跳过每餐去哪家食堂图表生成。')
        return
    if '交易地点' not in df.columns:
        print('缺少交易地点字段，跳过每餐去哪家食堂图表生成。')
        return

    df['日期'] = df['交易时间'].dt.date
    dedup = df.drop_duplicates(subset=['日期', '餐次', '交易地点'])
    meal_canteen = (
        dedup.groupby(['餐次', '交易地点'])
        .size()
        .unstack(fill_value=0)
        .reindex(labels)
    )
    if meal_canteen.empty or meal_canteen.values.sum() == 0:
        print('没有可用于统计的食堂数据，跳过每餐去哪家食堂图表生成。')
        return

    totals = meal_canteen.sum()
    totals = totals[totals > 0].sort_values(ascending=False)
    canteens = totals.index.tolist()
    need_others = False
    for meal in labels:
        row = meal_canteen.loc[meal]
        row = row[row > 0]
        if row.empty:
            continue
        if row.sum() < 3:
            continue
        if not row[row < 3].empty:
            need_others = True
            break
    if not canteens:
        print('没有可用于统计的食堂数据，跳过每餐去哪家食堂图表生成。')
        return

    cmap = plt.get_cmap('tab20')
    canteen_colors = {
        canteen: cmap(idx % cmap.N)
        for idx, canteen in enumerate(canteens)
    }
    if need_others:
        canteens.append('其他')
        canteen_colors['其他'] = 'lightgray'

    fig, ax = plt.subplots(figsize=(10, 6))
    x = list(range(len(labels)))
    for i, meal in enumerate(labels):
        row = meal_canteen.loc[meal]
        row = row[row > 0]
        if row.sum() >= 3:
            small = row[row < 3]
            row = row[row >= 3]
            if not small.empty:
                row = pd.concat([row, pd.Series({'其他': small.sum()})])
        row = row.sort_values(ascending=False)
        if '其他' in row.index:
            others_value = row.loc['其他']
            row = row.drop('其他')
            row = pd.concat([row, pd.Series({'其他': others_value})])
        bottom = 0
        for canteen, count in row.items():
            color = canteen_colors[canteen]
            ax.bar(
                i,
                count,
                bottom=bottom,
                color=color,
                edgecolor='white',
                linewidth=0.5,
            )
            r, g, b, _ = to_rgba(color)
            luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
            text_color = 'white' if luminance < 0.5 else 'black'
            ax.text(
                i,
                bottom + count / 2,
                f'{int(count)}',
                ha='center',
                va='center',
                fontsize=9,
                color=text_color,
            )
            bottom += count

    chart_title = '每餐去哪家食堂（次数）'
    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel('餐次', fontsize=12)
    ax.set_ylabel('次数', fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.tick_params(axis='x', rotation=0)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    legend_handles = [Patch(facecolor=canteen_colors[c], label=c) for c in canteens]
    ax.legend(
        handles=legend_handles,
        title='食堂',
        loc='center left',
        bbox_to_anchor=(1.02, 0.5),
    )
    fig.tight_layout()
    fig.savefig(output_dir / f'{chart_title}.png')



def main():
    # 1. 读取数据 (跳过前两行元数据)
    input_path = sys.argv[1] if len(sys.argv) > 1 else os.getenv('DATA_FILE')
    file_path = find_data_file(input_path)
    df = load_data(file_path)
    df = preprocess_data(df)

    # 2. 筛选出“持卡人消费”和“离线码在线消费”的记录
    spending_df = df[df['交易事件'].isin(['持卡人消费', '离线码在线消费'])].copy()

    # 3. 按交易地点统计总金额，并按金额从高到低排序
    location_spending = (
        spending_df.groupby('交易地点')['交易金额（元）']
        .sum()
        .sort_values(ascending=False)
    )

    # 4. 设置绘图参数（解决中文显示问题）
    setup_chinese_font()
    plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

    # 5. 保存并展示
    output_dir = Path('Outputs')
    output_dir.mkdir(parents=True, exist_ok=True)
    create_bar_chart(location_spending, output_dir)
    create_pie_chart(location_spending, output_dir)
    create_yearly_spending_chart(df, output_dir)
    create_monthly_daily_avg_chart(spending_df, output_dir)
    create_meal_period_chart(spending_df, output_dir)
    create_meal_count_chart(spending_df, output_dir)
    create_meal_avg_chart(spending_df, output_dir)
    create_canteen_avg_per_meal_chart(spending_df, output_dir)
    create_yearly_extremes_table(spending_df, output_dir)
    create_meal_first_last_time_chart(spending_df, output_dir)
    create_meal_attendance_chart(spending_df, output_dir)
    create_meal_canteen_chart(spending_df, output_dir)

    print("统计结果：")
    print(location_spending)


if __name__ == "__main__":
    main()
