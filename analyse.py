"""
analyse.py
Runs KPI analysis on the CRM data and produces:
  - output/sales_dashboard.xlsx  (formatted Excel dashboard)
  - output/charts/*.png          (matplotlib charts)
"""

import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# ── Paths ──────────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR  = os.path.join(BASE, 'data')
OUT_DIR   = os.path.join(BASE, 'output')
CHART_DIR = os.path.join(OUT_DIR, 'charts')
os.makedirs(CHART_DIR, exist_ok=True)

# ── Colour palette ────────────────────────────────────────────────────────
NAVY   = '1F3557'
TEAL   = '2E86AB'
ORANGE = 'E07B39'
GREEN  = '27AE60'
LIGHT  = 'EAF0FB'
WHITE  = 'FFFFFF'
GREY   = 'D5D8DC'

# ── Load data ─────────────────────────────────────────────────────────────
def load():
    tx = pd.read_csv(os.path.join(DATA_DIR, 'transactions.csv'), parse_dates=['date'])
    cu = pd.read_csv(os.path.join(DATA_DIR, 'customers.csv'))
    return tx, cu

# ── Matplotlib helpers ────────────────────────────────────────────────────
def save_chart(fig, name):
    path = os.path.join(CHART_DIR, name)
    fig.savefig(path, dpi=140, bbox_inches='tight')
    plt.close(fig)
    print(f'  chart → {path}')

def style_ax(ax, title, xlabel='', ylabel='€'):
    ax.set_title(title, fontsize=13, fontweight='bold', color='#1F3557', pad=10)
    ax.set_xlabel(xlabel, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.spines[['top', 'right']].set_visible(False)
    ax.tick_params(labelsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f'€{x:,.0f}' if abs(x) >= 1 else f'{x:.0%}'
    ))

# ── Analysis ──────────────────────────────────────────────────────────────
def analyse(tx, cu):
    won = tx[tx['stage'] == 'Won'].copy()

    # 1. Monthly revenue
    monthly = (
        won.groupby('month')['amount'].sum()
        .reset_index().sort_values('month')
    )
    monthly['mom_growth'] = monthly['amount'].pct_change() * 100

    # 2. Revenue by segment
    seg = won.groupby('segment')['amount'].sum().reset_index()
    seg.columns = ['Segment', 'Revenue']
    seg = seg.sort_values('Revenue', ascending=False)

    # 3. Revenue by product
    prod = won.groupby('product')['amount'].sum().reset_index()
    prod.columns = ['Product', 'Revenue']
    prod = prod.sort_values('Revenue', ascending=False)

    # 4. Rep performance
    rep = won.groupby('sales_rep').agg(
        Revenue=('amount', 'sum'),
        Deals=('transaction_id', 'count')
    ).reset_index().sort_values('Revenue', ascending=False)
    rep.columns = ['Sales Rep', 'Revenue', 'Deals Won']
    rep['Avg Deal (€)'] = (rep['Revenue'] / rep['Deals Won']).round(0)

    # 5. Win rate
    total_by_rep = tx.groupby('sales_rep')['transaction_id'].count()
    won_by_rep   = won.groupby('sales_rep')['transaction_id'].count()
    win_rate = (won_by_rep / total_by_rep * 100).round(1).reset_index()
    win_rate.columns = ['Sales Rep', 'Win Rate (%)']
    rep = rep.merge(win_rate, on='Sales Rep')

    # 6. Revenue by region
    region = won.groupby('region')['amount'].sum().reset_index()
    region.columns = ['Region', 'Revenue']
    region = region.sort_values('Revenue', ascending=False)

    # 7. Top 10 customers
    top_cust = (
        won.groupby('customer_id')['amount'].sum()
        .reset_index().sort_values('amount', ascending=False).head(10)
        .merge(cu[['customer_id', 'company_name', 'segment']], on='customer_id')
    )
    top_cust.columns = ['Customer ID', 'Revenue', 'Company', 'Segment']
    top_cust = top_cust[['Company', 'Segment', 'Revenue']]

    # 8. KPI summary
    total_rev  = won['amount'].sum()
    total_deals = won['transaction_id'].count()
    overall_wr = (len(won) / len(tx) * 100)
    avg_deal   = total_rev / total_deals
    best_month = monthly.loc[monthly['amount'].idxmax(), 'month']
    new_custs  = cu.shape[0]

    kpis = {
        'Total Revenue (€)':   f'€{total_rev:,.0f}',
        'Deals Won':           f'{total_deals:,}',
        'Avg Deal Size (€)':   f'€{avg_deal:,.0f}',
        'Overall Win Rate':    f'{overall_wr:.1f}%',
        'Best Month':          best_month,
        'Total Customers':     str(new_custs),
    }

    return monthly, seg, prod, rep, region, top_cust, kpis

# ── Chart generation ──────────────────────────────────────────────────────
def make_charts(monthly, seg, prod, rep, region):
    months_short = [m[-2:] + '/' + m[2:4] for m in monthly['month']]

    # Chart 1 — Monthly revenue trend
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.fill_between(range(len(monthly)), monthly['amount'], alpha=0.15, color='#2E86AB')
    ax.plot(range(len(monthly)), monthly['amount'], color='#2E86AB', linewidth=2.5, marker='o', markersize=5)
    ax.set_xticks(range(len(monthly)))
    ax.set_xticklabels(months_short, rotation=45, ha='right')
    style_ax(ax, 'Monthly Revenue (Won Deals) — 2024', ylabel='Revenue (€)')
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'€{x/1000:.0f}k'))
    plt.tight_layout()
    save_chart(fig, 'monthly_revenue.png')

    # Chart 2 — Revenue by segment (bar)
    fig, ax = plt.subplots(figsize=(6, 4))
    colours = ['#1F3557', '#2E86AB', '#E07B39']
    bars = ax.bar(seg['Segment'], seg['Revenue'], color=colours[:len(seg)], edgecolor='white', linewidth=0.5)
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 500,
                f'€{bar.get_height()/1000:.0f}k', ha='center', va='bottom', fontsize=8, fontweight='bold')
    ax.set_title('Revenue by Segment', fontsize=13, fontweight='bold', color='#1F3557')
    ax.spines[['top', 'right']].set_visible(False)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'€{x/1000:.0f}k'))
    ax.tick_params(labelsize=9)
    plt.tight_layout()
    save_chart(fig, 'revenue_by_segment.png')

    # Chart 3 — Revenue by product (horizontal bar)
    fig, ax = plt.subplots(figsize=(7, 4))
    colours_prod = plt.cm.Blues(np.linspace(0.4, 0.9, len(prod)))[::-1]
    bars = ax.barh(prod['Product'], prod['Revenue'], color=colours_prod, edgecolor='white')
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'€{x/1000:.0f}k'))
    ax.set_title('Revenue by Product', fontsize=13, fontweight='bold', color='#1F3557')
    ax.spines[['top', 'right']].set_visible(False)
    ax.tick_params(labelsize=9)
    for bar in bars:
        ax.text(bar.get_width() + 200, bar.get_y() + bar.get_height()/2,
                f'€{bar.get_width()/1000:.0f}k', va='center', fontsize=8)
    plt.tight_layout()
    save_chart(fig, 'revenue_by_product.png')

    # Chart 4 — Sales rep performance (grouped bar: revenue + win rate)
    fig, ax1 = plt.subplots(figsize=(8, 4))
    x = np.arange(len(rep))
    ax1.bar(x - 0.2, rep['Revenue'], width=0.35, color='#2E86AB', label='Revenue', alpha=0.9)
    ax1.set_ylabel('Revenue (€)', color='#2E86AB', fontsize=9)
    ax1.tick_params(axis='y', labelcolor='#2E86AB')
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'€{v/1000:.0f}k'))
    ax2 = ax1.twinx()
    ax2.bar(x + 0.2, rep['Win Rate (%)'], width=0.35, color='#E07B39', label='Win Rate %', alpha=0.9)
    ax2.set_ylabel('Win Rate (%)', color='#E07B39', fontsize=9)
    ax2.tick_params(axis='y', labelcolor='#E07B39')
    ax2.set_ylim(0, 100)
    ax1.set_xticks(x)
    ax1.set_xticklabels(rep['Sales Rep'], rotation=20, ha='right', fontsize=8)
    ax1.set_title('Sales Rep Performance — Revenue & Win Rate', fontsize=13, fontweight='bold', color='#1F3557')
    ax1.spines[['top']].set_visible(False)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right', fontsize=8)
    plt.tight_layout()
    save_chart(fig, 'rep_performance.png')

    # Chart 5 — Revenue by region (pie)
    fig, ax = plt.subplots(figsize=(6, 5))
    palette = ['#1F3557', '#2E86AB', '#E07B39', '#27AE60', '#8E44AD']
    wedges, texts, autotexts = ax.pie(
        region['Revenue'], labels=region['Region'],
        colors=palette[:len(region)], autopct='%1.1f%%',
        startangle=140, pctdistance=0.78
    )
    for at in autotexts:
        at.set_fontsize(8)
    ax.set_title('Revenue by Region', fontsize=13, fontweight='bold', color='#1F3557')
    plt.tight_layout()
    save_chart(fig, 'revenue_by_region.png')

    print('All charts saved.')

# ── Excel dashboard ────────────────────────────────────────────────────────
def excel_fill(ws, row, col, value, bold=False, font_colour=None,
               bg=None, align='left', num_fmt=None, border=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, color=font_colour or '000000', size=10)
    if bg:
        cell.fill = PatternFill('solid', fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    if num_fmt:
        cell.number_format = num_fmt
    if border:
        thin = Side(style='thin', color='D5D8DC')
        cell.border = Border(bottom=thin)
    return cell

def build_excel(monthly, seg, prod, rep, region, top_cust, kpis):
    wb = Workbook()
    # ── Sheet 1: Summary Dashboard ─────────────────────────────────────────
    ws = wb.active
    ws.title = 'Dashboard'
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2

    # Title banner
    ws.merge_cells('B1:K2')
    cell = ws['B1']
    cell.value = 'SALES & CRM ANALYTICS DASHBOARD  —  2024 FULL YEAR'
    cell.font = Font(bold=True, size=16, color=WHITE)
    cell.fill = PatternFill('solid', fgColor=NAVY)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # KPI cards (row 4-7)
    kpi_items = list(kpis.items())
    kpi_cols = [2, 4, 6, 8, 10, 12]
    for i, (label, value) in enumerate(kpi_items):
        c = kpi_cols[i]
        # Background
        for r in range(4, 8):
            ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=LIGHT)
            if c + 1 <= 13:
                ws.cell(row=r, column=c+1).fill = PatternFill('solid', fgColor=LIGHT)
        ws.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+1)
        ws.merge_cells(start_row=5, start_column=c, end_row=6, end_column=c+1)
        label_cell = ws.cell(row=4, column=c, value=label)
        label_cell.font = Font(size=9, color='666666')
        label_cell.alignment = Alignment(horizontal='center', vertical='center')
        label_cell.fill = PatternFill('solid', fgColor=LIGHT)
        val_cell = ws.cell(row=5, column=c, value=value)
        val_cell.font = Font(bold=True, size=16, color=NAVY)
        val_cell.alignment = Alignment(horizontal='center', vertical='center')
        val_cell.fill = PatternFill('solid', fgColor=LIGHT)
        ws.row_dimensions[4].height = 18
        ws.row_dimensions[5].height = 28

    # Section: Monthly Revenue (row 9+)
    ws.merge_cells('B9:K9')
    h = ws['B9']
    h.value = 'MONTHLY REVENUE — WON DEALS'
    h.font = Font(bold=True, size=10, color=WHITE)
    h.fill = PatternFill('solid', fgColor=TEAL)
    h.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[9].height = 18

    headers = ['Month', 'Revenue (€)', 'MoM Growth (%)']
    for ci, hdr in enumerate(headers, start=2):
        c = ws.cell(row=10, column=ci, value=hdr)
        c.font = Font(bold=True, size=9, color=NAVY)
        c.fill = PatternFill('solid', fgColor=LIGHT)
        c.alignment = Alignment(horizontal='center', vertical='center')

    for ri, row in enumerate(monthly.itertuples(), start=11):
        ws.cell(row=ri, column=2, value=str(row.month)).font = Font(size=9)
        rev_cell = ws.cell(row=ri, column=3, value=row.amount)
        rev_cell.number_format = '€#,##0'
        rev_cell.font = Font(size=9)
        rev_cell.alignment = Alignment(horizontal='right')
        if pd.notna(row.mom_growth):
            g = ws.cell(row=ri, column=4, value=round(row.mom_growth, 1))
            g.number_format = '#,##0.0"%"'
            g.font = Font(size=9, color=GREEN if row.mom_growth >= 0 else 'C0392B')
            g.alignment = Alignment(horizontal='right')

    # ── Sheet 2: Rep Performance ──────────────────────────────────────────
    ws2 = wb.create_sheet('Rep Performance')
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions['A'].width = 2

    ws2.merge_cells('B1:H2')
    t2 = ws2['B1']
    t2.value = 'SALES REP PERFORMANCE — 2024'
    t2.font = Font(bold=True, size=14, color=WHITE)
    t2.fill = PatternFill('solid', fgColor=NAVY)
    t2.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 28

    rep_headers = ['Sales Rep', 'Revenue (€)', 'Deals Won', 'Avg Deal (€)', 'Win Rate (%)']
    for ci, hdr in enumerate(rep_headers, start=2):
        c = ws2.cell(row=4, column=ci, value=hdr)
        c.font = Font(bold=True, size=9, color=WHITE)
        c.fill = PatternFill('solid', fgColor=TEAL)
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws2.column_dimensions[get_column_letter(ci)].width = 18

    for ri, row in enumerate(rep.itertuples(), start=5):
        bg = LIGHT if ri % 2 == 0 else WHITE
        ws2.cell(row=ri, column=2, value=row._1).font = Font(size=9, bold=True)
        for ci, (val, fmt) in enumerate([
            (row.Revenue, '€#,##0'),
            (row._3, '#,##0'),
            (row._4, '€#,##0'),
            (row._5, '#,##0.0"%"'),
        ], start=3):
            c2 = ws2.cell(row=ri, column=ci, value=val)
            c2.number_format = fmt
            c2.font = Font(size=9)
            c2.fill = PatternFill('solid', fgColor=bg)
            c2.alignment = Alignment(horizontal='right')

    # ── Sheet 3: Segment & Region ─────────────────────────────────────────
    ws3 = wb.create_sheet('Segments & Products')
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions['A'].width = 2

    ws3.merge_cells('B1:J2')
    t3 = ws3['B1']
    t3.value = 'REVENUE BY SEGMENT, PRODUCT & REGION'
    t3.font = Font(bold=True, size=14, color=WHITE)
    t3.fill = PatternFill('solid', fgColor=NAVY)
    t3.alignment = Alignment(horizontal='center', vertical='center')
    ws3.row_dimensions[1].height = 28

    def mini_table(ws_t, start_r, start_c, title, headers, rows, col_widths):
        ws_t.cell(row=start_r, column=start_c, value=title).font = Font(bold=True, size=11, color=NAVY)
        for ci, hdr in enumerate(headers, start=start_c):
            c = ws_t.cell(row=start_r+1, column=ci, value=hdr)
            c.font = Font(bold=True, size=9, color=WHITE)
            c.fill = PatternFill('solid', fgColor=TEAL)
            c.alignment = Alignment(horizontal='center')
            ws_t.column_dimensions[get_column_letter(ci)].width = col_widths[ci - start_c]
        for ri_off, row_data in enumerate(rows):
            bg = LIGHT if ri_off % 2 == 0 else WHITE
            for ci, (val, fmt) in enumerate(row_data, start=start_c):
                c = ws_t.cell(row=start_r+2+ri_off, column=ci, value=val)
                if fmt:
                    c.number_format = fmt
                c.font = Font(size=9)
                c.fill = PatternFill('solid', fgColor=bg)
                c.alignment = Alignment(horizontal='right' if fmt else 'left')

    seg_rows = [((r.Segment, None), (r.Revenue, '€#,##0')) for r in seg.itertuples()]
    mini_table(ws3, 4, 2, 'Revenue by Segment', ['Segment', 'Revenue (€)'],
               seg_rows, [18, 16])

    prod_rows = [((r.Product, None), (r.Revenue, '€#,##0')) for r in prod.itertuples()]
    mini_table(ws3, 4, 6, 'Revenue by Product', ['Product', 'Revenue (€)'],
               prod_rows, [20, 16])

    region_rows = [((r.Region, None), (r.Revenue, '€#,##0')) for r in region.itertuples()]
    mini_table(ws3, 16, 2, 'Revenue by Region', ['Region', 'Revenue (€)'],
               region_rows, [18, 16])

    top_cust_rows = [((r.Company, None), (r.Segment, None), (r.Revenue, '€#,##0'))
                     for r in top_cust.itertuples()]
    mini_table(ws3, 16, 6, 'Top 10 Customers', ['Company', 'Segment', 'Revenue (€)'],
               top_cust_rows, [22, 14, 14])

    out_path = os.path.join(OUT_DIR, 'sales_dashboard.xlsx')
    wb.save(out_path)
    print(f'Excel dashboard → {out_path}')


# ── Main ──────────────────────────────────────────────────────────────────
def main():
    print('Loading data...')
    tx, cu = load()
    print(f'  {len(tx)} transactions, {len(cu)} customers')

    print('Analysing...')
    monthly, seg, prod, rep, region, top_cust, kpis = analyse(tx, cu)

    print('Building charts...')
    make_charts(monthly, seg, prod, rep, region)

    print('Building Excel dashboard...')
    build_excel(monthly, seg, prod, rep, region, top_cust, kpis)

    print('\n=== KPI Summary ===')
    for k, v in kpis.items():
        print(f'  {k}: {v}')

    print('\nDone.')


if __name__ == '__main__':
    main()
