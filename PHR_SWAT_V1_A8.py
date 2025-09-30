# -*- coding: utf-8 -*-
"""
Created on Tue Jul  1 13:35:14 2025
@author: anegrete (with SWAT enhancements)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
from pathlib import Path
import os
import platform
import subprocess
import logging
import traceback
from datetime import datetime
import threading
import queue

# --- SETUP: Error Logging ---
logging.basicConfig(
    filename='error_log.txt',
    level=logging.ERROR,
    format='%(asctime)s %(levelname)s %(message)s'
)

# --- UI COMPONENT: Yearly Comparison Dialog ---
class YearlyComparisonDialog:
    def __init__(self, parent, years):
        self.top = tk.Toplevel(parent)
        self.years = sorted(years)
        self.result = None
        self.top.title("Yearly Price & Volume Comparison")
        self.top.resizable(False, False)
        main = ttk.Frame(self.top, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        ttk.Label(main, text="Select a year range and a target year to compare.", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=2, pady=5)
        range_frame = ttk.LabelFrame(main, text="Reference Year Range", padding=10)
        range_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky="ew")
        ttk.Label(range_frame, text="From:").grid(row=0, column=0, padx=5, pady=5)
        self.start_year = ttk.Combobox(range_frame, values=self.years, width=10, state="readonly")
        self.start_year.grid(row=0, column=1, padx=5, pady=5)
        self.start_year.set(self.years[0])
        ttk.Label(range_frame, text="To:").grid(row=0, column=2, padx=5, pady=5)
        self.end_year = ttk.Combobox(range_frame, values=self.years, width=10, state="readonly")
        self.end_year.grid(row=0, column=3, padx=5, pady=5)
        self.end_year.set(self.years[-1])
        target_frame = ttk.LabelFrame(main, text="Target Year", padding=10)
        target_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")
        ttk.Label(target_frame, text="Compare with:").grid(row=0, column=0, padx=5, pady=5)
        self.target_year = ttk.Combobox(target_frame, values=self.years, width=10, state="readonly")
        self.target_year.grid(row=0, column=1, padx=5, pady=5)
        self.target_year.set(self.years[-1])
        btn_frm = ttk.Frame(main)
        btn_frm.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frm, text="Compare", command=self.ok).pack(side="left", padx=5)
        ttk.Button(btn_frm, text="Cancel", command=self.cancel).pack(side="left", padx=5)
        self.top.update_idletasks()
        width, height = self.top.winfo_width(), self.top.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
        self.top.geometry(f'{width}x{height}+{x}+{y}')
        self.top.transient(parent)
        self.top.grab_set()
        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        parent.wait_window(self.top)

    def ok(self, event=None):
        try:
            start_y = int(self.start_year.get())
            end_y   = int(self.end_year.get())
            target_y= int(self.target_year.get())
            if start_y > end_y:
                messagebox.showerror("Error", "Start Year must be before or the same as End Year.", parent=self.top)
                return
            self.result = {'start': start_y, 'end': end_y, 'target': target_y}
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please select valid years.", parent=self.top)

    def cancel(self, event=None):
        self.result = None
        self.top.destroy()
        
# --- UI COMPONENT: Fiscal Month Dialog for SWAT ---
class FiscalMonthDialog:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.result = None
        self.top.title("SWAT Analysis Period")
        self.top.resizable(False, False)
        main = ttk.Frame(self.top, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        
        ttk.Label(main, text="Select the fiscal period for Last Paid Price calculation.", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=4, pady=5)

        # Date Entry
        date_frame = ttk.LabelFrame(main, text="Date Range", padding=10)
        date_frame.grid(row=1, column=0, columnspan=4, pady=5, sticky="ew")

        ttk.Label(date_frame, text="Year:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.year_var = tk.StringVar(value=datetime.now().year)
        self.year_entry = ttk.Entry(date_frame, textvariable=self.year_var, width=8)
        self.year_entry.grid(row=0, column=1, padx=5, pady=5)
        
        months = [str(i) for i in range(1, 13)]
        days = [str(i) for i in range(1, 32)]

        ttk.Label(date_frame, text="Start (M/D):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.start_month = ttk.Combobox(date_frame, values=months, width=4, state="readonly")
        self.start_month.grid(row=1, column=1, padx=(5,0), pady=5, sticky="w")
        self.start_month.set('1')
        self.start_day = ttk.Combobox(date_frame, values=days, width=4, state="readonly")
        self.start_day.grid(row=1, column=2, padx=(0,5), pady=5, sticky="w")
        self.start_day.set('1')

        ttk.Label(date_frame, text="End (M/D):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.end_month = ttk.Combobox(date_frame, values=months, width=4, state="readonly")
        self.end_month.grid(row=2, column=1, padx=(5,0), pady=5, sticky="w")
        self.end_month.set(months[-1])
        self.end_day = ttk.Combobox(date_frame, values=days, width=4, state="readonly")
        self.end_day.grid(row=2, column=2, padx=(0,5), pady=5, sticky="w")
        self.end_day.set(days[-1])
        
        # Optional Name
        name_frame = ttk.LabelFrame(main, text="Optional Period Name", padding=10)
        name_frame.grid(row=2, column=0, columnspan=4, pady=5, sticky="ew")
        self.name_var = tk.StringVar()
        ttk.Entry(name_frame, textvariable=self.name_var, width=30).pack(padx=5, pady=5)

        # Buttons
        btn_frm = ttk.Frame(main)
        btn_frm.grid(row=3, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frm, text="Confirm", command=self.ok).pack(side="left", padx=5)
        ttk.Button(btn_frm, text="Cancel", command=self.cancel).pack(side="left", padx=5)
        
        self.top.update_idletasks()
        width, height = self.top.winfo_width(), self.top.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
        self.top.geometry(f'{width}x{height}+{x}+{y}')
        self.top.transient(parent)
        self.top.grab_set()
        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        parent.wait_window(self.top)

    def ok(self, event=None):
        try:
            year = int(self.year_var.get())
            start_m, start_d = int(self.start_month.get()), int(self.start_day.get())
            end_m, end_d = int(self.end_month.get()), int(self.end_day.get())
            
            start_date = datetime(year, start_m, start_d)
            end_date = datetime(year, end_m, end_d)

            if start_date > end_date:
                messagebox.showerror("Error", "Start date cannot be after the end date.", parent=self.top)
                return

            self.result = {
                'start_date': start_date,
                'end_date': end_date,
                'name': self.name_var.get().strip()
            }
            self.top.destroy()
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid date provided. Please check values.\nDetails: {e}", parent=self.top)

    def cancel(self, event=None):
        self.result = None
        self.top.destroy()
        
# --- HELPER FUNCTIONS ---
def find_column(df, aliases, friendly_name):
    aliases_clean = [str(a).strip().lower() for a in aliases]
    for alias, alias_l in zip(aliases, aliases_clean):
        for col in df.columns:
            if str(col).strip().lower() == alias_l:
                logging.info(f'Using column "{col}" for {friendly_name}')
                return col
    logging.warning(f'Could not find column for {friendly_name}. Searched for: {aliases}')
    return None

def open_file(file_path):
    file_path = str(file_path)
    if platform.system() == 'Windows':
        os.startfile(file_path)
    elif platform.system() == 'Darwin':
        subprocess.call(['open', file_path])
    else:
        subprocess.call(['xdg-open', file_path])

# --- DATA PROCESSING PIPELINE ---
def read_and_prepare_data(file_path):
    file_path = Path(file_path)
    file_ext = file_path.suffix.lower()
    read_kw = dict(sheet_name=0)
    if file_ext in [".xlsx", ".xlsm"]:
        read_kw["engine"] = "openpyxl"
    elif file_ext == ".xls":
        read_kw["engine"] = "xlrd"
    elif file_ext == ".xlsb":
        read_kw["engine"] = "pyxlsb"
    elif file_ext == ".ods":
        read_kw["engine"] = "odf"

    id_dtypes = {
        "Part Number": str, "Material": str, "Part": str,
        "Vendor Account Number": str, "Vendor #": str,
        "Vendor Number": str, "Supplier Number": str
    }
    if file_ext == ".csv":
        df = pd.read_csv(file_path, dtype=id_dtypes, keep_default_na=False)
    else:
        df = pd.read_excel(file_path, dtype=id_dtypes, keep_default_na=False, **read_kw)

    df.columns = [str(c).strip() for c in df.columns]

    pstng_col    = find_column(df, ["Pstng Date", "Posting Date", "Post Date"], "posting date")
    amount_col   = find_column(df, ["Amount in PO currency", "USD Invoiced", "Amount"], "amount")
    qty_col      = find_column(df, ["Net Qty in BUoM", "Units", "Quantity"], "quantity")
    part_num_col = find_column(df, ["Part Number", "Material", "Part"], "part number")

    if not all([pstng_col, amount_col, qty_col, part_num_col]):
        missing = [name for name, col in [
            ("Date", pstng_col), ("Amount", amount_col),
            ("Quantity", qty_col), ("Part Number", part_num_col)
        ] if not col]
        raise ValueError(f"Essential column(s) not found: {', '.join(missing)}.")

    df.rename(columns={part_num_col: 'Part Number'}, inplace=True)
    df.dropna(subset=['Part Number'], inplace=True)
    df = df[df['Part Number'].astype(str).str.strip() != '']

    for col in [amount_col, qty_col]:
        df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
    df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
    df[qty_col]    = pd.to_numeric(df[qty_col], errors='coerce')
    df[pstng_col]  = pd.to_datetime(df[pstng_col], errors='coerce')

    # Currency cleanup
    currency_aliases = ["Crcy.1", "Currency", "Crcy", "Curr."]
    found_crcy = find_column(df, currency_aliases, "currency")
    if found_crcy:
        aliases_lower = {a.lower() for a in currency_aliases}
        to_drop = [
            col for col in df.columns
            if col != found_crcy and str(col).strip().lower() in aliases_lower
        ]
        if to_drop:
            logging.info(f"Dropping extra currency columns: {to_drop}")
            df.drop(columns=to_drop, inplace=True)

    rename_map = {
        "Vendor": find_column(df, ["Vendor", "Vendor Name", "Supplier"], "vendor name"),
        "Vendor Number": find_column(df, ["Vendor Account Number", "Vendor #", "Vendor Number", "Supplier Number"], "vendor number"),
        "Plnt": find_column(df, ["Plant", "Plnt"], "plant"),
        "Tr./ev.type": find_column(df, ["Tr./ev.type", "Tr./Ev. type", "Transaction Event Type", "Event Type"], "transaction / event type"),
        "OUn": find_column(df, ["Order Unit", "OUn", "UoM"], "order unit"),
        "Crcy": found_crcy
    }
    inverted_rename_map = {v: k for k, v in rename_map.items() if v is not None}
    df.rename(columns=inverted_rename_map, inplace=True)

    if 'OUn' in df.columns:
        oun_map = df.groupby('Part Number')['OUn'].unique().apply(lambda x: '/'.join(sorted(x)))
        df['Aggregated OUn'] = df['Part Number'].map(oun_map)
    else:
        df['Aggregated OUn'] = 'N/A'

    standard_id_cols = ['Part Number', 'Vendor', 'Vendor Number', 'Aggregated OUn', 'Crcy', 'Plnt', 'Tr./ev.type']
    for col in standard_id_cols:
        if col not in df.columns:
            df[col] = 'N/A'

    denom = df[qty_col].replace(0, np.nan)
    df['P/U'] = (df[amount_col] / denom).replace([np.inf, -np.inf], np.nan)

    return df, standard_id_cols, pstng_col, qty_col

def generate_analysis_tables(df, id_cols, pstng_col, qty_col):
    df_for_calcs = df.dropna(subset=[pstng_col]).copy()
    tables = {}
    df_for_calcs['Manual Date'] = df_for_calcs[pstng_col].apply(lambda d: d.replace(day=1))

    price_id_cols  = id_cols
    volume_id_cols = [c for c in id_cols if c != 'Crcy']
    tables['price_id_cols']  = price_id_cols
    tables['volume_id_cols'] = volume_id_cols

    # Summary & MoM
    raw_summary = pd.pivot_table(df_for_calcs, index=price_id_cols,
                                 columns='Manual Date', values='P/U', aggfunc='mean')
    missing_before_ffill = raw_summary.isna()
    summary_ffill = raw_summary.ffill(axis=1)
    summary_ffill_mask = missing_before_ffill & ~summary_ffill.isna()
    tables['summary_ffill_mask'] = summary_ffill_mask.reset_index(drop=True)
    summary_df = summary_ffill.reset_index()
    summary_df.columns = [
        c.strftime("%m/%d/%Y") if isinstance(c, pd.Timestamp) else c
        for c in summary_df.columns
    ]
    tables['summary'] = summary_df

    mom = summary_ffill.pct_change(axis=1).replace([np.inf, -np.inf], np.nan)
    tables['mom_empty_mask'] = mom.isna()
    mom_df = mom.fillna(0).reset_index()
    mom_df.columns = [
        c.strftime("%m/%d/%Y") if isinstance(c, pd.Timestamp) else c
        for c in mom_df.columns
    ]
    tables['mom'] = mom_df

    # Monthly Volume
    vol_monthly = pd.pivot_table(df_for_calcs, index=volume_id_cols,
                                 columns='Manual Date', values=qty_col, aggfunc='sum')\
                    .fillna(0)
    vol_monthly_df = vol_monthly.reset_index()
    vol_monthly_df.columns = [
        c.strftime("%m/%d/%Y") if isinstance(c, pd.Timestamp) else c
        for c in vol_monthly_df.columns
    ]
    tables['vol_monthly'] = vol_monthly_df

    # Last Paid Price (overall)
    df_unique_last = df_for_calcs.sort_values(by=pstng_col, ascending=False)\
                      .drop_duplicates(subset=price_id_cols)
    last_paid_df = df_unique_last[price_id_cols + [pstng_col, 'P/U']]\
                    .rename(columns={pstng_col: 'Date', 'P/U': 'LastPaidPrice'})\
                    .reset_index(drop=True)
    last_paid_df['Date'] = last_paid_df['Date'].dt.strftime("%m/%d/%Y")
    tables['last_paid'] = last_paid_df

    return tables

def generate_yearly_comparison_tables(df, id_cols, pstng_col, qty_col, parent_window):
    df_for_calcs = df.dropna(subset=[pstng_col, 'P/U', qty_col]).copy()
    if df_for_calcs.empty:
        return {}
    df_for_calcs['Year'] = pd.to_datetime(df_for_calcs[pstng_col]).dt.year
    years = sorted(df_for_calcs['Year'].unique())
    if not years:
        messagebox.showwarning("Comparison Warning",
                               "No valid date data for yearly comparison.",
                               parent=parent_window)
        return {}

    dialog = YearlyComparisonDialog(parent_window, years)
    if not dialog.result:
        return {}
    params = dialog.result
    start_y, end_y, target_y = params['start'], params['end'], params['target']
    years_in_range = list(range(start_y, end_y + 1))
    all_years_needed = sorted(set(years_in_range + [target_y]))

    price_id_cols  = id_cols
    volume_id_cols = [c for c in id_cols if c != 'Crcy']

    yearly_avg_pivot = pd.pivot_table(df_for_calcs, index=price_id_cols,
                                      columns='Year', values='P/U', aggfunc='mean')
    yearly_vol_pivot = pd.pivot_table(df_for_calcs, index=volume_id_cols,
                                      columns='Year', values=qty_col, aggfunc='sum')\
                         .fillna(0)

    for y in all_years_needed:
        if y not in yearly_avg_pivot.columns:
            yearly_avg_pivot[y] = np.nan
        if y not in yearly_vol_pivot.columns:
            yearly_vol_pivot[y] = 0

    yearly_avg_pivot = yearly_avg_pivot[all_years_needed]
    yearly_vol_pivot = yearly_vol_pivot[all_years_needed]

    missing_before_ffill = yearly_avg_pivot.isna()
    yearly_avg_pivot_ffill = yearly_avg_pivot.ffill(axis=1)
    forward_filled_mask = missing_before_ffill & ~yearly_avg_pivot_ffill.isna()

    comparison_dfs = {}
    if target_y in yearly_avg_pivot_ffill.columns:
        for y in years_in_range:
            if y in yearly_avg_pivot_ffill.columns:
                comparison_dfs[f'Price Change % vs {target_y} [{y}]'] = (
                    yearly_avg_pivot_ffill[target_y] /
                    yearly_avg_pivot_ffill[y] - 1
                ).replace([np.inf, -np.inf], np.nan)

    yearly_prices_df  = yearly_avg_pivot_ffill.reset_index()
    yearly_volumes_df = yearly_vol_pivot.reset_index()
    yearly_comparison_df = pd.DataFrame(comparison_dfs,
                                        index=yearly_avg_pivot.index)\
                                .reset_index()

    return {
        'yearly_prices': yearly_prices_df,
        'yearly_volumes': yearly_volumes_df,
        'yearly_comparison': yearly_comparison_df,
        'yearly_ffill_mask': forward_filled_mask.reset_index(drop=True),
    }

def generate_last_paid_period_tables(df, id_cols, pstng_col, parent_window, gen_options):
    df2 = df.dropna(subset=[pstng_col, 'P/U']).copy()
    df2[pstng_col] = pd.to_datetime(df2[pstng_col], errors='coerce')
    df2['Year'] = df2[pstng_col].dt.year
    years = sorted(df2['Year'].dropna().unique())
    if not years:
        messagebox.showwarning(
            "Last-Paid Period Warning",
            "No valid date data for Last-Paid-Period tables.",
            parent=parent_window
        )
        return {}

    dialog = YearlyComparisonDialog(parent_window, years)
    if not dialog.result:
        return {}
    start_y, end_y = dialog.result['start'], dialog.result['end']

    df_range = df2[(df2['Year'] >= start_y) & (df2['Year'] <= end_y)].copy()
    if df_range.empty:
        return {}

    price_id_cols = id_cols
    out = {}

    # --- Yearly pivoted ---
    if gen_options.get('last_paid_year'):
        dfy = df_range.sort_values(by=pstng_col, ascending=False)
        subset_year = price_id_cols + ['Year']
        dfy = dfy.drop_duplicates(subset=subset_year)
        raw_year = dfy.set_index(price_id_cols + ['Year'])['P/U'] \
                      .unstack(level='Year', fill_value=np.nan)
        all_years = list(range(start_y, end_y + 1))
        for y in all_years:
            if y not in raw_year.columns:
                raw_year[y] = np.nan
        raw_year = raw_year[all_years]
        missing = raw_year.isna()
        ffill  = raw_year.ffill(axis=1)
        mask_filled = missing & ~ffill.isna()
        df_year = ffill.reset_index()
        df_year.columns = price_id_cols + [str(y) for y in all_years]
        out['last_paid_yearly']       = df_year
        out['last_paid_yearly_mask']  = mask_filled.reset_index(drop=True)

    # --- Monthly pivoted ---
    if gen_options.get('last_paid_month'):
        dfm = df_range.sort_values(by=pstng_col, ascending=False)
        dfm['YearMonth'] = dfm[pstng_col].dt.to_period('M').astype(str)
        subset_month = price_id_cols + ['YearMonth']
        dfm = dfm.drop_duplicates(subset=subset_month)
        raw_month = dfm.set_index(price_id_cols + ['YearMonth'])['P/U'] \
                       .unstack(level='YearMonth', fill_value=np.nan)
        all_months = pd.period_range(
            start=f"{start_y}-01", end=f"{end_y}-12", freq='M'
        ).astype(str).tolist()
        for m in all_months:
            if m not in raw_month.columns:
                raw_month[m] = np.nan
        raw_month = raw_month[all_months]
        missing_m = raw_month.isna()
        ffill_m   = raw_month.ffill(axis=1)
        mask_filled_m = missing_m & ~ffill_m.isna()
        df_month = ffill_m.reset_index()
        df_month.columns = price_id_cols + all_months
        out['last_paid_monthly']      = df_month
        out['last_paid_monthly_mask'] = mask_filled_m.reset_index(drop=True)

    return out

def write_formatted_excel_report(output_path, tables, gen_options):
    with pd.ExcelWriter(output_path, engine="xlsxwriter",
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        
        wb = writer.book
        align_left       = {'align': 'left'}
        fmt_money        = wb.add_format({'num_format': '$#,##0.0000', **align_left})
        fmt_light_text   = wb.add_format({'font_color': '#A9A9A9', **align_left})
        fmt_light_ffill  = wb.add_format({'font_color': '#A9A9A9', 'num_format': '$#,##0.0000', **align_left})
        fmt_percent      = wb.add_format({'num_format': '0.0%', **align_left})
        fmt_light_percent= wb.add_format({'font_color': '#A9A9A9', 'num_format': '0.0%', **align_left})
        fmt_red_pct      = wb.add_format({'num_format': '0.0%', 'font_color': 'red',   'bold': True, **align_left})
        fmt_green_pct    = wb.add_format({'num_format': '0.0%', 'font_color': 'green', 'bold': True, **align_left})
        fmt_volume       = wb.add_format({'num_format': '#,##0', **align_left})
        fmt_light_volume = wb.add_format({'num_format': '#,##0', 'font_color': '#A9A9A9', **align_left})
      
        # New formats for conditional PPV coloring
        fmt_red_money    = wb.add_format({'num_format': '$#,##0.0000', 'font_color': 'red',   'bold': True, **align_left})
        fmt_green_money  = wb.add_format({'num_format': '$#,##0.0000', 'font_color': 'green', 'bold': True, **align_left})
    
        # --- Write raw data and other sheets ---
        tables['raw_data'].to_excel(writer, sheet_name='Data', index=False)
        if gen_options.get('summary') and 'summary' in tables:
            tables['summary'].to_excel(writer, sheet_name='Summary', index=False)
        if gen_options.get('mom') and 'mom' in tables:
            tables['mom'].to_excel(writer, sheet_name='MoM Change', index=False)
        if 'vol_monthly' in tables:
            tables['vol_monthly'].to_excel(writer, sheet_name='Monthly Volume', index=False)
        if gen_options.get('last_paid') and 'last_paid' in tables:
            tables['last_paid'].to_excel(writer, sheet_name='Last Paid Price', index=False)
        if 'yearly_prices' in tables:
            tables['yearly_prices'].to_excel(writer, sheet_name='Yearly Avg Price', index=False)
        if 'yearly_volumes' in tables:
            tables['yearly_volumes'].to_excel(writer, sheet_name='Yearly Volume', index=False)
        if 'yearly_comparison' in tables:
            tables['yearly_comparison'].to_excel(writer, sheet_name='Yearly Comparison', index=False)
        if 'last_paid_yearly' in tables:
            tables['last_paid_yearly'].to_excel(writer, sheet_name='Last Paid Yearly', index=False)
        if 'last_paid_monthly' in tables:
            tables['last_paid_monthly'].to_excel(writer, sheet_name='Last Paid Monthly', index=False)

        # SWAT (with dynamic sheet name)
        swat_sheet_name_to_write = None
        for key in tables.keys():
            if str(key).startswith('SWAT'):
                swat_sheet_name_to_write = key
                break
        
        if gen_options.get('swat_cost') and swat_sheet_name_to_write:
            tables[swat_sheet_name_to_write].to_excel(writer, sheet_name=swat_sheet_name_to_write, index=False)
    
        # helper to autofit columns and add Excel table objects
        def _autofit_and_add_table(ws, df, table_name, style='Table Style Medium 2'):
            if df.empty:
                return
            for i, col in enumerate(df.columns):
                header_len = len(str(col))
                max_len = df[col].astype(str).str.len().max()
                if pd.isna(max_len):
                    max_len = 0
                ws.set_column(i, i, max(header_len, int(max_len)) + 2)
            n_rows, n_cols = df.shape
            opts = {
                'name': table_name,
                'style': style,
                'columns': [{'header': str(h)} for h in df.columns]
            }
            ws.add_table(0, 0, n_rows, n_cols-1, opts)
    
        # --- Format "Data" sheet ---
        ws_data = writer.sheets['Data']
        raw_df = tables['raw_data']
        # apply date format
        dt_cols = [c for c in raw_df.columns if pd.api.types.is_datetime64_any_dtype(raw_df[c])]
        fmt_date = wb.add_format({'num_format': 'mm/dd/yyyy', **align_left})
        for col in dt_cols:
            idx = raw_df.columns.get_loc(col)
            ws_data.set_column(idx, idx, 12, fmt_date)
        _autofit_and_add_table(ws_data, raw_df, 'DataTbl')
        if 'P/U' in raw_df.columns:
            pu_idx = raw_df.columns.get_loc('P/U')
            ws_data.set_column(pu_idx, pu_idx, None, fmt_money)
    
        # --- Format Summary ---
        if gen_options.get('summary') and 'Summary' in writer.sheets:
            ws = writer.sheets['Summary']
            df = tables['summary']
            mask = tables['summary_ffill_mask']
            id_len = len(tables['price_id_cols'])
            _autofit_and_add_table(ws, df, 'SummaryTbl')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val = df.iat[r, c]
                    ffill = mask.iat[r, c-id_len]
                    if pd.isna(val):
                        ws.write_string(r+1, c, 'NoData', fmt_light_text)
                    elif ffill:
                        ws.write_number(r+1, c, val, fmt_light_ffill)
                    else:
                        ws.write_number(r+1, c, val, fmt_money)
    
        # --- Format MoM Change ---
        if gen_options.get('mom') and 'MoM Change' in writer.sheets:
            ws = writer.sheets['MoM Change']
            df = tables['mom']
            mask = tables['mom_empty_mask']
            id_len = len(tables['price_id_cols'])
            _autofit_and_add_table(ws, df, 'MoMTbl')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val = df.iat[r, c]
                    if mask.iat[r, c-id_len]:
                        ws.write_string(r+1, c, 'NoData', fmt_light_text)
                    else:
                        fmt = fmt_red_pct if val>0 and abs(val)>=0.01 else fmt_green_pct if val<0 and abs(val)>=0.01 else fmt_light_percent
                        ws.write_number(r+1, c, val, fmt)
    
        # --- Format Monthly Volume ---
        if 'Monthly Volume' in writer.sheets:
            ws = writer.sheets['Monthly Volume']
            df = tables['vol_monthly']
            id_len = len(tables['volume_id_cols'])
            _autofit_and_add_table(ws, df, 'MonthlyVolTbl', style='Table Style Medium 3')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val = df.iat[r, c]
                    fmt = fmt_light_volume if pd.notna(val) and val==0 else fmt_volume
                    ws.write_number(r+1, c, val, fmt)
    
        # --- Format Yearly Avg Price ---
        if 'Yearly Avg Price' in writer.sheets:
            ws = writer.sheets['Yearly Avg Price']
            df = tables['yearly_prices']
            mask = tables['yearly_ffill_mask']
            id_len = len(tables['price_id_cols'])
            _autofit_and_add_table(ws, df, 'YearlyAvgPriceTbl')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val = df.iat[r, c]
                    if pd.isna(val):
                        ws.write_string(r+1, c, 'NoData', fmt_light_text)
                    elif mask.iat[r, c-id_len]:
                        ws.write_number(r+1, c, val, fmt_light_ffill)
                    else:
                        ws.write_number(r+1, c, val, fmt_money)
    
        # --- Format Yearly Volume ---
        if 'Yearly Volume' in writer.sheets:
            ws = writer.sheets['Yearly Volume']
            df = tables['yearly_volumes']
            id_len = len(tables['volume_id_cols'])
            _autofit_and_add_table(ws, df, 'YearlyVolTbl', style='Table Style Medium 3')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val = df.iat[r, c]
                    fmt = fmt_light_volume if pd.notna(val) and val==0 else fmt_volume
                    ws.write_number(r+1, c, val, fmt)
    
        # --- Format Yearly Comparison ---
        if 'Yearly Comparison' in writer.sheets:
            ws = writer.sheets['Yearly Comparison']
            df = tables['yearly_comparison']
            id_len = len(tables['price_id_cols'])
            # autofit
            for i,col in enumerate(df.columns):
                hdr=len(str(col)); mx=df[col].astype(str).str.len().max() or 0
                ws.set_column(i,i,max(hdr,int(mx))+2)
            n_rows,n_cols = df.shape
            ws.add_table(0,0,n_rows,n_cols-1,{
                'name':'YearlyChangesTbl','style':'Table Style Medium 9',
                'columns':[{'header':str(h)} for h in df.columns]
            })
            for r in range(n_rows):
                for c in range(id_len,n_cols):
                    val=df.iat[r,c]
                    cell=(r+1,c)
                    if pd.isna(val):
                        ws.write_string(*cell,'NoData',fmt_light_text)
                    else:
                        fmt = fmt_red_pct if val>0 and abs(val)>=0.01 else fmt_green_pct if val<0 and abs(val)>=0.01 else fmt_light_percent
                        ws.write_number(*cell,val,fmt)
    
        # --- Format Last Paid Price (all-time) ---
        if gen_options.get('last_paid') and 'Last Paid Price' in writer.sheets:
            ws = writer.sheets['Last Paid Price']
            df_lp = tables['last_paid']
            _autofit_and_add_table(ws, df_lp, 'LastPaidAllTimeTbl')
            if 'LastPaidPrice' in df_lp:
                idx = df_lp.columns.get_loc('LastPaidPrice')
                for r in range(len(df_lp)):
                    val = df_lp.iat[r,idx]
                    cell=(r+1, idx)
                    if pd.isna(val):
                        ws.write_string(*cell,'NoData',fmt_light_text)
                    else:
                        ws.write_number(*cell,val,fmt_money)
    
        # --- Format Last Paid Yearly ---
        if 'Last Paid Yearly' in writer.sheets:
            ws = writer.sheets['Last Paid Yearly']
            df = tables['last_paid_yearly']; mask = tables['last_paid_yearly_mask']
            id_len = len(tables['price_id_cols'])
            _autofit_and_add_table(ws, df, 'LastPaidYearlyTbl')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val=df.iat[r,c]; cell=(r+1, c)
                    if pd.isna(val):
                        ws.write_string(*cell,'NoData',fmt_light_text)
                    elif mask.iat[r,c-id_len]:
                        ws.write_number(*cell,val,fmt_light_ffill)
                    else:
                        ws.write_number(*cell,val,fmt_money)
    
        # --- Format Last Paid Monthly ---
        if 'Last Paid Monthly' in writer.sheets:
            ws = writer.sheets['Last Paid Monthly']
            df = tables['last_paid_monthly']; mask = tables['last_paid_monthly_mask']
            id_len = len(tables['price_id_cols'])
            _autofit_and_add_table(ws, df, 'LastPaidMonthlyTbl')
            for r in range(len(df)):
                for c in range(id_len, df.shape[1]):
                    val=df.iat[r,c]; cell=(r+1,c)
                    if pd.isna(val):
                        ws.write_string(*cell,'NoData',fmt_light_text)
                    elif mask.iat[r,c-id_len]:
                        ws.write_number(*cell,val,fmt_light_ffill)
                    else:
                        ws.write_number(*cell,val,fmt_money)
    
        # --- Format SWAT Cost analysis (with dynamic sheet name and conditional coloring) ---
        swat_sheet_name = None
        for key in tables.keys():
            if str(key).startswith('SWAT'):
                swat_sheet_name = key
                break
        
        if swat_sheet_name and swat_sheet_name in writer.sheets:
            ws = writer.sheets[swat_sheet_name]
            df = tables[swat_sheet_name]
            _autofit_and_add_table(ws, df, 'SWATCostTbl', style='Table Style Medium 4')
            
            col_idx = {c:i for i,c in enumerate(df.columns)}
            
            # --- Apply formatting row by row for maximum control ---
            for r in range(len(df)):
                # Format universal data columns
                for col_name in ["Vendor", "Vendor Number", "Aggregated OUn", "Crcy"]:
                    if col_name in col_idx:
                        val = df.iat[r, col_idx[col_name]]
                        if val == "Part number not found":
                            ws.write_string(r + 1, col_idx[col_name], val, fmt_light_text)

                # Format Last Paid Price & New Cost
                for col_name in ['Last Paid Price', 'New Cost']:
                    if col_name in col_idx:
                        val = df.iat[r, col_idx[col_name]]
                        cell = (r + 1, col_idx[col_name])
                        if val == "No transactions":
                            ws.write_string(*cell, val, fmt_light_text)
                        elif pd.notna(val):
                            ws.write_number(*cell, val, fmt_money)
                
                # Format PPV with conditional coloring
                if 'PPV' in col_idx:
                    val = df.iat[r, col_idx['PPV']]
                    cell = (r + 1, col_idx['PPV'])
                    if pd.notna(val):
                        if val > 0: ws.write_number(*cell, val, fmt_red_money)
                        elif val < 0: ws.write_number(*cell, val, fmt_green_money)
                        else: ws.write_number(*cell, val, fmt_money)
                    else: # If value is NaN
                        ws.write_string(*cell, "NoData", fmt_light_text)
                
                # Format Fiscal Month Volume
                if 'Fiscal Month Volume' in col_idx:
                    val = df.iat[r, col_idx['Fiscal Month Volume']]
                    cell = (r + 1, col_idx['Fiscal Month Volume'])
                    if pd.notna(val):
                        fmt = fmt_light_volume if val == 0 else fmt_volume
                        ws.write_number(*cell, val, fmt)
                    else: # If value is NaN
                        ws.write_string(*cell, "NoData", fmt_light_text)

                # Format Extended PPV with conditional coloring
                if 'Extended PPV' in col_idx:
                    val = df.iat[r, col_idx['Extended PPV']]
                    cell = (r + 1, col_idx['Extended PPV'])
                    if pd.notna(val):
                        if val > 0: ws.write_number(*cell, val, fmt_red_money)
                        elif val < 0: ws.write_number(*cell, val, fmt_green_money)
                        else: ws.write_number(*cell, val, fmt_money)
                    else: # If value is NaN (will also be NaN if PPV or Volume is NaN)
                        ws.write_string(*cell, "NoData", fmt_light_text)

                # Format % Difference with conditional coloring
                if '% Difference' in col_idx:
                    val = df.iat[r, col_idx['% Difference']]
                    cell = (r + 1, col_idx['% Difference'])
                    if pd.notna(val):
                        if val > 0.0001: ws.write_number(*cell, val, fmt_red_pct)
                        elif val < -0.0001: ws.write_number(*cell, val, fmt_green_pct)
                        else: ws.write_number(*cell, val, fmt_percent)
                    else: # If value is NaN
                        ws.write_string(*cell, "NoData", fmt_light_text)
                    
def process_file_in_background(file_path, gen_options, view_mode, parent_window, result_queue):
    try:
        raw_df, standard_id_cols, pstng_col, qty_col = read_and_prepare_data(file_path)
        simple_id_cols = ['Part Number', 'Vendor', 'Vendor Number', 'Aggregated OUn', 'Crcy']
        id_cols = (simple_id_cols if view_mode=='simple' else standard_id_cols)
        id_cols = [c for c in id_cols if c in raw_df.columns and c!='N/A']

        # 1) Prepare raw_data sheet
        raw_df_out = raw_df.copy()
        audit_col = f"{pstng_col} (dt)"
        raw_df_out[audit_col] = raw_df[pstng_col]
        if pd.api.types.is_datetime64_any_dtype(raw_df_out[pstng_col]):
            raw_df_out[pstng_col] = raw_df_out[pstng_col].dt.strftime("%m/%d/%Y").fillna("Invalid Date")

        # 2) analysis tables
        analysis_tables = generate_analysis_tables(raw_df, id_cols, pstng_col, qty_col)

        # 3) yearly comparison
        yearly_tables = {}
        if gen_options.get('yearly_comp'):
            yearly_tables = generate_yearly_comparison_tables(
                raw_df, id_cols, pstng_col, qty_col, parent_window
            )

        # 4) last-paid period tables
        period_tables = {}
        if gen_options.get('last_paid_year') or gen_options.get('last_paid_month'):
            period_tables = generate_last_paid_period_tables(
                raw_df, id_cols, pstng_col, parent_window, gen_options
            )

        # 5) SWAT Cost analysis
        swat_tbl = {}
        if gen_options.get('swat_cost'):
            # --- Get user input for date range ---
            dialog = FiscalMonthDialog(parent_window)
            if not dialog.result:
                # User cancelled, so we skip the rest of SWAT analysis
                gen_options['swat_cost'] = False # Prevents writing an empty sheet
            else:
                swat_params = dialog.result
                start_date = swat_params['start_date']
                end_date = swat_params['end_date']

                # --- Load CIP cost master with new columns ---
                cip_path = gen_options['cip_file']
                p = Path(cip_path)
                if p.suffix.lower() == '.csv':
                    cip_df = pd.read_csv(cip_path, keep_default_na=False, dtype=str)
                else:
                    cip_df = pd.read_excel(cip_path, keep_default_na=False, dtype=str)
                cip_df.columns = [str(c).strip() for c in cip_df.columns]
                
                part_col    = find_column(cip_df, ["Part Number", "Material", "Part#"], "part number")
                newcost_col = find_column(cip_df, ["New Cost", "Cost"], "new cost")
                pv_col      = find_column(cip_df, ["PV", "Planning Value"], "PV")
                desc_col    = find_column(cip_df, ["Description", "Desc"], "Description")
                
                required_cols = {'Part Number': part_col, 'New Cost': newcost_col}
                if not all(required_cols.values()):
                    missing = [k for k,v in required_cols.items() if not v]
                    raise ValueError(f"CIP master: missing {', '.join(missing)} column(s)")

                # Build the dataframe with all available columns
                swat_base_cols = {part_col: "Part Number", newcost_col: "New Cost"}
                final_cip_cols = ["Part Number", "New Cost"]
                if pv_col:
                    swat_base_cols[pv_col] = "PV"
                    final_cip_cols.append("PV")
                if desc_col:
                    swat_base_cols[desc_col] = "Description"
                    final_cip_cols.append("Description")
                
                cip_df = cip_df[list(swat_base_cols.keys())].copy()
                cip_df.rename(columns=swat_base_cols, inplace=True)
                
                cip_df["New Cost"] = (cip_df["New Cost"].astype(str)
                                      .str.replace(r'[$,]','',regex=True))
                cip_df["New Cost"] = pd.to_numeric(cip_df["New Cost"], errors='coerce')

                # --- Find Last Paid Price and Universal Data in separate steps ---
                if "Tr./ev.type" not in raw_df.columns:
                    raise ValueError("SWAT requires 'Tr./ev.type' column")
                
                df2 = raw_df.copy()
                df2[pstng_col] = pd.to_datetime(df2[pstng_col], errors='coerce')

                # STEP 1: Get Last Paid Price from WITHIN the specified date range
                df_period = df2[
                    (df2["Tr./ev.type"].astype(str).str.strip() == "2") &
                    (df2[pstng_col].notna()) &
                    (df2[pstng_col] >= start_date) &
                    (df2[pstng_col] <= end_date)
                ].copy()
                
                df_period_sorted = df_period.sort_values(by=pstng_col, ascending=False)
                last_paid_in_period = df_period_sorted.drop_duplicates(subset=["Part Number"])
                last_paid_price_df = last_paid_in_period[["Part Number", "P/U"]].rename(columns={"P/U": "Last Paid Price"})

                # NEW STEP 1b: Calculate total volume from WITHIN the specified date range
                volume_df = df_period.groupby('Part Number')[qty_col].sum().reset_index()
                volume_df.rename(columns={qty_col: 'Fiscal Month Volume'}, inplace=True)

                # STEP 2: Get the most recent UNIVERSAL data from the ENTIRE dataset
                df_all_time = df2[(df2["Tr./ev.type"].astype(str).str.strip() == "2") & (df2[pstng_col].notna())].copy()
                df_all_time_sorted = df_all_time.sort_values(by=pstng_col, ascending=False)
                latest_universal_data = df_all_time_sorted.drop_duplicates(subset=["Part Number"])
                
                universal_cols = ["Part Number", "Vendor", "Vendor Number", "Aggregated OUn", "Crcy"]
                for col in universal_cols:
                    if col not in latest_universal_data.columns:
                        latest_universal_data[col] = 'N/A'
                universal_df = latest_universal_data[universal_cols]

                # STEP 3: Merge everything together
                swat = pd.merge(cip_df[final_cip_cols], universal_df, on="Part Number", how="left")
                swat = pd.merge(swat, last_paid_price_df, on="Part Number", how="left")
                swat = pd.merge(swat, volume_df, on="Part Number", how="left") # Merge new volume data

                # STEP 4: Perform all numeric calculations
                swat["PPV"] = swat["Last Paid Price"] - swat["New Cost"]
                swat["% Difference"] = (swat["PPV"] / swat["New Cost"]).replace([np.inf, -np.inf], np.nan)
                # Note: Fiscal Month Volume NaN values are now preserved for formatting
                swat['Extended PPV'] = swat['PPV'] * swat['Fiscal Month Volume'] # NEW: Calculate Extended PPV


                # STEP 5: Replace NaNs with informative text labels
                universal_cols_to_fill = ["Vendor", "Vendor Number", "Aggregated OUn", "Crcy"]
                mask_not_found = swat['Vendor'].isnull()
                for col in universal_cols_to_fill:
                    if col in swat.columns:
                        swat.loc[mask_not_found, col] = "Part number not found"

                mask_no_transactions = swat['Last Paid Price'].isnull() & ~mask_not_found
                swat.loc[mask_no_transactions, 'Last Paid Price'] = "No transactions"
                
                # NEW STEP 6: Re-order columns to the desired final layout
                final_column_order = ['Part Number']
                if 'Description' in swat.columns: final_column_order.append('Description')
                if 'PV' in swat.columns: final_column_order.append('PV')
                
                final_column_order.extend(['Vendor', 'Vendor Number', 'Aggregated OUn', 'Crcy'])
                final_column_order.extend(['Last Paid Price', 'New Cost', 'PPV', 'Fiscal Month Volume', 'Extended PPV', '% Difference'])
                
                # Filter list to only include columns that actually exist, preventing errors
                final_column_order = [col for col in final_column_order if col in swat.columns]
                swat = swat[final_column_order]
                
               # --- Set dynamic sheet name AND STORE THE DATA ---
                sheet_name = "SWAT Cost analysis"
                if swat_params['name']:
                    safe_name = swat_params['name'].replace('/','-').replace('\\','-')[:20] # Clean name
                    sheet_name = f"SWAT - {safe_name}"
                swat_tbl[sheet_name] = swat

        # combine all
        all_tables = {
            'raw_data': raw_df_out,
            **analysis_tables,
            **yearly_tables,
            **period_tables,
            **swat_tbl
        }

        # 6) write output
        p = Path(file_path)
        out_path = p.parent / f"{p.stem}_processed_{view_mode}.xlsx"
        write_formatted_excel_report(out_path, all_tables, gen_options)

        result_queue.put(('success', out_path))
    except Exception:
        logging.error("process_file failed: %s", traceback.format_exc())
        result_queue.put(('error', str(traceback.format_exc())))
        
class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor")
        self.root.resizable(False, False)
        self.result_queue = queue.Queue()
        self.loading_window = None
        self._setup_ui()
        self._center_window()

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # View Type
        view_frame = ttk.LabelFrame(main_frame, text="Analysis Granularity", padding=10)
        view_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
        self.view_mode_var = tk.StringVar(value="detailed")
        ttk.Radiobutton(view_frame, text="Detailed View", variable=self.view_mode_var, value="detailed").pack(anchor='w')
        ttk.Radiobutton(view_frame, text="Simple View",   variable=self.view_mode_var, value="simple").pack(anchor='w')

        # Sheets
        options_frame = ttk.LabelFrame(main_frame, text="Sheets to Generate", padding=10)
        options_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)
        self.gen_summary_var        = tk.BooleanVar(value=True)
        self.gen_mom_var            = tk.BooleanVar(value=True)
        self.gen_last_paid_var      = tk.BooleanVar(value=True)
        self.gen_yearly_comp_var    = tk.BooleanVar(value=True)
        self.gen_last_paid_year_var = tk.BooleanVar(value=False)
        self.gen_last_paid_month_var= tk.BooleanVar(value=False)
        self.gen_swat_var           = tk.BooleanVar(value=False)

        ttk.Checkbutton(options_frame, text="Summary & Monthly Volume", variable=self.gen_summary_var).grid(row=0, column=0, sticky='w', columnspan=2)
        ttk.Checkbutton(options_frame, text="MoM Change",                 variable=self.gen_mom_var).grid(row=1, column=0, sticky='w')
        ttk.Checkbutton(options_frame, text="Last Paid Price",            variable=self.gen_last_paid_var).grid(row=1, column=1, sticky='w')
        ttk.Checkbutton(options_frame, text="Yearly Reports",             variable=self.gen_yearly_comp_var).grid(row=2, column=0, sticky='w', columnspan=2)
        ttk.Checkbutton(options_frame, text="Last Paid by Year",          variable=self.gen_last_paid_year_var).grid(row=3, column=0, sticky='w', columnspan=2)
        ttk.Checkbutton(options_frame, text="Last Paid by Month",         variable=self.gen_last_paid_month_var).grid(row=4, column=0, sticky='w', columnspan=2)
        ttk.Checkbutton(options_frame, text="SWAT Cost Analysis",         variable=self.gen_swat_var).grid(row=5, column=0, sticky='w', columnspan=2)

        # Buttons
        self.process_button = ttk.Button(main_frame, text="Select Excel / CSV File...", command=self.start_processing)
        self.process_button.grid(row=2, column=0, sticky="ew", padx=5, pady=10)
        ttk.Button(main_frame, text="Quit", command=self.root.destroy).grid(row=2, column=1, sticky="ew", padx=5, pady=10)
        main_frame.columnconfigure((0,1), weight=1)

    def _center_window(self):
        self.root.update_idletasks()
        w,h = self.root.winfo_width(), self.root.winfo_height()
        sw,sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{sw//2-w//2}+{sh//2-h//2}")

    def start_processing(self):
        gen_options = {
            'summary':         self.gen_summary_var.get(),
            'mom':             self.gen_mom_var.get(),
            'last_paid':       self.gen_last_paid_var.get(),
            'yearly_comp':     self.gen_yearly_comp_var.get(),
            'last_paid_year':  self.gen_last_paid_year_var.get(),
            'last_paid_month': self.gen_last_paid_month_var.get(),
            'swat_cost':       self.gen_swat_var.get(),
        }
        view_mode = self.view_mode_var.get()

        # main file
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="Select file to process",
            filetypes=[("Supported Files", "*.xlsx;*.xlsb;*.xls;*.xlsm;*.ods;*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        # if SWAT requested, also pick CIP master
        if gen_options['swat_cost']:
            cip_path = filedialog.askopenfilename(
                parent=self.root,
                title="Select CIP Cost master file",
                filetypes=[("Excel or CSV","*.xlsx;*.xls;*.xlsm;*.ods;*.csv"),("All files","*.*")]
            )
            if not cip_path:
                return
            gen_options['cip_file'] = cip_path

        # show loader and disable button
        self._show_loading_window()
        self.process_button.config(state="disabled")

        # start background thread
        threading.Thread(
            target=process_file_in_background,
            args=(file_path, gen_options, view_mode, self.root, self.result_queue),
            daemon=True
        ).start()
        self.root.after(100, self.check_queue)

    def check_queue(self):
        try:
            status, data = self.result_queue.get_nowait()
            self._hide_loading_window()
            self.process_button.config(state="normal")
            if status=='success':
                if messagebox.askyesno("Success!",
                                       f"Process finished.\nOutput file:\n{data}\n\nOpen now?",
                                       parent=self.root):
                    open_file(data)
            else:
                messagebox.showerror("Error",
                                     f"An error occurred:\n{data}\n\nSee 'error_log.txt'.",
                                     parent=self.root)
        except queue.Empty:
            self.root.after(100, self.check_queue)

    def _show_loading_window(self):
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("Please wait…")
        x = self.root.winfo_rootx() + self.root.winfo_width()//4
        y = self.root.winfo_rooty() + self.root.winfo_height()//4
        self.loading_window.geometry(f"+{x}+{y}")
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        self.loading_window.resizable(False, False)
        ttk.Label(self.loading_window, text="Processing file…", padding=20).pack()
        pb = ttk.Progressbar(self.loading_window, mode="indeterminate", length=250)
        pb.pack(padx=20, pady=(0,20))
        pb.start()

    def _hide_loading_window(self):
        if self.loading_window:
            self.loading_window.destroy()
            self.loading_window = None

def main():
    root = tk.Tk()
    ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()