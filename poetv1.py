import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.gridspec as gridspec
import xlsxwriter
import os
from scipy.stats import skew, kurtosis
import yfinance as yf

# GLOBAL SETTINGS
INPUT_DIR = "input"
OUTPUT_DIR = "output"
CONFIG_FILE = "Poetstrategies.csv"

# TEXT CONFIG
HEADER_TITLE = "ArchDAP"
HEADER_SUB = "Hypothetical Performance"
DISCLAIMER = "THIS COMPOSITE PERFORMANCE RECORD AND STATICS BELOW ARE HYPOTHETICAL.\nPLEASE VIEW THE IMPORTANT RISK DISCLAIMER REGARDING THIS PORTFOLIO."

def process_portfolio_group(group_name, group_data):
    """
    Blends multiple strategies using dynamic rebalancing based on start dates.
    Uses 'Net P&L %' to ensure correct scaling regardless of backtest position size.
    """
    print(f"--- Processing Portfolio: {group_name} ({len(group_data)} strategies) ---")

    # LOAD AND ALIGN EQUITY CURVES
    all_series = {}
    asset_meta = {} 
    
    header_title = group_data.iloc[0]['Title']
    header_sub = group_data.iloc[0]['Subtitle']
    risk_free_rate = float(group_data.iloc[0]['Risk_Free_Rate'])

    # Total Target Capital (Denominator for weights)
    total_target_capital = group_data['Allocated_Capital'].sum()

    for idx, row in group_data.iterrows():
        filename = row['Filename']
        target_allocation = float(row['Allocated_Capital'])
        start_date_str = str(row['Start_Date']) if pd.notnull(row['Start_Date']) else None
        
        file_path = os.path.join(INPUT_DIR, filename)
        
        if not os.path.exists(file_path):
            print(f"WARNING: {filename} not found. Skipping.")
            continue
            
        try:
            # Load Data
            df = pd.read_csv(file_path)
            
            # Determine Date Column
            date_col = 'Date/Time' if 'Date/Time' in df.columns else 'Date'
            df[date_col] = pd.to_datetime(df[date_col])
            
            # Filter Exits
            df_exits = df[df['Type'].str.contains('Exit', na=False)].copy()
            
            # USE P&L % INSTEAD OF USD
            pct_col = next((c for c in df.columns if "Net P&L %" in c), None)
            
            if pct_col:
                # Convert percent to decimal
                df_exits['Return'] = pd.to_numeric(df_exits[pct_col], errors='coerce') / 100.0
                daily_ret = df_exits.set_index(date_col)['Return'].resample('D').sum().fillna(0)
            else:
                print(f"Error: 'Net P&L %' column not found in {filename}. Cannot process reliably.")
                continue
            
            all_series[filename] = daily_ret
            
            # Parse start date
            if start_date_str and start_date_str.lower() != 'nan':
                s_date = pd.to_datetime(start_date_str)
            else:
                s_date = pd.Timestamp.min # Active from beginning
                
            asset_meta[filename] = {
                'target_alloc': target_allocation,
                'start_date': s_date,
                'name': row['Subtitle'] 
            }
            
        except Exception as e:
            print(f"ERROR reading {filename}: {e}")
            return

    if not all_series:
        print("Error: No valid data found.")
        return

    # Dynamic Blending Logic
    full_idx = pd.DatetimeIndex([])
    for s in all_series.values():
        full_idx = full_idx.union(s.index)
    full_idx = full_idx.sort_values()
    
    # DataFrame of Daily Returns (Not P&L USD)
    df_ret = pd.DataFrame(index=full_idx)
    for name, s in all_series.items():
        df_ret[name] = s
    df_ret = df_ret.fillna(0) 

    # Calculate Portfolio Equity Day by Day
    current_capital = total_target_capital 
    portfolio_equity_curve = [current_capital]
    portfolio_dates = [full_idx[0] - pd.Timedelta(days=1)]
    
    asset_names = list(all_series.keys())
    active_mask = pd.DataFrame(False, index=df_ret.index, columns=asset_names)
    
    for name in asset_names:
        start = asset_meta[name]['start_date']
        if start == pd.Timestamp.min:
            active_mask[name] = True
        else:
            active_mask.loc[df_ret.index >= start, name] = True
            
    target_allocs = np.array([asset_meta[n]['target_alloc'] for n in asset_names])
    
    for dt, row in df_ret.iterrows():
        is_active = active_mask.loc[dt].values
        
        if not is_active.any():
            portfolio_equity_curve.append(current_capital)
            portfolio_dates.append(dt)
            continue
            
        # Re-Normalize Weights based on who/what is active
        active_allocs = target_allocs * is_active
        total_active_alloc = np.sum(active_allocs)
        
        # leverage inactive capital
        if total_active_alloc == 0:
            portfolio_equity_curve.append(current_capital)
            portfolio_dates.append(dt)
            continue

        weights = active_allocs / total_active_alloc
        
        # Get returns for today
        day_rets = row[asset_names].values
        
        # Portfolio Weighted Return
        port_ret = np.sum(weights * day_rets)
        
        current_capital = current_capital * (1 + port_ret)
        portfolio_equity_curve.append(current_capital)
        portfolio_dates.append(dt)

    portfolio_equity = pd.Series(portfolio_equity_curve, index=portfolio_dates)
    portfolio_equity = portfolio_equity.iloc[1:] 

    # GENERATE REPORT
    generate_factsheet_from_equity(
        portfolio_equity, 
        total_target_capital, 
        group_name, 
        header_title, 
        header_sub, 
        risk_free_rate,
        group_data 
    )

def generate_factsheet_from_equity(equity_series, initial_capital, report_name, header_title, header_sub, risk_free_rate, portfolio_constituents):
    
    # Paths
    output_excel = os.path.join(OUTPUT_DIR, f"{report_name}_Dashboard.xlsx")
    output_pdf = os.path.join(OUTPUT_DIR, f"{report_name}_Factsheet.pdf")
    
    # CALCULATIONS

    # Resample to Monthly
    df_monthly_equity = equity_series.resample('ME').last().ffill().to_frame(name='Equity')
    df_monthly_equity['VAMI'] = (df_monthly_equity['Equity'] / initial_capital) * 1000
    df_monthly_returns = df_monthly_equity['Equity'].pct_change().fillna(0)
    
    clean_monthly_returns = df_monthly_returns.replace([np.inf, -np.inf], np.nan).dropna()
    if not clean_monthly_returns.empty and clean_monthly_returns.iloc[0] == 0 and len(clean_monthly_returns) > 1:
         clean_monthly_returns = clean_monthly_returns.iloc[1:]

    # Metrics
    start_date = equity_series.index.min()
    end_date = equity_series.index.max()
    years = (end_date - start_date).days / 365.25
    final_equity = equity_series.iloc[-1]
    
    if years > 0:
        cagr = (final_equity / initial_capital) ** (1 / years) - 1
    else:
        cagr = 0

    # Drawdowns (EOM & Peak)
    df_monthly_equity['Peak'] = df_monthly_equity['Equity'].cummax()
    df_monthly_equity['Drawdown'] = (df_monthly_equity['Equity'] / df_monthly_equity['Peak']) - 1
    max_drawdown_eom = df_monthly_equity['Drawdown'].min()

    hf_peak = equity_series.cummax()
    hf_dd = (equity_series / hf_peak) - 1
    max_drawdown_im = hf_dd.min()

    # SPY BENCHMARK
    print("Fetching S&P 500 Data...")
    sp_drawdown = pd.Series(dtype=float)
    sp_vami = pd.Series(dtype=float)
    try:
        buffer_start = start_date - pd.Timedelta(days=40)
        sp_data_raw = yf.download("SPY", start=buffer_start, end=end_date, progress=False)
        if not sp_data_raw.empty:
            if isinstance(sp_data_raw.columns, pd.MultiIndex):
                try: sp_s = sp_data_raw['Adj Close']
                except: sp_s = sp_data_raw['Close']
                if isinstance(sp_s, pd.DataFrame): sp_s = sp_s.squeeze()
            else:
                sp_s = sp_data_raw['Adj Close'] if 'Adj Close' in sp_data_raw.columns else sp_data_raw['Close']

            sp_monthly = sp_s.resample('ME').last().ffill()
            sp_monthly = sp_monthly[sp_monthly.index >= df_monthly_equity.index[0]]
            sp_monthly = sp_monthly[sp_monthly.index <= df_monthly_equity.index[-1]]
            
            sp_peak = sp_monthly.cummax()
            sp_drawdown = (sp_monthly / sp_peak) - 1
            if not sp_monthly.empty:
                sp_vami = (sp_monthly / sp_monthly.iloc[0]) * 1000
                
            common_idx = df_monthly_equity.index.intersection(sp_drawdown.index)
            sp_drawdown = sp_drawdown.loc[common_idx]
            sp_vami = sp_vami.loc[common_idx]
    except Exception as e:
        print(f"Warning: SPY data error: {e}")

    # STATS CALCULATIONS
    last_month_ret = clean_monthly_returns.iloc[-1] if len(clean_monthly_returns) > 0 else 0
    current_year = end_date.year
    ytd_data = clean_monthly_returns[clean_monthly_returns.index.year == current_year]
    ytd = (1 + ytd_data).prod() - 1 if len(ytd_data) > 0 else 0
    
    ror_3m = (1 + clean_monthly_returns.tail(3)).prod() - 1
    ror_12m = (1 + clean_monthly_returns.tail(12)).prod() - 1
    ror_36m = (1 + clean_monthly_returns.tail(36)).prod() - 1
    total_return = (final_equity / initial_capital) - 1
    compound_ror = cagr 
    
    win_months = clean_monthly_returns[clean_monthly_returns > 0]
    loss_months = clean_monthly_returns[clean_monthly_returns < 0]
    winning_months_pct = len(win_months) / len(clean_monthly_returns) if len(clean_monthly_returns) > 0 else 0
    avg_winning_month = win_months.mean() if len(win_months) > 0 else 0
    avg_losing_month = loss_months.mean() if len(loss_months) > 0 else 0
    
    annual_volatility = clean_monthly_returns.std() * np.sqrt(12)
    sharpe_ratio = (cagr - risk_free_rate) / annual_volatility if annual_volatility != 0 else 0
    downside_dev = clean_monthly_returns[clean_monthly_returns < 0].std() * np.sqrt(12)
    annual_loss_dev = downside_dev 
    sortino_ratio = (cagr - risk_free_rate) / downside_dev if downside_dev != 0 else 0
    
    ann_max_dd = df_monthly_equity['Drawdown'].groupby(df_monthly_equity.index.year).min().abs()
    avg_annual_max_dd = ann_max_dd.mean()
    sterling_ratio = cagr / (avg_annual_max_dd + 0.1) if not np.isnan(avg_annual_max_dd) else 0
    
    if years >= 3:
        calmar_ratio = cagr / abs(max_drawdown_eom) if max_drawdown_eom != 0 else 0
    else:
        calmar_ratio = cagr / abs(max_drawdown_eom) if max_drawdown_eom != 0 else 0
    
    mar_ratio = cagr / abs(max_drawdown_eom) if max_drawdown_eom != 0 else 0
    skew_val = skew(clean_monthly_returns)
    kurt_val = kurtosis(clean_monthly_returns)

    # TABLES SETUP

    # Monthly Performance
    df_table_std = clean_monthly_returns.rename('Value').to_frame()
    df_table_std['Year'] = df_table_std.index.year
    df_table_std['Month'] = df_table_std.index.strftime('%b')
    table_std = df_table_std.pivot(index='Year', columns='Month', values='Value')
    m_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    table_std = table_std[[m for m in m_order if m in table_std.columns]].fillna(0)
    table_std['Year Total'] = (clean_monthly_returns + 1).groupby(clean_monthly_returns.index.year).prod() - 1

    # Intramonth Drawdown
    im_dd_dict = {}
    months_range = pd.period_range(start=start_date, end=end_date, freq='M')
    for period in months_range:
        m_start, m_end = period.start_time, period.end_time
        subset = equity_series[m_start:m_end]
        prev = equity_series[equity_series.index < m_start]
        if not prev.empty:
            start_val = pd.Series([prev.iloc[-1]], index=[m_start - pd.Timedelta(seconds=1)])
            subset = pd.concat([start_val, subset])
        
        if not subset.empty:
            p_peak = subset.cummax()
            p_dd = (subset / p_peak) - 1
            im_dd_dict[period] = p_dd.min()
        else:
            im_dd_dict[period] = 0.0
            
    im_series = pd.Series(im_dd_dict)
    im_series.index = im_series.index.to_timestamp()
    
    df_table_im = im_series.rename('Value').to_frame()
    df_table_im['Year'] = df_table_im.index.year
    df_table_im['Month'] = df_table_im.index.strftime('%b')
    table_im = df_table_im.pivot(index='Year', columns='Month', values='Value')
    table_im = table_im[[m for m in m_order if m in table_im.columns]].fillna(0)
    if 1999 in table_im.index: table_im = table_im.drop(1999)

    # Drawdown Report
    drawdowns = []
    curr_dd = {}
    in_dd = False
    
    for dt, row in df_monthly_equity.iterrows():
        val = row['Drawdown']
        if val < 0:
            if not in_dd:
                in_dd = True
                curr_dd = {'start': dt, 'trough': dt, 'end': None, 'depth': val}
                loc = df_monthly_equity.index.get_loc(dt)
                if loc > 0: curr_dd['start'] = df_monthly_equity.index[loc-1]
            else:
                if val < curr_dd['depth']:
                    curr_dd['depth'] = val
                    curr_dd['trough'] = dt
        elif val == 0 and in_dd:
            in_dd = False
            curr_dd['end'] = dt
            p_ts = pd.Timestamp(curr_dd['start'])
            t_ts = pd.Timestamp(curr_dd['trough'])
            e_ts = pd.Timestamp(curr_dd['end'])
            l_mo = (t_ts.year - p_ts.year)*12 + (t_ts.month - p_ts.month)
            r_mo = (e_ts.year - t_ts.year)*12 + (e_ts.month - t_ts.month)
            curr_dd['len'] = max(1, l_mo)
            curr_dd['rec'] = max(0, r_mo)
            drawdowns.append(curr_dd)
            
    if in_dd:
        curr_dd['end'] = df_monthly_equity.index[-1]
        p_ts = pd.Timestamp(curr_dd['start'])
        t_ts = pd.Timestamp(curr_dd['trough'])
        l_mo = (t_ts.year - p_ts.year)*12 + (t_ts.month - p_ts.month)
        curr_dd['len'] = max(1, l_mo)
        curr_dd['rec'] = "Open"
        drawdowns.append(curr_dd)
        
    df_dd_rep = pd.DataFrame(drawdowns)
    if not df_dd_rep.empty:
        df_dd_rep = df_dd_rep.sort_values('depth', ascending=True).head(5)
        df_dd_rep['No.'] = range(1, len(df_dd_rep)+1)
        df_dd_rep = df_dd_rep[['No.', 'depth', 'len', 'rec', 'start', 'end']]
    else:
        df_dd_rep = pd.DataFrame(columns=['No.', 'depth', 'len', 'rec', 'start', 'end'])

    # Return Report
    ret_rep_data = []
    periods = {'1 Month':1, '3 Months':3, '6 Months':6, '1 Year':12, '2 Years':24, '3 Years':36, '5 Years':60}
    cum_ret_s = (1 + clean_monthly_returns).cumprod()
    
    for lbl, mo in periods.items():
        if len(clean_monthly_returns) >= mo:
            roll = (cum_ret_s / cum_ret_s.shift(mo)) - 1
            roll = roll.dropna()
            if not roll.empty:
                ret_rep_data.append([lbl, roll.max(), roll.min(), roll.mean(), roll.median(), roll.iloc[-1], (roll>0).mean()])
            else:
                ret_rep_data.append([lbl, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan])
        else:
            ret_rep_data.append([lbl, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan])
            
    df_ret_rep = pd.DataFrame(ret_rep_data, columns=['Period','Best','Worst','Avg','Med','Last','Win%'])

    # GENERATE EXCEL
    print("Generating Excel...")
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        wb = writer.book
        sh = wb.add_worksheet('Factsheet')
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white'})
        fmt_pct = wb.add_format({'num_format': '0.00%'})
        fmt_dec = wb.add_format({'num_format': '0.00'})
        
        sh.write(0, 0, header_title)
        sh.write(1, 0, header_sub)
        
        # Monthly Performance
        sh.write(4, 0, "Monthly Performance", fmt_head)
        sh.write_row(5, 1, table_std.columns, fmt_head)
        for r, (idx, row) in enumerate(table_std.iterrows()):
            sh.write(6+r, 0, idx, fmt_head)
            sh.write_row(6+r, 1, row, fmt_pct)
            
        cur = 6 + len(table_std) + 2
        
        # Stats
        sh.write(cur, 0, "Return Statistics", fmt_head)
        sh.write(cur, 1, "Value", fmt_head)
        
        ret_data_list = [
            ("Last Month", last_month_ret, fmt_pct),
            ("Year To Date", ytd, fmt_pct),
            ("3 Month ROR", ror_3m, fmt_pct),
            ("12 Months ROR", ror_12m, fmt_pct),
            ("36 Month ROR", ror_36m, fmt_pct),
            ("Total Return Cumulative", total_return, fmt_pct),
            ("Compound ROR", compound_ror, fmt_pct),
            ("Winning Months (%)", winning_months_pct, fmt_pct),
            ("Average Winning Month", avg_winning_month, fmt_pct),
            ("Average Losing Month", avg_losing_month, fmt_pct)
        ]
        for i, v in enumerate(ret_data_list):
            sh.write(cur+1+i, 0, v[0])
            sh.write(cur+1+i, 1, v[1], v[2])
            
        cur += len(ret_data_list) + 3
        sh.write(cur, 0, "Risk Statistics", fmt_head)
        sh.write(cur, 1, "Value", fmt_head)
        risk_data_list = [
            ("Sharpe Ratio", sharpe_ratio, fmt_dec),
            ("Sortino Ratio", sortino_ratio, fmt_dec),
            ("Sterling Ratio", sterling_ratio, fmt_dec),
            ("Calmar Ratio", calmar_ratio, fmt_dec),
            ("MAR Ratio", mar_ratio, fmt_dec),
            ("Skewness", skew_val, fmt_dec),
            ("Kurtosis", kurt_val, fmt_dec),
            ("Max Drawdown (EOM)", max_drawdown_eom, fmt_pct),
            ("Max Drawdown (Peak-to-Valley)", max_drawdown_im, fmt_pct),
            ("Standard Deviation (Ann.)", annual_volatility, fmt_pct),
            ("Loss Deviation (Ann.)", annual_loss_dev, fmt_pct)
        ]
        for i, v in enumerate(risk_data_list):
            sh.write(cur+1+i, 0, v[0])
            sh.write(cur+1+i, 1, v[1], v[2])
            
        # Intramonth Drawdown Table
        cur += len(risk_data_list) + 3
        sh.write(cur, 0, "Intramonth Max Drawdown", fmt_head)
        sh.write_row(cur+1, 1, table_im.columns, fmt_head)
        for r, (idx, row) in enumerate(table_im.iterrows()):
            sh.write(cur+2+r, 0, idx, fmt_head)
            sh.write_row(cur+2+r, 1, row, fmt_pct)
            
        # Drawdown Report
        cur += len(table_im) + 3
        sh.write(cur, 0, "Drawdown Report", fmt_head)
        sh.write_row(cur+1, 0, df_dd_rep.columns, fmt_head)
        for r, row in enumerate(df_dd_rep.values):
            sh.write(cur+2+r, 0, row[0])
            sh.write(cur+2+r, 1, row[1], fmt_pct)
            sh.write(cur+2+r, 2, row[2])
            sh.write(cur+2+r, 3, row[3])
            s_d = row[4].strftime('%m/%Y') if pd.notnull(row[4]) else "-"
            e_d = row[5].strftime('%m/%Y') if pd.notnull(row[5]) else "-"
            sh.write(cur+2+r, 4, s_d)
            sh.write(cur+2+r, 5, e_d)
            
    print(f"Excel Saved: {output_excel}")

    # GENERATE PDF
    print("Generating PDF...")
    with PdfPages(output_pdf) as pp:
        
        # P1
        fig1 = plt.figure(figsize=(8.27, 11.69))
        plt.subplots_adjust(left=0.08, right=0.92, top=0.95, bottom=0.05)
        gs1 = gridspec.GridSpec(3, 1, height_ratios=[1, 1, 10], hspace=0.2)
        
        ax_h = fig1.add_subplot(gs1[0])
        ax_h.axis('off')
        ax_h.text(0, 0.85, header_title, fontsize=20, weight='bold', color='#222222')
        #ax_h.text(0, 0.55, header_sub, fontsize=10, weight='bold', color='#444444')
        ax_h.text(0, 0.25, DISCLAIMER, fontsize=8, style='italic', color='#666666')
        
        ax_p = fig1.add_subplot(gs1[1])
        ax_p.axis('off')
        ax_p.set_title("Target Allocation (Final State)", fontsize=8, weight='bold', pad=10)
        
        # PORTFOLIO COMPOSITION TABLE
        p_data = []
        for idx, row in portfolio_constituents.iterrows():
            sub = row['Subtitle']
            alloc = float(row['Allocated_Capital'])
            weight = alloc / initial_capital
            s_date_label = "Inception"
            if pd.notnull(row['Start_Date']):
                s_date_label = str(row['Start_Date'])
            p_data.append([sub, f"${alloc:,.0f}", f"{weight:.1%}", s_date_label])
        
        p_data.append(["TOTAL", f"${initial_capital:,.0f}", "100.0%", "-"])
        col_labs_p = ["Strategy Component", "Target Capital", "Target Weight", "Added On"]
        
        tbl_p = ax_p.table(cellText=p_data, colLabels=col_labs_p, loc='center', cellLoc='center', bbox=[0, 0.1, 1, 0.8])
        tbl_p.auto_set_font_size(False); tbl_p.set_fontsize(7)
        for (r,c), cell in tbl_p.get_celld().items():
            if r==0: 
                cell.set_facecolor('#4472C4'); 
                cell.set_text_props(color='white', weight='bold')
            else:
                if r == len(p_data): 
                    cell.set_text_props(weight='bold')
                    cell.set_facecolor('#F2F2F2')
            
        ax_m = fig1.add_subplot(gs1[2])
        ax_m.axis('off')
        ax_m.set_title("Monthly Composite Performance", fontsize=8, weight='bold', pad=15)
        df_r = table_std.reset_index()
        txt = []
        for v in df_r.values:
            txt.append([str(int(v[0]))] + [f"{x:.2%}" for x in v[1:]])
        tbl_m = ax_m.table(cellText=txt, colLabels=df_r.columns, loc='center', bbox=[0,0,1,1])
        tbl_m.auto_set_font_size(False); tbl_m.set_fontsize(6)
        for (r,c), cell in tbl_m.get_celld().items():
            if r==0: cell.set_facecolor('#4472C4'); cell.set_text_props(color='white', weight='bold')
            else:
                if r%2==0: cell.set_facecolor('#F2F2F2')
                if c == len(df_r.columns)-1: cell.set_text_props(weight='bold')
        
        pp.savefig(fig1); plt.close(fig1)
        
        # P2 
        fig2 = plt.figure(figsize=(8.27, 11.69))
        plt.subplots_adjust(left=0.1, right=0.9, top=0.95, bottom=0.05)
        gs2 = gridspec.GridSpec(3, 1, height_ratios=[1.2, 2, 2], hspace=0.3)
        
        gs_stats = gridspec.GridSpecFromSubplotSpec(1, 2, subplot_spec=gs2[0], wspace=0.15)
        
        ax_ret = fig2.add_subplot(gs_stats[0])
        ax_ret.axis('off')
        ret_data = [
            ["Last Month", f"{last_month_ret:.2%}"],
            ["Year To Date", f"{ytd:.2%}"],
            ["3 Month ROR", f"{ror_3m:.2%}"],
            ["12 Months ROR", f"{ror_12m:.2%}"],
            ["36 Month ROR", f"{ror_36m:.2%}"],
            ["Total Return Cumulative", f"{total_return:.2%}"],
            ["Compound ROR", f"{compound_ror:.2%}"],
            ["Winning Months (%)", f"{winning_months_pct:.2%}"],
            ["Average Winning Month", f"{avg_winning_month:.2%}"],
            ["Average Losing Month", f"{avg_losing_month:.2%}"]
        ]
        tbl_ret = ax_ret.table(cellText=ret_data, colLabels=["Return Statistics", ""], loc='center', cellLoc='left', bbox=[0, 0, 1, 1])
        tbl_ret.auto_set_font_size(False)
        tbl_ret.set_fontsize(7)
        for (r, c), cell in tbl_ret.get_celld().items():
            if r == 0:
                cell.set_facecolor('#4472C4')
                cell.set_text_props(color='white', weight='bold', ha='center')
            else:
                if c == 1: cell.set_text_props(ha='right')
                if r % 2 == 0: cell.set_facecolor('#F2F2F2')

        ax_risk = fig2.add_subplot(gs_stats[1])
        ax_risk.axis('off')
        risk_data = [
            ["Sharpe Ratio", f"{sharpe_ratio:.2f}"],
            ["Sortino Ratio", f"{sortino_ratio:.2f}"],
            ["Sterling Ratio", f"{sterling_ratio:.2f}"],
            ["Calmar Ratio", f"{calmar_ratio:.2f}"],
            ["MAR Ratio", f"{mar_ratio:.2f}"],
            ["Skewness", f"{skew_val:.2f}"],
            ["Kurtosis", f"{kurt_val:.2f}"],
            ["Max Drawdown (EOM)", f"{max_drawdown_eom:.2%}"],
            ["Max Drawdown (Peak-to-Valley)", f"{max_drawdown_im:.2%}"],
            ["Standard Deviation (Ann.)", f"{annual_volatility:.2%}"],
            ["Loss Deviation (Ann.)", f"{annual_loss_dev:.2%}"]
        ]
        tbl_risk = ax_risk.table(cellText=risk_data, colLabels=["Risk Statistics", ""], loc='center', cellLoc='left', bbox=[0, 0, 1, 1])
        tbl_risk.auto_set_font_size(False)
        tbl_risk.set_fontsize(7)
        for (r, c), cell in tbl_risk.get_celld().items():
            if r == 0:
                cell.set_facecolor('#4472C4')
                cell.set_text_props(color='white', weight='bold', ha='center')
            else:
                if c == 1: cell.set_text_props(ha='right')
                if r % 2 == 0: cell.set_facecolor('#F2F2F2')
                
        # VAMI CHART SECTION
        ax_v = fig2.add_subplot(gs2[1])
        ax_v.plot(df_monthly_equity.index, df_monthly_equity['VAMI'], color='#4472C4', lw=1.5, label=header_title)
        if not sp_vami.empty: ax_v.plot(sp_vami.index, sp_vami, color='#7f8c8d', lw=1, ls='-', label='S&P 500')
        ax_v.set_title("Growth of $1,000 (VAMI)", fontsize=8, weight='bold')
        ax_v.grid(True, ls=':', alpha=0.5)
        ax_v.legend(loc='upper left', fontsize=8)

        # Rotate x-axis labels
        plt.setp(ax_v.get_xticklabels(), fontsize=8)

        # Resize y-axis labels
        plt.setp(ax_v.get_yticklabels(), fontsize=8)

        # VAMI Vertical lines
        colors = ['green', 'orange', 'purple', 'brown']
        ci = 0
        for idx, row in portfolio_constituents.iterrows():
            if pd.notnull(row['Start_Date']):
                dt_mark = pd.to_datetime(row['Start_Date'])
                if dt_mark > df_monthly_equity.index.min() and dt_mark < df_monthly_equity.index.max():
                      color_choice = colors[ci % len(colors)]
                      ax_v.axvline(x=dt_mark, color=color_choice, linestyle='--', alpha=0.7)
                      # Add Label
                      trans = ax_v.get_xaxis_transform()
                      ax_v.text(dt_mark, 0.95, f"  Added: {row['Subtitle']}", 
                                color=color_choice, rotation=90, verticalalignment='top', 
                                transform=trans, fontsize=7, weight='bold')
                      ci += 1

        ax_v.yaxis.set_major_formatter(mtick.StrMethodFormatter('{x:,.0f}'))
        
        # DRAWDOWN CHART
        ax_d = fig2.add_subplot(gs2[2])
        ax_d.fill_between(df_monthly_equity.index, df_monthly_equity['Drawdown'], 0, color="#0022FF", alpha=0.3)
        ax_d.plot(df_monthly_equity.index, df_monthly_equity['Drawdown'], color="#657AFF", lw=1, label=header_title)
        if not sp_drawdown.empty: ax_d.plot(sp_drawdown.index, sp_drawdown, color='#7f8c8d', lw=1, ls='-', label='S&P 500')
        ax_d.yaxis.set_major_formatter(mtick.PercentFormatter(1.0, decimals=0))
        ax_d.set_title("Drawdown Profile (vs S&P 500)", fontsize=8, weight='bold'); ax_d.grid(True, ls=':', alpha=0.5)

        # Rotate x-axis labels
        plt.setp(ax_d.get_xticklabels(), fontsize=8)

        # Resize y-axis labels
        plt.setp(ax_d.get_yticklabels(), fontsize=8)
        
        pp.savefig(fig2); plt.close(fig2)
        
        # P3
        fig3 = plt.figure(figsize=(8.27, 11.69))
        plt.subplots_adjust(left=0.08, right=0.92, top=0.95, bottom=0.05)
        gs3 = gridspec.GridSpec(2, 1, height_ratios=[1, 15], hspace=0.05)
        ax_t3 = fig3.add_subplot(gs3[0]); ax_t3.axis('off')
        ax_t3.text(0.5, 0.5, "Intramonth Max Drawdown", fontsize=12, weight='bold', ha='center', color='#222222')
        ax_tab3 = fig3.add_subplot(gs3[1]); ax_tab3.axis('off')
        df_r3 = table_im.reset_index()
        txt3 = []
        for v in df_r3.values:
            txt3.append([str(int(v[0]))] + [f"{x:.2%}" for x in v[1:]])
        t3 = ax_tab3.table(cellText=txt3, colLabels=df_r3.columns, loc='upper center', bbox=[0,0,1,1])
        t3.auto_set_font_size(False); t3.set_fontsize(6)
        for (r,c), cell in t3.get_celld().items():
            if r==0: cell.set_facecolor('#4472C4'); cell.set_text_props(color='white', weight='bold')
            else:
                if r%2==0: cell.set_facecolor('#F2F2F2')
                else: cell.set_facecolor('white')
                if c>0:
                    val = df_r3.iloc[r-1, c]
                    if pd.notnull(val) and val <= -0.10: cell.set_facecolor('#FFC7CE')
        pp.savefig(fig3); plt.close(fig3)
        
        # TABLES + GROWTH CHART MERGED
        fig4 = plt.figure(figsize=(8.27, 11.69))
        
        # Define GridSpec
        gs4 = gridspec.GridSpec(2, 2, height_ratios=[1, 1], width_ratios=[1, 1.3], wspace=0.15, hspace=0.3)
        
        # Apply margins
        plt.subplots_adjust(left=0.1, right=0.9, top=0.95, bottom=0.08)
        
        # DRAWDOWN REPORT
        ax_dr = fig4.add_subplot(gs4[0,0])
        ax_dr.axis('off')
        ax_dr.text(0.5, 1.05, "Drawdown Report", fontsize=10, weight='bold', ha='center', color='#4472C4', transform=ax_dr.transAxes)
        if not df_dd_rep.empty:
            d_dat = []
            cols = ['NO.','DEPTH','LENGTH','RECOVERY','START','END']
            for v in df_dd_rep.values:
                s_d = v[4].strftime('%m/%Y') if pd.notnull(v[4]) else "-"
                e_d = v[5].strftime('%m/%Y') if pd.notnull(v[5]) else "-"
                d_dat.append([str(int(v[0])), f"{v[1]:.2%}", str(int(v[2])), str(v[3]), s_d, e_d])
            t_dr = ax_dr.table(cellText=d_dat, colLabels=cols, loc='upper center', bbox=[0,0.5,1,0.5])
            t_dr.auto_set_font_size(False); t_dr.set_fontsize(6)
            t_dr.auto_set_column_width(col=list(range(len(cols))))
            for (r,c), cell in t_dr.get_celld().items():
                if r==0: cell.set_facecolor('#4472C4'); cell.set_text_props(color='white', weight='bold')
                else:
                    if r%2==0: cell.set_facecolor('#F2F2F2')
        
        # RETURN REPORT
        ax_rr = fig4.add_subplot(gs4[0,1])
        ax_rr.axis('off')
        ax_rr.text(0.5, 1.05, "Return Report", fontsize=10, weight='bold', ha='center', color='#4472C4', transform=ax_rr.transAxes)
        if not df_ret_rep.empty:
            r_dat = []
            cols = ['PERIOD','BEST','WORST','AVG','MED','LAST','WIN %']
            for v in df_ret_rep.values:
                row_str = [str(v[0])]
                for x in v[1:]: row_str.append(f"{x:.2%}" if pd.notnull(x) else "-")
                r_dat.append(row_str)
            t_rr = ax_rr.table(cellText=r_dat, colLabels=cols, loc='upper center', bbox=[0,0.3,1,0.7])
            t_rr.auto_set_font_size(False); t_rr.set_fontsize(6)
            t_rr.auto_set_column_width(col=list(range(len(cols))))
            for (r,c), cell in t_rr.get_celld().items():
                if r==0: cell.set_facecolor('#4472C4'); cell.set_text_props(color='white', weight='bold')
                else:
                    if r%2==0: cell.set_facecolor('#F2F2F2')
                    
        # REVENUE GROWTH CHART
        ax_g = fig4.add_subplot(gs4[1, :]) 
        
        ax_g.plot(equity_series.index, equity_series, color='#2E86C1', lw=2, label="Portfolio Value")
        ax_g.set_yscale('log')

        plt.subplots_adjust(left=0.15, right=0.88, top=0.92, bottom=0.08)
        
        ci = 0
        for idx, row in portfolio_constituents.iterrows():
            if pd.notnull(row['Start_Date']):
                dt_mark = pd.to_datetime(row['Start_Date'])
                if dt_mark > equity_series.index.min() and dt_mark < equity_series.index.max():
                      color_choice = colors[ci % len(colors)]
                      ax_g.axvline(x=dt_mark, color=color_choice, linestyle='--', alpha=0.8, linewidth=1.5, label=f"Added: {row['Subtitle']}")
                      ci += 1
                      
        ax_g.set_title("Portfolio Revenue Growth with Asset Additions", fontsize=10, weight='bold', pad=10)
        ax_g.set_ylabel("Portfolio Value (Log Scale)", fontsize=8)
        ax_g.set_xlabel("Year", fontsize=9)
        ax_g.grid(True, which="both", ls="-", alpha=0.2)
        ax_g.legend(loc='upper left', fontsize=8)
        
        fmt = mtick.StrMethodFormatter('${x:,.0f}')
        ax_g.yaxis.set_major_formatter(fmt)
        
        # Rotate x-axis labels
        plt.setp(ax_g.get_xticklabels(), rotation=45, ha='right', fontsize=8)

        # Resize y-axis labels
        plt.setp(ax_g.get_yticklabels(), fontsize=8)
        
        pp.savefig(fig4); plt.close(fig4)
        
    print(f"PDF Saved: {output_pdf}")

# MAIN EXECUTION
if __name__ == "__main__":
    if not os.path.exists(INPUT_DIR): os.makedirs(INPUT_DIR)
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    
    if not os.path.exists(CONFIG_FILE):
        print(f"Creating sample {CONFIG_FILE}...")
        samp = pd.DataFrame([
            ['Test_Portfolio', 'tradeList.csv', 100000, '2020-01-01', 0.00, 'ArchDAP', 'Test Portfolio'],
        ], columns=['Report_ID', 'Filename', 'Allocated_Capital', 'Start_Date', 'Risk_Free_Rate', 'Title', 'Subtitle'])
        samp.to_csv(CONFIG_FILE, index=False)
        print("Done. Add your files to 'input' folder and run again.")
    else:
        try:
            cfg = pd.read_csv(CONFIG_FILE)
            groups = cfg.groupby('Report_ID')
            print(f"Found {len(groups)} portfolios to process.")
            for grp_name, grp_data in groups:
                process_portfolio_group(grp_name, grp_data)
            print("\nAll done.")
        except Exception as e:
            print(f"Error: {e}")