import os
import datetime
import pandas as pd
import xlwings as xw

def update_masters(mapping_file, output_root, root_folder,
                   upd_last_monthly, upd_last_weekly, upd_last_daily,
                   upd_prev_monthly, upd_prev_weekly, upd_prev_daily,
                   shift_monthly, shift_weekly, shift_daily,
                   log_func, update_prog):
    """
    Updates master Excel sheets by shifting old values and writing new ones.
    """

    # Build mapping DataFrame
    mapping_sheets = pd.read_excel(mapping_file, sheet_name=None)
    maps = []
    for sheet_name, df in mapping_sheets.items():
        prog = sheet_name.split()[0]
        folder = " ".join(sheet_name.split()[1:])
        df = df.rename(columns=str.strip)
        if 'Manual_label' in df.columns and 'Manual_Label' not in df.columns:
            df = df.rename(columns={'Manual_label': 'Manual_Label'})
        maps.append(df.assign(Program=prog, Folder=folder))
    mapping_df = pd.concat(maps, ignore_index=True)

    mapping_lookup = {}
    for key, g in mapping_df.groupby(["Program", "Folder"]):
        mapping_lookup[key] = g

    with xw.App(visible=False) as app:
        app.api.ScreenUpdating = False
        app.api.DisplayAlerts  = False
        app.api.EnableEvents   = False
        app.api.Calculation    = xw.constants.Calculation.xlCalculationManual

        masters = [
            os.path.join(root_folder, f)
            for f in os.listdir(root_folder)
            if f.lower().startswith("study_nl_") and f.lower().endswith(".xlsx")
        ]

        for mf in masters:
            log_func(f"Updating (xlwings) {os.path.basename(mf)}...\n")
            wb = app.books.open(mf)

            for sht in wb.sheets:
                # ... determine folder/program as before ...
                prog = sht.name.split()[0]
                subset = mapping_lookup.get((prog, folder), pd.DataFrame())

                col_b = sht.range("B3").expand("down").value or []
                max_row = sht.range("B3").expand("down").last_cell.row

                # Prepare batch update lists for each output column
                col_f = [None] * len(col_b)  # Last Monthly (F)
                col_e = [None] * len(col_b)  # Prev Monthly (E)
                col_m = [None] * len(col_b)  # Last Weekly (M)
                col_l = [None] * len(col_b)  # Prev Weekly (L)
                col_q = [None] * len(col_b)  # Last Daily (Q)
                col_p = [None] * len(col_b)  # Prev Daily (P)

                sticker_to_idx = {str(cell).strip(): idx for idx, cell in enumerate(col_b)}

                for _, row in subset.iterrows():
                    sticker = str(row.iloc[0]).strip()
                    stem    = str(row.iloc[1]).strip().removesuffix(".pdf")
                    data_file = os.path.join(output_root, prog, folder, f"{stem}.xlsx")
                    if not os.path.isfile(data_file):
                        log_func(f"[WARN] Missing data: {data_file}\n")
                        continue

                    data = pd.read_excel(data_file, sheet_name=0).iloc[0]
                    ml  = data.get('monthly_last')
                    mp  = data.get('monthly_prev')
                    wl  = data.get('weekly_last')
                    wp  = data.get('weekly_prev')
                    dl  = data.get('daily_last', data.get('daily_prev'))
                    dp  = data.get('daily_prev')
                    dp2 = data.get('daily_prev2', dp)

                    idx = sticker_to_idx.get(sticker)
                    if idx is None:
                        continue

                    # Determine which daily to use
                    pdf_p = os.path.join(root_folder, prog, folder, stem + ".pdf")
                    try:
                        hr = datetime.datetime.fromtimestamp(os.path.getmtime(pdf_p)).hour
                    except:
                        hr = 0
                    if hr >= 8:
                        last_val = dp
                        prev_val = dp2
                    else:
                        last_val = dl
                        prev_val = dp

                    if upd_last_monthly and ml is not None:
                        col_f[idx] = ml
                    if upd_prev_monthly and mp is not None:
                        col_e[idx] = mp
                    if upd_last_weekly and wl is not None:
                        col_m[idx] = wl
                    if upd_prev_weekly and wp is not None:
                        col_l[idx] = wp
                    if upd_last_daily and last_val is not None:
                        col_q[idx] = last_val
                    if upd_prev_daily and prev_val is not None:
                        col_p[idx] = prev_val

                # Now batch write to Excel in one go for each column (None stays unmodified)
                if upd_last_monthly:
                    sht.range(f"F3:F{max_row}").value = [[v] for v in col_f]
                if upd_prev_monthly:
                    sht.range(f"E3:E{max_row}").value = [[v] for v in col_e]
                if upd_last_weekly:
                    sht.range(f"M3:M{max_row}").value = [[v] for v in col_m]
                if upd_prev_weekly:
                    sht.range(f"L3:L{max_row}").value = [[v] for v in col_l]
                if upd_last_daily:
                    sht.range(f"Q3:Q{max_row}").value = [[v] for v in col_q]
                if upd_prev_daily:
                    sht.range(f"P3:P{max_row}").value = [[v] for v in col_p]


            wb.save()
            wb.close()
            update_prog()

        # Restore Excel defaults
        app.api.Calculation    = xw.constants.Calculation.xlCalculationAutomatic
        app.api.DisplayAlerts  = True
        app.api.EnableEvents   = True
        app.api.ScreenUpdating = True
