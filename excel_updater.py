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
                name = sht.name
                if "Items_FCPA" in name:
                    folder = "FCPA"
                elif "Items_WO CAL" in name:
                    folder = "WO CAL"
                elif "Items_CAL" in name:
                    folder = "CAL"
                else:
                    continue

                prog = name.split()[0]
                subset = mapping_df[
                    (mapping_df.Program == prog) &
                    (mapping_df.Folder  == folder)
                ]

                col_b   = sht.range("B3").expand("down").value or []
                max_row = sht.range("B3").expand("down").last_cell.row

                # 1) Shift existing values
                if shift_monthly:
                    sht.range(f"C3:E{max_row}").value = sht.range(f"D3:F{max_row}").value
                    sht.range(f"F3:F{max_row}").clear_contents()
                if shift_weekly:
                    sht.range(f"H3:L{max_row}").value = sht.range(f"I3:M{max_row}").value
                    sht.range(f"M3:M{max_row}").clear_contents()
                if shift_daily:
                    sht.range(f"O3:P{max_row}").value = sht.range(f"P3:Q{max_row}").value
                    sht.range(f"Q3:Q{max_row}").clear_contents()

                # 2) Write new values
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

                    for row_idx, cell in enumerate(col_b, start=3):
                        if str(cell).strip() == sticker:
                            if upd_last_monthly and ml is not None:
                                sht.cells(row_idx, 6).value = ml
                            if upd_prev_monthly and mp is not None:
                                sht.cells(row_idx, 5).value = mp
                            if upd_last_weekly and wl is not None:
                                sht.cells(row_idx, 13).value = wl
                            if upd_prev_weekly and wp is not None:
                                sht.cells(row_idx, 12).value = wp

                            # Unified daily logic
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

                            if upd_last_daily:
                                sht.cells(row_idx, 17).value = last_val
                            if upd_prev_daily:
                                sht.cells(row_idx, 16).value = prev_val
                            break

            wb.save()
            wb.close()
            update_prog()

        # Restore Excel defaults
        app.api.Calculation    = xw.constants.Calculation.xlCalculationAutomatic
        app.api.DisplayAlerts  = True
        app.api.EnableEvents   = True
        app.api.ScreenUpdating = True
