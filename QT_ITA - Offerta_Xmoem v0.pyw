
import pandas as pd
import xlsxwriter
import math
import glob
import os
import tkinter as tk
from tkinter import messagebox

def main():
    # Mostra popup iniziale
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Elaborazione in corso", "Inizio elaborazione dei file Excel...")

    # Ottieni la directory corrente dello script
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # Cerca Database Excel nella directory corrente
    listino_df = pd.read_excel(os.path.join(current_dir, "00_Database_XMOEM.xlsx"), sheet_name=0, engine='openpyxl', header=None)
    sconti_df = pd.read_excel(os.path.join(current_dir, "00_Database_XMOEM.xlsx"), sheet_name=1, engine='openpyxl', header=None)

    # Salta la prima riga del listino (intestazione)
    listino_dict = {str(row.iloc[0]).strip(): row for _, row in listino_df.iloc[1:].iterrows()}
    sconti_dict = {fam: idx + 2 for idx, fam in enumerate(sconti_df.iloc[:, 0].astype(str).str.strip().str.upper())}

    # Inizializza scrittura file Output
    start_row = 1
    output_path = os.path.join(current_dir, "Output_XMOEM.xlsx")
    workbook = xlsxwriter.Workbook(output_path, {'nan_inf_to_errors': True})
    worksheet = workbook.add_worksheet()

    # Formati celle excel
    highlight = workbook.add_format({'bg_color': '#FFFF00', 'align': 'center'})
    text = workbook.add_format({'num_format': '@', 'align': 'center'})
    left_align = workbook.add_format({'align': 'left'})
    highlight_left = workbook.add_format({'bg_color': '#FFFF00', 'align': 'left'})
    percent_format = workbook.add_format({'num_format': '0%', 'align': 'center'})
    general_format = workbook.add_format({'align': 'center'})
    blu_ctr_format = workbook.add_format({'bold': True, 'align': 'center', 'font_color': '#0066CC'})
    blu_rth_format = workbook.add_format({'bold': True, 'align': 'right', 'font_color': '#0066CC'})

    # Cerca file con i vari Codici del Quadro nella directory corrente
    excel_files = [f for f in glob.glob(os.path.join(current_dir, "*.xls*")) if os.path.basename(f) != "Database_XMOEM.xlsx"]

    # Cicla su tutti i file dei quadri
    for file_path in excel_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        df_full = pd.read_excel(file_path, sheet_name=0, engine='openpyxl', header=None)

        target_col = None
        for col_idx in range(min(30, df_full.shape[1])):
            for row_idx in range(min(7, df_full.shape[0])):
                cell_value = str(df_full.iat[row_idx, col_idx])
                if cell_value in ["CODICI", "Codice"]:
                    target_col = col_idx
                    break
            if target_col is not None:
                break

        if target_col is not None and target_col + 1 < df_full.shape[1]:
            input_df = df_full.iloc[:, [target_col, target_col + 1]]
        else:
            continue

        output_data = []
        highlight_rows = set()

        for idx, code in enumerate(input_df.iloc[:, 0]):
            if pd.isna(code) or str(code).strip().upper() in ["CODICE", "CODICI", "QTY", "QUANTITÀ", ""]:
                continue

            code_str = str(code).strip()
            row = listino_dict.get(code_str)
            qty = input_df.iloc[idx, 1] if input_df.shape[1] > 1 else ""
            excel_row = (start_row + 1) + (len(output_data) + 1)

            if row is not None:
                evidenzia = isinstance(row.iloc[7], str) and row.iloc[7].strip() != ""
                if evidenzia:
                    highlight_rows.add(len(output_data))

                row_data = [
                    "",  # A
                    code_str,  # B
                    row.iloc[1],  # C
                    f"=IF(O{excel_row}=0,IF(N{excel_row}=0,ROUND(J{excel_row}*(1-K{excel_row})*(1-L{excel_row})*(1-M{excel_row}),2),ROUND(J{excel_row}*(1-N{excel_row}),2)),O{excel_row})",  # D
                    qty,  # E
                    f"=D{excel_row}*E{excel_row}",  # F
                    row.iloc[2],  # G
                    row.iloc[3],  # H
                    row.iloc[6],  # I
                    row.iloc[4],  # J
                    f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 3, FALSE)",  # K
                    f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 4, FALSE)",  # L
                    f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 5, FALSE)",  # M
                    "",  # N
                    "",  # O
                    row.iloc[8],  # P
                    f"=P{excel_row}*E{excel_row}"  # Q
                ]
            else:
                row_data = ["", code_str, "", f"=J{excel_row}", qty, f"=D{excel_row}*E{excel_row}"] + [""] * 10 + [f"=P{excel_row}*E{excel_row}"]

            while len(row_data) < 17:
                row_data.append("")

            output_data.append(row_data)

        worksheet.set_column(2, 2, 50, left_align)

        for row_idx, row_data in enumerate(output_data):
            for col_idx, cell in enumerate(row_data):
                if isinstance(cell, float) and (math.isnan(cell) or math.isinf(cell)):
                    cell = ""

                if col_idx == 2:
                    fmt = highlight_left if row_idx in highlight_rows else left_align
                elif col_idx == 1 and row_idx in highlight_rows:
                    fmt = highlight
                elif col_idx in [10, 11, 12, 13]:
                    fmt = percent_format
                else:
                    fmt = general_format

                if isinstance(cell, str) and cell.startswith("="):
                    worksheet.write_formula(start_row + row_idx + 1, col_idx, cell, fmt)
                else:
                    worksheet.write(start_row + row_idx + 1, col_idx, cell, fmt)

        worksheet.write(start_row, 0, file_name, blu_ctr_format)
        worksheet.write(start_row + row_idx + 2, 2, "Totale", blu_rth_format)
        worksheet.write(start_row + row_idx + 2, 3, f"=A{start_row + 1}", blu_ctr_format)
        worksheet.write(start_row + row_idx + 2, 5, f"=SUM(F{start_row + 2}:F{start_row + row_idx + 2})", blu_ctr_format)
        start_row = start_row + row_idx + 4

    workbook.close()

    # Mostra popup finale
    messagebox.showinfo("Completato", "Elaborazione completata con successo!\nFile generato: Output_XMOEM.xlsx")

if __name__ == "__main__":
    main()
