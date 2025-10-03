import pandas as pd
import xlsxwriter
import math
import glob
import os
import tkinter as tk
from tkinter import messagebox

# Mostra popup iniziale
root = tk.Tk()
root.withdraw()
messagebox.showinfo("Elaborazione in corso", "Inizio elaborazione dei file Excel...")

# Ottieni la directory corrente dello script
current_dir = os.path.dirname(os.path.abspath(__file__))

# Carica i dati da Excel
listino_df = pd.read_excel(os.path.join(current_dir, "00_Database_XMOEM.xlsx"), sheet_name=0, engine='openpyxl', header=None)
sconti_df = pd.read_excel(os.path.join(current_dir, "00_Database_XMOEM.xlsx"), sheet_name=1, engine='openpyxl', header=None)
rubrica_df = pd.read_excel(os.path.join(current_dir, "00_Database_XMOEM.xlsx"), sheet_name=2, engine='openpyxl', header=None)
# Crea dizionari
listino_dict = {str(row.iloc[0]).strip(): row for _, row in listino_df.iloc[1:].iterrows()}
sconti_dict = {fam: idx + 2 for idx, fam in enumerate(sconti_df.iloc[:, 0].astype(str).str.strip().str.upper())}

# Percorso file output
output_path = os.path.join(current_dir, "Output_XMOEM.xlsx")
riga_quadri = 19
start_row = riga_quadri + 2
start_row_rias = riga_quadri + 1
totaleOfferta_str = "="

# Crea file Excel con xlsxwriter
workbook = xlsxwriter.Workbook(output_path, {'nan_inf_to_errors': True})
riassuntivoSheet = workbook.add_worksheet("Riassuntivo")
worksheet = workbook.add_worksheet("Preventivo")
famiglieSheet = workbook.add_worksheet("Famiglie Sconto")
rubricaSheet = workbook.add_worksheet("Rubrica")
rubricaSheet.set_tab_color('red')

# Formati celle excel
formato_data = workbook.add_format({'num_format': 'dd/mm/yyyy','align': 'left'})
highlight = workbook.add_format({'bg_color': '#FFFF00', 'align': 'center'})
highlight_left = workbook.add_format({'bg_color': '#FFFF00', 'align': 'left'})
percent_format = workbook.add_format({'num_format': '0%', 'align': 'center'})
left_align = workbook.add_format({'align': 'left'})
center_align = workbook.add_format({'align': 'center'})
right_align = workbook.add_format({'align': 'right'})
left_align_bold = workbook.add_format({'bold': True,'align': 'left'})
highlight_grn_lft = workbook.add_format({'align': 'left','bg_color': '#EEECE0'})
highlight_grn_lft_bold = workbook.add_format({'bold': True,'align': 'left','bg_color': '#EEECE0'})
blu_lft_format = workbook.add_format({'bold': True, 'align': 'left', 'font_color': '#0066CC'})
blu_ctr_format = workbook.add_format({'bold': True, 'align': 'center', 'font_color': '#0066CC'})
blu_rth_format = workbook.add_format({'bold': True, 'align': 'right', 'font_color': '#0066CC'})
euro = workbook.add_format({'num_format': '€#,##0.00', 'align': 'right'})
euro_blu = workbook.add_format({'bold': True,'num_format': '€#,##0.00', 'align': 'right','font_color': '#0066CC'})
euro_bold = workbook.add_format({'bold': True,'num_format': '€#,##0.00', 'align': 'right'})
highlight_blu_lft = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0066CC','align': 'left' })
highlight_blu = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#0066CC',
    'align': 'center', 'valign': 'vcenter', 'text_wrap': True  })
highlight_nero = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#404040',
    'align': 'center', 'valign': 'vcenter', 'text_wrap': True  })


# Lista di nominativi
nominativi_QE = ['Michele Angelastri','Giuseppe Aru', 'Gianni Bruni','Fabrizio Genchi','Ivan Mazzarella','Chiara Spanu']

def formattazione_iniziale(riga):
    riassuntivoSheet.set_column(0, 0, 2, left_align)
    riassuntivoSheet.set_column(1, 1, 14, left_align)
    riassuntivoSheet.set_column(2, 2, 40, left_align)
    riassuntivoSheet.set_column(3, 3, 27, left_align)
    riassuntivoSheet.set_column(4, 5, 15, left_align)
    worksheet.set_column(2, 2, 50, left_align)
    worksheet.set_column(5, 5,12, left_align)
    famiglieSheet.set_column(1, 1, 50, left_align)
    famiglieSheet.set_column(7,7, 15, left_align)

    # Inserisce l'immagine nella cella A1 (può sbordare)
    image_path = os.path.join(current_dir,"Eaton-Logo.jpg")
    riassuntivoSheet.insert_image('B1', image_path,{'x_scale': 0.27, 'y_scale': 0.35})
    worksheet.insert_image('B1', image_path,{'x_scale': 0.27, 'y_scale': 0.35})

    # Info Eaton
    riassuntivoSheet.write(0,3, "EATON INDUSTRIES (ITALY) S.r.l.", left_align)
    riassuntivoSheet.write(1,3, "Via San Bovio, 3", left_align)
    riassuntivoSheet.write(2,3, "20090 Segrate (MI)", left_align)
    riassuntivoSheet.write(3,3, "Tel: +39 02 959501", left_align)
    riassuntivoSheet.write(4,3, "www.eaton.it", left_align)
    # Info OFFERTA
    riassuntivoSheet.write(7,1, "Spett.le:", highlight_grn_lft)
    riassuntivoSheet.write(7,2, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(8,1, "Alla c.a.:", highlight_grn_lft)
    riassuntivoSheet.write(8,2, "", highlight_grn_lft)
    riassuntivoSheet.write(10,1, "Progetto:", highlight_grn_lft)
    riassuntivoSheet.write(10,2, "", highlight_grn_lft)
    riassuntivoSheet.write(11,1, "Offerta n.:", highlight_grn_lft)
    riassuntivoSheet.write(11,2, "", highlight_grn_lft)
    riassuntivoSheet.write(12,1, "", highlight_grn_lft)
    riassuntivoSheet.write(12,2, "", highlight_grn_lft)
    riassuntivoSheet.write(13,1, "Data:", highlight_grn_lft)
    riassuntivoSheet.write(13,2, "", highlight_grn_lft)
    riassuntivoSheet.write(15,1, "Con riferimento alla Vs. gradita richiesta, Vi sottoponiamo la nostra migliore quotazione", left_align)
    riassuntivoSheet.write(16,1, "relativa agli articoli di Vostro interesse:", left_align)

    riassuntivoSheet.write(7,3, "Offerta redatta da:", highlight_grn_lft)   
    riassuntivoSheet.write(8,3, "", highlight_grn_lft_bold)
    # Aggiungi la data validation al Quotation Engineer
    riassuntivoSheet.data_validation(8, 3, 8, 3, {
        'validate': 'list',
        'source': nominativi_QE,
        'input_message': 'Inserire nome Quotation Engineer',
        'error_message': 'Valore non valido. Seleziona dalla lista.'
    })

    riassuntivoSheet.write(10,3, "Venditore:",highlight_grn_lft)
    riassuntivoSheet.write(11,3, "", highlight_grn_lft_bold)
    # Scrivi i dati di rubrica_df nel foglio "Rubrica"
    for row_idx, row in rubrica_df.iterrows():
        for col_idx, value in enumerate(row):
            rubricaSheet.write(row_idx, col_idx, value)
    # Scrivi i valori unici della colonna A (indice 0) nel foglio Rubrica, colonna di appoggio (es. colonna U = indice 20)
    validation_values = rubrica_df.iloc[:, 0].dropna().astype(str).unique()
    for i, val in enumerate(validation_values):
        rubricaSheet.write(i, 20, val)  # Scrive in colonna U
    # Aggiungi la data validation in Riassuntivo!D12 usando l'intervallo Rubrica!U1:U{n}
    riassuntivoSheet.data_validation(11, 3, 11, 3, {
        'validate': 'list',
        'source': f'=Rubrica!U1:U{len(validation_values)}',
        'input_message': 'Inserire nome venditore',
        'error_message': 'Valore non valido. Seleziona dalla lista.'
    })
    riassuntivoSheet.write_formula(12,3, f'=VLOOKUP(Riassuntivo!D12,Rubrica!A1:G300,4,FALSE)', highlight_grn_lft)
    riassuntivoSheet.write_formula(13,3, f'=VLOOKUP(Riassuntivo!D12,Rubrica!A1:G300,5,FALSE)', highlight_grn_lft)

    riassuntivoSheet.write(18,1, "Rif.", highlight_blu)
    riassuntivoSheet.write(18,2, "Denominazione quadro", highlight_blu)
    riassuntivoSheet.write(18,3, "Prezzo", highlight_blu)

    # Inserisce formule in A1:D14 del foglio 'Preventivo'
    for row in range(14):  # righe da 0 a 13 (A1-A14)
        for col in range(4):  # colonne da 0 a 4 (A-D)
            cella_riassuntivo = xlsxwriter.utility.xl_rowcol_to_cell(row, col)
            formula = f'=IF(Riassuntivo!{cella_riassuntivo}<>"",Riassuntivo!{cella_riassuntivo}&"","")'
            formula_data = f'=IF(Riassuntivo!{cella_riassuntivo}<>"",Riassuntivo!{cella_riassuntivo},"")'
            if row == 13 and col == 2:
                worksheet.write_formula(row, col, formula_data, formato_data)
            else:
                worksheet.write_formula(row, col, formula)

    worksheet.write(15,1, "Con riferimento alla Vs. gradita richiesta, Vi sottoponiamo la nostra migliore quotazione", left_align)
    worksheet.write(16,1, "relativa agli articoli di Vostro interesse:", left_align)

    # Valori da inserire
    intestazione = [
        'Pos.', 'Codice', 'Tipo / Descrizione', 'Listino 2025', 'Q.tà', 'Tot. Listino',
        'Minimo Ordine', 'Lead Time', 'Famiglia Statistica', 'Listino', 'Sc1%', 'Sc2%',
        'Sc3%', 'ScX%', 'Prezzo Netto', 'UM', 'U.M x Qtà'   ]

    # Scrittura dei valori con formattazione
    for col, valore in enumerate(intestazione):
        formato = highlight_blu if col <= 8 else highlight_nero
        worksheet.write(riga, col, valore, formato)

    # Impostazioni opzionali per migliorare la leggibilità
    worksheet.set_row(riga, 25)
    #worksheet.set_column('A:Q', 15)

def formattazione_finale_riassuntivo(wb, ws, riga):
    ws.write(riga,1, "Attenzione:", highlight_blu_lft)
    ws.write(riga,2, "", highlight_blu_lft)
    ws.write(riga,3, "", highlight_blu_lft)
    ws.write(riga,4, "", highlight_blu_lft)

    ws.write(riga+1,1,"• Il preventivo non comprende elementi di cablaggio e collegamenti di terra.")
    ws.write(riga+2,1,"• Si declina ogni responsabilità dovuta al mancato controllo della presente da parte dell'interessato.")
    ws.write(riga+3,1,"• Al fine di evitare inconvenienti Vi chiediamo di controllare attentamente gli articoli quotati")
    ws.write(riga+4,1,"   rispetto alla Vostra richiesta.")
    ws.write(riga+5,1,"• Quanto non espressamente indicato deve ritenersi escluso.")
    riga += 7
    ws.write(riga,1, "Note:", highlight_blu_lft)
    ws.write(riga,2, "", highlight_blu_lft)
    ws.write(riga,3, "", highlight_blu_lft)
    ws.write(riga,4, "", highlight_blu_lft)

    riga += 7
    ws.write(riga,1, "Condizioni di fornitura:", highlight_blu_lft)
    ws.write(riga,2, "", highlight_blu_lft)
    ws.write(riga,3, "", highlight_blu_lft)
    ws.write(riga,4, "", highlight_blu_lft)

    ws.write(riga+1,1,"Consegna:")
    ws.write(riga+2,1,"Pagamento:")
    ws.write(riga+3,1,"Sconto:")
    ws.write(riga+4,1,"Validità offerta:")
    ws.write(riga+1,3,"Da definire")
    ws.write(riga+2,3,"Da definire")
    ws.write(riga+3,3,"Da definire")
    ws.write(riga+4,3,"60gg")
    ws.write(riga+6,1,"I prezzi si intendono IVA esclusa e sono comprensivi di eco-contributo RAEE ove applicabile.")
    ws.write(riga+7,1,"Listino di riferimento 05-2025")
    ws.write(riga+11,1,"Con la speranza che quanto sopra possa essere di Vs. gradimento e in attesa di un Vs. gentile riscontro,")
    ws.write(riga+12,1,"cogliamo l'occasione per porgerVi cordiali saluti")
    ws.write(riga+14,3,"EATON INDUSTRIES (Italy) s.r.l.", blu_lft_format)
    ws.write_formula(riga+15,3,"=D9",left_align_bold)
    riga += 19
    ws.write(riga,1,"• Questa offerta è soggetta all'applicazione delle Condizioni Generali di EATON INDUSTRIES (ITALY) SRL,")
    ws.write(riga+1,1,"   che trovate al link www.eaton.eu/termsconditions.")
    ws.write(riga+2,1,"• Qualsiasi ordine o accordo, basato o risultante da questa offerta o alle successive revisioni della stessa,")
    ws.write(riga+3,1,"   saranno governate dalle Condizioni succitate, con l'esclusione di qualsiasi altra condizione o termine,")
    ws.write(riga+4,1,"   in particolare le condizioni generali di acquisto del Cliente.")
    ws.write(riga+5,1,"• Nessuna variazione o modifica a queste Condizioni Generali di Vendita di Servizi EATON INDUSTRIES (ITALY) SRL")
    ws.write(riga+6,1,"   sarà valida se non esplicitamente accettata per iscritto dalla EATON INDUSTRIES (ITALY) SRL.")
    ws.write(riga+8,1,"www.eaton.eu/termsconditions", blu_lft_format)

def formattazione_finale_preventivo(wb, ws, riga):
    ws.write(riga,0, "LEADTIME")
    ws.write(riga+1,0, "(valori indicativi)")
    ws.write(riga,2, "5 = 5 gg    7 = 7gg   A = 2 sett.   B = 3 sett.   C = 4 sett.   D =  5 sett.   E = 6 sett.   F = 7 sett.   G = 8 sett.")
    ws.write(riga+1,2,"H = 9 sett.   I = 10 sett.   J = 11 sett.   K = 12 sett.   L = 13 sett.   M = 14 sett.   N = 15 sett.   O = 16 sett.")
    ws.write(riga+2,2,"P =  17 sett.   Q = 18 sett.   R = 19 sett.   S = 20 sett.   T = 21 sett.   U = 22 sett.   V = 23 sett.   X = 24 sett. ")
    ws.write(riga+3,2,"Y = 25 sett.   Z = 26 sett.   0 = da richiedere")

formattazione_iniziale(riga_quadri)

# Scrivi i dati di sconti_df nel foglio "Famiglie Sconto"
for row_idx, row in sconti_df.iterrows():
    for col_idx, value in enumerate(row[:5]):  # Solo colonne A–E
        if col_idx < 2:
            famiglieSheet.write(row_idx, col_idx, value)
        else:
            famiglieSheet.write(row_idx, col_idx, value, percent_format)
        if row_idx == 0:
            famiglieSheet.write(row_idx, 0, "FAMIGLIA", highlight)
            famiglieSheet.write(row_idx, 1, "DESCRIZIONE", highlight)
            famiglieSheet.write(row_idx, 2, "Sc1%", highlight)
            famiglieSheet.write(row_idx, 3, "Sc2%", highlight)
            famiglieSheet.write(row_idx, 4, "Sc3%", highlight)
            famiglieSheet.write(row_idx, 5, "IMPORTO", highlight)
            famiglieSheet.write(row_idx, 7, "TOTALE", highlight)
            famiglieSheet.write_formula(1, 7, f'=SUM(F2:F1000)', euro_bold)
        else:
            famiglieSheet.write_formula(row_idx, 5, f'=SUMIF(Preventivo!I:I,A{row_idx + 1},Preventivo!F:F)')

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
                fmt = center_align

            if isinstance(cell, str) and cell.startswith("="):
                worksheet.write_formula(start_row + row_idx + 1, col_idx, cell, fmt)
            else:
                worksheet.write(start_row + row_idx + 1, col_idx, cell, fmt)

    worksheet.write(start_row, 0, file_name, blu_lft_format)
    worksheet.write(start_row, 2, "", blu_lft_format)
    actual_row_excel = start_row + row_idx + 2
    worksheet.write(actual_row_excel, 2, "Totale", blu_rth_format)
    worksheet.write(actual_row_excel, 3, f"=A{start_row + 1}", blu_lft_format)
    worksheet.write(actual_row_excel, 5, f"=SUM(F{start_row + 2}:F{actual_row_excel})", euro_blu)

    riassuntivoSheet.write_formula(start_row_rias-1, 1, f'=Preventivo!A{start_row + 1}')
    riassuntivoSheet.write_formula(start_row_rias-1, 2, f'=Preventivo!C{start_row + 1}')
    riassuntivoSheet.write_formula(start_row_rias-1, 3, f'=Preventivo!F{actual_row_excel+1}', euro)
    totaleOfferta_str += f"+F{actual_row_excel+1}"
    start_row = actual_row_excel + 2
    start_row_rias += 1

worksheet.write(start_row + 2, 2, "Totale Offerta", blu_rth_format)
worksheet.write(start_row + 2, 5, totaleOfferta_str, euro_blu)
riassuntivoSheet.write_formula(start_row_rias+1, 2, f'=Preventivo!C{start_row + 3}', euro_bold)
riassuntivoSheet.write_formula(start_row_rias+1, 3, f'=Preventivo!F{start_row + 3}', euro_bold)

formattazione_finale_riassuntivo(workbook, riassuntivoSheet, start_row_rias + 5)
formattazione_finale_preventivo(workbook, worksheet, start_row + 6)
workbook.close()

# Mostra popup finale
messagebox.showinfo("Completato", "Elaborazione completata con successo!\nFile generato: Output_XMOEM.xlsx")
