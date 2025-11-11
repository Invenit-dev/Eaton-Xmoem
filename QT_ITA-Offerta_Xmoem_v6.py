"""
V6 : codice convertito per usare itertuples() invece che iterrows, sia per iterare su input_df (ciclo for riga 331) che per accedere ai dati del listino_dict (riga 40), che ora contiene tuple.
    riga 155 inserito valore Default nominativo Quotation Engineer in base a suo e_number
"""

import pandas as pd
import xlsxwriter
import math
import glob
import os
import sys
from datetime import datetime
import re
user_path= os.path.expanduser("~")
e_number = os.path.basename(user_path)


# Percorso per i file inclusi nel .exe (come il logo)
if getattr(sys, 'frozen', False):
    embedded_dir = sys._MEIPASS
    current_dir = os.path.dirname(sys.executable)
else:
    embedded_dir = os.path.dirname(os.path.abspath(__file__))
    current_dir = embedded_dir

# Percorso del logo incluso nel .exe
image_path = os.path.join(embedded_dir, "Eaton-Logo.jpg")

#print(f"Cerco in: {current_dir}")
start_time = datetime.now()

print("Caricamento Listino in corso...")
# Trova il file che inizia con "00_Listino" e termina con ".xlsx", poi estrai MESE e ANNO
file_listino = next((f for f in os.listdir(current_dir) if f.startswith("00_Listino") and f.endswith(".xlsx")), None)
listino_mese = file_listino[11:13]
listino_anno = file_listino[14:18]

if False:
    file_listino_csv = next((f for f in os.listdir(current_dir) if f.startswith("00_Listino") and f.endswith(".csv")), None)
    listino_mese = file_listino_csv[11:13]
    listino_anno = file_listino_csv[14:18]
    # Carica i dati da CSV
    listino_df = pd.read_csv(os.path.join(current_dir, file_listino_csv), header=None)

# Carica i dati del Listino
listino_df = pd.read_excel(os.path.join(current_dir, file_listino), sheet_name=0, engine='openpyxl', header=None)
# Crea dizionari
listino_dict = {
    str(row[0]).strip(): row
    for row in listino_df.iloc[1:].itertuples(index=False, name=None)
}
# Carica i dati delle Famiglie Sconto e Rubrica venditori
print("Caricamento Famiglie Sconto e Rubrica venditori in corso...")
sconti_df = pd.read_excel(os.path.join(current_dir, file_listino), sheet_name=1, engine='openpyxl', header=None)
rubrica_df = pd.read_excel(os.path.join(current_dir, file_listino), sheet_name=2, engine='openpyxl', header=None)


end_time = datetime.now()
#print(f"Durata: {end_time - start_time}")

# Crea file Excel in OUTPUT con xlsxwriter
output_path = os.path.join(current_dir, "01_OUTPUT_XMOEM.xlsx")
workbook = xlsxwriter.Workbook(output_path, {'nan_inf_to_errors': True})
riassuntivoSheet = workbook.add_worksheet("Riassuntivo")
worksheet = workbook.add_worksheet("Preventivo")
if listino_mese == "99":
   famiglieSheet = workbook.add_worksheet("Famiglie")
else:
    famiglieSheet = workbook.add_worksheet("Famiglie Sconto")
rubricaSheet = workbook.add_worksheet("Rubrica")
rubricaSheet.set_tab_color('red')

# Imposta dimensioni foglio e formato pagina
worksheet.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
worksheet.set_portrait()
worksheet.set_paper(9)  # 9 = formato A4
# Imposta scalatura per adattare tutte le colonne in una pagina
worksheet.fit_to_pages(1, 0)  # 1 pagina in larghezza, altezza automatica

# Info utili alla compilazione
riga_quadri = 19
start_row = riga_quadri + 2
start_row_rias = riga_quadri + 1
totaleOfferta_str = "="
nominativi_QE = ['Michele Angelastri','Giuseppe Aru', 'Gianni Bruni','Fabrizio Genchi','Ivan Mazzarella','Chiara Spanu']
eNumber_QE = ["e1502763","e0520769","e1502777","e0415636","e0645807","e0721898"]
# Formati celle excel
formato_data = workbook.add_format({'bold': True, 'num_format': 'dd/mm/yyyy','align': 'left'})
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
    riassuntivoSheet.insert_image('B1', image_path,{'x_scale': 0.27, 'y_scale': 0.35})
    worksheet.insert_image('B1', image_path,{'x_scale': 0.27, 'y_scale': 0.35})

    col_right_num = 4
    # Info Eaton
    riassuntivoSheet.write(0,col_right_num, "EATON INDUSTRIES (ITALY) S.r.l.", left_align_bold)
    riassuntivoSheet.write(1,col_right_num, "Via San Bovio, 3", left_align)
    riassuntivoSheet.write(2,col_right_num, "20090 Segrate (MI)", left_align)
    riassuntivoSheet.write(3,col_right_num, "Tel: +39 02 959501", left_align)
    riassuntivoSheet.write(4,col_right_num, "www.eaton.it", left_align)
    # Info OFFERTA
    riassuntivoSheet.write(7,1, "Spett.le:", highlight_grn_lft)
    riassuntivoSheet.write(7,2, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(7,3, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(8,1, "Alla c.a.:", highlight_grn_lft)
    riassuntivoSheet.write(8,2, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(8,3, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(10,1, "Progetto:", highlight_grn_lft)
    riassuntivoSheet.write(10,2, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(10,3, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(11,1, "Offerta n.:", highlight_grn_lft)
    riassuntivoSheet.write(11,2, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(11,3, "", highlight_grn_lft_bold)
    riassuntivoSheet.write(12,1, "", highlight_grn_lft)
    riassuntivoSheet.write(12,2, "", highlight_grn_lft)
    riassuntivoSheet.write(12,3, "", highlight_grn_lft)
    # Ottieni la data e ora attuale
    oggi = datetime.now()
    riassuntivoSheet.write(13,1, "Data:", highlight_grn_lft)
    riassuntivoSheet.write(13,2, f"{oggi.day:02}/{oggi.month:02}/{oggi.year}", highlight_grn_lft_bold)
    riassuntivoSheet.write(13,3, "", highlight_grn_lft)
    riassuntivoSheet.write(15,1, "Con riferimento alla Vs. gradita richiesta, Vi sottoponiamo la nostra migliore quotazione", left_align)
    riassuntivoSheet.write(16,1, "relativa agli articoli di Vostro interesse:", left_align)
    # Imposta il valore di default nella cella (riga 8, colonna col_right_num)
    riassuntivoSheet.write(7,col_right_num, "Offerta redatta da:", highlight_grn_lft)
    try:
        QE_def = nominativi_QE[eNumber_QE.index(e_number)]
    except ValueError:
        QE_def = ""
    riassuntivoSheet.write(8,col_right_num, QE_def, highlight_grn_lft_bold)
    # Aggiungi la data validation al Quotation Engineer
    riassuntivoSheet.data_validation(8, col_right_num, 8, col_right_num, {
        'validate': 'list',
        'source': nominativi_QE,
        'input_message': 'Inserire nome Quotation Engineer',
        'error_message': 'Valore non valido. Seleziona dalla lista.'
    })
    

    riassuntivoSheet.write(10,col_right_num, "Venditore:",highlight_grn_lft)
    riassuntivoSheet.write(11,col_right_num, "", highlight_grn_lft_bold)
    # Scrivi i dati di rubrica_df nel foglio "Rubrica"
    for row_idx, row in rubrica_df.iterrows():
        for col_idx, value in enumerate(row):
            rubricaSheet.write(row_idx, col_idx, value)
    # Scrivi i valori unici della colonna A (indice 0) nel foglio Rubrica, colonna di appoggio (es. colonna U = indice 20)
    validation_values = rubrica_df.iloc[:, 0].dropna().astype(str).unique()
    for i, val in enumerate(validation_values):
        rubricaSheet.write(i, 20, val)  # Scrive in colonna U
    # Aggiungi la data validation in Riassuntivo!E12 usando l'intervallo Rubrica!U1:U{n}
    riassuntivoSheet.data_validation(11, col_right_num, 11, col_right_num, {
        'validate': 'list',
        'source': f'=Rubrica!U1:U{len(validation_values)}',
        'input_message': 'Inserire nome venditore',
        'error_message': 'Valore non valido. Seleziona dalla lista.'
    })
    riassuntivoSheet.write_formula(12,col_right_num, f'=VLOOKUP(Riassuntivo!E12,Rubrica!A1:G300,4,FALSE)', highlight_grn_lft)
    riassuntivoSheet.write_formula(13,col_right_num, f'=VLOOKUP(Riassuntivo!E12,Rubrica!A1:G300,5,FALSE)', highlight_grn_lft)

    riassuntivoSheet.write(18,1, "Rif.", highlight_blu)
    riassuntivoSheet.write(18,2, "Denominazione quadro", highlight_blu)
    riassuntivoSheet.write(18,3, "Prezzo", highlight_blu)

    # Inserisce formule in A1:D14 del foglio 'Preventivo'
    for row in range(14):  # righe da 0 a 13 (A1-A14)
        for col in range(5):  # colonne da 0 a 4 (A-E)
            cella_riassuntivo = xlsxwriter.utility.xl_rowcol_to_cell(row, col)
            formula = f'=IF(Riassuntivo!{cella_riassuntivo}<>"",Riassuntivo!{cella_riassuntivo}&"","")'
            formula_data = f'=IF(Riassuntivo!{cella_riassuntivo}<>"",Riassuntivo!{cella_riassuntivo},"")'
            if row == 13 and col == 2:
                worksheet.write_formula(row, col, formula_data, formato_data)
            elif col == 2:
                worksheet.write_formula(row, col, formula, left_align_bold)
            elif (row == 0 or row == 8 or row == 11) and col == col_right_num:
                worksheet.write_formula(row, col, formula, left_align_bold)
            else:
                worksheet.write_formula(row, col, formula, left_align)

    worksheet.write(15,1, "Con riferimento alla Vs. gradita richiesta, Vi sottoponiamo la nostra migliore quotazione", left_align)
    worksheet.write(16,1, "relativa agli articoli di Vostro interesse:", left_align)

    # Valori da inserire
    intestazione = [
        'Pos.', 'Codice', 'Tipo / Descrizione', f"Listino {listino_mese}-{listino_anno}", 'Q.tà', 'Tot. Listino',
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
    ws.write(riga+7,1,f"Listino di riferimento {listino_mese}-{listino_anno}")
    ws.write(riga+11,1,"Con la speranza che quanto sopra possa essere di Vs. gradimento e in attesa di un Vs. gentile riscontro,")
    ws.write(riga+12,1,"cogliamo l'occasione per porgerVi cordiali saluti")
    ws.write(riga+14,3,"EATON INDUSTRIES (Italy) s.r.l.", blu_lft_format)
    ws.write_formula(riga+15,3,"=E9",left_align_bold)
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
    # Imposta l'area di stampa da A1 a I{ultima riga scritta}
    worksheet.print_area(f'A1:I{riga+3}') 
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
excel_files = [f for f in glob.glob(os.path.join(current_dir, "*.xls*")) 
if not ( os.path.basename(f).startswith("00_Listino") or
         os.path.basename(f).startswith("01_OUTPUT")  )
]

# Cicla su tutti i file dei quadri
for file_path in excel_files:
    file_quadro = os.path.splitext(os.path.basename(file_path))[0]
    df_full = pd.read_excel(file_path, sheet_name=0, engine='openpyxl', header=None)
    print(f"Elaborazione: {file_quadro}")
    
    # Estrazione nome Quadro con regex
    match = re.match(r"^(.*?)\s*\((.*?)\)$", file_quadro)
    if match:
        riferimentoQuadro_raw = match.group(1).strip()
        denominazioneQuadro = match.group(2).strip()
        # Se il riferimento inizia con due cifre + underscore, rimuovili (servono solo per enumerazione)
        if re.match(r"^\d{2}_", riferimentoQuadro_raw):
            riferimentoQuadro = riferimentoQuadro_raw[3:]
        else:
            riferimentoQuadro = riferimentoQuadro_raw
    else:
        print("\t\tNome file con formato non standard: Numero_SiglaQuadro (Nome Quadro)")
        riferimentoQuadro = file_quadro
        denominazioneQuadro = ""

    # Identifica quale sono le colonne con Codici e quantita'
    target_col = None
    for col_idx in range(min(30, df_full.shape[1])):
        for row_idx in range(min(7, df_full.shape[0])):
            cell_value = str(df_full.iat[row_idx, col_idx])
            if cell_value in ["CODICI", "Codice"]:
                target_col = col_idx
                break
        if target_col is not None and target_col + 1 < df_full.shape[1]:
            input_df = df_full.iloc[:, [target_col, target_col + 1]]
            break

    output_data = []
    highlight_rows = set()

    for idx, row_input in enumerate(input_df.itertuples(index=False, name=None)):
        code = row_input[0]

        # Salta righe vuote o che contengono intestazioni comuni
        if pd.isna(code) or str(code).strip().upper() in ["CODICI", "CODICE", "QTY", "QUANTITÀ", ""]:
            continue

        code_str = str(code).strip()
        row = listino_dict.get(code_str)  # cerca il codice nel dizionario listino_dict
        qty = row_input[1] if len(row_input) > 1 else ""
        excel_row = (start_row + 1) + (len(output_data) + 1)

        if row is not None:
            # row è un TUPLE (sequenza ordinata di elementi non modificabile), quindi si accede con gli indici
            evidenzia = (
                isinstance(row[7], str) and
                "non compatibile con il nuovo sistema ProfiSNAP" in row[7]
            )
            if evidenzia:
                highlight_rows.add(len(output_data))
            
            row_data = [
                "",  # A
                code_str,  # B
                row[1],  # C
                f"=IF(O{excel_row}=0,IF(N{excel_row}=0,ROUND(J{excel_row}*(1-K{excel_row})*(1-L{excel_row})*(1-M{excel_row}),2),ROUND(J{excel_row}*(1-N{excel_row}),2)),O{excel_row})",  # D
                qty,  # E
                f"=D{excel_row}*E{excel_row}",  # F
                row[2],  # G
                row[3],  # H
                row[6],  # I
                str("{:.2f}".format(float(row[4])).replace('.', ',')) if not pd.isna(row[4]) else "N/A",  # J  Prezzo Listino
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 3, FALSE)",  # K
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 4, FALSE)",  # L
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 5, FALSE)",  # M
                "",  # N
                "",  # O
                row[8],  # P
                f"=P{excel_row}*E{excel_row}"  # Q
            ]
        else:
            # codice articolo non trovato nel listino
            row_data = [
                "", code_str, "", f"=J{excel_row}", qty, f"=D{excel_row}*E{excel_row}", "", "", "",
                "XXX",
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 3, FALSE)",  # K
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 4, FALSE)",  # L
                f"=VLOOKUP(I{excel_row}, 'Famiglie Sconto'!A:E, 5, FALSE)",  # M
                "", "", "", f"=P{excel_row}*E{excel_row}"
            ]

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

    worksheet.write(start_row, 0, f"=Riassuntivo!B{start_row_rias}", blu_lft_format)
    worksheet.write_formula(start_row, 2, f"=VLOOKUP(A{start_row + 1}, Riassuntivo!B:C, 2, FALSE)",blu_lft_format)

    actual_row_excel = start_row + row_idx + 2
    worksheet.write(actual_row_excel, 2, "Totale", blu_rth_format)
    worksheet.write(actual_row_excel, 3, f"=A{start_row + 1}", blu_lft_format)
    worksheet.write(actual_row_excel, 5, f"=SUM(F{start_row + 2}:F{actual_row_excel})", euro_blu)

    riassuntivoSheet.write(start_row_rias-1, 1, riferimentoQuadro)    # Rif.
    riassuntivoSheet.write(start_row_rias-1, 2, denominazioneQuadro)  # Denominazione quadro
    riassuntivoSheet.write_formula(start_row_rias-1, 3, f'=Preventivo!F{actual_row_excel+1}', euro) # Prezzo

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
print("\nFile Output generato con successo!")