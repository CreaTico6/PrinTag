# TiSoft 2026
# PrinTag v1.6.5

import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
from decimal import Decimal, ROUND_HALF_UP

PRINTER_PATH = r"\\ps10\etiquetas"

def formatar_preco_epl(valor):
    try:
        f_valor = float(str(valor).replace(',', '.'))
        return f"{f_valor:.2f}".replace('.', ',')
    except:
        return "0,00"

def converter_para_float(valor):
    try:
        if not valor or str(valor).strip() == "": return 0.0
        val = str(valor).replace(',', '.')
        return float(val)
    except:
        return 0.0

def arredondar_excel(valor):
    """Aplica o algoritmo de Arredondamento Comercial do Excel"""
    try:
        return float(Decimal(str(valor)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    except:
        return 0.0

def dividir_nome_inteligente(nome):
    nome = str(nome).strip()
    if len(nome) <= 20:
        return nome, ""
    
    if len(nome) > 20 and nome[20] == ' ':
        return nome[:20].strip(), nome[20:].strip()[:20]
        
    ultimo_espaco = nome[:20].rfind(' ')
    if ultimo_espaco == -1:
        return nome[:20], nome[20:40]
    else:
        return nome[:ultimo_espaco], nome[ultimo_espaco+1:].strip()[:20]

def enviar_para_zebra(nome1, nome2, codigo, preco_ant, preco_act, perc_desc=0):
    ant = converter_para_float(preco_ant)
    act = arredondar_excel(preco_act)
    
    p_ant_str = formatar_preco_epl(ant)
    p_act_str = f"{formatar_preco_epl(act)} EUR"
    
    p_desc = "0"
    if converter_para_float(perc_desc) > 0:
        p_desc = str(int(converter_para_float(perc_desc)))
    elif ant > 0 and act < ant:
        p_desc = str(math.floor(((ant - act) / ant) * 100))
    
    texto_promo = f"-{p_desc}%"
    offset_promo = 260

    largura_preco_final = len(p_act_str) * 24
    offset_direita = 400 - largura_preco_final
    if offset_direita < 80: offset_direita = 80

    epl = f"""N
A80,15,0,2,1,1,N,"{str(nome1)[:20]}"
A80,40,0,2,1,1,N,"{str(nome2)[:20]}"
A80,65,0,1,1,1,N,"{str(codigo)[:15]}"
A81,65,0,1,1,1,N,"{str(codigo)[:15]}"
A{offset_promo},75,0,4,1,1,N,"{texto_promo}"
A80,120,0,2,1,1,N,"Antes: {p_ant_str} EUR"
A{offset_direita},160,0,4,1,1,N,"{p_act_str}"
P1
"""
    temp_file = os.path.join(os.environ['TEMP'], "printag_job.txt")
    try:
        with open(temp_file, "w", encoding="cp1252", errors="replace") as f:
            f.write(epl)
        subprocess.run(f'copy /b "{temp_file}" "{PRINTER_PATH}"', shell=True, check=True)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro: {e}")
    finally:
        if os.path.exists(temp_file): os.remove(temp_file)

def processar_e_abrir_editavel():
    f_path = filedialog.askopenfilename(title="Selecionar export.csv", filetypes=[("CSV", "*.csv")])
    if not f_path: return
    try:
        try:
            df_raw = pd.read_csv(f_path, sep=';', encoding='utf-8', quotechar='"').fillna('')
        except UnicodeDecodeError:
            df_raw = pd.read_csv(f_path, sep=';', encoding='cp1252', quotechar='"').fillna('')
            
        df_raw.columns = [c.strip().upper().replace('"', '') for c in df_raw.columns]
        
        col_designacao = [c for c in df_raw.columns if 'DESIGNA' in c][0]
        col_codigo = [c for c in df_raw.columns if 'C' in c and 'DIGO' in c][0]
        col_pvp = [c for c in df_raw.columns if 'PVP' in c][0]

        df_ready = pd.DataFrame()
        nomes_separados = df_raw[col_designacao].apply(dividir_nome_inteligente)
        df_ready['Nome1'] = nomes_separados.apply(lambda x: x[0])
        df_ready['Nome2'] = nomes_separados.apply(lambda x: x[1])
        
        df_ready['Codigo'] = df_raw[col_codigo].astype(str)
        df_ready['Preco_Ant'] = df_raw[col_pvp]
        df_ready['%'] = 0
        df_ready['Preco_Act'] = 0

        temp_xlsx = "Etiquetas_para_Imprimir.xlsx"
        writer = pd.ExcelWriter(temp_xlsx, engine='openpyxl')
        df_ready.to_excel(writer, index=False, sheet_name='Etiquetas')
        
        worksheet = writer.book['Etiquetas']
        larguras = {'A': 22, 'B': 22, 'C': 15, 'D': 12, 'E': 8, 'F': 12}
        for col_letter, width in larguras.items():
            worksheet.column_dimensions[col_letter].width = width

        for row_num in range(2, len(df_ready) + 2):
            formula = f"=ROUND(D{row_num}*(1-E{row_num}/100), 2)"
            worksheet[f'F{row_num}'] = formula

        writer.close()
        os.startfile(temp_xlsx)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro: {e}")

def imprimir_lote():
    f_path = filedialog.askopenfilename(title="Escolher Ficheiro", filetypes=[("Excel ou CSV", "*.xlsx *.csv")])
    if not f_path: return
    try:
        if f_path.endswith('.xlsx'):
            df = pd.read_excel(f_path, engine='openpyxl')
        else:
            try:
                df = pd.read_csv(f_path, sep=None, engine='python', encoding='utf-8')
            except:
                df = pd.read_csv(f_path, sep=None, engine='python', encoding='cp1252')

        for _, row in df.iterrows():
            p_ant = converter_para_float(row['Preco_Ant'])
            
            raw_act = row.get('Preco_Act', 0)
            p_act = converter_para_float(raw_act)
            perc_val = converter_para_float(row.get('%', 0))
            
            if p_act <= 0 and perc_val > 0:
                p_act = arredondar_excel(p_ant * (1 - (perc_val / 100)))
            elif p_act <= 0:
                p_act = p_ant
            
            enviar_para_zebra(str(row['Nome1']), str(row.get('Nome2','')), str(row['Codigo']), p_ant, p_act, perc_val)
            
        messagebox.showinfo("Sucesso", "Etiquetas impressas impacábelmente!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro: {e}")

# --- INTERFACE ---
def auto_calcular(event):
    p_ant = converter_para_float(entry_ant.get())
    if p_ant <= 0: return
    if event.widget == entry_desc_val:
        d_val = converter_para_float(entry_desc_val.get())
        entry_act.delete(0, tk.END)
        entry_act.insert(0, formatar_preco_epl(arredondar_excel(p_ant - d_val)))
    elif event.widget == entry_desc_perc:
        d_perc = converter_para_float(entry_desc_perc.get())
        entry_act.delete(0, tk.END)
        entry_act.insert(0, formatar_preco_epl(arredondar_excel(p_ant * (1 - (d_perc / 100)))))

def limitar_input(P): return len(P) <= 20

root = tk.Tk()
root.title("PrinTag v1.6.5")
root.geometry("420x760")
vcmd = (root.register(limitar_input), '%P')

tk.Label(root, text="PRINTAG 2026", font=("Arial", 11, "bold")).pack(pady=10)

f1 = tk.LabelFrame(root, text=" Identificação ", padx=10, pady=10)
f1.pack(pady=5, fill="x", padx=20)
tk.Label(f1, text="Nome Linha 1:").pack(anchor="w")
entry_n1 = tk.Entry(f1, width=45, validate='key', validatecommand=vcmd); entry_n1.pack()
tk.Label(f1, text="Nome Linha 2:").pack(anchor="w")
entry_n2 = tk.Entry(f1, width=45, validate='key', validatecommand=vcmd); entry_n2.pack()
tk.Label(f1, text="Código:").pack(anchor="w")
entry_cod = tk.Entry(f1, width=45); entry_cod.pack()

f2 = tk.LabelFrame(root, text=" Valores ", padx=10, pady=10)
f2.pack(pady=5, fill="x", padx=20)
tk.Label(f2, text="PREÇO ANTERIOR:").grid(row=0, column=0, sticky="w")
entry_ant = tk.Entry(f2, width=15); entry_ant.grid(row=0, column=1)
tk.Label(f2, text="Desc. Valor:").grid(row=1, column=0, sticky="w")
entry_desc_val = tk.Entry(f2, width=15); entry_desc_val.grid(row=1, column=1)
entry_desc_val.bind("<KeyRelease>", auto_calcular)
tk.Label(f2, text="Desc. %:").grid(row=2, column=0, sticky="w")
entry_desc_perc = tk.Entry(f2, width=15); entry_desc_perc.grid(row=2, column=1)
entry_desc_perc.bind("<KeyRelease>", auto_calcular)
tk.Label(f2, text="PREÇO FINAL:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w", pady=5)
entry_act = tk.Entry(f2, width=15, font=("Arial", 11, "bold"), bg="#ffffcc"); entry_act.grid(row=3, column=1)

tk.Button(root, text="Imprimir Individual", 
          command=lambda: enviar_para_zebra(entry_n1.get(), entry_n2.get(), entry_cod.get(), entry_ant.get(), entry_act.get(), entry_desc_perc.get()), 
          bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=30).pack(pady=10)

f_csv = tk.LabelFrame(root, text=" Gestão de Lotes (Excel/CSV) ", padx=10, pady=10)
f_csv.pack(pady=5, fill="x", padx=20)
tk.Button(f_csv, text="1. Editar Export (Gera Excel c/ Fórmulas)", command=processar_e_abrir_editavel, bg="#e3f2fd").pack(fill="x", pady=2)
tk.Button(f_csv, text="2. Imprimir Tudo (Escolher Ficheiro)", command=imprimir_lote, bg="#fce4ec").pack(fill="x", pady=2)

root.mainloop()