from tkinter import *
from tkinter import ttk
from tkinter import Canvas
from tkinter import messagebox
from PIL import Image, ImageTk
from DB import completar_informacoes, consolidar_dados, Processar_Demandas, limpar_erros, obter_erros, adicionar_erro
import pandas as pd
import re
import os
import sys
import threading


#------------------------------- COMERNTS --------------------------------- \\

# 800012939/20812 Edcha - Viajante novo calcula volume da funilaria. Exemplo, edscha do viajante novo saiu volume maior do que do acionamento antigo porque esse ultimo não calculava o volume da funilaria. 
#522705650 Mareli (1097/1092) - Viajante novo calcula mas o antigo não calculava, por isso o volume do viajante novo saiu maior do que do antigo.



def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # When running from the .exe
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        # When running from source
        return os.path.join(os.path.abspath("."), relative_path)

import warnings # <-- 1. Import the library

# 2. Add these lines to ignore the specific warnings from the Excel reader
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message=".*OLE2.*"
)
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message=".*CODEPAGE.*"
)
warnings.filterwarnings(
    "ignore",
    message="^WARNING .*" # Hides the file size warnings which don't have a category
)

# Always resolve paths relative to the exe's location (or source root), never os.getcwd()
if getattr(sys, 'frozen', False):
    caminho_base = os.path.dirname(sys.executable)
else:
    caminho_base = os.path.dirname(os.path.abspath(__file__))
# --- START: Global variables for filtering ---
# Stores the complete, unfiltered data from the Treeview
original_tree_data = []
# A dictionary to hold the filter Combobox widgets
filter_widgets = {}
# --- END: Global variables ---


# ------------------- carregar veículos dinâmicos -------------------
def load_veiculos(caminho_base):
    possible_files = [
        os.path.join(caminho_base, "BD", "VEÍCULOS.xlsx"),
        os.path.join(caminho_base, "BD", "VEICULOS.xlsx"),
        os.path.join(caminho_base, "BD", "Veiculos.xlsx"),
        os.path.join(caminho_base, "BD", "VEICULOS.xls")
    ]
    for fpath in possible_files:
        if os.path.exists(fpath):
            try:
                df_veh = pd.read_excel(fpath, sheet_name=0, dtype=str)  # read as str to be safe
                # normalize column names (case-insensitive)
                cols = {c.strip().upper(): c for c in df_veh.columns}
                # find code column (prefer "COD VEICULO" or similar)
                code_col = None
                desc_col = None
                for key_upper, orig in cols.items():
                    if "COD" in key_upper and "VEIC" in key_upper:
                        code_col = orig
                    if "DESCR" in key_upper or "DESC" in key_upper:
                        desc_col = orig
                # fallback: use first column as code and second (or next) as desc
                if code_col is None and len(df_veh.columns) >= 1:
                    code_col = df_veh.columns[0]
                if desc_col is None and len(df_veh.columns) >= 2:
                    # try second column
                    desc_col = df_veh.columns[1]
                if code_col is None or desc_col is None:
                    # can't map properly from this file
                    continue
                veic_map = {}
                for _, r in df_veh.iterrows():
                    desc = str(r.get(desc_col, "")).strip()
                    code_raw = r.get(code_col, "")
                    # try to convert code to int if possible, else keep as string
                    try:
                        code = int(float(str(code_raw).strip()))
                    except Exception:
                        code = str(code_raw).strip()
                    if desc:
                        # store only the original description as display key
                        if desc not in veic_map:
                            veic_map[desc] = code
                if veic_map:
                    return veic_map
            except Exception as e:
                print(f"[WARN] Could not read vehicles file {fpath}: {e}")
    # If we get here, no good file found
    return None

# Keep your original static mapping as a fallback so behavior remains unchanged if file is missing.
_FALLBACK_VEICULOS_DISPLAY = {
    'BIG SIDER': 6, 'BITREM': 7, 'CARRETA': 4, 'CARRETA LINE HAUL': 14,
    'CARRETA REBAIXADA': 9, 'CTNR 20': 15, 'CTNR 40': 16, 'FIORINO': 11,
    'RODOTREM': 8, 'TRUCK 3M': 3, 'TRUCK 3M ALONGADO': 18, 'TRUCK 3M PLUS': 13,
    'TRUCK ALONGADO': 17, 'TRUCK VIAGEM': 2, 'TRUCK VIAGEM PLUS': 12, 'VAN': 10,
    'VANDERLEA': 5, 'VEÍCULO 3/4': 1, 'TRUCK SIDER': 2
}

# load display dict (used for labels) and build lookup dict (case-insensitive) used for mapping
veiculos_display = load_veiculos(caminho_base) or _FALLBACK_VEICULOS_DISPLAY
# build lookup dict (includes uppercase keys for robustness)
veiculos_lookup = {}
for k, v in veiculos_display.items():
    veiculos_lookup[k] = v
    veiculos_lookup[k.upper()] = v
# ------------------------------------------------------------------


def show_temporary_message(master, title, mensagem, kind="info", timeout=10000):
    try:
        top = Toplevel(master)
        top.title(title)
        top.transient(master)
        top.attributes("-topmost", True)
        top.resizable(False, False)

        # Compact frame to mimic the native messagebox layout
        frm = Frame(top, padx=12, pady=8)
        frm.pack(fill=BOTH, expand=True)

        # Use a simple Label with wraplength to resemble messagebox text
        lbl = Label(frm, text=mensagem, justify=LEFT, anchor='w', wraplength=420)
        lbl.pack(fill=BOTH, expand=True)

        # Button frame to center the OK button similar to messagebox
        btn_frm = Frame(frm)
        btn_frm.pack(fill=X, pady=(8, 0))
        btn = Button(btn_frm, text="OK", width=10, command=top.destroy)
        btn.pack(side=RIGHT)

        # center relative to master if possible
        try:
            top.update_idletasks()
            mw = master.winfo_width()
            mh = master.winfo_height()
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            w = top.winfo_width()
            h = top.winfo_height()
            x = mx + (mw // 2) - (w // 2)
            y = my + (mh // 2) - (h // 2)
            top.geometry(f"+{x}+{y}")
        except Exception:
            pass

        top.after(timeout, top.destroy)
    except Exception:
        # fallback to blocking messagebox
        if kind == "warning":
            messagebox.showwarning(title, mensagem)
        else:
            messagebox.showinfo(title, mensagem)



def get_vehicle_code(nome_veiculo):
    """
    Robust lookup: try exact name, stripped, upper-case, and finally fallback to None.
    Uses veiculos_lookup (built from the display mapping).
    """
    if nome_veiculo is None:
        return None
    s = str(nome_veiculo).strip()
    if s in veiculos_lookup:
        return veiculos_lookup[s]
    su = s.upper()
    if su in veiculos_lookup:
        return veiculos_lookup[su]
    return None


def normalizar_codigos(campo):
    if pd.isna(campo):
        return []
    return [c.strip() for c in re.split(r'\s*,\s*', str(campo).strip()) if c.strip()]


def input_demanda(sheet_name=None):
    """
    Simplified demand processing: just read Excel files and return all columns.
    No FLUXO matching, no filtering - the Excel file already has all required info.
    
    Args:
        sheet_name: optional sheet name to read from demand files (Geral, Sábado, Domingo)
    
    Returns:
        DataFrame with all columns from Excel files
    """
    # Clear previous errors
    limpar_erros()
    
    # Just read the Excel files - all columns are already there
    df_final = Processar_Demandas(sheet_name=sheet_name)
    
    if df_final.empty:
        adicionar_erro("Nenhum dado foi processado. Verifique os arquivos MVM.", "ERRO")
        return pd.DataFrame()
    
    # # Display summary for user validation
    # print(f"\n{'='*60}")
    # print(f"DADOS CARREGADOS: {len(df_final)} linhas processadas")
    # print(f"{'='*60}")
    # print(f"\nColunas disponíveis: {', '.join(df_final.columns.tolist())}")
    # print(f"\nPrimeiras linhas:")
    # print(df_final.head(10).to_string())
    # print(f"\n{'='*60}\n")
    
    # Save to Template.xlsx for further processing
    df_final.to_excel("Template.xlsx", index=False)
    
    return df_final




def apply_filters(event=None):
    """
    Filters the Treeview using "contains" logic for typed text.
    Also handles dropdown selections.
    """
    if event and event.widget.get() == "-- All --":
        event.widget.set('')

    tree.delete(*tree.get_children())

    filters = {col: widget.get() for col, widget in filter_widgets.items()}
    
    column_ids = tree["columns"]

    for row_values in original_tree_data:
        match = True
        row_dict = dict(zip(column_ids, row_values))

        for col_id, filter_value in filters.items():
            if filter_value:
                cell_value = str(row_dict.get(col_id, "")).lower()
                text_to_find = filter_value.lower()
                if text_to_find not in cell_value:
                    match = False
                    break
        
        if match:
            tree.insert("", END, values=row_values)


# --------------------- GUI (mantive seu design e cores originais) ---------------------
janela = Tk()
try:
    img = Image.open(resource_path("carreta.png")).resize((140, 100))
    caminhao_img = ImageTk.PhotoImage(img)
except Exception as e:
    print(f"Erro ao carregar imagem da carreta: {e}")
    caminhao_img = None
janela.title("SATURAÇÃO REALIZADA")
janela.geometry("1400x700")
janela.state('zoomed')
janela.config(bg="#002855")

frame_principal = Frame(janela, bg="#002855")
frame_principal.pack(fill=BOTH, expand=True, pady=(0, 0))

frame_top = Frame(frame_principal, bg="#002855")
frame_top.pack(fill=X, padx=10, pady=5)

# Configure grid columns for proper spacing
frame_top.grid_columnconfigure(0, weight=0, minsize=500)
frame_top.grid_columnconfigure(1, weight=1, minsize=360)
frame_top.grid_columnconfigure(2, weight=0)

frame_selecao = Frame(frame_top, bg="#002855")
frame_selecao.grid(row=0, column=0, sticky='nw', padx=10)

# Simplified controls: only Flechinha dropdown and Atualizar button
Label(frame_selecao, text="Processar Arquivos MVM:", font=("Arial", 12, "bold"), bg="#002855", fg="#FFCC00").grid(row=0, column=0, columnspan=2, pady=(5, 10), sticky='w')

frame_controls = Frame(frame_selecao, bg="#002855")
frame_controls.grid(row=1, column=0, columnspan=2, sticky='w')

style = ttk.Style()
style.theme_use('clam')
style.configure("Highlight.TButton", font=("Arial", 10, "bold"), background="#FFCC00",
                foreground="#002855", padding=(12, 8), borderwidth=2, relief="raised")
style.map("Highlight.TButton", background=[('active', '#FFD633'), ('!disabled', '#FFCC00')])

# Atualizar button
btn_atualizar = ttk.Button(frame_controls, text="Atualizar Dados",
                           command=lambda: atualizar(), style="Highlight.TButton")
btn_atualizar.pack(side=RIGHT, padx=15)

# Configure the dropdown list colors
janela.option_add('*TCombobox*Listbox.background', '#FFCC00')
janela.option_add('*TCombobox*Listbox.foreground', '#002855')
janela.option_add('*TCombobox*Listbox.selectBackground', '#FFD633')
janela.option_add('*TCombobox*Listbox.selectForeground', '#002855')


frame_resumo = Frame(frame_top, bg="#002855")
frame_resumo.grid(row=0, column=1, sticky='w', padx=(0, 10))


items = ["Ocupação Total", "Qtd Veículos", "Volume Total", "Peso Total", "Embalagens"]

# Split items into groups of 2 per widget
groups = [items[i:i+2] for i in range(0, len(items), 2)]

for group in groups:
    tree_resumo = ttk.Treeview(frame_resumo, columns=("Info", "Valor"), show="headings", height=2)
    tree_resumo.heading("Info", text="Info")
    tree_resumo.heading("Valor", text="Valor")
    tree_resumo.column("Info", width=140, anchor='center')
    tree_resumo.column("Valor", width=120, anchor='center')
    tree_resumo.pack(side="left", padx=10)

    for item in group:
        tree_resumo.insert("", END, values=(item, ""))





frame_bottom = Frame(frame_principal, bg="white")
frame_bottom.pack(fill=BOTH, expand=True, padx=10, pady=(0, 0))

# Loading label - position in data area
loading_label = Label(frame_bottom, text="Processando... Por favor, aguarde.",
                      font=("Arial", 14, "bold"), bg="#002855", fg="#FFCC00",
                      relief="solid", borderwidth=2, padx=15, pady=8)

frame_filters = Frame(frame_bottom, bg="#f0f0f0")
frame_filters.pack(fill=X, pady=(5, 2))

# Create a frame for the treeview with scrollbars
tree_frame = Frame(frame_bottom, bg="white")
tree_frame.pack(fill=BOTH, expand=True)

scroll_y = Scrollbar(tree_frame, orient=VERTICAL)
scroll_y.pack(side=RIGHT, fill=Y)

scroll_x = Scrollbar(tree_frame, orient=HORIZONTAL)
scroll_x.pack(side=BOTTOM, fill=X)

tree = ttk.Treeview(tree_frame, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
tree.pack(fill=BOTH, expand=True)

scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)

style.configure("Treeview.Heading", background="#002855", foreground="#FFCC00",
                font=("Arial", 8, "bold"), relief="flat")
style.map("Treeview.Heading", background=[('active', '#004080')])


def atualizar():
    # --- Start spinner ---
    start_loading()

    def processar():
        try:
            # Limpa erros anteriores antes de processar
            limpar_erros()
            
            # Get selected sheet name from Flechinha dropdown (only if not empty)
            # selected_sheet = flechinha_var.get() if flechinha_var.get() else None
            selected_sheet = None
            
            # Load data - simplified call, no parameters needed except sheet_name
            df_final = input_demanda(sheet_name=selected_sheet)
            
            if df_final.empty:
                janela.after(0, lambda: messagebox.showerror("Erro", "Nenhum dado foi carregado. Verifique os arquivos MVM."))
                loading_label.spinning = False
                janela.after(0, lambda: finalizar_status("Erro: Nenhum dado carregado", "red"))
                return
            
            # Validate required columns exist
            required_cols = ['DESENHO', 'QTDE', 'VEÍCULO']
            missing_cols = [col for col in required_cols if col not in df_final.columns]
            if missing_cols:
                error_msg = f"Colunas obrigatórias faltando: {', '.join(missing_cols)}"
                janela.after(0, lambda: messagebox.showerror("Erro", error_msg))
                loading_label.spinning = False
                janela.after(0, lambda msg=error_msg: finalizar_status(f"Erro: {msg}", "red"))
                return
            
            # For now, use vehicle from first row as default (until we update completar_informacoes)
            # TODO: Update completar_informacoes to not require vehicle parameter
            default_veiculo = df_final['VEÍCULO'].iloc[0] if 'VEÍCULO' in df_final.columns and not df_final.empty else 1
            
            # Process data with database mappings
            completar_informacoes(
                tree, default_veiculo, tree_resumo, caminhao_img, usar_manual=False
            )

            global original_tree_data
            original_tree_data = [tree.item(child)['values'] for child in tree.get_children()]

            columns_to_filter = ['COD FORNECEDOR', 'FORNECEDOR', 'DESENHO']
            all_table_columns = list(tree["columns"])

            if not filter_widgets:
                for widget in frame_filters.winfo_children():
                    widget.destroy()

                for col_id in columns_to_filter:
                    if col_id in all_table_columns:
                        col_frame = Frame(frame_filters)
                        col_frame.pack(side=LEFT, padx=2, fill=X, expand=True)
                        Label(col_frame, text=col_id, font=("Arial", 8)).pack(anchor='w')
                        combo = ttk.Combobox(col_frame, font=("Arial", 9))
                        combo.pack(fill=X)
                        combo.bind('<KeyRelease>', apply_filters)
                        combo.bind('<<ComboboxSelected>>', apply_filters)
                        filter_widgets[col_id] = combo

            for col_id, combo in filter_widgets.items():
                col_index = all_table_columns.index(col_id)
                unique_values = sorted(
                    list(set(str(row[col_index]) for row in original_tree_data if str(row[col_index]).strip()))
                )
                combo['values'] = ["-- All --"] + unique_values
                combo.set('')

            # Consolidate data
            consolidar_dados(use_manual=False, manual_veiculo=None)
            
            # Mostra erros/avisos se houver
            erros = obter_erros()
            if erros:
                # Separa erros e avisos
                erros_criticos = [e for e in erros if '[ERRO]' in e]
                avisos = [e for e in erros if '[AVISO]' in e]
                
                mensagem = ""
                if erros_criticos:
                    mensagem += "ERROS ENCONTRADOS:\n" + "\n".join(erros_criticos) + "\n\n"
                if avisos:
                    mensagem += "AVISOS:\n" + "\n".join(avisos)
                
                # Mostra popup com os erros (auto-closing)
                if erros_criticos:
                    janela.after(0, lambda: show_temporary_message(janela, "Atenção - Problemas Detectados", mensagem, kind="warning", timeout=10000))
                else:
                    janela.after(0, lambda: show_temporary_message(janela, "Avisos de Processamento", mensagem, kind="info", timeout=10000))

            # --- Stop spinner and show success ---
            loading_label.spinning = False
            janela.after(0, lambda: finalizar_status("Concluído com sucesso!", "#2e8b57"))

        except Exception as e:
            error_msg = str(e)  # Capture error message before list comprehensions shadow 'e'
            adicionar_erro(error_msg, "AVISO")
            # Mostra erros/avisos se houver
            erros = obter_erros()
            if erros:
                # Separa erros e avisos
                erros_criticos = [e for e in erros if '[ERRO]' in e]
                avisos = [e for e in erros if '[AVISO]' in e]
                
                mensagem = ""
                if erros_criticos:
                    mensagem += "ERROS ENCONTRADOS:\n" + "\n".join(erros_criticos) + "\n\n"
                if avisos:
                    mensagem += "AVISOS:\n" + "\n".join(avisos)
                
                # Mostra popup com os erros (auto-closing)
                if erros_criticos:
                    janela.after(0, lambda: show_temporary_message(janela, "Atenção - Problemas Detectados", mensagem, kind="warning", timeout=10000))
                else:
                    janela.after(0, lambda: show_temporary_message(janela, "Avisos de Processamento", mensagem, kind="info", timeout=10000))
            loading_label.spinning = False
            janela.after(0, lambda msg=error_msg: finalizar_status(f"Erro: {msg}", "red"))

    threading.Thread(target=processar, daemon=True).start()


def start_loading():
    spinner_chars = ['|', '/', '--', '\\']
    loading_label.place(relx=0.5, rely=0.5, anchor='center')
    loading_label.lift()
    janela.update_idletasks()

    def spin():
        i = 0
        while getattr(loading_label, "spinning", False):
            loading_label.config(text=f"Processando... {spinner_chars[i % len(spinner_chars)]}")
            i += 1
            janela.update_idletasks()
            threading.Event().wait(0.1)  # short delay for animation

    loading_label.spinning = True
    threading.Thread(target=spin, daemon=True).start()

  
def finalizar_status(msg, color):
    """Atualiza o texto e esconde após 2 segundos"""
    # Check if Flechinha was selected
    # flechinha_selected = flechinha_var.get() != ''
    
    if "sucesso" in msg.lower():
        
        loading_label.config(text=msg, fg="#FFCC00", bg="#2e8b57", relief="solid", borderwidth=2)
    else:
        loading_label.config(text=msg, fg="#FFCC00", bg="#002855", relief="solid", borderwidth=2)
    janela.after(2000, loading_label.place_forget)


footer_frame = Frame(janela, bg="#002855", height=18)
footer_frame.pack(side=BOTTOM, fill=X)
footer_frame.pack_propagate(False)

footer_left = Label(footer_frame, text="DHL → STELLANTIS", 
                    font=("Arial", 7, "bold"), bg="#002855", fg="#FFCC00", 
                    anchor="w", padx=8, pady=0)
footer_left.pack(side=LEFT, fill=Y)

footer_right = Label(footer_frame, text="Developer: Vincent Pernarh", 
                     font=("Arial", 7), bg="#002855", fg="#FFCC00", 
                     anchor="e", padx=8, pady=0)
footer_right.pack(side=RIGHT, fill=Y)

# ------------------- Database Update Check -------------------
# Check and update database files from SharePoint if needed
# This runs after GUI is created so we can show progress in the loading_label

def update_progress_callback(message):
    """Callback to update the loading label with progress messages"""
    loading_label.config(text=message, bg="#002855", fg="#FFCC00")
    loading_label.place(relx=0.5, rely=0.5, anchor='center')
    loading_label.lift()
    janela.update_idletasks()

def check_database_updates():
    """Check and update database files in a thread"""
    # Use resource_path to handle both dev and PyInstaller paths
    update_db_path = resource_path('Update DataBase')
    if update_db_path not in sys.path:
        sys.path.insert(0, update_db_path)
    try:
        from Update_Manager import check_and_update_files
        
        # Show initial message
        update_progress_callback("Verificando atualizações do banco de dados...")
        
        # Check files and update if older than 5 days
        update_result = check_and_update_files(
            max_age_days=5, 
            silent=False,
            progress_callback=update_progress_callback
        )
        
        if update_result.get("updated"):
            janela.after(0, lambda: finalizar_status("✓ Banco de dados atualizado!", "#2e8b57"))
        else:
            janela.after(0, lambda: finalizar_status("✓ Banco de dados atualizado!", "#2e8b57"))
            
    except Exception as e:
        print(f"⚠️ Aviso: Não foi possível verificar atualizações: {e}")
        janela.after(0, lambda: loading_label.place_forget())

# Start update check in background thread
threading.Thread(target=check_database_updates, daemon=True).start()
# ---------------------------------------------------------------

janela.mainloop()

