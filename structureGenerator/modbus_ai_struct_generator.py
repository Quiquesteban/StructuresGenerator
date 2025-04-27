import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import difflib
import os
import re
import uuid
from typing import Dict, List, Optional

# Constantes
CONFIG_FILE = "modbus_ai_config.json"
DEFAULT_VALUES = {
    'address': "0",
    'name': "UnnamedVariable",
    'datatype': "UINT",
    'offset': "0",
    'unit': "",
    'description': ""
}

# Funciones auxiliares
def load_configuration() -> Dict[str, List[str]]:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def find_best_column(possible_aliases: List[str], columns: List[str]) -> Optional[str]:
    for alias in possible_aliases:
        best_match = difflib.get_close_matches(alias.lower(), [col.lower() for col in columns], n=1, cutoff=0.7)
        if best_match:
            return best_match[0]
    return None

def clean_variable_name(name: str) -> str:
    name = re.sub(r'[^\w ]', '', name)
    name = re.sub(r'^[^A-Za-z_]+', '', name)
    name = ''.join(word.capitalize() for word in name.split())
    if not name:
        name = "InvalidName"
    return name

def validate_address(address: str) -> bool:
    try:
        int(address)
        return True
    except (ValueError, TypeError):
        return False

def process_sheet(df_raw: pd.DataFrame, sheet_name: str, excel_name: str) -> None:
    df = df_raw.iloc[1:].copy()
    df.columns = [str(col).strip().lower() for col in df_raw.iloc[0]]

    column_aliases = load_configuration()
    col_map = {}
    for key, aliases in column_aliases.items():
        col_map[key] = find_best_column(aliases, df.columns)

    for field, col_name in col_map.items():
        if col_name and field in DEFAULT_VALUES:
            df[col_name] = df[col_name].fillna(DEFAULT_VALUES[field])

    if col_map['address']:
        df = df[df[col_map['address']].astype(str).apply(validate_address)]

    variable_names = set()
    variables = []

    for _, row in df.iterrows():
        name = str(row.get(col_map['name'], DEFAULT_VALUES['name']))
        name = clean_variable_name(name)
        original_name = name
        counter = 1
        while name in variable_names:
            name = f"{original_name}_{counter}"
            counter += 1
        variable_names.add(name)

        address = int(row.get(col_map['address'], DEFAULT_VALUES['address']))
        scale = row.get(col_map.get('scale', ''), '1')
        unit = row.get(col_map.get('unit', ''), '')
        offset = row.get(col_map.get('offset', ''), '0')

        if pd.isna(scale):
            scale = '1'
        if pd.isna(unit):
            unit = ''
        if pd.isna(offset):
            offset = '0'

        variables.append({
            'name': name,
            'address': address,
            'scale': str(scale).strip(),
            'unit': str(unit).strip(),
            'offset': str(offset).strip()
        })

    max_name_length = max(len(v['name']) for v in variables)

    struct_lines = []
    for var in variables:
        spaces_needed = (max_name_length - len(var['name']) + 4)
        comment = f"(* ModbusAddress: {var['address']};  Scale: {var['scale']};  Unit: {var['unit']};  Offset: {var['offset']}; *)"
        struct_lines.append(f"    {var['name']}{' ' * spaces_needed}: UINT; {comment}")

    dut_name = clean_variable_name(f"{excel_name}_{sheet_name}")
    struct_text = f"TYPE {dut_name} :\nSTRUCT\n" + "\n".join(struct_lines) + "\nEND_STRUCT\nEND_TYPE"

    output_base = f"{excel_name}_{sheet_name}_modbus_struct"

    with open(output_base + ".txt", "w", encoding="utf-8") as f_txt:
        f_txt.write(struct_text)

    guid = str(uuid.uuid4())
    xml_content = f"""<?xml version=\"1.0\" encoding=\"utf-8\"?>
<TcPlcObject Version=\"1.1.0.1\">
  <DUT Name=\"{dut_name}\" Id=\"{{{guid}}}\">
    <Declaration><![CDATA[
{struct_text}
]]></Declaration>
  </DUT>
</TcPlcObject>
"""

    with open(output_base + ".TcDUT", "w", encoding="utf-8") as f_tcdut:
        f_tcdut.write(xml_content)

def process_excel(file_path: str):
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names

        def generate_from_sheet(sheet_index: int):
            try:
                df_raw = pd.read_excel(file_path, sheet_name=sheet_names[sheet_index], header=None)
                excel_name = os.path.splitext(os.path.basename(file_path))[0]
                safe_sheet_name = re.sub(r'[^\w]', '', sheet_names[sheet_index])
                process_sheet(df_raw, safe_sheet_name, excel_name)
                messagebox.showinfo("Éxito", f"Archivos generados para la hoja: {sheet_names[sheet_index]}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar la hoja:\n{str(e)}")

        sheet_window = tk.Toplevel(root)
        sheet_window.title("Selecciona una hoja")
        sheet_window.geometry("400x300")

        tk.Label(sheet_window, text="Selecciona la hoja de Excel:", font=("Arial", 10)).pack(pady=10)

        for idx, name in enumerate(sheet_names):
            tk.Button(
                sheet_window,
                text=name,
                command=lambda i=idx: [generate_from_sheet(i), sheet_window.destroy()],
                padx=10,
                pady=5
            ).pack(fill='x', padx=20, pady=2)

    except Exception as e:
        messagebox.showerror("Error", f"Error procesando archivo:\n{str(e)}")

def seleccionar_archivo():
    file_path = filedialog.askopenfilename(
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
    )
    if file_path:
        process_excel(file_path)

def create_main_window():
    global root
    root = tk.Tk()
    root.title("Generador de Estructuras Modbus")
    root.geometry("500x200")

    tk.Label(root, text="Generador de Estructuras Modbus", font=("Arial", 14, "bold")).pack(pady=20)

    tk.Button(root, text="Seleccionar archivo Excel", command=seleccionar_archivo, font=("Arial", 10), padx=20, pady=5).pack(pady=10)
    tk.Button(root, text="Salir", command=root.quit, font=("Arial", 10), padx=20, pady=5).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    try:
        load_configuration()
        create_main_window()
    except Exception as e:
        messagebox.showerror("Error Inicial", f"No se pudo iniciar la aplicación:\n{str(e)}")