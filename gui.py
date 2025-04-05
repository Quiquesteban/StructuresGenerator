import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from file_processor import ExcelProcessor
import logging

class ModbusGUI:
    """Clase encargada de manejar la interfaz gráfica."""
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Excel Modbus")
        self.excel_entry = None
        self.sheet_var = None
        self.sheet_dropdown = None
        self.column_entries = {}
        self.fixed_comment_entry = None
        self.create_gui()

    def create_gui(self):
        """Crea la interfaz gráfica."""
        tk.Label(self.root, text="Archivo Excel:").grid(row=0, column=0, sticky=tk.W)
        self.excel_entry = tk.Entry(self.root, width=40)
        self.excel_entry.grid(row=0, column=1)
        tk.Button(self.root, text="Seleccionar", command=self.select_excel_file).grid(row=0, column=2)

        tk.Label(self.root, text="Seleccionar hoja:").grid(row=1, column=0, sticky=tk.W)
        self.sheet_var = tk.StringVar(self.root)
        self.sheet_dropdown = tk.OptionMenu(self.root, self.sheet_var, '')
        self.sheet_dropdown.grid(row=1, column=1)

        columns = ['Name/Description', 'Data Type', 'ADDRESS (DEC)', 'Information']
        for i, col in enumerate(columns, start=2):
            optional = " (opcional)" if col == 'Information' else ""
            tk.Label(self.root, text=f"Columna para {col}{optional}:").grid(row=i, column=0, sticky=tk.W)
            entry = tk.Entry(self.root, width=5)
            entry.grid(row=i, column=1)
            self.column_entries[col] = entry

        tk.Label(self.root, text="Comentario fijo (opcional):").grid(row=i+1, column=0, sticky=tk.W)
        self.fixed_comment_entry = tk.Entry(self.root, width=40)
        self.fixed_comment_entry.grid(row=i+1, column=1)

        tk.Button(self.root, text="Procesar archivo", command=self.process_file).grid(row=i+2, column=0, columnspan=2)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)
            self.load_sheets(file_path)

    def load_sheets(self, file_path):
        try:
            excel_data = pd.ExcelFile(file_path)
            self.sheet_dropdown['menu'].delete(0, 'end')
            for sheet in excel_data.sheet_names:
                self.sheet_dropdown['menu'].add_command(label=sheet, command=tk._setit(self.sheet_var, sheet))
            self.sheet_var.set(excel_data.sheet_names[0])
        except Exception as e:
            logging.error(f"Error al cargar las hojas: {e}")
            messagebox.showerror("Error", f"Error al cargar las hojas: {e}")

    def process_file(self):
        file_path = self.excel_entry.get()
        sheet_name = self.sheet_var.get()
        try:
            name_idx = self.column_letter_to_index(self.column_entries['Name/Description'].get())
            type_idx = self.column_letter_to_index(self.column_entries['Data Type'].get())
            address_idx = self.column_letter_to_index(self.column_entries['ADDRESS (DEC)'].get())
            info_idx = self.column_entries['Information'].get()
            info_idx = self.column_letter_to_index(info_idx) if info_idx else None
            fixed_comment = self.fixed_comment_entry.get()

            processor = ExcelProcessor(name_idx, type_idx, address_idx, info_idx, fixed_comment)
            modbus_mappings = processor.process(Path(file_path), sheet_name)
            self.save_to_file(modbus_mappings)

        except ValueError:
            messagebox.showerror("Error", "Por favor, asegúrate de ingresar letras de columna válidas.")
        except Exception as e:
            logging.error(f"Error al procesar el archivo: {e}")
            messagebox.showerror("Error", f"Error al procesar el archivo: {e}")

    def save_to_file(self, modbus_mappings):
        output_file = 'modbus_mappings_cleaned_fixed.txt'
        try:
            with open(output_file, 'w', encoding='utf-8') as file:
                for mapping in modbus_mappings:
                    file.write(mapping + '\n')
            logging.info(f"Se han guardado los resultados en el archivo {output_file}.")
            messagebox.showinfo("Éxito", f"Se han escrito los resultados en el archivo {output_file}.")
        except Exception as e:
            logging.error(f"Error al guardar el archivo: {e}")
            messagebox.showerror("Error", f"Error al guardar el archivo: {e}")

    @staticmethod
    def column_letter_to_index(letter):
        return sum((ord(char) - ord('A') + 1) * 26**i for i, char in enumerate(letter.upper()[::-1])) - 1
