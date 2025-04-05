from pathlib import Path
import pandas as pd
import logging
from typing import List, Optional
from dataclasses import dataclass, field

from config import TWINCAT_TYPE_MAPPING
# Abstracción de procesamiento de archivos
class FileProcessor:
    """Interfaz para el procesamiento de diferentes tipos de archivos."""
    def process(self, file_path: Path, sheet_name: str):
        raise NotImplementedError("Esta función debe ser implementada por subclases.")

@dataclass
class ExcelProcessor(FileProcessor):
    """Procesador de archivos Excel que genera mapeos Modbus."""
    
    name_idx: int
    type_idx: int
    address_idx: int
    info_idx: Optional[int] = None
    fixed_comment: str = ''

    map_data: pd.DataFrame = field(default=None, init=False)

    def process(self, file_path: Path, sheet_name: str) -> List[str]:
        """Carga y procesa el archivo Excel."""
        logging.info("Iniciando el procesamiento del archivo Excel.")
        self._load_excel(file_path, sheet_name)
        self._process_data()
        return self._generate_modbus_mappings()

    def _load_excel(self, file_path: Path, sheet_name: str):
        """Carga el archivo Excel y renombra las columnas según los índices."""
        try:
            self.map_data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            self._rename_columns()
        except FileNotFoundError:
            logging.error("No se encontró el archivo especificado.")
            raise RuntimeError("No se encontró el archivo especificado.")
        except Exception as e:
            logging.error(f"Error al cargar el archivo Excel: {e}")
            raise RuntimeError(f"Error al cargar el archivo Excel: {e}")

    def _rename_columns(self):
        """Renombra las columnas según los índices proporcionados."""
        column_mapping = {
            self.name_idx: 'Name/Description',
            self.type_idx: 'Data Type',
            self.address_idx: 'ADDRESS (DEC)'
        }
        if self.info_idx is not None:
            column_mapping[self.info_idx] = 'Information'
        self.map_data.rename(columns=column_mapping, inplace=True)

        # Si no hay columna de información, agregar comentario fijo
        self.map_data['Information'] = self.map_data.get('Information', self.fixed_comment)
        self.map_data['Information'].fillna('', inplace=True)

    def _process_data(self):
        """Procesa los datos del archivo: filtra y convierte los tipos de datos."""
        self.map_data['Data Type'] = self.map_data['Data Type'].apply(self.convert_to_twincat_type)
        self.map_data = self.map_data[self.map_data['Data Type'] != 'UNKNOWN']

    def _generate_modbus_mappings(self) -> List[str]:
        """Genera los mapeos Modbus en base a los datos procesados."""
        return self.map_data.apply(self.generate_modbus_mapping, axis=1).tolist()
