# utils.py

import re
import pandas as pd
from config import TWINCAT_TYPE_MAPPING

class Utils:
    @staticmethod
    def format_modbus_address(address) -> str:
        """Formatea la dirección Modbus."""
        if pd.isna(address):
            return 'N/A'
        try:
            address = float(address)
            return str(int(address)) if address.is_integer() else str(address)
        except ValueError:
            return str(address)

    @staticmethod
    def to_camel_case(name: str) -> str:
        """Convierte una cadena a formato CamelCase."""
        if pd.isna(name):
            return ''
        name = re.sub(r'[^a-zA-Z0-9 ]', '', str(name))
        parts = name.split()
        camel_case_name = ''.join(word.capitalize() for word in parts)
        return f'Var{camel_case_name}' if camel_case_name and not camel_case_name[0].isalpha() else camel_case_name

    @staticmethod
    def convert_to_twincat_type(data_type: str) -> str:
        """Convierte un tipo de dato a formato TwinCAT."""
        if pd.isna(data_type) or not isinstance(data_type, str):
            return 'UNKNOWN'
        return TWINCAT_TYPE_MAPPING.get(data_type.strip().lower(), 'UNKNOWN')

    @staticmethod
    def generate_modbus_mapping(row) -> str:
        """Genera un mapeo Modbus para una fila dada."""
        nombre_variable = Utils.to_camel_case(row.get('Name/Description'))
        tipo_dato = row.get('Data Type', 'UNKNOWN')
        direccion_modbus = Utils.format_modbus_address(row.get('ADDRESS (DEC)'))
        informacion = row.get('Information', '')

        return (f"{nombre_variable:<30}: {tipo_dato:<10}; "
                f"(*ModbusAddress: {direccion_modbus:<10}; "
                f"DataType: {tipo_dato:<10}; Information: {informacion};*)")
