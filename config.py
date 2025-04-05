import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Mapeo de tipos de datos TwinCAT
TWINCAT_TYPE_MAPPING = {
    'uint16': 'UINT',
    'uint32': 'UDINT',
    'int16': 'INT',
    'int32': 'DINT',
    'float': 'REAL',
    'double': 'LREAL',
    'string': 'STRING',
    'uint': 'UINT',
    '/': 'UNKNOWN'
}
