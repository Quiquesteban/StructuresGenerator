
# 🧠 Generador de Estructuras Modbus AI (Offline)

Este proyecto es una aplicación en Python que interpreta mapas Modbus en Excel (.xlsx)
y genera automáticamente:

- Un `.txt` Structured Text para PLCs.
- Un `.xml` listo para importar en TwinCAT 3.

Funciona 100% offline.

---

## 📂 Estructura

```
modbus_ai_struct_generator/
├── modbus_ai_struct_generator.py
├── modbus_ai_config.json
├── README.md
```

---

## ⚙️ Instalación

1. Instala Python: https://www.python.org/
2. Terminal → Ejecuta:

```bash
pip install pandas openpyxl
```

---

## 🚀 Uso

1. Ejecuta:

```bash
python modbus_ai_struct_generator.py
```

2. Selecciona tu archivo Excel.
3. Selecciona la hoja correcta.
4. Se generarán automáticamente:

- `[NombreExcel]_[NombreHoja]_modbus_struct.txt`
- `[NombreExcel]_[NombreHoja]_modbus_struct.xml`

---

## 🛠️ Funciones

- Limpieza de caracteres no válidos.
- Corrección automática de duplicados.
- Detección de errores y advertencias.

---

## 👨‍💻 Autor

Proyecto desarrollado para entornos industriales exigentes.
