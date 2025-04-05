import tkinter as tk
from gui import ModbusGUI

if __name__ == "__main__":
    root = tk.Tk()
    gui = ModbusGUI(root)
    root.mainloop()
