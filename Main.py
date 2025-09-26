import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pyodbc
import pandas as pd

# Configuración conexión BD
SERVER = "SERVIDOR_SQL"
DATABASE = "SOFTLAND_DB"
USERNAME = "USUARIO"
PASSWORD = "CONTRASEÑA"

def generar_reporte():
    try:
        fecha_inicio = start_date.get_date().strftime("%Y-%m-%d")
        fecha_fin = end_date.get_date().strftime("%Y-%m-%d")

        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}"
        )

        query = f"""
        SELECT 
            f.CodAux AS Cliente,
            f.NumDoc AS Documento,
            f.FechaEmision,
            f.MontoTotal
        FROM Softland.cwtFactura f
        WHERE f.FechaEmision BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
        """

        df = pd.read_sql(query, conn)
        conn.close()

        if df.empty:
            messagebox.showinfo("Reporte", "No se encontraron registros en ese rango de fechas.")
            return

        # Seleccionar dónde guardar
        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte como"
        )
        if archivo:
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Éxito", f"Reporte guardado en:\n{archivo}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# Interfaz
root = tk.Tk()
root.title("Generador de Reportes Softland")
root.geometry("400x200")

ttk.Label(root, text="Fecha inicio:").pack(pady=5)
start_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
start_date.pack(pady=5)

ttk.Label(root, text="Fecha fin:").pack(pady=5)
end_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
end_date.pack(pady=5)

ttk.Button(root, text="Generar Reporte", command=generar_reporte).pack(pady=20)

root.mainloop()
