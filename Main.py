import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pyodbc
import pandas as pd
import logging
from datetime import datetime

# ==========================
# CONFIGURACIÓN DE LOGGING
# ==========================
logging.basicConfig(
    filename="reportes_guia.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ==========================
# CONFIGURACIÓN BD
# ==========================
SERVER = "SQLREPLICA01.SOFTLANDCLOUD.CL"
DATABASE = "KROPSYS1"
USERNAME = "KROPSYS"
PASSWORD = "B*KGsV1eb)C(Z))q"

# ==========================
# FUNCIÓN GENERAR REPORTE
# ==========================
def generar_reporte():
    try:
        fecha_inicio = start_date.get_date().strftime("%Y-%m-%d")
        fecha_fin = end_date.get_date().strftime("%Y-%m-%d")
        productos_input = productos_entry.get().strip()
        productos = [p.strip() for p in productos_input.split(",") if p.strip()]

        if not productos:
            messagebox.showwarning("Advertencia", "Debe ingresar al menos un producto.")
            return

        logging.info(f"Iniciando conexión a BD: {SERVER}, DB: {DATABASE}, Usuario: {USERNAME}")
        print(f"[INFO] Conectando a {SERVER} / {DATABASE} como {USERNAME}...")

        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 18 for SQL Server};"
            f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
            "Encrypt=no;TrustServerCertificate=yes;"
        )

        placeholders = ",".join("'" + p + "'" for p in productos)

        query = f"""
        SELECT 
            g.NumGuia AS GuiaDespacho,
            g.Fecha AS FechaGuia,
            p.CodProd AS CodigoProducto,
            p.Descripcion AS Producto,
            m.Cantidad AS Cantidad,
            g.CodAux AS Cliente
        FROM softland.iw_gsaen g
        JOIN softland.iw_gmovi m ON g.ID_Guia = m.ID_Guia
        JOIN softland.iw_tprod p ON m.CodProd = p.CodProd
        WHERE m.CodProd IN ({placeholders})
        AND g.Fecha BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
        ORDER BY g.Fecha, g.NumGuia
        """

        logging.info(f"Ejecutando consulta para productos: {productos_input}, fechas {fecha_inicio} - {fecha_fin}")
        print(f"[INFO] Ejecutando consulta...")

        df = pd.read_sql(query, conn)
        conn.close()
        logging.info("Conexión cerrada correctamente.")
        print("[INFO] Conexión cerrada correctamente.")

        if df.empty:
            messagebox.showinfo("Reporte", "No se encontraron registros con esos productos en el rango de fechas.")
            logging.warning("Consulta sin resultados.")
            print("[WARN] No se encontraron registros.")
            return

        # Guardar archivo Excel
        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte como"
        )
        if archivo:
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Éxito", f"Reporte guardado en:\n{archivo}")
            logging.info(f"Reporte generado y guardado en {archivo} con {len(df)} registros.")
            print(f"[INFO] Reporte guardado en {archivo} ({len(df)} registros).")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        logging.error(f"Error en generar_reporte: {str(e)}")
        print(f"[ERROR] {str(e)}")


# ==========================
# INTERFAZ
# ==========================
root = tk.Tk()
root.title("Generador de Reportes de Guías - Softland")
root.geometry("600x350")

ttk.Label(root, text="Fecha inicio:").pack(pady=5)
start_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
start_date.pack(pady=5)

ttk.Label(root, text="Fecha fin:").pack(pady=5)
end_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
end_date.pack(pady=5)

ttk.Label(root, text="Productos (separados por coma):").pack(pady=5)
productos_entry = ttk.Entry(root, width=50)
productos_entry.pack(pady=5)

ttk.Button(root, text="Generar Reporte", command=generar_reporte).pack(pady=20)

root.mainloop()
