import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pyodbc
import pandas as pd
import logging
import os
import threading

# ==========================
# Logging
# ==========================
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reporte_guias.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ==========================
# Conexión predefinida
# ==========================
SERVER = "SQLREPLICA01.SOFTLANDCLOUD.CL"
DATABASE = "KROPSYS1"
USERNAME = "KROPSYS"
PASSWORD = "B*KGsV1eb)C(Z))q"
TIMEOUT = 30

conn = None

# ==========================
# Funciones
# ==========================
def conectar_sql():
    global conn
    try:
        logging.info(f"Intentando conexión a {SERVER} / {DATABASE} como {USERNAME}")
        print(f"[INFO] Intentando conexión...")

        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};"
            f"Encrypt=no;TrustServerCertificate=yes;Timeout={TIMEOUT};"
        )

        conn = pyodbc.connect(conn_str)
        logging.info("Conexión exitosa")
        messagebox.showinfo("Conexión", "Conexión exitosa!")

    except Exception as e:
        logging.error(f"Error conectando a SQL Server: {e}")
        messagebox.showerror("Error conexión", f"No se pudo conectar: {e}")


def desconectar_sql():
    global conn
    if conn:
        try:
            conn.close()
            logging.info("Desconexión de SQL Server exitosa")
            messagebox.showinfo("Desconexión", "Desconexión exitosa.")
        except Exception as e:
            logging.error(f"Error al desconectar: {e}")
            messagebox.showerror("Error", f"No se pudo desconectar: {e}")
        finally:
            conn = None
    else:
        messagebox.showinfo("Desconexión", "No hay conexión activa.")


def generar_reporte_thread():
    threading.Thread(target=generar_reporte).start()


def generar_reporte():
    global conn
    if conn is None:
        messagebox.showwarning("Sin conexión", "Debe conectarse primero a la base de datos.")
        return

    try:
        fecha_inicio = start_date.get_date().strftime("%Y-%m-%d")
        fecha_fin = end_date.get_date().strftime("%Y-%m-%d")

        query = f"""
        SELECT 
            g.Folio AS GuiaDespacho,
            g.Fecha AS FechaGuia,
            p.CodProd AS CodigoProducto,
            p.DesProd AS Producto,
            m.CantDespachada AS Cantidad,
            p.CantPieza AS StockActual
        FROM softland.iw_gsaen g
        JOIN softland.iw_gmovi m ON g.NroInt = m.NroInt
        JOIN softland.iw_tprod p ON m.CodProd = p.CodProd
        WHERE g.Fecha BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
        ORDER BY g.Fecha, g.Folio
        """

        logging.info(f"Ejecutando consulta desde {fecha_inicio} hasta {fecha_fin}")
        print("[INFO] Ejecutando consulta...")

        progress['value'] = 0
        root.update_idletasks()

        df = pd.read_sql(query, conn)
        progress['value'] = 50
        root.update_idletasks()

        if df.empty:
            messagebox.showinfo("Reporte", "No se encontraron registros en ese rango de fechas.")
            logging.info("Consulta ejecutada pero sin resultados.")
            progress['value'] = 0
            return

        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar reporte como"
        )
        if not archivo:
            progress['value'] = 0
            return

        # Crear un Excel con dos hojas
        with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
            # Hoja principal con detalle de guías
            df.to_excel(writer, sheet_name="Detalle", index=False)

            # Hoja resumen por producto
            resumen = df.groupby(["CodigoProducto", "Producto"]).agg(
                CantidadDespachada=pd.NamedAgg(column="Cantidad", aggfunc="sum"),
                StockActual=pd.NamedAgg(column="StockActual", aggfunc="max")
            ).reset_index()

            resumen.to_excel(writer, sheet_name="Resumen", index=False)

        logging.info(f"Reporte guardado en {archivo}")
        messagebox.showinfo("Éxito", f"Reporte guardado en:\n{archivo}")
        progress['value'] = 100
        root.update_idletasks()
        progress['value'] = 0

    except Exception as e:
        logging.error(f"Error generando reporte: {e}")
        messagebox.showerror("Error", f"No se pudo generar el reporte: {e}")
        progress['value'] = 0
        print(f"[ERROR] {e}")

# ==========================
# Manejo de cierre seguro
# ==========================
def on_closing():
    if messagebox.askyesno("Salir", "¿Desea salir de la aplicación?"):
        if conn:
            try:
                conn.close()
                logging.info("Conexión cerrada automáticamente al salir.")
            except Exception as e:
                logging.error(f"Error al cerrar conexión al salir: {e}")
        root.destroy()

# ==========================
# Interfaz Tkinter
# ==========================
root = tk.Tk()
root.title("Generador Reporte Guías de Despacho")
root.geometry("650x500")

ttk.Button(root, text="Conectar SQL", command=conectar_sql).pack(pady=10)
ttk.Button(root, text="Desconectar SQL", command=desconectar_sql).pack(pady=5)

ttk.Label(root, text="Fecha inicio:").pack(pady=5)
start_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
start_date.pack(pady=5)

ttk.Label(root, text="Fecha fin:").pack(pady=5)
end_date = DateEntry(root, width=20, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
end_date.pack(pady=5)

ttk.Button(root, text="Generar Reporte", command=generar_reporte_thread).pack(pady=20)

progress = ttk.Progressbar(root, orient='horizontal', length=550, mode='determinate')
progress.pack(pady=10)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
