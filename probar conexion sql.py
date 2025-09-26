import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import logging

# ==========================
# Configuración logging
# ==========================
logging.basicConfig(
    filename="conexion_sql_diagnostico.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ==========================
# Función para probar conexión
# ==========================
def probar_conexion():
    servidor = servidor_entry.get().strip()
    bd = bd_entry.get().strip()
    usuario = usuario_entry.get().strip()
    contrasena = contrasena_entry.get().strip()
    timeout = timeout_spin.get()

    if not all([servidor, bd, usuario, contrasena]):
        messagebox.showwarning("Advertencia", "Todos los campos son obligatorios.")
        return

    try:
        logging.info(f"Intentando conexión a {servidor} / {bd} como {usuario}")
        print(f"[INFO] Intentando conexión a {servidor} / {bd} como {usuario}...")

        conn_str = (
            f"DRIVER={{ODBC Driver 18 for SQL Server}};"
            f"SERVER={servidor};"
            f"DATABASE={bd};"
            f"UID={usuario};"
            f"PWD={contrasena};"
            f"Encrypt=no;"
            f"TrustServerCertificate=yes;"
            f"Timeout={timeout};"
        )

        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute("SELECT GETDATE()")  # consulta simple para probar
        result = cursor.fetchone()[0]
        conn.close()

        logging.info(f"Conexión exitosa a {bd}. Servidor respondió: {result}")
        print(f"[INFO] Conexión exitosa. Servidor respondió: {result}")
        messagebox.showinfo("Éxito", f"Conexión exitosa a {bd}.\nServidor respondió: {result}")

    except pyodbc.InterfaceError as e:
        logging.error(f"Error de conexión: {e}")
        print(f"[ERROR] Error de conexión: {e}")
        messagebox.showerror("Error de conexión", str(e))

    except pyodbc.Error as e:
        logging.error(f"Error ODBC: {e}")
        print(f"[ERROR] Error ODBC: {e}")
        messagebox.showerror("Error ODBC", str(e))

    except Exception as e:
        logging.error(f"Error inesperado: {e}")
        print(f"[ERROR] Error inesperado: {e}")
        messagebox.showerror("Error inesperado", str(e))


# ==========================
# Interfaz Tkinter
# ==========================
root = tk.Tk()
root.title("Diagnóstico Conexión SQL Server")
root.geometry("500x400")

ttk.Label(root, text="Servidor SQL:").pack(pady=5)
servidor_entry = ttk.Entry(root, width=50)
servidor_entry.pack(pady=5)
servidor_entry.insert(0, "SQLREPLICA01.SOFTLANDCLOUD.CL")  # valor por defecto

ttk.Label(root, text="Base de datos:").pack(pady=5)
bd_entry = ttk.Entry(root, width=50)
bd_entry.pack(pady=5)
bd_entry.insert(0, "KROPSYS1")

ttk.Label(root, text="Usuario:").pack(pady=5)
usuario_entry = ttk.Entry(root, width=50)
usuario_entry.pack(pady=5)
usuario_entry.insert(0, "KROPSYS")

ttk.Label(root, text="Contraseña:").pack(pady=5)
contrasena_entry = ttk.Entry(root, width=50, show="*")
contrasena_entry.pack(pady=5)

ttk.Label(root, text="Timeout (segundos):").pack(pady=5)
timeout_spin = ttk.Spinbox(root, from_=5, to=120, increment=5, width=10)
timeout_spin.pack(pady=5)
timeout_spin.set(30)

ttk.Button(root, text="Probar Conexión", command=probar_conexion).pack(pady=20)

root.mainloop()
