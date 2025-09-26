import pyodbc
import logging
import os

# ==========================
# Logging
# ==========================
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "relaciones_tablas.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ==========================
# Conexión
# ==========================
SERVER = "SQLREPLICA01.SOFTLANDCLOUD.CL"
DATABASE = "KROPSYS1"
USERNAME = "KROPSYS"
PASSWORD = "B*KGsV1eb)C(Z))q"

conn = None
try:
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD};"
        f"Encrypt=no;"
        f"TrustServerCertificate=yes;"
    )
    conn = pyodbc.connect(conn_str)
    logging.info("Conexión exitosa a SQL Server")
    print("[INFO] Conexión exitosa.")
except Exception as e:
    logging.error(f"Error conectando a SQL Server: {e}")
    print(f"[ERROR] No se pudo conectar: {e}")
    exit()

# ==========================
# Función para mostrar columnas
# ==========================
def mostrar_columnas_y_ejemplos(tabla, top=5):
    cursor = conn.cursor()
    try:
        cursor.execute(f"SELECT TOP {top} * FROM {tabla}")
        columnas = [column[0] for column in cursor.description]
        filas = cursor.fetchall()

        logging.info(f"Columnas de {tabla}: {columnas}")
        print(f"[INFO] Columnas de {tabla}: {columnas}")
        print(f"[INFO] Primeras {top} filas de {tabla}:")
        for fila in filas:
            print(fila)
            logging.info(f"{fila}")

        return columnas, filas
    except Exception as e:
        logging.error(f"Error leyendo tabla {tabla}: {e}")
        print(f"[ERROR] {e}")
        return [], []

# ==========================
# Revisar tablas
# ==========================
tablas = ["softland.iw_gsaen", "softland.iw_gmovi", "softland.iw_tprod"]
columnas_data = {}

for tabla in tablas:
    cols, filas = mostrar_columnas_y_ejemplos(tabla)
    columnas_data[tabla] = cols

# ==========================
# Sugerir posibles relaciones
# ==========================
# Comparar columnas comunes entre gsaen y gmovi
gsaen_cols = set(columnas_data["softland.iw_gsaen"])
gmovi_cols = set(columnas_data["softland.iw_gmovi"])
columnas_comunes = gsaen_cols.intersection(gmovi_cols)

logging.info(f"Columnas comunes entre iw_gsaen y iw_gmovi: {columnas_comunes}")
print(f"[INFO] Columnas comunes entre iw_gsaen y iw_gmovi: {columnas_comunes}")

# Revisar coincidencias para las primeras filas
print("[INFO] Posibles relaciones entre guías y movimientos:")
logging.info("Posibles relaciones entre guías y movimientos:")

# Tomar las primeras 10 filas de cada tabla para pruebas
gsaen_rows = mostrar_columnas_y_ejemplos("softland.iw_gsaen", top=10)[1]
gmovi_rows = mostrar_columnas_y_ejemplos("softland.iw_gmovi", top=10)[1]

for col in columnas_comunes:
    coincidencias = []
    for g_row in gsaen_rows:
        for m_row in gmovi_rows:
            try:
                if getattr(g_row, col) == getattr(m_row, col):
                    coincidencias.append(getattr(g_row, col))
            except Exception:
                continue
    if coincidencias:
        logging.info(f"Columna '{col}' tiene coincidencias: {set(coincidencias)}")
        print(f"[INFO] Columna '{col}' tiene coincidencias: {set(coincidencias)}")
    else:
        logging.info(f"Columna '{col}' no tiene coincidencias visibles en las primeras filas.")
        print(f"[INFO] Columna '{col}' no tiene coincidencias visibles en las primeras filas.")

# ==========================
# Cerrar conexión
# ==========================
conn.close()
logging.info("Conexión cerrada.")
print("[INFO] Conexión cerrada.")
