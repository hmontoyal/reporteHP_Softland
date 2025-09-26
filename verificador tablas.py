import pyodbc
import logging

# ==========================
# Configuración logging
# ==========================
logging.basicConfig(
    filename="verificador_tablas.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ==========================
# Configuración conexión
# ==========================
SERVER = "SQLREPLICA01.SOFTLANDCLOUD.CL"
DATABASE = "KROPSYS1"
USERNAME = "KROPSYS"
PASSWORD = "B*KGsV1eb)C(Z))q"  # contraseña con caracteres especiales

# ==========================
# Conectar y listar tablas
# ==========================
try:
    logging.info("Intentando conexión a SQL Server...")
    print("[INFO] Conectando a SQL Server...")

    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )

    logging.info("Conexión exitosa.")
    print("[INFO] Conexión exitosa.")

    cursor = conn.cursor()
    cursor.execute("SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_SCHEMA, TABLE_NAME")
    tablas = cursor.fetchall()

    print("\n=== Tablas disponibles ===")
    for schema, table in tablas:
        print(f"{schema}.{table}")
    logging.info(f"Se listaron {len(tablas)} tablas correctamente.")

    conn.close()
    logging.info("Conexión cerrada correctamente.")
    print("[INFO] Conexión cerrada correctamente.")

except pyodbc.InterfaceError as e:
    logging.error(f"Error de conexión: {e}")
    print(f"[ERROR] Error de conexión: {e}")

except pyodbc.Error as e:
    logging.error(f"Error ODBC: {e}")
    print(f"[ERROR] Error ODBC: {e}")

except Exception as e:
    logging.error(f"Error inesperado: {e}")
    print(f"[ERROR] Error inesperado: {e}")
