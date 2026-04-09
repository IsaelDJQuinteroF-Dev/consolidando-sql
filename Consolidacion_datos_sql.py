'''
Reporte Consolidado de Ventas - es un script que toma los reportes mensuales de ventas,  desde 
una base de datos SQL, los consolida, limpia los datos y genera un resumen anual por producto. 
El resultado se guarda en un archivo Excel llamado 'Resumen_Consolidado_Ventas.xlsx'.
Elaborado por: Isael D'Jesús Quintero Fuentes.
'''
import pandas as pd
from sqlalchemy import create_engine
import sys

# 1. Función para conectar a SQL Server y extraer datos comerciales.
def conectar_y_extraer_ventas():
 
# 2. PARÁMETROS DE CONEXIÓN (Información interna de la empresa)
    config = {
        "driver": "ODBC Driver 17 for SQL Server",
        "server": "000.000.000.000",  # IP del servidor de la empresa
        "database": "DB_COMERCIAL", # Nombre de la base de datos
        "user": "usuario",        # Usuario con permisos de lectura
        "pass": "*******"         # Contraseña del usuario
    }

# 3. CONSTRUCCIÓN DE LA CADENA DE CONEXIÓN
    url_conexion = (
        f"mssql+pyodbc://{config['user']}:{config['pass']}@{config['server']}/"
        f"{config['database']}?driver={config['driver']}"
    )

    try:
        engine = create_engine(url_conexion)
# 4. EXTRACCIÓN MEDIANTE SQL
        engine = create_engine(url_conexion, connect_args={'timeout': 10})
        
        query = """
            SELECT Producto, Cantidad, Precio, Fecha_Venta 
            FROM Ventas_Historico 
            WHERE Fecha_Venta >= '2026-01-01'
        """
        
        print("Intentando conectar al servidor SQL...")
        df_ventas = pd.read_sql(query, engine)
        if df_ventas.empty:
            print("AVISO: Conexión establecida, pero no se encontraron registros para los criterios buscados.")
            return None
        print("Conexión exitosa. Datos extraídos correctamente.")
        
        return df_ventas

    except Exception as e:
        # CONTROL DE SEGURIDAD: Error de red, credenciales o servidor caído
        print("\n" + "!"*50)
        print("ERROR CRÍTICO: No se pudo establecer la conexión.")
        print(f"DETALLE TÉCNICO: {e}")
        print("!"*50)
        return None
    
# 5. Función para limpiar y dar formato a los datos extraídos.
def limpiar_datos_extraidos(df):

    if df is not None:
        print("Iniciando fase de normalización y limpieza...")
        
        df = df.drop_duplicates()
        
        df['Precio'] = pd.to_numeric(df['Precio'], errors='coerce')
        df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce')
        
        df = df.dropna(subset=['Precio', 'Cantidad'])
        
        df['Subtotal'] = df['Cantidad'] * df['Precio']
        
        print("Datos normalizados y listos para análisis.")
        return df
    return None

# 6. Desarrollo
if __name__ == "__main__":

    datos_crudos = conectar_y_extraer_ventas()
    
    if datos_crudos is not None:
        datos_limpios = limpiar_datos_extraidos(datos_crudos)
    
        if datos_limpios is not None:
            datos_limpios.to_excel("Reporte_SQL_Consolidado.xlsx", index=False)
            print("\nEl reporte ejecutivo ha sido generado desde la Base de Datos y guardado como 'Reporte_SQL_Consolidado.xlsx'.")
        else:
            print("\nEl proceso de limpieza resultó en un dataset vacío.")
    else:
        print("\nLa operación fue cancelada por falta de fuente de datos.") 