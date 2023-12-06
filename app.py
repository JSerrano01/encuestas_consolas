from flask import Flask, render_template, request, abort, Response
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pymysql
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import pandas as pd
import openpyxl
from decimal import Decimal


app = Flask(__name__)


# -------------------------------------------------------------------RUTAS DE VISTAS SIMPLES---------------------------------------------------------------------------------------------------------
# Ruta al Index
@app.route("/")
def index():
    
    try:
        connection = conectar_base_datos()
        
        cursor = connection.cursor()

        # Realizar la consulta con límite y desplazamiento para la paginación
        cursor.execute("SELECT FECHA FROM informe ORDER BY ID DESC LIMIT 1;")
        resultados = cursor.fetchone()

        # Asegúrate de que resultados no sea None antes de manipularlo
        if resultados:
            # Obtener el primer elemento de la tupla y eliminar paréntesis y comillas
            fecha = resultados[0][0:-8]
        else:
            fecha = None

        return render_template("index.html", resultados=fecha)
    
    except Exception as e:
        # Manejo de la excepción
        raise e
    finally:
        # Cerrar la conexión
        connection.close()
    


# Ruta a la vista generar todas las dependencias
@app.route("/generar_todas")
def generar_todas():
    return render_template("generar_todas.html")


# Ruta a la vista generar por dependencias
@app.route("/generar_por_dependencia")
def generar_por_dependencia():
    return render_template("generar_por_dependencia.html")

#-------------------------------------------------------------------------------------CONEXION BASE DE DATOS------------------------------------------------------------------------------------------

def conectar_base_datos():
    # Configurar la conexión a la base de datos
    db_host = "localhost"
    db_user = "root"
    db_password = ""
    db_name = "encuestas_consolas"

    # Intentar establecer la conexión
    try:
        connection = pymysql.connect(
            host=db_host, user=db_user, password=db_password, database=db_name
        )
        print("Conexión exitosa a la base de datos")
        return connection
    except pymysql.Error as e:
        print(f"Error al conectar a la base de datos: {e}")
        return None

# -------------------------------------------------------------------FUNCIONES A RUTAS-----------------------------------------------------------------------------------------------------------------


# -----------------------------------------------------------------------------FUNCION DE GENERAR Y MOSTRAR TODA LA BASE DE DATOS COMPLETA-------------------------------------------------------------
@app.route("/completa", methods=["GET", "POST"])
def mostrar_completa():
    try:
        # Obtener los valores de fecha_inicio y fecha_fin del formulario
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_fin = request.form.get("fecha_fin")

        # print(f"Fecha de Inicio: {fecha_inicio}")
        # print(f"Fecha de Fin: {fecha_fin}")

        # Obtener el número total de registros en la base de datos
        connection = conectar_base_datos()
        cursor = connection.cursor()
        
        # Consulta SQL para obtener el promedio_total
        cursor.execute(
            """
            SELECT ROUND(
                SUM(GREATEST(0,
                    COALESCE(
                        CASE AMABILIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE PUNTUALIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE EFECTIVIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    )
                )) / COUNT(*) / 3, 1) AS promedio_total
            FROM informe
        """
        )
        promedio_total = cursor.fetchone()[0]
        # print(promedio_total)

        # Consulta SQL para contar la cantidad de cada valor en cada columna
        cursor.execute(
            """
SELECT
    COUNT(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_amabilidad,
    COUNT(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_amabilidad,
    COUNT(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_amabilidad,
    COUNT(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_amabilidad,
    COUNT(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_amabilidad,

    COUNT(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_puntualidad,
    COUNT(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_puntualidad,
    COUNT(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_puntualidad,
    COUNT(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_puntualidad,
    COUNT(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_puntualidad,

    COUNT(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_efectividad,
    COUNT(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_efectividad,
    COUNT(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_efectividad,
    COUNT(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_efectividad,
    COUNT(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_efectividad,

    SUM(CASE WHEN AMABILIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_amabilidad,
    SUM(CASE WHEN PUNTUALIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_puntualidad,
    SUM(CASE WHEN EFECTIVIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_efectividad,

    SUM(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) AS total_excelente,

    SUM(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE 0 END) AS total_bueno,

    SUM(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE 0 END) AS total_regular,

    SUM(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) AS total_deficiente,

    SUM(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE 0 END) +
    SUM(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE 0 END) AS total_malo

FROM informe;

        """
        )
        cantidad_valores = cursor.fetchall()

        cantidad_valores = cantidad_valores[0]
        # print(cantidad_valores)

        connection.close()

        resultados = consultar_datos_completo()

        return render_template(
            "completa.html",
            resultados=resultados,
            promedio_total=promedio_total,
            cantidad_valores=cantidad_valores,
        )
    except Exception as e:
        # Manejo de la excepción, por ejemplo, imprimir el error o enviar un mensaje al usuario.
        return render_template("error.html", error=str(e))


def consultar_datos_completo():
    try:
        connection = conectar_base_datos()
        cursor = connection.cursor()

        # Realizar la consulta con límite y desplazamiento para la paginación
        cursor.execute(f"SELECT * FROM informe")
        resultados = cursor.fetchall()

        return resultados
    except Exception as e:
        # Manejo de la excepción
        raise e
    finally:
        # Cerrar la conexión
        connection.close()

# --------------------------------------------------------------------------------------GENERADOR DE TODAS LAS DEPENDENCIAS CON FECHAS----------------------------------------------------------------------------


@app.route("/completa_fechas", methods=["GET", "POST"])
def completa_fechas():
    try:
        # Obtener los valores de fecha_inicio y fecha_fin del formulario
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_fin = request.form.get("fecha_fin")

        # Convertir las fechas al formato deseado (de 'YYYY-MM-DD' a 'DD/MM/YYYY')
        fecha_inicio_obj = datetime.strptime(fecha_inicio, "%Y-%m-%d")
        fecha_fin_obj = datetime.strptime(fecha_fin, "%Y-%m-%d")

        fecha_inicio_formateada = fecha_inicio_obj.strftime("%d/%m/%Y")
        fecha_fin_formateada = fecha_fin_obj.strftime("%d/%m/%Y")

        # Imprimir las fechas formateadas para verificar
        # print(f'Fecha de inicio formateada: {fecha_inicio_formateada}')
        # print(f'Fecha de fin formateada: {fecha_fin_formateada}')

        # Obtener el número total de registros en la base de datos
        connection = conectar_base_datos()
        cursor = connection.cursor()
        # Consulta SQL para obtener el promedio_total con parámetros
        query = """
            SELECT ROUND(
                SUM(GREATEST(0,
                    COALESCE(
                        CASE AMABILIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE PUNTUALIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE EFECTIVIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    )
                )) / COUNT(*) / 3, 1) AS promedio_total
            FROM informe WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
            AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
        """

        cursor.execute(query, (fecha_inicio_formateada, fecha_fin_formateada))

        resultado = cursor.fetchone()
        promedio_total = resultado[0] if resultado else None
        # print(f'Promedio total: {promedio_total}')

        # Consulta SQL para obtener la cantidad de valores con parámetros
        query = """
            SELECT
                COUNT(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_amabilidad,

                COUNT(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_puntualidad,

                COUNT(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_efectividad,

                SUM(CASE WHEN AMABILIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_amabilidad,
                SUM(CASE WHEN PUNTUALIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_puntualidad,
                SUM(CASE WHEN EFECTIVIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_efectividad,

                SUM(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) AS total_excelente,

                SUM(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE 0 END) AS total_bueno,

                SUM(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE 0 END) AS total_regular,

                SUM(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) AS total_deficiente,

                SUM(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE 0 END) AS total_malo

            FROM informe WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
            AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
        """
        cursor.execute(query, (fecha_inicio_formateada, fecha_fin_formateada))
        cantidad_valores = cursor.fetchall()

        cantidad_valores = cantidad_valores[0]
        # print(cantidad_valores)

        connection.close()

        resultados = consultar_datos_completo_fecha(
            fecha_inicio_formateada, fecha_fin_formateada
        )

        return render_template(
            "completa.html",
            resultados=resultados,
            promedio_total=promedio_total,
            cantidad_valores=cantidad_valores,
        )
    except Exception as e:
        # Manejo de la excepción, por ejemplo, imprimir el error o enviar un mensaje al usuario.
        return render_template("error.html", error=str(e))


def consultar_datos_completo_fecha(fecha_inicio_formateada, fecha_fin_formateada):
    try:
        connection = conectar_base_datos()
        cursor = connection.cursor()

        # Utilizar parámetros en la consulta para evitar la inyección de SQL
        query = (
            "SELECT * FROM informe "
            "WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s') "
            "AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')"
        )
        cursor.execute(query, (fecha_inicio_formateada, fecha_fin_formateada))

        resultados = cursor.fetchall()
        # print(resultados)

        return resultados
    except Exception as e:
        # Manejo de la excepción
        raise e
    finally:
        # Cerrar la conexión
        connection.close()


# --------------------------------------------------------------------------------------GENERADOR POR DEPENDENCIAS CON FECHAS----------------------------------------------------------------------------


@app.route("/por_dependencia_fechas", methods=["GET", "POST"])
def por_dependencia_fechas():
    try:
        # Obtener los valores de fecha_inicio y fecha_fin del formulario
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_fin = request.form.get("fecha_fin")
        dependencia = request.form.get("dependencia")

        # Convertir las fechas al formato deseado (de 'YYYY-MM-DD' a 'DD/MM/YYYY')
        fecha_inicio_obj = datetime.strptime(fecha_inicio, "%Y-%m-%d")
        fecha_fin_obj = datetime.strptime(fecha_fin, "%Y-%m-%d")

        fecha_inicio_formateada = fecha_inicio_obj.strftime("%d/%m/%Y")
        fecha_fin_formateada = fecha_fin_obj.strftime("%d/%m/%Y")

        # Imprimir las fechas formateadas para verificar
        # print(f'Fecha de inicio formateada: {fecha_inicio_formateada}')
        # print(f'Fecha de fin formateada: {fecha_fin_formateada}')

        # Obtener el número total de registros en la base de datos
        connection = conectar_base_datos()
        cursor = connection.cursor()
        # Consulta SQL para obtener el promedio_total con parámetros
        query = """
            SELECT ROUND(
                SUM(GREATEST(0,
                    COALESCE(
                        CASE AMABILIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE PUNTUALIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    ) +
                    COALESCE(
                        CASE EFECTIVIDAD
                            WHEN 'EXCELENTE' THEN 5
                            WHEN 'BUENO' THEN 4
                            WHEN 'REGULAR' THEN 3
                            WHEN 'DEFICIENTE' THEN 2
                            WHEN 'MALO' THEN 1
                            ELSE 0
                        END, 0
                    )
                )) / COUNT(*) / 3, 1) AS promedio_total
            FROM informe WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
            AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')AND DEPENDENCIA = %s
        """

        cursor.execute(
            query, (fecha_inicio_formateada, fecha_fin_formateada, dependencia)
        )

        resultado = cursor.fetchone()
        promedio_total = resultado[0] if resultado else None
        # print(f'Promedio total: {promedio_total}')

        # Consulta SQL para obtener la cantidad de valores con parámetros
        query = """
            SELECT
                COUNT(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_amabilidad,
                COUNT(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_amabilidad,

                COUNT(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_puntualidad,
                COUNT(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_puntualidad,

                COUNT(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE NULL END) AS excelente_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE NULL END) AS bueno_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE NULL END) AS regular_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE NULL END) AS deficiente_efectividad,
                COUNT(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE NULL END) AS malo_efectividad,

                SUM(CASE WHEN AMABILIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_amabilidad,
                SUM(CASE WHEN PUNTUALIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_puntualidad,
                SUM(CASE WHEN EFECTIVIDAD IN ('EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE', 'MALO') THEN 1 ELSE 0 END) AS suma_efectividad,

                SUM(CASE WHEN AMABILIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'EXCELENTE' THEN 1 ELSE 0 END) AS total_excelente,

                SUM(CASE WHEN AMABILIDAD = 'BUENO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'BUENO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'BUENO' THEN 1 ELSE 0 END) AS total_bueno,

                SUM(CASE WHEN AMABILIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'REGULAR' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'REGULAR' THEN 1 ELSE 0 END) AS total_regular,

                SUM(CASE WHEN AMABILIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'DEFICIENTE' THEN 1 ELSE 0 END) AS total_deficiente,

                SUM(CASE WHEN AMABILIDAD = 'MALO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN PUNTUALIDAD = 'MALO' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN EFECTIVIDAD = 'MALO' THEN 1 ELSE 0 END) AS total_malo

            FROM informe WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s')
            AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s') AND DEPENDENCIA = %s;
        """
        cursor.execute(
            query, (fecha_inicio_formateada, fecha_fin_formateada, dependencia)
        )
        cantidad_valores = cursor.fetchall()

        cantidad_valores = cantidad_valores[0]
        # print(cantidad_valores)

        connection.close()

        resultados = consultar_datos_por_dependencia_fecha(
            fecha_inicio_formateada, fecha_fin_formateada, dependencia
        )

        return render_template(
            "completa.html",
            resultados=resultados,
            promedio_total=promedio_total,
            cantidad_valores=cantidad_valores,
        )
    except Exception as e:
        # Manejo de la excepción, por ejemplo, imprimir el error o enviar un mensaje al usuario.
        return render_template("error.html", error=str(e))


def consultar_datos_por_dependencia_fecha(
    fecha_inicio_formateada, fecha_fin_formateada, dependencia
):
    try:
        connection = conectar_base_datos()
        cursor = connection.cursor()

        # Utilizar parámetros en la consulta para evitar la inyección de SQL
        query = (
            "SELECT * FROM informe "
            "WHERE STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') >= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s') "
            "AND STR_TO_DATE(FECHA, '%%d/%%m/%%Y %%H:%%i:%%s') <= STR_TO_DATE(%s, '%%d/%%m/%%Y %%H:%%i:%%s') AND DEPENDENCIA = %s"
        )
        cursor.execute(
            query, (fecha_inicio_formateada, fecha_fin_formateada, dependencia)
        )

        resultados = cursor.fetchall()
        # print(resultados)

        return resultados
    except Exception as e:
        # Manejo de la excepción
        raise e
    finally:
        # Cerrar la conexión
        connection.close()


# -----------------------------------------------------------------------------FUNCION DE ACTUALIZAR LA BASE DE DATOS CON LA NUBE DE GOOGLE SHEET---------------------------------------------------------------------------------------


@app.route("/actualizar", methods=["POST"])
def actualizar():
    try:
        connection = conectar_base_datos()
        cursor = connection.cursor()

        # Cargar credenciales del archivo JSON descargado
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            "JSON/api-encuesta-conosolas-0e96bc3cd058.json", scope
        )
        client = gspread.authorize(credentials)

        # Abrir la hoja de cálculo por su URL
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1Nbw91Az8ffithGzooS8tZYcffWR57C_7PbL7_fmO6OM/edit"
        spreadsheet = client.open_by_url(spreadsheet_url)

        # Seleccionar la hoja por nombre
        worksheet = spreadsheet.get_worksheet(0)

        # Obtener todos los valores en forma de lista de listas
        data = worksheet.get_all_values()

        # Imprimir los datos
        # print(data)

        query = f"TRUNCATE `encuestas_consolas`.`informe`"
        cursor.execute(query)

        # Iterar sobre los datos y almacenarlos en la base de datos
        for row in data[1:]:
            # Suponiendo que tu hoja de cálculo tiene columnas: FECHA, AMABIILIDAD, EFECTIVIDAD, DEPENDENCIA, ...
            query = "INSERT INTO informe (FECHA, AMABILIDAD, PUNTUALIDAD, EFECTIVIDAD, DEPENDENCIA) VALUES (%s, %s, %s, %s, %s)"
            cursor.execute(query, (row[0], row[1], row[2], row[3], row[4]))

        # Commit y cerrar la conexión
        connection.commit()

        return render_template("actualizado.html")

    except Exception as e:
        # Manejo de la excepción, por ejemplo, imprimir el error o enviar un mensaje al usuario.
        return render_template("error.html", error=str(e))
    finally:
        # Cerrar la conexión en el bloque finally para asegurarse de que siempre se cierre, incluso si hay una excepción.
        connection.close()
        

# ------------------------------------------------------------------------------------------RUTA PARA GENERAR EXCEL---------------------------------------------------------------------------------------------------


@app.route("/completa/export_excel", methods=["POST"])
def exportar_excel():
    try:
        # connection = conectar_base_datos()
        # cursor = connection.cursor()

        # Obtener los datos del formulario
        resultados = request.form.get("resultados")
        promedio_total = request.form.get("promedio_total")
        cantidad_valores = request.form.get("cantidad_valores")

        # print(resultados)
        # print(promedio_total)
        # print(cantidad_valores)

        # Asegurarse de que los valores no sean nulos o vacíos antes de usarlos
        if resultados is None or promedio_total is None or cantidad_valores is None:
            return "Alguno de los parámetros es nulo o vacío", 400

        # Decodificar los valores de cadena y manejar los Decimal correctamente
        resultados = eval(resultados)
        promedio_total = float(promedio_total)
        cantidad_valores = eval(cantidad_valores)

        # Limpiar los valores y convertirlos a números de punto flotante
        cantidad_valores = [
            float(str(val).replace(",", "").strip())
            if not isinstance(val, Decimal)
            else float(val)
            for val in cantidad_valores
        ]

        denominador = float(
            cantidad_valores[15]
            + cantidad_valores[16]
            + cantidad_valores[17]
            + cantidad_valores[22]
        )

        data = {
            "Categoría": [
                "AMABILIDAD",
                "PUNTUALIDAD",
                "EFECTIVIDAD",
                "TOTAL",
                "PROMEDIOS",
            ],
            "Excelente": [
                float(cantidad_valores[0]),
                float(cantidad_valores[5]),
                float(cantidad_valores[10]),
                float(cantidad_valores[18]),
                round((float(cantidad_valores[18]) / denominador) * 100, 2),
            ],
            "Bueno": [
                float(cantidad_valores[1]),
                float(cantidad_valores[6]),
                float(cantidad_valores[11]),
                float(cantidad_valores[19]),
                round((float(cantidad_valores[19]) / denominador) * 100, 2),
            ],
            "Regular": [
                float(cantidad_valores[2]),
                float(cantidad_valores[7]),
                float(cantidad_valores[12]),
                float(cantidad_valores[20]),
                round((float(cantidad_valores[20]) / denominador) * 100, 2),
            ],
            "Deficiente": [
                float(cantidad_valores[3]),
                float(cantidad_valores[8]),
                float(cantidad_valores[13]),
                float(cantidad_valores[21]),
                round((float(cantidad_valores[21]) / denominador) * 100, 2),
            ],
            "Malo": [
                float(cantidad_valores[4]),
                float(cantidad_valores[9]),
                float(cantidad_valores[14]),
                float(cantidad_valores[22]),
                round((float(cantidad_valores[22]) / denominador) * 100, 2),
            ],
            "TOTALES": [
                float(cantidad_valores[15]),
                float(cantidad_valores[16]),
                float(cantidad_valores[17]),
                float(cantidad_valores[15])
                + float(cantidad_valores[16])
                + float(cantidad_valores[17]),
                (denominador / denominador) * 100,
            ],
        }

        # Crea un DataFrame de pandas
        df = pd.DataFrame(data)


        # Convert the data to a pandas DataFrame
        df_additional = pd.DataFrame(
            resultados,
            columns=[
                "ID",
                "FECHA",
                "AMABILIDAD",
                "PUNTUALIDAD",
                "EFECTIVIDAD",
                "DEPENDENCIA",
            ]
        )
        
        # Convert the data to a pandas DataFrame
        df_additional_promedio = pd.DataFrame([promedio_total], columns=["NOTA FINAL"])

        # Create an in-memory Excel writer
        excel_data_additional = BytesIO()

        # Write the DataFrame to the Excel writer
        with pd.ExcelWriter(excel_data_additional, engine="openpyxl") as writer:
            df_additional.to_excel(writer, index=False, sheet_name="Sheet1")

        # Combine the two Excel writers
        excel_data_combined = BytesIO()
        with pd.ExcelWriter(excel_data_combined, engine="openpyxl") as writer:
            df_additional.to_excel(writer, index=False, sheet_name="REGISTROS")
            df.to_excel(writer, index=False, sheet_name="RESULTADOS")
            df_additional_promedio.to_excel(writer, index=False, sheet_name="NOTA FINAL")

        # Retorna el archivo Excel combinado como respuesta
        excel_data_combined.seek(0)
        response = Response(
            excel_data_combined,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers[
            "Content-Disposition"
        ] = "attachment; filename=Registro General.xlsx"

        return response

    except Exception as e:
        # Manejo de la excepción, por ejemplo, imprimir el error o enviar un mensaje al usuario.
        return render_template("error.html", error=str(e))


# ----------------------------------------------------------------------------------------PAGINA DE ERROR-------------------------------------------------------------------------------------------


@app.errorhandler(404)
def page_not_found(error):
    return render_template("404.html"), 404


if __name__ == "__main__":
    app.run(debug=True)
