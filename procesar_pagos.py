import pandas as pd
import os
from datetime import datetime

# Diccionario de meses para conversión
MESES = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL",
    "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO",
    "09": "SEPTIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE"
}

# Función para formatear el concepto
def formatear_concepto(base, mes, año):
    if base.upper() == "LUZ":
        return f"LUZ {MESES[mes].lower()}{año}"
    elif base.upper() == "ALQ":
        return f"ALQ {MESES[mes]}{año}"
    return base

# Cargar datos desde el archivo de Excel
def cargar_datos_excel(ruta_archivo, nombre_hoja):
    df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, engine="openpyxl")
    df.columns = df.columns.str.strip().str.upper()
    return df

# Detectar nuevas transacciones y asignar conceptos
def procesar_pagos(df_pagos, df_listainq):
    df_pagos["FECHA"] = pd.to_datetime(df_pagos["FECHA"], errors='coerce').dt.date
    df_pagos = df_pagos.dropna(subset=["FECHA"])
    resultados = []
    
    for _, row in df_pagos.iterrows():
        remitente = row.get("REMITENTE", "")
        monto = row.get("MONTO", 0)
        concepto = ""
        observacion = ""

        filtro = df_listainq[df_listainq["REMITENTE"] == remitente]
        if not filtro.empty:
            precio = filtro.iloc[0]["PRECIO"]
            dpto = filtro.iloc[0]["DPTO"]
            edif = filtro.iloc[0]["EDIF"]
            año = filtro.iloc[0]["AÑO"]

            if monto == precio:
                concepto = "ALQ01"  # Asignar primera cuota
            elif monto > precio:
                if monto % precio == 0:
                    cuotas = monto // precio
                    concepto = " + ".join([f"ALQ{str(i).zfill(2)}" for i in range(1, cuotas + 1)])
                else:
                    concepto = "ALQ01"  # Si hay excedente, se asume pago adicional pero sin determinar luz
            else:
                if "LUZ" in str(row.get("OBSERVACION", "")).upper():
                    concepto = "LUZ01"
                else:
                    observacion = f"Pago parcial, falta {precio - monto}"

        else:
            observacion = "Revisar con inquilino"

        # Si el monto no es redondo, dejar en blanco el concepto
        if not (monto % 1000 == 0):
            concepto = ""

        resultados.append({
            "FECHA": row["FECHA"],
            "COMPROBANTE": row["COMPROBANTE"],
            "ED": edif if 'edif' in locals() else "",
            "COMERC": dpto if 'dpto' in locals() else "",
            "AÑO": año if 'año' in locals() else "",
            "CONCEPTO": concepto,
            "OBSERVACIONES": observacion,
            "REMITENTE": remitente,
            "MONTO": monto
        })
    
    return pd.DataFrame(resultados)

# Ejecutar el procesamiento
def main():
    ruta_excel_pagos = "ALQ_PAGOS.xlsx"
    ruta_excel_listainq = "LISTAINQ.xlsx"
    salida_excel = "PROCESADO_PAGOS.xlsx"

    df_pagos = cargar_datos_excel(ruta_excel_pagos, "Sheet1")
    df_listainq = cargar_datos_excel(ruta_excel_listainq, "Sheet1")
    df_resultado = procesar_pagos(df_pagos, df_listainq)

    df_resultado.to_excel(salida_excel, index=False, engine="openpyxl")
    print(f"Procesamiento completado. Archivo guardado en {salida_excel}")

if __name__ == "__main__":
    main()