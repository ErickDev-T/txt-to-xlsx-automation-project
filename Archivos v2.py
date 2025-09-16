import pandas as pd
import glob
import os

ruta_actual = os.getcwd()
archivos = glob.glob(os.path.join(ruta_actual, "*.Z*"))  # leerá todos los archivos .Z* en el directorio actual

dataframes = []

for archivo in archivos:
    try:
        df = pd.read_csv(
            archivo,
            sep=">",
            header=None,
            names=["Empleado", "FechaHora"],
            dtype={"Empleado": "string"},
            engine="python"
        )

        # Limpieza inicial
        df["Empleado"] = df["Empleado"].str.strip()
        df["FechaHora"] = df["FechaHora"].str.strip()

        # Convertir a datetime
        df["FechaHora"] = pd.to_datetime(
            df["FechaHora"],
            format="%Y/%m/%d:%H:%M",
            errors="coerce"
        )

        # Quitar inválidos
        df = df.dropna(subset=["FechaHora"])
        if df.empty:
            print(f"{archivo} no tenía filas válidas.")
            continue

        # Quitar ceros a la izquierda en Empleado
        df["Empleado"] = df["Empleado"].astype(str).str.lstrip("0")

        # Quitar duplicados exactos dentro del archivo
        df = df.drop_duplicates(subset=["Empleado", "FechaHora"], keep="first")

        # Columnas adicionales
        df["Código de Trabajo"] = 1
        df["Motivo"] = "Manual"
        df["Comentarios"] = ""
        df["Fecha"] = df["FechaHora"].dt.date

        # Ordenar y renombrar
        df = df.sort_values(["Empleado", "FechaHora"])
        df = df.rename(columns={"FechaHora": "Fecha y Hora de Checada"})

        # --- MARCAR ENTRADA / SALIDA (vectorizado) ---
        g = df.groupby(["Empleado", "Fecha"])

        primera_ts = g["Fecha y Hora de Checada"].transform("min")
        ultima_ts  = g["Fecha y Hora de Checada"].transform("max")
        n_uniq     = g["Fecha y Hora de Checada"].transform("nunique")

        df["Estado de Asistencia"] = pd.NA
        df.loc[df["Fecha y Hora de Checada"].eq(primera_ts), "Estado de Asistencia"] = 0
        mask_salida = df["Fecha y Hora de Checada"].eq(ultima_ts) & (n_uniq >= 2)
        df.loc[mask_salida, "Estado de Asistencia"] = 1

        # Conservar solo entrada/salida
        df = df[df["Estado de Asistencia"].notna()].copy()

        # Orden final y selección de columnas
        df = df.sort_values(["Empleado", "Fecha y Hora de Checada"]).reset_index(drop=True)
        df = df[[
            "Empleado",
            "Fecha y Hora de Checada",
            "Estado de Asistencia",
            "Código de Trabajo",
            "Motivo",
            "Comentarios",
        ]]

        dataframes.append(df)

    except Exception as e:
        print(f"Error leyendo {archivo}: {e}")

# Unir todos los archivos procesados
if dataframes:
    df_final = pd.concat(dataframes, ignore_index=True)
    df_final = df_final.sort_values(["Empleado", "Fecha y Hora de Checada"]).reset_index(drop=True)

    # Quitar duplicados globales (por si se repiten entre archivos)
    df_final = df_final.drop_duplicates(
        subset=["Empleado", "Fecha y Hora de Checada", "Estado de Asistencia"],
        keep="first"
    )

    salida = os.path.join(ruta_actual, "todos_los_txt.xlsx")

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Checadas")

        worksheet = writer.sheets["Checadas"]
        # Ajustar ancho de columnas
        for i, col in enumerate(df_final.columns):
            max_len = max(df_final[col].astype(str).map(len).max(), len(col)) + 4
            worksheet.set_column(i, i, max_len)

    print(f"Archivos convertidos: {salida}")
else:
    print("Archivos inválidos.")
