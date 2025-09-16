import pandas as pd
import glob
import os

ruta_actual = os.getcwd()
archivos = glob.glob(os.path.join(ruta_actual, "*.Z*"))

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

        df["Empleado"] = df["Empleado"].str.strip().str.lstrip("0")
        df["FechaHora"] = pd.to_datetime(
            df["FechaHora"].str.strip(),
            format="%Y/%m/%d:%H:%M",
            errors="coerce"
        )

        df = df.dropna(subset=["FechaHora"])
        if df.empty:
            print(f"{archivo} no tenía filas válidas.")
            continue

        df = df.drop_duplicates(subset=["Empleado", "FechaHora"], keep="first")
        df["Código de Trabajo"] = 1
        df["Motivo"] = "Manual"
        df["Comentarios"] = ""
        df["Fecha"] = df["FechaHora"].dt.date
        df = df.sort_values(["Empleado", "FechaHora"])
        df = df.rename(columns={"FechaHora": "Fecha y Hora de Checada"})
        df["Estado de Asistencia"] = pd.NA

        g = df.groupby(["Empleado", "Fecha"])
        idx_ultima = g["Fecha y Hora de Checada"].idxmax()
        idx_primera = g["Fecha y Hora de Checada"].idxmin()

        df.loc[idx_ultima, "Estado de Asistencia"] = 1
        df.loc[idx_primera, "Estado de Asistencia"] = 0

        df = df[df["Estado de Asistencia"].notna()].copy()
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
        print(f"error leyendo {archivo}: {e}")

if dataframes:
    df_final = pd.concat(dataframes, ignore_index=True)
    df_final = df_final.sort_values(
        ["Empleado", "Fecha y Hora de Checada"]
    ).reset_index(drop=True)

    salida = os.path.join(ruta_actual, "todos_los_txt.csv")
    df_final.to_csv(salida, index=False, encoding="utf-8-sig")

    print(f"Archivos convertidos: {salida}")
else:
    print("Archivos inválidos.")
