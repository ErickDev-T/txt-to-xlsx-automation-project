import os
import re
import sys
from datetime import datetime, date, time
from collections import defaultdict

try:
    import pandas as pd
except ImportError:
    print("Necesitas instalar pandas y xlsxwriter")
    print("  pip install pandas xlsxwriter")
    sys.exit(1)

LINE_RE = re.compile(
    r"""^\s*
        (?P<emp>\d+)\s*>\s*
        (?P<y>\d{4})[/-](?P<m>\d{1,2})[/-](?P<d>\d{1,2})
        \s*[:\s]\s*
        (?P<H>\d{1,2}):(?P<M>\d{2})
        \s*$""",
    re.VERBOSE
)

def find_input_files():
    files = []
    for name in os.listdir("."):
        if not os.path.isfile(name):
            continue
        lower = name.lower()
        if lower.endswith(".txt"):
            files.append(name)
        else:
            # coincide con .Z20, .Z38, .Z70, etc
            if re.search(r"\.z[a-z0-9]+$", lower) and not lower.endswith(".zip"):
                files.append(name)
    return sorted(files)

def normalize_emp(emp_raw: str) -> str:
    emp = emp_raw.lstrip("0")
    return emp if emp else "0"

def parse_line(line: str):
    m = LINE_RE.match(line)
    if not m:
        return None
    emp = normalize_emp(m.group("emp"))
    y, mth, d = int(m.group("y")), int(m.group("m")), int(m.group("d"))
    H, M = int(m.group("H")), int(m.group("M"))

    try:
        dt = datetime(y, mth, d, H, M)
    except ValueError:
        return None

    return emp, dt.date(), dt.time()

def read_points(files):
    rows = []
    for fname in files:
        try:
            with open(fname, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    parsed = parse_line(line)
                    if parsed:
                        rows.append(parsed)
        except Exception as e:
            print(f"Advertencia: no se pudo leer '{fname}': {e}")
    return rows

def dedupe_exact(rows):
    seen = set()
    deduped = []
    for emp, fch, hr in rows:
        key = (emp, fch, hr)
        if key in seen:
            continue
        seen.add(key)
        deduped.append((emp, fch, hr))
    return deduped

def keep_first_and_last_per_day(rows):
    grouped = defaultdict(list)
    for emp, fch, hr in rows:
        grouped[(emp, fch)].append(hr)

    out = []
    for (emp, fch), horas in grouped.items():
        horas_uniq_sorted = sorted(set(horas))
        if not horas_uniq_sorted:
            continue
        if len(horas_uniq_sorted) == 1:
            out.append((emp, fch, horas_uniq_sorted[0], 0))
        else:
            out.append((emp, fch, horas_uniq_sorted[0], 0))
            out.append((emp, fch, horas_uniq_sorted[-1], 1))
    out.sort(key=lambda r: (r[0], r[1], r[2], r[3]))
    return out

def to_dataframe(rows_with_flag):
    data = []
    for emp, fch, hr, flag in rows_with_flag:
        data.append({
            "Empleado": emp,
            "Fecha": fch.strftime("%Y-%m-%d"),
            "Hora": f"{hr.hour:02d}:{hr.minute:02d}",
            "EntradaSalida": flag
        })
    df = pd.DataFrame(data, columns=["Empleado", "Fecha", "Hora", "EntradaSalida"])
    # orden final por empleado fecha hora y etrada salida
    df.sort_values(by=["Empleado", "Fecha", "Hora", "EntradaSalida"], inplace=True, kind="stable")
    df.reset_index(drop=True, inplace=True)
    return df

def main():
    files = find_input_files()
    if not files:
        print("No se encontraron archivos .Z* ni .txt en la carpeta actual.")
        return

    print(f"Archivos detectados ({len(files)}):")
    for f in files:
        print(f"  - {f}")

    raw_rows = read_points(files)
    print(f"\nLíneas válidas encontradas: {len(raw_rows)}")

    rows_nodup = dedupe_exact(raw_rows)
    print(f"Después de eliminar duplicados exactos: {len(rows_nodup)}")

    kept = keep_first_and_last_per_day(rows_nodup)
    print(f"Registros finales (primer y último ponche por empleado/día): {len(kept)}")

    df = to_dataframe(kept)

    out_name = "ponches.xlsx"
    try:
        with pd.ExcelWriter(out_name, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Ponches")
        print(f"\nArchivo generado: {os.path.abspath(out_name)}")
    except Exception:
        df.to_excel(out_name, index=False)
        print(f"\nArchivo generado (engine por defecto): {os.path.abspath(out_name)}")

if __name__ == "__main__":
    main()
