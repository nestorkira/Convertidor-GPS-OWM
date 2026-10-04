"""Convertidor GPS - App Streamlit
Replica y mejora el flujo del Excel con macros Convertidor GPS.xlsm.

Flujo original Excel:
  1. LIMPIAR (Modulo2.Limpiar): Columns("A:H").ClearContents en hoja CSV
  2. Usuario copia desde fila 8 del CSV origen (datos sin metadata ni header)
     y pega en columna A de hoja CSV
  3. CONVERTIR (Modulo1.Convertir): TextToColumns (Tab, ; , ,) + copia I->D
     donde I = IF(D-E=0,"",D-E)  => D_final = ELEV - LENGTH
  4. CONTINUIDAD (Modulo2.Continuidad): J=A2-A1, K=C2-C1, L=B2-B1, M=SQRT(K^2+L^2)
  5. Hoja GPS: CONCATENATE(CSV!A:H con ",") => archivo "Z Obj.csv" sin header.

Esta app automatiza todo: subir CSV -> detectar header -> convertir -> validar -> descargar.
"""
import csv
import io
import math
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Convertidor GPS", layout="wide")

COLUMNAS = ["ID", "EAST", "NORTH", "ELEV", "LENGTH", "AZIMUTH", "DIP", "MATERIAL"]


def detectar_encoding(raw: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            raw.decode(enc)
            return enc
        except Exception:
            continue
    return "utf-8"


def detectar_header_y_delimitador(texto: str):
    """Devuelve (fila_header_idx, delimitador, nombres_columnas)."""
    lineas = [l for l in texto.splitlines() if l.strip() != ""]
    candidatos = [",", ";", "\t", "|"]
    mejor = None  # (score, idx, delim, cols)
    for i, lin in enumerate(lineas[:30]):
        for d in candidatos:
            try:
                cols = next(csv.reader([lin], delimiter=d))
            except Exception:
                continue
            cols = [c.strip().strip('"').strip("'").upper() for c in cols]
            # score: cuantas de ID,EAST,NORTH aparecen
            score = sum(k in cols for k in ("ID", "EAST", "NORTH"))
            if score >= 2:
                if mejor is None or score > mejor[0]:
                    mejor = (score, i, d, cols)
        # tambien probar deteccion automatica con Sniffer en esa linea
    if mejor is None:
        # fallback: buscar linea que contenga ID y EAST como texto
        for i, lin in enumerate(lineas[:30]):
            up = lin.upper()
            if "ID" in up and "EAST" in up and "NORTH" in up:
                d = ";" if lin.count(";") > lin.count(",") else ","
                cols = [c.strip() for c in next(csv.reader([lin], delimiter=d))]
                return i, d, cols
        return None, ",", COLUMNAS
    _, idx, delim, cols = mejor
    return idx, delim, cols


def parsear_csv_origen(raw: bytes) -> tuple[pd.DataFrame, dict]:
    enc = detectar_encoding(raw)
    texto = raw.decode(enc, errors="replace")
    idx_header, delim, cols_raw = detectar_header_y_delimitador(texto)
    lineas = texto.splitlines()
    info = {"encoding": enc, "delimitador": repr(delim), "fila_header": idx_header}

    if idx_header is None:
        # Sin header: asumir que todo es datos con delimiter coma
        filas = [r for r in csv.reader(lineas, delimiter=",") if any(c.strip() for c in r)]
        df = pd.DataFrame(filas)
        # si tiene 8 columnas, asignar nombres
        if df.shape[1] >= 8:
            df = df.iloc[:, :8]
            df.columns = COLUMNAS
        info["metadata"] = []
        return df, info

    metadata = lineas[:idx_header]
    info["metadata"] = metadata
    cuerpo = "\n".join(lineas[idx_header:])
    # leer con pandas, engine python para tolerar filas irregulares
    df = pd.read_csv(
        io.StringIO(cuerpo),
        delimiter=delim,
        dtype=str,
        keep_default_na=False,
        engine="python",
    )
    # normalizar nombres: mapear a COLUMNAS por posicion si coinciden
    df.columns = [c.strip().strip('"').upper() for c in df.columns.astype(str)]
    # Si las columnas detectadas coinciden en orden, renombrar a estandar
    rename = {}
    for i, c in enumerate(df.columns):
        if i < len(COLUMNAS):
            # aceptar variantes: ESTE->EAST, etc.
            rename[c] = COLUMNAS[i] if c in (COLUMNAS[i],) or True else c
    # mapeo simple por posicion (el origen siempre trae el mismo orden)
    df.columns = COLUMNAS[: len(df.columns)]
    # quedarnos con 8 primeras, eliminar filas totalmente vacias
    df = df.iloc[:, :8]
    df.columns = COLUMNAS
    df = df[df.apply(lambda r: "".join(str(v).strip() for v in r) != "", axis=1)]
    df = df.reset_index(drop=True)
    return df, info


def es_numero(s) -> bool:
    try:
        float(str(s).strip())
        return str(s).strip() != ""
    except Exception:
        return False


def recortar_ceros(s: str) -> str:
    """Replica el formato General de Excel: 3853.500->3853.5, 0.000->0, -90.000->-90."""
    t = str(s).strip()
    if t == "" or t.upper() in ("NAN", "NONE"):
        return ""
    # quitar comillas
    t = t.strip('"').strip("'").strip()
    # intentar numerico
    try:
        f = float(t.replace(",", ""))
        if math.isnan(f) or math.isinf(f):
            return t
        # si el texto original tenia punto decimal, recortar ceros tipo string
        # para evitar artefactos binarios (8443458.07 -> 8443458.07 y no 8443458.0699999)
        base = t.replace(",", "")
        # manejar notacion cientifica: convertir via repr controlado
        if "e" in base.lower():
            if f.is_integer():
                return str(int(f))
            return ("%f" % f).rstrip("0").rstrip(".")
        if "." in base:
            rec = base.rstrip("0").rstrip(".")
            # validar que rec == f (por si base tenia espacios raros)
            try:
                if float(rec) == f:
                    # "-0" -> "0"
                    return "0" if rec in ("-0",) else rec
            except Exception:
                pass
            if f.is_integer():
                return str(int(f))
            # fallback redondeo a 6 decimales maximo
            return ("%f" % round(f, 6)).rstrip("0").rstrip(".")
        else:
            # entero
            if f.is_integer():
                return str(int(f))
            return str(f)
    except Exception:
        return t


def convertir(df: pd.DataFrame, modo_cota: str) -> pd.DataFrame:
    """Aplica la logica de la macro Convertir + I->D.
    modo_cota: 'fondo' => D = ELEV - LENGTH (macro original).
               'original' => D = ELEV sin tocar.
    """
    out = df.copy()
    for c in COLUMNAS:
        out[c] = out[c].astype(str).str.strip()
    # MATERIAL: mantener texto tal cual (Geometry, DBH, etc.)
    if modo_cota == "fondo (ELEV - LENGTH) - replica macro":
        nuevos_elev = []
        for _, r in out.iterrows():
            e, l = r["ELEV"], r["LENGTH"]
            if es_numero(e) and es_numero(l):
                fe, fl = float(str(e).replace(",", "")), float(str(l).replace(",", ""))
                val = fe - fl
                # Comportamiento macro original: si D-E==0 deja blanco.
                # Lo mejoramos: dejamos "0" salvo que el usuario pida replica exacta.
                nuevos_elev.append(val)
            elif es_numero(e):
                nuevos_elev.append(float(str(e).replace(",", "")))
            else:
                nuevos_elev.append(r["ELEV"])
        # formatear numeros
        out["ELEV"] = [
            recortar_ceros(("%f" % round(v, 3)) if isinstance(v, float) else v)
            for v in nuevos_elev
        ]
    else:
        out["ELEV"] = out["ELEV"].apply(recortar_ceros)

    # Normalizar resto de numericos igual que Excel General
    for c in ["ID", "EAST", "NORTH", "LENGTH", "AZIMUTH", "DIP"]:
        out[c] = out[c].apply(recortar_ceros)
    out["MATERIAL"] = df["MATERIAL"].astype(str).str.strip()
    return out


def tabla_continuidad(df: pd.DataFrame, umbral_salto: float = 10.0) -> pd.DataFrame:
    """Replica columnas J,K,L,M: dID, dNorte, dEste, distancia."""
    t = df.copy()
    for c in ["EAST", "NORTH"]:
        t[c + "_num"] = pd.to_numeric(t[c], errors="coerce")
    t["ID_num"] = pd.to_numeric(t["ID"], errors="coerce")
    t["dID"] = t["ID_num"].diff()
    t["dNORTE"] = t["NORTH_num"].diff()
    t["dESTE"] = t["EAST_num"].diff()
    t["DIST_m"] = (t["dNORTE"] ** 2 + t["dESTE"] ** 2) ** 0.5
    t["alerta"] = ""
    t.loc[(t["dID"] != 1) & t["dID"].notna(), "alerta"] += "ID no consecutivo; "
    t.loc[t["DIST_m"] > umbral_salto, "alerta"] += f"salto > {umbral_salto} m; "
    t.loc[(t["DIST_m"] == 0) & t["DIST_m"].notna(), "alerta"] += "duplicado XY; "
    return t


def a_csv_gps(df: pd.DataFrame) -> str:
    """Hoja GPS: concatena A:H con comas, sin header (formato Z Obj.csv)."""
    lineas = []
    for _, r in df.iterrows():
        lineas.append(",".join(str(r[c]) for c in COLUMNAS))
    return "\n".join(lineas) + ("\n" if lineas else "")


# ---------------- UI ----------------
st.title("Convertidor GPS — CSV Ayudas ➜ Z Obj")
st.caption("Sube el CSV y descarga el Z Obj.csv para GPS.")

with st.sidebar:
    st.header("Opciones")
    modo_cota = st.radio(
        "Columna ELEV de salida",
        ["fondo (ELEV - LENGTH) - replica macro", "original (mantener ELEV)"],
        index=0,
        help="La macro hace I=D-E y pega valores en D. Si LENGTH=0 el resultado es igual al ELEV.",
    )
    replica_exacta = st.checkbox(
        "Réplica exacta del bug macro (blanco si ELEV==LENGTH)",
        value=False,
        help="La macro deja vacío cuando D-E=0. Por defecto lo corregimos y dejamos 0.",
    )

archivo = st.file_uploader("Sube el CSV de ayudas", type=["csv", "txt"])

if archivo is None:
    st.info("Sube un archivo para empezar.")
    st.stop()

raw = archivo.getvalue()
nombre = archivo.name.rsplit(".", 1)[0] + " Z Obj.csv"

df_in, info = parsear_csv_origen(raw)

df_out = convertir(df_in, modo_cota)
if replica_exacta:
    # aplicar bug original: ELEV==LENGTH -> vacio
    mask = pd.to_numeric(df_in["ELEV"], errors="coerce") == pd.to_numeric(
        df_in["LENGTH"], errors="coerce"
    )
    df_out.loc[mask.fillna(False), "ELEV"] = ""

st.write(f"**Taladros detectados:** {len(df_out)}")
st.dataframe(df_out, use_container_width=True)

csv_gps = a_csv_gps(df_out)
st.download_button(
    "Descargar Z Obj.csv",
    data=csv_gps.encode("utf-8"),
    file_name=nombre,
    mime="text/csv",
)
