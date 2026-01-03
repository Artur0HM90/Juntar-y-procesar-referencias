# import pandas as pd
# import os

# # Pedir nombre del archivo Excel
# archivo_entrada = input("Ingrese el nombre del archivo Excel: ").strip()

# if not os.path.isfile(archivo_entrada):
#     raise ValueError(f"❌ El archivo '{archivo_entrada}' no existe.")

# # Archivo de salida
# archivo_salida = "solo_referencias.xlsx"

# # Leer Excel
# df = pd.read_excel(archivo_entrada, engine="openpyxl")

# # Verificar que exista la columna TIPO
# if "TIPO" not in df.columns:
#     raise ValueError("❌ La columna 'TIPO' no existe en el archivo.")

# # Filtrar solo REFERENCIA
# df_referencias = df[
#     df["TIPO"].astype(str).str.strip().str.upper() == "REFERENCIA"
# ]

# # Guardar resultado
# df_referencias.to_excel(archivo_salida, index=False, engine="openpyxl")

# print("\n✔ Archivo creado:", archivo_salida)
# print("✔ Total de REFERENCIAS:", len(df_referencias))

import pandas as pd
import os

# Pedir nombre del archivo Excel
archivo_entrada = input("Ingrese el nombre del archivo Excel: ").strip()

if not os.path.isfile(archivo_entrada):
    raise ValueError(f"❌ El archivo '{archivo_entrada}' no existe.")

# Archivo de salida
archivo_salida = "solo_referenciassssssss.xlsx"

# Leer Excel
df = pd.read_excel(archivo_entrada, engine="openpyxl")

# Verificar columnas necesarias
if "TIPO" not in df.columns:
    raise ValueError("❌ La columna 'TIPO' no existe en el archivo.")

if "TIPO DIAG." not in df.columns:
    raise ValueError("❌ La columna 'TIPO DIAG.' no existe en el archivo.")

# -------------------------------
# 1️⃣ FILTRO: solo REFERENCIA
# -------------------------------
df_filtrado = df[
    df["TIPO"].astype(str).str.strip().str.upper() == "REFERENCIA"
].copy()   # ← AQUÍ ESTÁ LA SOLUCIÓN

# -------------------------------
# 2️⃣ FILTRO: conservar D y P (eliminar R)
# -------------------------------
df_filtrado.loc[:, "TIPO DIAG."] = (
    df_filtrado["TIPO DIAG."].astype(str).str.strip().str.upper()
)

df_filtrado = df_filtrado[
    df_filtrado["TIPO DIAG."].isin(["D", "P"])
]

# Guardar resultado
df_filtrado.to_excel(archivo_salida, index=False, engine="openpyxl")

print("\n✔ Archivo creado:", archivo_salida)
print("✔ Total de registros:", len(df_filtrado))

