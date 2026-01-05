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

#========================================================================


# import pandas as pd
# import os

# # Pedir nombre del archivo Excel
# archivo_entrada = input("Ingrese el nombre del archivo Excel: ").strip()

# if not os.path.isfile(archivo_entrada):
#     raise ValueError(f"❌ El archivo '{archivo_entrada}' no existe.")

# # Archivo de salida
# archivo_salida = "solo_referenciassssssss.xlsx"

# # Leer Excel
# df = pd.read_excel(archivo_entrada, engine="openpyxl")

# # Verificar columnas necesarias
# if "TIPO" not in df.columns:
#     raise ValueError("❌ La columna 'TIPO' no existe en el archivo.")

# if "TIPO DIAG." not in df.columns:
#     raise ValueError("❌ La columna 'TIPO DIAG.' no existe en el archivo.")

# # -------------------------------
# # 1️⃣ FILTRO: solo REFERENCIA
# # -------------------------------
# df_filtrado = df[
#     df["TIPO"].astype(str).str.strip().str.upper() == "REFERENCIA"
# ].copy()   # ← AQUÍ ESTÁ LA SOLUCIÓN

# # -------------------------------
# # 2️⃣ FILTRO: conservar D y P (eliminar R)
# # -------------------------------
# df_filtrado.loc[:, "TIPO DIAG."] = (
#     df_filtrado["TIPO DIAG."].astype(str).str.strip().str.upper()
# )

# df_filtrado = df_filtrado[
#     df_filtrado["TIPO DIAG."].isin(["D", "P"])
# ]

# # Guardar resultado
# df_filtrado.to_excel(archivo_salida, index=False, engine="openpyxl")

# print("\n✔ Archivo creado:", archivo_salida)
# print("✔ Total de registros:", len(df_filtrado))




#========================================================================

# import pandas as pd
# import os

# # Pedir nombre del archivo Excel
# archivo_entrada = input("Ingrese el nombre del archivo Excel: ").strip()

# if not os.path.isfile(archivo_entrada):
#     raise ValueError(f"❌ El archivo '{archivo_entrada}' no existe.")

# archivo_salida = "solo_referencias_procesado.xlsx"

# # Leer Excel
# df = pd.read_excel(archivo_entrada, engine="openpyxl")

# # Verificar columnas necesarias
# columnas_necesarias = ["TIPO", "TIPO DIAG.", "NRO REFERENCIA", "NRO DOC", "SEXO", "EDAD", "TIPO EDAD", "COD. UNICO DESTINO", "UPS DESTINO"]
# for col in columnas_necesarias:
#     if col not in df.columns:
#         raise ValueError(f"❌ La columna '{col}' no existe.")

# # 1️⃣ Filtrar solo REFERENCIA
# df = df[df["TIPO"].astype(str).str.upper() == "REFERENCIA"].copy()

# # 2️⃣ Conservar solo D y P (eliminar R)
# df["TIPO DIAG."] = df["TIPO DIAG."].astype(str).str.upper().str.strip()
# df = df[df["TIPO DIAG."].isin(["D", "P"])]

# # 3️⃣ Columna B → dejar máximo 2 códigos iguales
# df = df.groupby("NRO REFERENCIA").head(2).reset_index(drop=True)

# # 4️⃣ Columna Y (DNI) tal cual
# dni_1 = df["NRO DOC"]

# # 5️⃣ Nueva columna con 1
# col_uno = pd.Series(1, index=df.index)

# # 6️⃣ Segunda columna DNI
# dni_2 = df["NRO DOC"]

# # 7️⃣ Columna AE → sexo
# sexo = df["SEXO"].astype(str).str.upper().map({
#     "MASCULINO": 1,
#     "FEMENINO": 2
# })

# # 8️⃣ Unir AC y AD con "-"
# ac_ad = df["EDAD"].astype(str) + "-" + df["TIPO EDAD"].astype(str)

# # 9️⃣ Nueva columna con 220000
# col_220000 = pd.Series(220000, index=df.index)

# # 🔟 Columna V → 4 ceros delante
# # v_formateado = df["COD. UNICO DESTINO"].astype(str).str.zfill(4 + df["COD. UNICO DESTINO"].astype(str).str.len())

# ups_destino = df["NRO DOC"]


# # 1️⃣1️⃣ Columna AO → solo números
# ao_numerico = df["UPS DESTINO"].astype(str).str.extract(r"(\d+)")[0]

# # 🔢 Crear DataFrame final ORDENADO
# df_final = pd.DataFrame({
#     "CODIGO": df["NRO REFERENCIA"],
#     "DNI": dni_1,
#     "UNO": col_uno,
#     "DNI_2": dni_2,
#     "SEXO": sexo,
#     "AC-AD": ac_ad,
#     "COD_FIJO": col_220000,
#     #"V": v_formateado,
#     "AO": ao_numerico
# })

# # Guardar Excel
# df_final.to_excel(archivo_salida, index=False, engine="openpyxl")

# print("\n✔ Archivo creado:", archivo_salida)
# print("✔ Total de registros:", len(df_final))
















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
].copy()

# -------------------------------
# 2️⃣ FILTRO: conservar D y P
# -------------------------------
df_filtrado.loc[:, "TIPO DIAG."] = (
    df_filtrado["TIPO DIAG."].astype(str).str.strip().str.upper()
)

df_filtrado = df_filtrado[
    df_filtrado["TIPO DIAG."].isin(["D", "P"])
]

# ==========================================================
# 🔽 TODO LO SIGUIENTE ES NUEVO (SIN TOCAR LO ANTERIOR)
# ==========================================================

# 3️⃣ Conservar máximo 2 registros por NRO REFERENCIA
# df_filtrado = (
#     df_filtrado
#     .groupby("NRO REFERENCIA")
#     .head(2)
#     .reset_index(drop=True)
# )

df_filtrado["__contador__"] = df_filtrado.groupby("NRO REFERENCIA").cumcount()

df_filtrado = df_filtrado[df_filtrado["__contador__"] < 2].copy()

df_filtrado.drop(columns="__contador__", inplace=True)


# 4️⃣ Crear columnas solicitadas en el orden requerido
df_salida = pd.DataFrame()

df_salida["NRO REFERENCIA"] = df_filtrado["NRO REFERENCIA"]
df_salida["NRO DOC"] = df_filtrado["NRO DOC"]

df_salida["Tipo documento de identidad del paciente"] = 1

df_salida["NRO DOC.1"] = df_filtrado["NRO DOC"]

# 5️⃣ SEXO → 1 masculino / 2 femenino
df_salida["SEXO"] = df_filtrado["SEXO"].str.upper().map({
    "MASCULINO": 1,
    "FEMENINO": 2
})

# 6️⃣ EDAD + TIPO EDAD
df_salida["Edad del paciente"] = (
    df_filtrado["EDAD"].astype(str) + "-" +
    df_filtrado["TIPO EDAD"].astype(str).str.lower()
)

# 7️⃣ Servicio Asistencial Origen Catalogo UPS
df_salida["Servicio Asistencial Origen Catalogo UPS"] = 220000

# 8️⃣ COD. UNICO DESTINO
df_salida["COD. UNICO DESTINO"] = df_filtrado["COD. UNICO DESTINO"]

# 9️⃣ UPS DESTINO
df_salida["UPS DESTINO"] = df_filtrado["UPS DESTINO"]


# 🔟 COD CIEX/CPT (VERSIÓN ROBUSTA)
def formatear_codigo(codigo):
    if pd.isna(codigo):
        return ""

    codigo = str(codigo).strip().replace(" ", "")

    if len(codigo) >= 4:
        return f"{codigo[:-1]}.{codigo[-1]}"

    return codigo

df_salida["COD CIEX/CPT"] = df_filtrado["COD CIEX/CPT"].apply(formatear_codigo)


# 1️⃣1️⃣ TIPO DIAG.
df_salida["TIPO DIAG."] = df_filtrado["TIPO DIAG."].map({
    "P": "01",
    "D": "02"
})

# 1️⃣2️⃣ Diagnóstico Secundario (vacío)
df_salida["Diagnóstico Secundario Motivo de la Referencia"] = ""

# 1️⃣3️⃣ Diagnóstico Secundario (vacío, segunda vez)
df_salida["Diagnóstico Secundario Motivo de la Referencia.1"] = ""

# 1️⃣4️⃣ FECHA. REGISTRO (solo fecha)
df_salida["FECHA. REGISTRO"] = pd.to_datetime(
    df_filtrado["FECHA. REGISTRO"], errors="coerce"
).dt.date

# 1️⃣5️⃣ FECHA ENVIO (solo fecha)
df_salida["FECHA ENVIO"] = pd.to_datetime(
    df_filtrado["FECHA ENVIO"], errors="coerce"
).dt.date

# Guardar resultado final
df_salida.to_excel(archivo_salida, index=False, engine="openpyxl")

print("\n✔ Archivo creado:", archivo_salida)
print("✔ Total de registros:", len(df_salida))














