import pandas as pd
import glob
import os

carpeta = input("Ingrese el nombre de la carpeta donde están los Excel: ").strip()

if not os.path.isdir(carpeta):
    raise ValueError(f"❌ La carpeta '{carpeta}' no existe.")

archivos = [
    f for f in (
        glob.glob(os.path.join(carpeta, "*.xlsx")) +
        glob.glob(os.path.join(carpeta, "*.xls"))
    )
    if not os.path.basename(f).startswith("~$")
]

print("\nArchivos encontrados:")
for a in archivos:
    print(" -", a)

if not archivos:
    raise ValueError("❌ No se encontraron archivos Excel válidos.")

tablas = []

for archivo in archivos:
    print(f"\nLeyendo: {archivo}")
    df = pd.read_excel(
        archivo,
        header=5,
        engine="openpyxl" if archivo.endswith("xlsx") else None
    )
    tablas.append(df)

tabla_final = pd.concat(tablas, ignore_index=True)
tabla_final.to_excel("referencias_unificadas.xlsx", index=False)

print("\n✔ Archivo 'Las tablas se unificarón' correctamente.")
