import pandas as pd
import requests
from io import BytesIO

# Lista de departamentos (nombres en orden del 1 al 24)
departamentos = [
    "Amazonas", "Ancash", "Apurímac", "Arequipa", "Ayacucho", "Cajamarca",
    "Cusco", "Huancavelica", "Huánuco", "Ica", "Junín", "La Libertad",
    "Lambayeque", "Lima", "Loreto", "Madre de Dios", "Moquegua", "Pasco",
    "Piura", "Puno", "San Martín", "Tacna", "Tumbes", "Ucayali"
]

# Lista de URLs directas a los archivos Excel
urls = []

for i in range(1, 24):  # dep01 a dep23 (con sufijo _15)
    num = str(i).zfill(2)
    url = f"https://m.inei.gob.pe/media/MenuRecursivo/indices_tematicos/pbi_dep{num}_15.xlsx"
    urls.append(url)

# Último archivo (dep24 con sufijo _16)
urls.append("https://m.inei.gob.pe/media/MenuRecursivo/indices_tematicos/pbi_dep24_16.xlsx")

# Verificación rápida
assert len(urls) == len(departamentos)

# Crear escritor de Excel
writer = pd.ExcelWriter("VAB_departamentos_cuadro2.xlsx", engine="openpyxl")

# Descargar y extraer hoja 'cuadro2' de cada archivo
for nombre, url in zip(departamentos, urls):
    try:
        print(f"Descargando: {nombre}")
        r = requests.get(url)
        xls = pd.ExcelFile(BytesIO(r.content))
        if "cuadro2" not in xls.sheet_names:
            print(f"No se encontró 'cuadro2' en {nombre}")
            continue
        df = pd.read_excel(xls, sheet_name="cuadro2")
        df.to_excel(writer, sheet_name=nombre[:31], index=False)
        print(f"Agregado: {nombre}")
    except Exception as e:
        print(f"Error con {nombre}: {e}")

# Guardar archivo final
writer.close()
print("Y si, se logró unir. Todo está en 'VAB_departamentos_cuadro2.xlsx'. Si no te gusta el nombre cambialo nomas..")

