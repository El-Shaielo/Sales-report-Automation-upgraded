from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import logging
import re
import tempfile
import zipfile

import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response

app = FastAPI(title="Generador de cortes de venta")

TEMPLATE_PATH = Path(__file__).resolve().with_name("Plantilla Cortes de ventas Oficial.docx")
REQUIRED_COLUMNS = [
    "Referencia",
    "Código",
    "Nombre",
    "Variantes",
    "Cantidad",
    "Total Costos",
]
MESES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]


@app.get("/", response_class=HTMLResponse)
def formulario() -> str:
    return """<!doctype html>
<html lang="es">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Generador de cortes de venta</title>
    <style>
      body { margin: 0; min-height: 100vh; display: grid; place-items: center;
        background: #edf2ef; color: #172b25; font: 16px system-ui, sans-serif; }
      main { width: min(480px, calc(100% - 40px)); padding: 28px;
        background: white; border: 1px solid #c9d6cf; border-radius: 8px; }
      h1 { margin: 0 0 8px; font-size: 24px; }
      p { margin: 0 0 24px; color: #53665d; }
      label { display: block; margin: 16px 0 8px; font-weight: 600; }
      input { display: block; width: 100%; box-sizing: border-box; padding: 12px;
        border: 1px solid #9caca3; border-radius: 4px; }
      button { margin-top: 24px; padding: 11px 16px; border: 0; border-radius: 4px;
        background: #176b50; color: white; font: inherit; font-weight: 600; cursor: pointer; }
      button:hover { background: #10543e; }
    </style>
  </head>
  <body>
    <main>
      <h1>Generar cortes de venta</h1>
      <p>Sube el reporte Excel y elige la fecha de corte.</p>
      <form action="/process" method="post" enctype="multipart/form-data">
        <label for="fecha_corte">Fecha de corte</label>
        <input id="fecha_corte" name="fecha_corte" type="date" required>
        <label for="file">Reporte de ganancias por ítems (.xlsx)</label>
        <input id="file" name="file" type="file" accept=".xlsx" required>
        <button type="submit">Generar y descargar reportes</button>
      </form>
    </main>
  </body>
</html>"""


def agregar_total_por_prefijo(grupo: pd.DataFrame) -> pd.DataFrame:
    fila_total = grupo.iloc[:1].copy()
    fila_total["Código"] = ""
    fila_total["Nombre"] = "Total"
    fila_total["Variantes"] = ""
    fila_total["Cantidad"] = grupo["Cantidad"].sum()
    fila_total["Total Costos"] = grupo["Total Costos"].sum()
    return pd.concat([grupo, fila_total])


def preparar_dataframe(df_rvin: pd.DataFrame) -> pd.DataFrame:
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df_rvin.columns]
    if missing_columns:
        raise HTTPException(
            status_code=400,
            detail=f"Faltan columnas requeridas: {', '.join(missing_columns)}",
        )

    df_rvin_1 = df_rvin[REQUIRED_COLUMNS]
    grupos = (
        agregar_total_por_prefijo(grupo)
        for _, grupo in df_rvin_1.sort_values("Referencia").groupby("Referencia")
    )
    return pd.concat(grupos) if not df_rvin_1.empty else df_rvin_1


def nombre_seguro(valor: object) -> str:
    nombre = re.sub(r"[^\w.-]+", "_", str(valor), flags=re.UNICODE).strip("._")
    return nombre or "sin_referencia"


def generar_paquete(df_totalizado: pd.DataFrame, fecha_corte: date) -> bytes:
    fecha = datetime.today().strftime("Bogotá DC, %d/%m/%Y")
    fecha_de_corte = f"{fecha_corte.day} de {MESES[fecha_corte.month - 1]}"
    periodo = f"{MESES[fecha_corte.month - 1].upper()} {fecha_corte.year % 100:02d}"

    with tempfile.TemporaryDirectory(prefix="cortes_de_ventas_") as temporary_directory:
        temporary_path = Path(temporary_directory)
        word_directory = temporary_path / "Word"
        pdf_directory = temporary_path / "PDF"
        word_directory.mkdir()
        pdf_directory.mkdir()

        for referencia, grupo in df_totalizado.groupby("Referencia"):
            contenido = {
                "reporte_de_ventas": grupo.to_dict(orient="records"),
                "nombre_marca": referencia,
                "fecha": fecha,
                "fecha_de_corte": fecha_de_corte,
            }
            nombre_archivo = (
                f"Corte de ventas realizadas en {periodo}_{nombre_seguro(referencia)}.docx"
            )
            documento = DocxTemplate(TEMPLATE_PATH)
            documento.render(contenido)
            documento.save(word_directory / nombre_archivo)

        try:
            convert(str(word_directory), str(pdf_directory))
        except Exception as error:
            logging.exception("Falló la conversión de los reportes Word a PDF")
            raise HTTPException(
                status_code=503,
                detail=(
                    "No se pudieron convertir los reportes a PDF. Verifica que Microsoft Word "
                    "esté instalado y disponible en este equipo."
                ),
            ) from error

        word_files = sorted(word_directory.glob("*.docx"))
        pdf_files = sorted(pdf_directory.glob("*.pdf"))
        if not word_files or len(word_files) != len(pdf_files):
            raise HTTPException(
                status_code=500,
                detail="No se generó un PDF por cada marca. Revisa la instalación de Microsoft Word.",
            )

        excel_package = BytesIO()
        df_totalizado.to_excel(excel_package, index=False)

        package = BytesIO()
        with zipfile.ZipFile(package, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for path in word_files:
                zip_file.write(path, arcname=f"Word/{path.name}")
            for path in pdf_files:
                zip_file.write(path, arcname=f"PDF/{path.name}")
            zip_file.writestr(
                f"Reporte Ganancias por Items {periodo.title()}.xlsx",
                excel_package.getvalue(),
            )
        return package.getvalue()


@app.post("/process")
def procesar_archivo(
    file: UploadFile = File(...),
    fecha_corte: date = Form(...),
) -> Response:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Sube un archivo Excel con extensión .xlsx.")
    if not TEMPLATE_PATH.is_file():
        raise HTTPException(
            status_code=500,
            detail=f"No se encontró la plantilla de Word: {TEMPLATE_PATH.name}",
        )

    try:
        df_rvin = pd.read_excel(BytesIO(file.file.read()))
    except Exception as error:
        raise HTTPException(
            status_code=400,
            detail="No se pudo leer el archivo. Verifica que sea un Excel .xlsx válido.",
        ) from error

    df_totalizado = preparar_dataframe(df_rvin)
    if df_totalizado.empty:
        raise HTTPException(status_code=400, detail="El Excel no contiene filas para procesar.")

    paquete = generar_paquete(df_totalizado, fecha_corte)
    periodo = f"{MESES[fecha_corte.month - 1].upper()} {fecha_corte.year % 100:02d}"
    return Response(
        content=paquete,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="Cortes_de_ventas_{periodo}.zip"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
