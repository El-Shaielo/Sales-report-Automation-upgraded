# Generador de cortes de venta

API FastAPI para generar reportes de cortes de venta en Word, PDF y Excel.
La imagen Docker instala LibreOffice Writer para convertir los documentos Word
a PDF en el servidor.

## Despliegue en Render

1. Sube estos archivos y la plantilla `Plantilla Cortes de ventas Oficial.docx`
   al repositorio conectado con Render.
2. En Render, crea un **Blueprint** desde ese repositorio. Render leerá
   `render.yaml` y construirá el servicio con `Dockerfile`.
3. Abre la URL del servicio. La ruta `/` muestra el formulario de carga.

Render proporciona la variable `PORT`; el contenedor la usa automáticamente.
Los archivos generados se devuelven en una descarga ZIP y los temporales se
eliminan después de cada solicitud, por lo que no se requiere disco persistente.

## Prueba local con Docker

```sh
docker build -t cortes-de-venta .
docker run --rm -p 10000:10000 -e PORT=10000 cortes-de-venta
```

Luego abre <http://localhost:10000>.
