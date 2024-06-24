# Powerpoineator

Powerpoineator es un script de Python que genera automáticamente presentaciones de PowerPoint a partir de un tema dado por el usuario, utilizando la Inteligencia Artificial.

## Descripción

El script solicita al usuario un tema para la presentación. Luego, utiliza una API de Replicate para obtener una respuesta del modelo de lenguaje GPT-3, que proporciona una estructura para la presentación. A continuación, se generan las imágenes para cada diapositiva utilizando la misma API de Replicate y se guardan localmente. Finalmente, se crea y abre la presentación de PowerPoint con las imágenes generadas y el texto proporcionado por el modelo de lenguaje.

## Requisitos

- Sistema operativo Windows con PowerPoint
- Python 3 o superior
- Bibliotecas Python: pptx, subprocess, os, requests, replicate, PIL, time
- Una API en Replicate: https://replicate.com/

## Configuración

Define tu token de API de Replicate en tu entorno:

```python
os.environ["REPLICATE_API_TOKEN"] = "escriba_su_api_aquí"
```
##Uso
Guarda el archivo Powerpoineator.py dentro de una carpeta.
Abre una terminal desde dentro de la misma carpeta.
Ejecuta el script con el siguiente comando: python Powerpoineator.py
Sigue las instrucciones del script para generar tu presentación.

##Licencia y Soporte
Este proyecto se ha creado de manera Open-Source bajo la licencia GPL (Licencia Pública General de GNU). Esto significa que puedes copiar, modificar y distribuir el código, siempre y cuando mantengas la misma licencia y hagas público cualquier cambio que realices.

Si tienes algún problema o duda con respecto a esta guía o al Powerpoineator, no dudes en comunicarlo. Estamos aquí tanto yo como el co-creador del proyecto para ayudar y mejorar continuamente este recurso para la comunidad.

Por favor, ten en cuenta que este proyecto se mantiene con la intención de ser un recurso útil y profesional. Cualquier contribución o sugerencia para mejorar es siempre bienvenida.

# Créditos

- Este proyecto ha sido desarrollado por Kevin Adán Zamora y Diego Martínez Fernández (@Dgmtnz) https://github.com/Dgmtnz

- Todos los enlaces proporcionados anteriormente.

- Gracias por utilizar esta guía y script de Python.
