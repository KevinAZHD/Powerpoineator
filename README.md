# Powerpoineador-PhotomakerEdition

Powerpoineador-PhotomakerEdition es un script de Python que genera automáticamente presentaciones de PowerPoint a partir de un tema dado por el usuario, utilizando una serie de imágenes con cara dadas por el mismo usuario, y por último utiliza la Inteligencia Artificial para su creación.

## (CAMBIAR) Descripción

El script solicita al usuario un tema para la presentación. Luego, utiliza dos modelos de Inteligencia Artificial alojados en Replicate: Llama 3 y SDXL-Lightning by ByteDance.

- **Llama 3** es un modelo de lenguaje que proporciona una estructura para la presentación.
- **SDXL-Lightning by ByteDance** es un modelo que genera imágenes para cada diapositiva.

Las imágenes se generan en base a las imágenes principales (`Imagen1.jpg` e `Imagen2.jpg`) en unas nuevas que corresponden a los contenidos de las diapositivas. Finalmente, se crea y abre la presentación de PowerPoint con las imágenes generadas y el texto proporcionado por el modelo de lenguaje.

## Requisitos

- Python 3 o superior
- Bibliotecas Python: python-pptx, subprocess, os, requests, replicate, PIL, time
- Una API en Replicate: https://replicate.com/
- Dos imágenes nombradas como `Imagen1.jpg` e `Imagen2.jpg` guardadas dentro del mismo directorio del script
- Sistema operativo Windows con PowerPoint (opcional)

## Configuración

Define tu token de API de Replicate en tu entorno:

```python
os.environ["REPLICATE_API_TOKEN"] = "escriba_su_api_aquí"
```
## Uso
-Guarde el archivo `Powerpoineador-PhotomakerEdition.py` dentro de una carpeta.

-Abre una terminal dentro de la misma carpeta.

-Ejecute el script con el siguiente comando:

```python
python Powerpoineador-PhotomakerEdition.py
```

-Sigue las instrucciones del script para generar tu presentación.

## Licencia

Este proyecto se ha creado de manera Open-Source bajo la licencia GPL (Licencia Pública General de GNU). Esto significa que puedes copiar, modificar y distribuir el código, siempre y cuando mantengas la misma licencia y hagas público cualquier cambio que realices.

## Soporte

Si tiene algún problema o duda con respecto a esta guía o al Powerpoineador-PhotomakerEdition, no dude en comunicarlo. Estamos aquí tanto yo como el co-creador del proyecto para ayudar y mejorar continuamente este recurso para la comunidad.

Por favor, tenga en cuenta que este proyecto se mantiene con la intención de ser un recurso útil y profesional. Cualquier contribución o sugerencia para mejorar es siempre bienvenida.

# Créditos

- Este proyecto ha sido desarrollado por Kevin Adán Zamora (@KevinAZHD) https://github.com/KevinAZHD y Diego Martínez Fernández (@Dgmtnz) https://github.com/Dgmtnz

- Todos los enlaces proporcionados anteriormente.

- Gracias por utilizar esta guía y script de Python.
