# Powerpoineator

Esto es un script de Python que genera automáticamente presentaciones de PowerPoint gracias a la Inteligencia Artificial, a partir de un tema dado por el usuario.

## Descripción

El script solicita al usuario un tema para la presentación. Luego, utiliza una API del Replicate para obtener una respuesta del modelo de lenguaje GPT, que proporciona una estructura para la presentación. A continuación, se generan las imágenes para cada diapositiva utilizando la misma API del Replicate y las guarda localmente. Finalmente, se crea y abre la presentación de PowerPoint con las imágenes generadas y el texto proporcionado por el modelo de lenguaje.

## Requisitos

- Sistema operativo Windows con PowerPoint
- Python 3 o superior
  Instalar las siguientes bibliotecas:
- pptx
- subprocess
- os
- requests
- replicate
- PIL
- time

## Utilización

1. Guarda el `Powerpoineator.py` dentro de una carpeta.
2. Abre desde dentro de la misma carpeta una terminal.
3. Ejecuta el script con el siguiente comando: `python main.py`
4. Siga las instrucciones del script para generar su presentación.

## Licencia y Soporte

Este proyecto se ha creado de manera Open-Source bajo la licencia GPL (Licencia Pública General de GNU). Esto significa que puedes copiar, modificar y distribuir el código, siempre y cuando mantengas la misma licencia y hagas público cualquier cambio que realices.

Si tienes algún problema o duda con respecto a esta guía o al Powerpoineator, no dudes en comunicarte. Estoy aquí para ayudar y mejorar continuamente este recurso para la comunidad.

Por favor, ten en cuenta que este proyecto se mantiene con la intención de ser un recurso útil y profesional. Cualquier contribución o sugerencia para mejorar es siempre bienvenida.

Gracias por usar esta guía y script de Python.

# Créditos

- Agradecimientos y colaborador principal del proyecto: @Dgmtnz https://github.com/Dgmtnz

- KevinAZ.
