# fus-bitcli

Herramienta CLI para generar bitacoras diarias para Fusap

![Mockupimage](./assets/mockup.png "mockup")

# Features

- Pregunta por linea de comando la cantidad de integrantes de cada proyecto y autocompleta las filas con cada uno.
- Automaticamente colorea fines de semana de todos los meses del año actual.
- La formula `=SUM` se agrega en cada sheet del mes para sumar el total del horas de todos los integrantes.
- Logre entender como estilizar el excel asi que supongamos que es **customizable**

# Como utilizar

## Pre-requisitos

- Tener instalada la ultima version de [node](https://nodejs.org/en) LTS.
- Clonar el proyecto y ejecutar `npm install` dentro de la carpeta del proyecto.

### Sin compilar

*Para no tener que compilar el proyecto, subo la carpeta `/dist` que contiene el programa compilado (**fusbitcli.js)***

1. Dentro de la carpeta del proyecto ejecutar: `npm run start` dentro de la carpeta del proyecto.

### Con compilacion

Si les interesa hacer pruebas, para compilar el proyecto tienen que:

1. Dentro de la carpeta del proyecto ejecutar `npm run build` dentro de la carpeta del proyecto.