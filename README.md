# Instrucciones de ejecución

## Requisitos

- Nodejs
- npm

## Ejecución

1. Ejecutar npm install
2. Modificar la ruta del archivo xsl en la variable *xlsxFile* en el archivo *app.js*
3. Modificar las celdas, de acuerdo a la estructura del excel.
4. Modificar el rango inicial en las variables *initial* y *limit*. Se recomienda utilizar un rango de 10 elementos.
5. Ejecutar el comando `npm run nodemon`
6. Cada vez que se actualiza el archivo app.js, se enviara nuevamente la petición. Por lo que basta con modificar nuevamente las variables *initial* y *limit* en un rango de 10 en 10
