# Valorizaci-n-Retiros-Mercado-Libre

La idea de este proyecto es que desde un archivo de Google Sheets se pueda leer un archivo guardado en Drive que contenga toda la información buscada (ID de producto y valor) de cada uno de los más de un millón al momento.

Dentro del script se puede ver que se va a buscar un archivo que tenga como nombre la fecha actual y desde ahí va a ir copiando y pegando la información de a varias partes para que no se sature la página de Sheets. Una vez finalizado se notificará al usuario que se ha actualizado todo correctamente aunque dentro del proyecto hay un trigger que va a ejecutar la función una vez por día en determinado horario, por lo que lo único que se debe hacer es cargar el archivo correspondiente en Drive.
