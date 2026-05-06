DLSR MAPS - versión optimizada v5

Archivos principales:
- index.html
- acceso.html
- placa.html
- contrato.html
- ubicacion.html
- styles.css
- app-core.js
- data-worker.js
- data-config.js
- sw.js
- manifest.json
- logo.png
- logo_ui.png
- logo_fondo_blanco.png

Mejoras aplicadas:
1. La lectura y búsqueda sobre Excel ahora se ejecuta en un Web Worker.
   Esto evita que la página se congele mientras el usuario escribe.
2. En búsqueda por placa se agregaron sugerencias tipo buscador.
   Permite buscar por referencia completa, parcial o por número contenido.
3. Las sugerencias ahora forman parte del flujo visual de la página.
   Ya no quedan superpuestas ni ocultas detrás del mapa o de los resultados.
4. El mapa se inicializa solo cuando se necesita mostrar resultados.
5. El inicio no carga Excel ni mapas, por lo que debe abrir más rápido.
6. Se suavizaron los colores y botones para una apariencia más profesional.
7. Se mantuvieron las funciones de Google Maps, Waze, geolocalización, contrato, placa y ubicación.

Para actualizar la data:
- Edite data-config.js.
- Agregue el nombre del nuevo Excel en la lista.
- El sistema elegirá automáticamente el archivo más reciente según la fecha del nombre.

Recomendación al publicar:
- Subir todos los archivos a la raíz del repositorio.
- Confirmar que los archivos Excel estén también en la raíz.
- Después de publicar, abrir la página y presionar Ctrl + F5 para limpiar caché.
