DLSR MAPS - Versión v7 profesional

Archivos principales:
- index.html: menú principal compacto y adaptable a teléfono.
- buscar.html: búsqueda general inteligente.
- placa.html: búsqueda por placa con sugerencias.
- contrato.html: búsqueda por contrato con sugerencias.
- ubicacion.html: búsqueda de cercanos por coordenadas, referencia, alimentador o ubicación actual.
- campo.html: modo campo para celular.
- favoritos.html: favoritos y recientes guardados localmente.
- estado.html: revisión del sistema.
- convertidor.html: genera JSON optimizado desde Excel.

Datos:
- La app intenta usar data_placas.json, data_contratos.json o data_index.json si existen.
- Si no existen, usa automáticamente el Excel más reciente declarado en data-config.js.
- Para mejor rendimiento móvil, genere los JSON desde convertidor.html y súbalos a la raíz del repositorio.

Importante después de subir:
1. Reemplazar todos los archivos anteriores en la raíz del repositorio.
2. Mantener los Excel en la raíz si aún no usa JSON.
3. Hacer recarga forzada en el navegador o limpiar caché desde estado.html.
