# Revisión del catálogo — 14/10/2026

Se revisó visualmente el PDF completo y se contrastó con catalogo_productos.xlsx y las fichas actuales de la tienda.

## Resultado

El diseño se ve ordenado y uniforme. Hay 18 páginas: portada más 17 páginas de productos, con 6 tarjetas por página. Los 102 productos están presentes, sin SKU duplicados, imágenes ausentes ni archivos de imagen dañados. Los precios impresos de los 102 productos coinciden con el Excel. Las direcciones, horarios y datos de WhatsApp coinciden con los publicados por el comercio.

Los 98 productos que tienen ficha identificada en la tienda muestran el mismo SKU y precio que el Excel actualizado. Los otros cuatro precios no se pudieron contrastar porque no se encontró su ficha. Los 98 enlaces de producto están en el PDF, en imagen, título y botón; los enlaces de WhatsApp usan el número 5492234173945.

## Correcciones aplicadas

El usuario confirmó que se deben usar los precios de la web. Se corrigieron estos tres importes y se regeneró salida/catalogo.pdf:

| SKU | Producto | Precio anterior | Precio corregido |
| --- | --- | --- | --- |
| 5000 | Lavandina Palacio 5 litros | $1.000 | $4.300 |
| 5002 | Desodorante Palacio 5 litros, variante lavanda | $4.500 | $4.000 |
| 6221 | Toalla Ecco x2500 | $26.000 | $25.999 |

Para las bolsas SKU 3521 y 5014 se cambió la presentación “Paquete estándar” por “Unidad (venta x50)”. Los precios publicados son $220 y $190 respectivamente; las fichas indican compra en múltiplos de 50. Se conservó el precio por unidad para que coincida con la web.

Se guardaron copias del Excel y del PDF previos a estas correcciones en esta carpeta. Se conservaron las otras celdas, los estilos y los enlaces del Excel.

## Datos para confirmar antes de entregar

- **Fragancias del desodorante 5002:** la ficha enlazada y el precio corregido de $4.000 corresponden a lavanda. La tarjeta del catálogo lista varias fragancias y no todas tienen ese precio: la variante limón figura a $4.500 entre los productos relacionados y citronella a $5.500 en la portada de la tienda. Si la tarjeta representa exclusivamente el SKU 5002, conviene indicar “Lavanda” en Fragancias. Si representa todas las variantes, conviene mostrar sus precios por separado o aclarar que el precio corresponde a lavanda. No se cambiaron las fragancias originales del Excel.
- **Foto de la toalla 6211:** en la página 5 del PDF (página 4 de productos), el título identifica la toalla como Palacio, pero la foto muestra un envase Campanita. Confirmar si es una imagen ilustrativa o si corresponde reemplazarla. La página completa está en pagina_05.png.
- **Cuatro URLs pendientes:** SKU 6211, 952, 100130 y 3269. Sus botones llevan a la página principal mientras no se complete la URL. Después de guardar los enlaces manuales en el Excel, ejecutar nuevamente generar_catalogo.py o ejecutar_catalogo.bat para incorporarlos al PDF.

## Detalles menores de diseño

Algunos títulos largos quedan cerca del borde derecho de su tarjeta; especialmente Pantene Shampoo Rizos (952) y Sedal Shampoo Sachets (1927). Se leen completos y no se observaron superposiciones que impidan leer los datos. La descripción del detergente 5001 y las listas largas de fragancias terminan en puntos suspensivos por el espacio disponible. Los precios y botones se leen completos.

Las fotografías tienen tamaños y fondos diferentes, pero todas se cargan correctamente. La portada es una imagen y sus textos no son seleccionables; visualmente se ve correctamente. El archivo pesa aproximadamente 3,59 MB.

## Fuentes y comprobaciones

- [Lavandina Palacio 5 litros](https://www.elpalaciodelalimpieza.com.ar/lavandina-5l-palacio-5000).
- [Desodorante lavanda 5 litros](https://www.elpalaciodelalimpieza.com.ar/desodorante-piso-5l-lavanda-5002).
- [Toalla Ecco x2500](https://www.elpalaciodelalimpieza.com.ar/toalla-blanca-ecco-x2500-18-4x24cm-6221).
- [Bolsas verdes 70×100](https://www.elpalaciodelalimpieza.com.ar/bolsas-70x100-verdes-3521) y [bolsas negras 70×100](https://www.elpalaciodelalimpieza.com.ar/bolsas-70x100-negras-5014).
- [Portada de la tienda: desodorante citronella a $5.500](https://www.elpalaciodelalimpieza.com.ar/).
- comparacion_web.json: detalle de las 98 fichas y sus precios.
- comprobaciones.json: comprobaciones del PDF, imágenes y medidas de los títulos.
- cambios_revision.json: las cinco celdas modificadas durante la revisión.
