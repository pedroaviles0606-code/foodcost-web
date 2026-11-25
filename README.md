# Food Cost – Control de costos de ventas y operaciones

Web para controlar el food cost de un restaurante con los mismos reportes que ya salen del sistema de ventas (Inforest). Subes los Excel y la web arma sola los reportes, los guarda y los deja listos para exportar.

**Ver la web:** https://foodcost-mesacriolla.onrender.com
Al crear una cuenta nueva se carga una demo con datos de ejemplo (Liberty Bistró).

## Qué tiene

| Pestaña | Para qué sirve |
|---|---|
| Resumen | Food cost teórico y real, margen y alerta de desvío |
| Reportes | Ingeniería de menú, alertas, venta por mozo, Rappi y proyección |
| Insumos | Lista de insumos con costo de compra, merma y costo aprovechable |
| Recetas | Ficha técnica por plato: costo por porción, food cost y precio sugerido |
| Ventas y consumo real | Unidades vendidas e inventario inicial, compras y final por insumo |
| Carga de datos | Subir los archivos de Inforest o plantillas propias |

## Qué mide

- **Food cost teórico:** lo que debió costar lo vendido según las recetas.
- **Food cost con compras:** lo que entró a almacén frente a la venta de alimentos.
- **Ingeniería de menú:** clasifica cada plato en estrella, caballo, puzzle o perro según su venta y su margen.
- **Alertas:** platos sobre el food cost máximo, productos vendidos que no descargan inventario e insumos que subieron de precio.
- **Venta por mozo:** venta neta, pedidos, comensales, ticket y porcentaje de bebidas.
- **Rappi:** venta y food cost del canal delivery, comparando su precio con el del salón.

## Qué proyecta

- Venta de la próxima semana por plato, día por día, con el promedio de cada día de la semana.
- Insumos que se van a necesitar para esa venta y su costo estimado.

## Archivos de Inforest que lee

1. **Analítico de productos por mesero** (paloteo diario). Se pueden subir varios días juntos y cada día reemplaza al anterior si ya existía.
2. **Resumen de ingresos** (compras de almacén).
3. **Hoja de costos – Recetas de venta** (o el Excel de costeo con la hoja MODELO COSTO INFOREST y CARTA ACTUAL).
4. **Maestro de insumos** (hoja INSUMOS con precio promedio, factor y unidad de costo).

## Cómo está armada

```
index.html          estructura de la página (pestañas, tablas, botones)
css/estilos.css     colores, tamaños y diseño
js/config.js        dirección y clave pública de Supabase
js/app.js           cálculos de food cost, pantallas de insumos, recetas y ventas, sesión
js/inforest.js      lectura de los Excel de Inforest
js/reportes.js      cálculo y dibujo de los reportes, exportación a Excel
js/cargas.js        guardado de ventas, mozos y compras; demo; zona de carga
demo/liberty.json   datos de la demo
supabase/esquema.sql tablas y seguridad de la base de datos
```

Tres herramientas, cada una con su trabajo:

- **Supabase** guarda los datos. Cada usuario entra con su correo y solo ve lo suyo.
- **GitHub** guarda el código y el historial de cambios.
- **Render** publica la web. Cada vez que se sube un cambio a GitHub, se actualiza sola.

## Cómo publicar un cambio

1. Edita el archivo que quieras en tu computadora.
2. En GitHub: **Add file → Upload files**, arrastra el archivo (o la carpeta) y dale **Commit changes**.
3. Espera unos 30 segundos y recarga la web con **Ctrl + F5**.

No subas Excel con datos reales al repositorio: es público y cualquiera puede verlo.
