/* =====================================================================
   LECTURA DE ARCHIVOS DE INFOREST
   ---------------------------------------------------------------------
   Este archivo solo "lee" los Excel que exportas de Inforest y los
   convierte en listas ordenadas. No guarda nada ni dibuja pantallas.

   Reconoce 4 reportes:
   1. Analítico de productos por mesero (paloteo diario)  -> ventas
   2. Resumen de ingresos (compras de almacén)             -> compras
   3. Hoja de costos - Recetas de venta                    -> recetas
   4. Maestro de insumos (hoja INSUMOS)                    -> insumos
   ===================================================================== */

/* Convierte una hoja de Excel en una lista de filas (cada fila es una lista de celdas). */
function filasDeHoja(hoja) {
  return XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "", raw: true });
}

/* Devuelve solo las celdas que tienen algo escrito, con su número de columna. */
function celdasLlenas(fila) {
  const out = [];
  fila.forEach((c, k) => {
    if (c !== null && c !== undefined && String(c).trim() !== "") out.push([k, c]);
  });
  return out;
}

const MESES = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, set: 9, sep: 9, oct: 10, nov: 11, dic: 12 };

/* "Del 01/nov/2026" -> "2026-11-01" */
function fechaInforest(texto) {
  const m = String(texto).match(/(\d{1,2})\/([a-zA-Z]{3})\/(\d{4})/);
  if (!m) return null;
  const mes = MESES[m[2].toLowerCase()];
  if (!mes) return null;
  return `${m[3]}-${String(mes).padStart(2, "0")}-${m[1].padStart(2, "0")}`;
}

/* "2/11/2026" o fecha de Excel -> "2026-11-02" */
function fechaDMA(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : null;
}

/* ---------------------------------------------------------------------
   ¿Qué reporte es? Mira el título y las hojas del archivo.
   --------------------------------------------------------------------- */
function detectarTipo(libro) {
  const nombres = libro.SheetNames.map(norm);
  if (nombres.includes("modelocostoinforest")) return "recetas";
  if (nombres.includes("insumos") && libro.SheetNames.length > 1) return "maestro";
  const filas = filasDeHoja(libro.Sheets[libro.SheetNames[0]]).slice(0, 40);
  const texto = norm(filas.map((f) => f.join(" ")).join(" "));
  if (texto.includes("analiticodeproductos")) return "paloteo";
  if (texto.includes("resumendeingresos")) return "compras";
  if (texto.includes("hojadecostos") && texto.includes("recetaventa")) return "recetas";
  if (texto.includes("descripciondeinsumo") || texto.includes("insumosdisponibles")) return "maestro";
  return null;
}

/* ---------------------------------------------------------------------
   1. PALOTEO DIARIO (Analítico de productos por mesero)
   Cada mozo tiene un bloque: "Mesero: Nombre  Comensales N  Pedidos N"
   y debajo sus productos con Grupo, SubGrupo, Producto, Cantidad y Neto.
   El "Neto" ya viene sin IGV ni recargo.
   --------------------------------------------------------------------- */
function leerPaloteo(filas) {
  let fecha = null;
  let mozo = null;
  let columnas = null;
  const lineas = [];
  const mozos = [];

  for (const fila of filas) {
    const llenas = celdasLlenas(fila);
    if (!llenas.length) continue;
    const primera = String(llenas[0][1]).trim();

    if (!fecha) fecha = fechaInforest(llenas.map((c) => c[1]).join(" "));

    if (primera.startsWith("Mesero")) {
      const valores = llenas.map((c) => c[1]);
      mozo = { mozo: String(valores[1]).trim(), comensales: 0, pedidos: 0 };
      valores.forEach((v, k) => {
        if (String(v).startsWith("Comensales")) mozo.comensales = num(valores[k + 1]);
        if (String(v).startsWith("Pedidos")) mozo.pedidos = num(valores[k + 1]);
      });
      mozos.push(mozo);
      continue;
    }

    if (primera === "Grupo") {
      columnas = {};
      fila.forEach((c, k) => (columnas[norm(c)] = k));
      continue;
    }

    if (!columnas || !mozo || primera.startsWith("Ticket") || primera === "Total") continue;
    const dato = (nombre) => fila[columnas[nombre]];
    const cantidad = dato("cantidad");
    const netoLinea = dato("neto");
    if (typeof cantidad !== "number" || typeof netoLinea !== "number") continue;
    const grupo = String(dato("grupo")).trim();
    lineas.push({
      mozo: mozo.mozo,
      grupo,
      subgrupo: String(dato("subgrupo")).trim(),
      producto: String(dato("producto")).trim(),
      canal: grupo === "Carta Rappi" ? "Rappi" : "Salón",
      cantidad,
      neto: netoLinea,
    });
  }
  lineas.forEach((l) => (l.fecha = fecha));
  mozos.forEach((m) => (m.fecha = fecha));
  return { fecha, lineas, mozos };
}

/* ---------------------------------------------------------------------
   2. COMPRAS (Resumen de ingresos agrupado por cuenta contable)
   Las filas de "Familia ... SubFamilia ..." marcan el grupo de los
   artículos que vienen debajo. Se usa el SubTotal (sin IGV).
   --------------------------------------------------------------------- */
function leerCompras(filas) {
  let columnas = null;
  let familia = "";
  let subfamilia = "";
  const lineas = [];

  for (const fila of filas) {
    const llenas = celdasLlenas(fila);
    if (!llenas.length) continue;
    const primera = String(llenas[0][1]).trim();

    if (primera.startsWith("Correlativ")) {
      columnas = {};
      llenas.forEach(([k, c]) => (columnas[norm(c)] = k));
      continue;
    }
    if (primera.startsWith("Familia")) {
      const t = llenas.map((c) => String(c[1]).trim());
      familia = t[1] || "";
      subfamilia = t[3] || "";
      continue;
    }
    if (!columnas || !/^\d{6}-\d+/.test(primera)) continue;

    const dato = (n) => fila[columnas[n]];
    const fecha = fechaDMA(dato("fecha"));
    const subtotal = num(dato("subtotal"));
    if (!fecha || !dato("articulo")) continue;
    // la razón social está en la primera celda con texto después del RUC
    const proveedor = llenas.find(([k, c]) => k > columnas["razonsocial"] && isNaN(Number(c)))?.[1] || "";
    lineas.push({
      fecha,
      familia,
      subfamilia,
      codigo: String(dato("codigo") || "").trim(),
      articulo: String(dato("articulo")).trim(),
      unidad: String(dato("unidad") || "").trim(),
      precio: num(dato("precio")),
      cantidad: num(dato("cantidad")),
      subtotal,
      proveedor: String(proveedor).trim(),
      documento: `${dato("serie") || ""}-${dato("document") || ""}`,
    });
  }
  return lineas;
}

/* ---------------------------------------------------------------------
   3. RECETAS (Hoja de costos - Recetas de venta)
   Cada receta empieza con "Codigo :" y "Receta Venta:", sigue la lista
   de insumos y cierra con "Venta :" (precio de carta con impuestos).
   Si el libro trae la hoja CARTA ACTUAL, se marca cuáles están en carta.
   --------------------------------------------------------------------- */
function leerRecetas(filas, enCarta) {
  const recetas = [];
  let actual = null;
  let grupo = "";
  let pos = null;

  for (const fila of filas) {
    const llenas = celdasLlenas(fila);
    if (!llenas.length) continue;
    const primera = String(llenas[0][1]).trim();

    if (llenas.length === 1 && primera.includes(" / ")) {
      grupo = primera;
      continue;
    }
    if (primera.startsWith("Codigo")) {
      actual = { cod: String(llenas[1]?.[1] ?? "").trim(), grupo, nom: "", precio: 0, ing: [] };
      recetas.push(actual);
      continue;
    }
    if (actual && primera.startsWith("Receta Venta")) {
      actual.nom = String(llenas[1]?.[1] ?? "").trim();
      continue;
    }
    if (primera === "Item") {
      // columnas en orden: Item, Insumo, Unidad compra, Factor, Unidad costo, Cantidad, Precio, Sub Total
      const k = llenas.map((c) => c[0]);
      pos = { insumo: k[1], uCompra: k[2], factor: k[3], uCosto: k[4], cantidad: k[5], precio: k[6] };
      continue;
    }
    if (actual && pos && /^\d{3}$/.test(primera)) {
      actual.ing.push({
        insumo: String(fila[pos.insumo]).trim(),
        uCompra: String(fila[pos.uCompra]).trim(),
        factor: num(fila[pos.factor]) || 1,
        uCosto: String(fila[pos.uCosto]).trim(),
        cantidad: num(fila[pos.cantidad]),
        precio: num(fila[pos.precio]),
      });
      continue;
    }
    if (actual && primera.startsWith("Venta")) actual.precio = num(llenas[1]?.[1]);
  }

  recetas.forEach((r) => {
    const partes = r.grupo.split(" / ");
    r.area = partes[0] || "";
    r.familia = partes.slice(1).join(" / ") || "";
    r.enCarta = enCarta ? enCarta.has(r.cod) : null;
  });
  return recetas.filter((r) => r.nom);
}

/* Lee la hoja CARTA ACTUAL (si existe) y devuelve los códigos marcados "Sí". */
function codigosEnCarta(libro) {
  const hoja = libro.Sheets["CARTA ACTUAL"];
  if (!hoja) return null;
  const filas = filasDeHoja(hoja);
  const cab = filas.findIndex((f) => norm(f[0]) === "codigo");
  if (cab < 0) return null;
  const col = filas[cab].findIndex((c) => norm(c).startsWith("encarta"));
  const set = new Set();
  filas.slice(cab + 1).forEach((f) => {
    if (f[0] && norm(f[col]) === "si") set.add(String(f[0]).trim());
  });
  return set;
}

/* Convierte una línea de receta a la unidad base de la app (kg, lt o und).
   Ej.: 150 Gramos de un insumo que se compra en Kilos a S/ 7.48
        -> cantidad 0.150 kg, costo S/ 7.48 por kg */
function lineaABase(l) {
  const conv = UCONV[norm(l.uCosto)] || ["und", 1];
  return {
    uni: conv[0],
    cantidad: l.cantidad / conv[1],
    costo: (l.precio / (l.factor || 1)) * conv[1],
  };
}
