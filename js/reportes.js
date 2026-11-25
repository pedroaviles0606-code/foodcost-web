/* =====================================================================
   REPORTES
   ---------------------------------------------------------------------
   Aquí se calculan y se dibujan los reportes de la pestaña "Reportes".
   Todo sale de 3 listas que se cargan desde los archivos de Inforest:
     DATOS.ventas   -> líneas del paloteo diario (producto, mozo, neto...)
     DATOS.mozos    -> comensales y pedidos por mozo y día
     DATOS.compras  -> líneas del resumen de ingresos
   y de las recetas (S.rec) y los insumos (S.ins) de la app.

   La función calcularReportes() hace todas las cuentas y devuelve un
   objeto con los resultados. renderReportes() solo los muestra, y
   exportarReportes() los manda a Excel. Así cada parte es fácil de leer.
   ===================================================================== */

/* Subgrupos y grupos que no entran en la ingeniería de menú */
const NO_MENU = ["adicionales", "mascotas"];
const NO_VENTA = ["total", "hotelypromociones", "promociones"];
/* Familias de compras que NO son comida (para el food cost real) */
const NO_ALIMENTOS = ["bebidasalcoholicas", "bebidassinalcohol", "vinosyespumantes", "suministrosdiversos"];
const DIAS = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];

/* ---------- ayudas ---------- */
const enPeriodo = (f) => (!S.cfg.ini || f >= S.cfg.ini) && (!S.cfg.fin || f <= S.cfg.fin);
const sinR = (p) => String(p).replace(/^\(R\)\s*/i, "");
const diaSemana = (f) => new Date(f + "T12:00:00").getDay();
function sumarDias(f, n) {
  const d = new Date(f + "T12:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}
function agrupar(lista, clave, inicial, sumar) {
  const m = new Map();
  lista.forEach((x) => {
    const k = clave(x);
    if (!m.has(k)) m.set(k, inicial(x));
    sumar(m.get(k), x);
  });
  return [...m.values()];
}

/* Busca la receta de un producto del paloteo por su nombre (sin el "(R)" de Rappi) */
function mapaRecetas() {
  const m = new Map();
  S.rec.forEach((r) => m.set(norm(r.nom), r));
  return m;
}

/* Pone en cada receta las unidades vendidas del período (salón + Rappi) */
function unidadesDesdeVentas() {
  if (!DATOS.ventas.length) return;
  const mapa = mapaRecetas();
  const u = new Map();
  DATOS.ventas.forEach((l) => {
    if (!enPeriodo(l.fecha)) return;
    const r = mapa.get(norm(sinR(l.producto)));
    if (r) u.set(r.id, (u.get(r.id) || 0) + l.cantidad);
  });
  S.rec.forEach((r) => (r.ven = +(u.get(r.id) || 0).toFixed(2)));
}

/* =====================================================================
   CÁLCULOS
   ===================================================================== */
function calcularReportes() {
  const mapa = mapaRecetas();
  const costoCache = new Map();
  const costoDe = (r) => {
    if (!costoCache.has(r.id)) costoCache.set(r.id, costoRec(r));
    return costoCache.get(r.id);
  };

  const ventas = DATOS.ventas.filter((l) => enPeriodo(l.fecha) && !NO_VENTA.includes(norm(l.grupo)));
  const compras = DATOS.compras.filter((c) => enPeriodo(c.fecha));
  const mozosDia = DATOS.mozos.filter((m) => enPeriodo(m.fecha));
  const dias = [...new Set(ventas.map((l) => l.fecha))].sort();

  // cada línea con su receta y su costo
  ventas.forEach((l) => {
    l._r = mapa.get(norm(sinR(l.producto))) || null;
    l._c = l._r ? costoDe(l._r) : 0;
    l._alim = l.canal === "Rappi" || norm(l.grupo) === "alimentos";
  });

  /* ---------- 1. Venta y food cost ---------- */
  const total = ventas.reduce((s, l) => s + l.neto, 0);
  const alim = ventas.filter((l) => l._alim);
  const ventaAlim = alim.reduce((s, l) => s + l.neto, 0);
  const costoTeo = alim.reduce((s, l) => s + l.cantidad * l._c, 0);
  const ventaSinCosto = alim.filter((l) => !l._c).reduce((s, l) => s + l.neto, 0);
  const compraAlim = compras
    .filter((c) => !NO_ALIMENTOS.includes(norm(c.familia)))
    .reduce((s, c) => s + c.subtotal, 0);
  const fcTeo = ventaAlim ? costoTeo / ventaAlim : 0;
  const fcReal = ventaAlim ? compraAlim / ventaAlim : 0;
  const porCanal = agrupar(ventas, (l) => l.canal, (l) => ({ canal: l.canal, venta: 0 }), (a, l) => (a.venta += l.neto));
  const comprasFamilia = agrupar(
    compras,
    (c) => c.familia + "|" + c.subfamilia,
    (c) => ({ familia: c.familia, subfamilia: c.subfamilia, monto: 0, lineas: 0 }),
    (a, c) => {
      a.monto += c.subtotal;
      a.lineas++;
    },
  ).sort((a, b) => b.monto - a.monto);

  /* ---------- 2. Ingeniería de menú (salón, alimentos, con receta) ---------- */
  const lineasMenu = alim.filter(
    (l) => l.canal === "Salón" && l._r && l._c > 0 && l.neto > 0 && !NO_MENU.includes(norm(l.subgrupo)),
  );
  let menu = agrupar(
    lineasMenu,
    (l) => l._r.id,
    (l) => ({ plato: l._r.nom, familia: l.subgrupo, costo: l._c, cant: 0, venta: 0 }),
    (a, l) => {
      a.cant += l.cantidad;
      a.venta += l.neto;
    },
  );
  const totCant = menu.reduce((s, x) => s + x.cant, 0);
  menu.forEach((x) => {
    x.precioNeto = x.venta / x.cant;
    x.margen = x.precioNeto - x.costo;
    x.margenTotal = x.margen * x.cant;
    x.mix = totCant ? x.cant / totCant : 0;
    x.fc = x.costo / x.precioNeto;
  });
  const corteMix = menu.length ? 0.7 / menu.length : 0;
  const corteMargen = totCant ? menu.reduce((s, x) => s + x.margenTotal, 0) / totCant : 0;
  menu.forEach((x) => {
    const pop = x.mix >= corteMix,
      ren = x.margen >= corteMargen;
    x.clase = pop && ren ? "Estrella" : pop ? "Caballo" : ren ? "Puzzle" : "Perro";
  });
  menu.sort((a, b) => b.cant - a.cant);

  /* ---------- 3. Alertas ---------- */
  const alerta = (num(S.cfg.alerta) || 35) / 100;
  const sobreAlerta = menu
    .filter((x) => x.fc > alerta)
    .map((x) => ({ ...x, sugerido: sugerido(x.costo) }))
    .sort((a, b) => b.fc - a.fc);

  const sinCosto = agrupar(
    ventas.filter((l) => !l._c && l.neto > 0 && !NO_MENU.includes(norm(l.subgrupo))),
    (l) => norm(l.producto),
    (l) => ({ producto: l.producto, grupo: l.grupo, canal: l.canal, receta: !!l._r, cant: 0, venta: 0 }),
    (a, l) => {
      a.cant += l.cantidad;
      a.venta += l.neto;
    },
  ).sort((a, b) => b.venta - a.venta);

  const alzas = agrupar(
    [...compras].sort((a, b) => a.fecha.localeCompare(b.fecha)),
    (c) => norm(c.articulo),
    (c) => ({ articulo: c.articulo, unidad: c.unidad, familia: c.familia, primero: c.precio, ultimo: c.precio, compras: 0, cantidad: 0 }),
    (a, c) => {
      a.ultimo = c.precio;
      a.compras++;
      a.cantidad += c.cantidad;
    },
  )
    .filter((a) => a.compras > 1 && a.primero > 0)
    .map((a) => ({ ...a, variacion: a.ultimo / a.primero - 1, impacto: (a.ultimo - a.primero) * a.cantidad }))
    .filter((a) => a.variacion >= 0.05)
    .sort((a, b) => b.impacto - a.impacto);

  /* ---------- 4. Venta por mozo ---------- */
  const mozos = agrupar(
    ventas,
    (l) => l.mozo,
    (l) => ({ mozo: l.mozo, venta: 0, alimentos: 0, bebidas: 0, comensales: 0, pedidos: 0 }),
    (a, l) => {
      a.venta += l.neto;
      if (l._alim) a.alimentos += l.neto;
      else a.bebidas += l.neto;
    },
  );
  mozosDia.forEach((m) => {
    const a = mozos.find((x) => x.mozo === m.mozo);
    if (a) {
      a.comensales += m.comensales;
      a.pedidos += m.pedidos;
    }
  });
  mozos.forEach((a) => {
    a.ticket = a.pedidos ? a.venta / a.pedidos : 0;
    a.porComensal = a.comensales ? a.venta / a.comensales : 0;
    a.pctBebidas = a.venta ? a.bebidas / a.venta : 0;
  });
  mozos.sort((a, b) => b.venta - a.venta);

  /* ---------- 5. Rappi ---------- */
  const rappiL = ventas.filter((l) => l.canal === "Rappi");
  const rappi = agrupar(
    rappiL,
    (l) => norm(sinR(l.producto)),
    (l) => ({ producto: sinR(l.producto), costo: l._c, cant: 0, venta: 0 }),
    (a, l) => {
      a.cant += l.cantidad;
      a.venta += l.neto;
    },
  ).sort((a, b) => b.venta - a.venta);
  rappi.forEach((x) => {
    x.fc = x.costo && x.venta ? (x.costo * x.cant) / x.venta : null;
    const s = menu.find((m) => norm(m.plato) === norm(x.producto));
    x.precioSalon = s ? s.precioNeto : null;
    x.precioRappi = x.venta / x.cant;
  });
  const rappiVenta = rappiL.reduce((s, l) => s + l.neto, 0);
  const rappiCosto = rappiL.reduce((s, l) => s + l.cantidad * l._c, 0);

  /* ---------- 6. Proyección de la próxima semana ---------- */
  const cuentaDia = [0, 0, 0, 0, 0, 0, 0];
  dias.forEach((f) => cuentaDia[diaSemana(f)]++);
  const inicioProy = dias.length ? sumarDias(dias[dias.length - 1], 1) : null;
  const fechasProy = inicioProy ? [0, 1, 2, 3, 4, 5, 6].map((n) => sumarDias(inicioProy, n)) : [];
  const porPlato = agrupar(
    alim.filter((l) => !NO_MENU.includes(norm(l.subgrupo)) && l.neto > 0),
    (l) => (l._r ? "r" + l._r.id : norm(sinR(l.producto))),
    (l) => ({ plato: l._r ? l._r.nom : sinR(l.producto), receta: l._r, porDia: [0, 0, 0, 0, 0, 0, 0], cant: 0, venta: 0 }),
    (a, l) => {
      a.porDia[diaSemana(l.fecha)] += l.cantidad;
      a.cant += l.cantidad;
      a.venta += l.neto;
    },
  );
  const promedioDiario = (a) => (dias.length ? a.cant / dias.length : 0);
  porPlato.forEach((a) => {
    a.proy = fechasProy.map((f) => {
      const d = diaSemana(f);
      return cuentaDia[d] ? a.porDia[d] / cuentaDia[d] : promedioDiario(a);
    });
    a.proy = a.proy.map((v) => Math.round(v));
    a.proyTotal = a.proy.reduce((s, v) => s + v, 0);
    a.ventaProy = a.proyTotal * (a.cant ? a.venta / a.cant : 0);
  });
  porPlato.sort((a, b) => b.proyTotal - a.proyTotal);

  const insumosProy = new Map();
  porPlato.forEach((a) => {
    if (!a.receta) return;
    a.receta.ing.forEach((g) => {
      const i = insById(g.iid);
      if (!i) return;
      const cant = (a.proyTotal * g.q) / Math.max(0.0001, 1 - i.merma / 100);
      const x = insumosProy.get(i.id) || { insumo: i.nom, uni: i.uni, cantidad: 0, costo: 0 };
      x.cantidad += cant;
      x.costo += cant * i.costo;
      insumosProy.set(i.id, x);
    });
  });
  const compraProy = [...insumosProy.values()].sort((a, b) => b.costo - a.costo);

  return {
    dias,
    total,
    porCanal,
    ventaAlim,
    costoTeo,
    ventaSinCosto,
    compraAlim,
    fcTeo,
    fcReal,
    comprasFamilia,
    menu,
    corteMix,
    corteMargen,
    sobreAlerta,
    alerta,
    sinCosto,
    alzas,
    mozos,
    rappi,
    rappiVenta,
    rappiCosto,
    fechasProy,
    porPlato,
    compraProy,
  };
}

/* =====================================================================
   PANTALLA
   ===================================================================== */
const CLASES = {
  Estrella: { css: "ok", que: "Mantener: cuidar porción y presentación" },
  Caballo: { css: "warn", que: "Se vende mucho pero deja poco: revisar costo o subir precio" },
  Puzzle: { css: "info", que: "Deja buen margen pero se vende poco: darle visibilidad" },
  Perro: { css: "bad", que: "Poca venta y poco margen: reformular o sacar de carta" },
};

function tabla(cabeceras, filas, clase = "") {
  return `<div class="scroll"><table class="${clase}"><thead><tr>${cabeceras
    .map((c) => `<th${c.n ? ' class="n"' : ""}>${c.t}</th>`)
    .join("")}</tr></thead><tbody>${filas.join("")}</tbody></table></div>`;
}

function renderReportes() {
  const box = $("#repBox");
  if (!DATOS.ventas.length) {
    box.innerHTML = `<div class="card"><h2>Aún no hay ventas cargadas</h2><p class="hint">Sube tus paloteos diarios de Inforest en la pestaña Carga de datos y los reportes aparecen aquí.</p></div>`;
    return;
  }
  const R = calcularReportes();
  const brecha = R.fcReal - R.fcTeo;
  const cuenta = (c) => R.menu.filter((x) => x.clase === c);
  const ventaProy = R.porPlato.reduce((s, a) => s + a.ventaProy, 0);
  const costoProy = R.compraProy.reduce((s, a) => s + a.costo, 0);

  box.innerHTML = `
  <div class="card row"><div><h2>Reportes del ${fmtD(S.cfg.ini)} al ${fmtD(S.cfg.fin)}</h2>
    <p class="hint">${R.dias.length} días con venta. Cambia las fechas en el encabezado para ver otro período.</p></div>
    <span class="sp"></span><button class="btn" id="expRep">Exportar reportes a Excel</button></div>

  <nav class="subnav">
    <a href="#r-fc">Food cost</a><a href="#r-menu">Ingeniería de menú</a><a href="#r-alertas">Alertas</a>
    <a href="#r-mozos">Mozos</a><a href="#r-rappi">Rappi</a><a href="#r-proy">Proyección</a>
  </nav>

  <div class="kpis">
    ${[
      ["Venta neta", money(R.total), R.porCanal.map((c) => `${c.canal} ${money(c.venta)}`).join(" · ")],
      ["Food cost teórico", pct(R.fcTeo), "alimentos, según recetas"],
      ["Food cost con compras", pct(R.fcReal), `${brecha >= 0 ? "+" : ""}${(brecha * 100).toFixed(1)} pts vs teórico`],
      ["Platos estrella", cuenta("Estrella").length, `de ${R.menu.length} platos analizados`],
      ["Vendidos sin costo", R.sinCosto.length, `${money(R.sinCosto.reduce((s, x) => s + x.venta, 0))} no descargan inventario`],
      ["Venta proyectada", money(ventaProy), `próximos 7 días · insumos ${money(costoProy)}`],
    ]
      .map(([a, b, c]) => `<div class="kpi"><small>${a}</small><b>${b}</b><span>${c}</span></div>`)
      .join("")}
  </div>

  <section class="card" id="r-fc">
    <h2>Food cost teórico vs food cost con compras</h2>
    <p class="hint">El teórico sale de lo vendido por sus recetas. El de compras toma lo que entró a almacén como si se hubiera consumido, porque no hay inventario inicial y final del período.</p>
    <div class="cmp">
      <div><small>Venta neta de alimentos</small><b>${money(R.ventaAlim)}</b></div>
      <div><small>Costo según recetas</small><b>${money(R.costoTeo)}</b><span class="tag ok">${pct(R.fcTeo)}</span></div>
      <div><small>Compras de alimentos</small><b>${money(R.compraAlim)}</b><span class="tag ${brecha > 0.01 ? "bad" : "ok"}">${pct(R.fcReal)}</span></div>
      <div><small>Diferencia</small><b>${money(R.compraAlim - R.costoTeo)}</b><span>${(brecha * 100).toFixed(1)} puntos</span></div>
    </div>
    <p class="hint">Qué puede explicar la diferencia: stock que quedó en almacén, mermas, comida de personal, platos vendidos sin receta (${money(R.ventaSinCosto)} de venta sin costo) y recetas que no reflejan la porción real.</p>
    <h3 style="margin:14px 0 8px">Compras por familia</h3>
    ${tabla(
      [{ t: "Familia" }, { t: "Subfamilia" }, { t: "Líneas", n: 1 }, { t: "Monto", n: 1 }],
      R.comprasFamilia.slice(0, 15).map((c) => `<tr><td>${esc(c.familia)}</td><td>${esc(c.subfamilia)}</td><td class="n">${c.lineas}</td><td class="n">${money(c.monto)}</td></tr>`),
    )}
  </section>

  <section class="card" id="r-menu">
    <h2>Ingeniería de menú</h2>
    <p class="hint">Platos de salón con receta. Popular = se vende más del 70% del promedio (mix ≥ ${pct(R.corteMix)}). Rentable = deja más margen que el promedio (≥ ${money(R.corteMargen)} por plato).</p>
    <div class="cuad">
      ${["Estrella", "Caballo", "Puzzle", "Perro"]
        .map(
          (c) => `<div class="q ${CLASES[c].css}"><div class="row"><b>${c}</b><span class="sp"></span><span class="tag ${CLASES[c].css}">${cuenta(c).length}</span></div>
        <p class="hint" style="margin:4px 0 8px">${CLASES[c].que}</p>
        <ul>${cuenta(c)
          .slice(0, 6)
          .map((x) => `<li>${esc(x.plato)} <span class="hint">${f2(x.cant)} u</span></li>`)
          .join("")}</ul></div>`,
        )
        .join("")}
    </div>
    ${tabla(
      [{ t: "Plato" }, { t: "Vendidos", n: 1 }, { t: "Mix", n: 1 }, { t: "Precio neto", n: 1 }, { t: "Costo", n: 1 }, { t: "FC", n: 1 }, { t: "Margen", n: 1 }, { t: "Margen total", n: 1 }, { t: "Clase" }],
      R.menu.map(
        (x) => `<tr><td>${esc(x.plato)}</td><td class="n">${f2(x.cant)}</td><td class="n">${pct(x.mix)}</td><td class="n">${money(x.precioNeto)}</td><td class="n">${money(x.costo)}</td><td class="n">${pct(x.fc)}</td><td class="n">${money(x.margen)}</td><td class="n">${money(x.margenTotal)}</td><td><span class="tag ${CLASES[x.clase].css}">${x.clase}</span></td></tr>`,
      ),
    )}
  </section>

  <section class="card" id="r-alertas">
    <h2>Alertas de control</h2>
    <h3 style="margin:12px 0 8px">Platos sobre ${pct(R.alerta)} de food cost <span class="tag ${R.sobreAlerta.length ? "bad" : "ok"}">${R.sobreAlerta.length}</span></h3>
    ${
      R.sobreAlerta.length
        ? tabla(
            [{ t: "Plato" }, { t: "Vendidos", n: 1 }, { t: "Precio neto", n: 1 }, { t: "Costo", n: 1 }, { t: "FC", n: 1 }, { t: `Precio sugerido (${S.cfg.meta}%)`, n: 1 }],
            R.sobreAlerta.map((x) => `<tr><td>${esc(x.plato)}</td><td class="n">${f2(x.cant)}</td><td class="n">${money(x.precioNeto)}</td><td class="n">${money(x.costo)}</td><td class="n loss">${pct(x.fc)}</td><td class="n">${money(x.sugerido)}</td></tr>`),
          )
        : '<p class="hint">Ningún plato pasa la alerta.</p>'
    }
    <h3 style="margin:18px 0 8px">Vendidos que no descargan inventario <span class="tag ${R.sinCosto.length ? "warn" : "ok"}">${R.sinCosto.length}</span></h3>
    <p class="hint">Se vendieron pero no tienen receta o su receta está en costo 0, así que no restan insumos del almacén.</p>
    ${tabla(
      [{ t: "Producto" }, { t: "Grupo" }, { t: "Canal" }, { t: "Situación" }, { t: "Vendidos", n: 1 }, { t: "Venta neta", n: 1 }],
      R.sinCosto.slice(0, 25).map((x) => `<tr><td>${esc(x.producto)}</td><td>${esc(x.grupo)}</td><td>${x.canal}</td><td>${x.receta ? '<span class="tag warn">Receta en costo 0</span>' : '<span class="tag bad">Sin receta</span>'}</td><td class="n">${f2(x.cant)}</td><td class="n">${money(x.venta)}</td></tr>`),
    )}
    <h3 style="margin:18px 0 8px">Insumos que subieron de precio en el período <span class="tag ${R.alzas.length ? "warn" : "ok"}">${R.alzas.length}</span></h3>
    ${
      R.alzas.length
        ? tabla(
            [{ t: "Insumo" }, { t: "Unidad" }, { t: "Primer precio", n: 1 }, { t: "Último precio", n: 1 }, { t: "Variación", n: 1 }, { t: "Impacto en el período", n: 1 }],
            R.alzas.slice(0, 15).map((a) => `<tr><td>${esc(a.articulo)}</td><td>${esc(a.unidad)}</td><td class="n">${money(a.primero)}</td><td class="n">${money(a.ultimo)}</td><td class="n loss">+${pct(a.variacion)}</td><td class="n">${money(a.impacto)}</td></tr>`),
          )
        : '<p class="hint">No hay alzas mayores a 5% entre compras del período.</p>'
    }
  </section>

  <section class="card" id="r-mozos">
    <h2>Venta por mozo</h2>
    ${tabla(
      [{ t: "Mozo" }, { t: "Venta neta", n: 1 }, { t: "Pedidos", n: 1 }, { t: "Comensales", n: 1 }, { t: "Ticket por pedido", n: 1 }, { t: "Por comensal", n: 1 }, { t: "% bebidas", n: 1 }],
      R.mozos.map((m) => `<tr><td>${esc(m.mozo === "SIN MOZO" ? "Delivery / sin mozo" : m.mozo)}</td><td class="n">${money(m.venta)}</td><td class="n">${f2(m.pedidos)}</td><td class="n">${f2(m.comensales)}</td><td class="n">${money(m.ticket)}</td><td class="n">${m.comensales ? money(m.porComensal) : "—"}</td><td class="n">${pct(m.pctBebidas)}</td></tr>`),
    )}
  </section>

  <section class="card" id="r-rappi">
    <h2>Rappi como canal</h2>
    <div class="cmp">
      <div><small>Venta neta Rappi</small><b>${money(R.rappiVenta)}</b><span>${pct(R.total ? R.rappiVenta / R.total : 0)} de la venta</span></div>
      <div><small>Costo según recetas</small><b>${money(R.rappiCosto)}</b><span class="tag ok">${pct(R.rappiVenta ? R.rappiCosto / R.rappiVenta : 0)}</span></div>
    </div>
    ${tabla(
      [{ t: "Producto" }, { t: "Vendidos", n: 1 }, { t: "Venta neta", n: 1 }, { t: "Precio neto Rappi", n: 1 }, { t: "Precio neto salón", n: 1 }, { t: "FC Rappi", n: 1 }],
      R.rappi.slice(0, 15).map((x) => `<tr><td>${esc(x.producto)}</td><td class="n">${f2(x.cant)}</td><td class="n">${money(x.venta)}</td><td class="n">${money(x.precioRappi)}</td><td class="n">${x.precioSalon ? money(x.precioSalon) : "—"}</td><td class="n">${x.fc === null ? '<span class="tag bad">Sin costo</span>' : pct(x.fc)}</td></tr>`),
    )}
  </section>

  <section class="card" id="r-proy">
    <h2>Proyección de la próxima semana</h2>
    <p class="hint">Promedio de lo vendido cada día de la semana en el período (${R.dias.length} días). Sirve para el pedido de cocina y para programar personal.</p>
    ${tabla(
      [{ t: "Plato" }, ...R.fechasProy.map((f) => ({ t: `${DIAS[diaSemana(f)]} ${f.slice(8)}`, n: 1 })), { t: "Total", n: 1 }, { t: "Venta", n: 1 }],
      R.porPlato.slice(0, 20).map((a) => `<tr><td>${esc(a.plato)}</td>${a.proy.map((v) => `<td class="n">${v}</td>`).join("")}<td class="n"><b>${a.proyTotal}</b></td><td class="n">${money(a.ventaProy)}</td></tr>`),
    )}
    <h3 style="margin:18px 0 8px">Insumos que vas a necesitar</h3>
    ${tabla(
      [{ t: "Insumo" }, { t: "Cantidad", n: 1 }, { t: "Unidad" }, { t: "Costo estimado", n: 1 }],
      R.compraProy.slice(0, 25).map((x) => `<tr><td>${esc(x.insumo)}</td><td class="n">${f2(x.cantidad)}</td><td>${x.uni}</td><td class="n">${money(x.costo)}</td></tr>`),
    )}
  </section>`;

  $("#expRep").onclick = () => exportarReportes(R);
}

/* =====================================================================
   EXPORTAR A EXCEL (una hoja por reporte)
   ===================================================================== */
function exportarReportes(R) {
  if (!window.XLSX) return toast("Lector de Excel no disponible");
  const r2 = (v) => +(+v || 0).toFixed(2);
  const p1 = (v) => +((+v || 0) * 100).toFixed(1);
  const hojas = [
    [
      "Resumen",
      [
        ["Indicador", "Valor"],
        ["Restaurante", S.cfg.nombre],
        ["Período", `${fmtD(S.cfg.ini)} al ${fmtD(S.cfg.fin)}`],
        ["Días con venta", R.dias.length],
        ["Venta neta total", r2(R.total)],
        ["Venta neta alimentos", r2(R.ventaAlim)],
        ["Costo según recetas", r2(R.costoTeo)],
        ["Food cost teórico %", p1(R.fcTeo)],
        ["Compras de alimentos", r2(R.compraAlim)],
        ["Food cost con compras %", p1(R.fcReal)],
        ["Venta sin costo", r2(R.ventaSinCosto)],
      ],
    ],
    [
      "Ingenieria de menu",
      [["Plato", "Vendidos", "Mix %", "Precio neto", "Costo", "FC %", "Margen", "Margen total", "Clase"]].concat(
        R.menu.map((x) => [x.plato, r2(x.cant), p1(x.mix), r2(x.precioNeto), r2(x.costo), p1(x.fc), r2(x.margen), r2(x.margenTotal), x.clase]),
      ),
    ],
    [
      "Sobre alerta",
      [["Plato", "Vendidos", "Precio neto", "Costo", "FC %", "Precio sugerido"]].concat(
        R.sobreAlerta.map((x) => [x.plato, r2(x.cant), r2(x.precioNeto), r2(x.costo), p1(x.fc), x.sugerido]),
      ),
    ],
    [
      "Sin descarga",
      [["Producto", "Grupo", "Canal", "Situación", "Vendidos", "Venta neta"]].concat(
        R.sinCosto.map((x) => [x.producto, x.grupo, x.canal, x.receta ? "Receta en costo 0" : "Sin receta", r2(x.cant), r2(x.venta)]),
      ),
    ],
    [
      "Alzas de compra",
      [["Insumo", "Unidad", "Primer precio", "Último precio", "Variación %", "Impacto"]].concat(
        R.alzas.map((a) => [a.articulo, a.unidad, r2(a.primero), r2(a.ultimo), p1(a.variacion), r2(a.impacto)]),
      ),
    ],
    [
      "Mozos",
      [["Mozo", "Venta neta", "Pedidos", "Comensales", "Ticket por pedido", "Por comensal", "% bebidas"]].concat(
        R.mozos.map((m) => [m.mozo, r2(m.venta), m.pedidos, m.comensales, r2(m.ticket), r2(m.porComensal), p1(m.pctBebidas)]),
      ),
    ],
    [
      "Rappi",
      [["Producto", "Vendidos", "Venta neta", "Precio neto Rappi", "Precio neto salón", "FC %"]].concat(
        R.rappi.map((x) => [x.producto, r2(x.cant), r2(x.venta), r2(x.precioRappi), x.precioSalon ? r2(x.precioSalon) : "", x.fc === null ? "" : p1(x.fc)]),
      ),
    ],
    [
      "Proyeccion platos",
      [["Plato", ...R.fechasProy.map((f) => `${DIAS[diaSemana(f)]} ${fmtD(f)}`), "Total", "Venta"]].concat(
        R.porPlato.map((a) => [a.plato, ...a.proy, a.proyTotal, r2(a.ventaProy)]),
      ),
    ],
    [
      "Proyeccion insumos",
      [["Insumo", "Cantidad", "Unidad", "Costo estimado"]].concat(R.compraProy.map((x) => [x.insumo, +x.cantidad.toFixed(3), x.uni, r2(x.costo)])),
    ],
  ];
  saveFile(`reportes_${S.cfg.ini}_${S.cfg.fin}.xlsx`, wbBlob(hojas));
}
