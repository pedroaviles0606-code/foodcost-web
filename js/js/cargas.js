/* =====================================================================
   CARGAS: ventas, mozos y compras
   ---------------------------------------------------------------------
   DATOS guarda en memoria lo que viene de los reportes de Inforest.
   - Se guarda en el navegador (localStorage) como respaldo.
   - Si hay sesión en Supabase, también se guarda en la nube.
   Cada vez que subes un paloteo de un día, se reemplaza ese día
   (no se duplica). Con las compras pasa igual con su rango de fechas.
   ===================================================================== */
const DATOS = { ventas: [], mozos: [], compras: [] };
const claveDatos = () => (user ? KEY + "-datos-" + user.id : KEY + "-datos");

function guardarDatosLocal() {
  try {
    localStorage.setItem(claveDatos(), JSON.stringify(DATOS));
  } catch (e) {}
}
function leerDatosLocal() {
  try {
    const d = JSON.parse(localStorage.getItem(claveDatos()) || "null");
    if (d && d.ventas) Object.assign(DATOS, d);
  } catch (e) {}
}

/* ---------- nube ---------- */
async function leerTodo(tabla, orden) {
  let out = [];
  for (let k = 0; ; k += 1000) {
    const d = chk(await sb.from(tabla).select("*").order(orden).order("id").range(k, k + 999));
    out = out.concat(d);
    if (d.length < 1000) break;
  }
  return out;
}

async function cargarDatosNube() {
  const [v, m, c] = await Promise.all([leerTodo("ventas", "fecha"), leerTodo("mozo_dia", "fecha"), leerTodo("compras", "fecha")]);
  DATOS.ventas = v
    .filter((x) => x.fecha && x.mozo)
    .map((x) => ({
      fecha: x.fecha,
      mozo: x.mozo,
      grupo: x.grupo || "",
      subgrupo: x.subgrupo || "",
      producto: x.plato,
      canal: x.canal,
      cantidad: Number(x.unidades),
      neto: Number(x.neto),
    }));
  DATOS.mozos = m.map((x) => ({ fecha: x.fecha, mozo: x.mozo, comensales: Number(x.comensales), pedidos: Number(x.pedidos) }));
  DATOS.compras = c.map((x) => ({
    fecha: x.fecha,
    familia: x.familia || "",
    subfamilia: x.subfamilia || "",
    codigo: x.codigo || "",
    articulo: x.articulo,
    unidad: x.unidad || "",
    precio: Number(x.precio),
    cantidad: Number(x.cantidad),
    subtotal: Number(x.subtotal),
    proveedor: x.proveedor || "",
    documento: x.documento || "",
  }));
  guardarDatosLocal();
}

/* Reemplaza en la nube los días de venta indicados */
async function subirVentas(fechas, lineas, mozos, archivo) {
  if (!sb || !user) return;
  const f = [...fechas];
  await inChunks(f, 50, (c) => sb.from("ventas").delete().in("fecha", c));
  await inChunks(f, 50, (c) => sb.from("mozo_dia").delete().in("fecha", c));
  const lote = new Date().toISOString();
  await inChunks(
    lineas.map((l) => ({
      user_id: user.id,
      lote,
      archivo,
      fecha: l.fecha,
      mozo: l.mozo,
      grupo: l.grupo,
      subgrupo: l.subgrupo,
      plato: l.producto,
      canal: l.canal,
      unidades: l.cantidad,
      neto: +l.neto.toFixed(2),
    })),
    500,
    (c) => sb.from("ventas").insert(c),
  );
  await inChunks(
    mozos.map((m) => ({ user_id: user.id, fecha: m.fecha, mozo: m.mozo, comensales: m.comensales, pedidos: m.pedidos })),
    500,
    (c) => sb.from("mozo_dia").upsert(c, { onConflict: "user_id,fecha,mozo" }),
  );
}

/* Reemplaza en la nube las compras del rango de fechas */
async function subirCompras(desde, hasta, lineas, archivo) {
  if (!sb || !user) return;
  chk(await sb.from("compras").delete().gte("fecha", desde).lte("fecha", hasta));
  await inChunks(
    lineas.map((c) => ({ user_id: user.id, archivo, ...c, subtotal: +c.subtotal.toFixed(2) })),
    500,
    (x) => sb.from("compras").insert(x),
  );
}

async function borrarDatosNube() {
  if (!sb || !user) return;
  for (const t of ["ventas", "mozo_dia", "compras"]) chk(await sb.from(t).delete().eq("user_id", user.id));
}

/* =====================================================================
   SUBIR ARCHIVOS DE INFOREST (se detecta solo qué reporte es)
   ===================================================================== */
let pendInf = null;

/* Lee un archivo del computador y devuelve sus bytes */
function leerBuffer(f) {
  return new Promise((ok, mal) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result);
    fr.onerror = () => mal(fr.error);
    fr.readAsArrayBuffer(f);
  });
}

async function leerArchivosInforest(archivos) {
  if (!window.XLSX) return toast("No se pudo cargar el lector de Excel. Revisa tu conexión.");
  const res = { paloteos: [], compras: [], recetas: null, desconocidos: [] };
  for (const f of archivos) {
    try {
      const buf = new Uint8Array(await leerBuffer(f));
      const libro = XLSX.read(buf, { type: "array", cellDates: false });
      const tipo = detectarTipo(libro);
      if (tipo === "maestro") {
        const hoja = libro.Sheets["INSUMOS"] || libro.Sheets[libro.SheetNames[0]];
        prevIns(filasDeHoja(hoja), f.name);
        return;
      }
      if (tipo === "paloteo") {
        const p = leerPaloteo(filasDeHoja(libro.Sheets[libro.SheetNames[0]]));
        if (p.fecha) res.paloteos.push({ ...p, archivo: f.name });
        else res.desconocidos.push(f.name);
      } else if (tipo === "compras") {
        res.compras.push({ archivo: f.name, lineas: leerCompras(filasDeHoja(libro.Sheets[libro.SheetNames[0]])) });
      } else if (tipo === "recetas") {
        const hoja = libro.Sheets["MODELO COSTO INFOREST"] || libro.Sheets[libro.SheetNames[0]];
        res.recetas = { archivo: f.name, lista: leerRecetas(filasDeHoja(hoja), codigosEnCarta(libro)) };
      } else res.desconocidos.push(f.name);
    } catch (e) {
      console.error(e);
      res.desconocidos.push(f.name);
    }
  }
  mostrarPrevInforest(res);
}

function mostrarPrevInforest(res) {
  pendInf = res;
  const partes = [];

  if (res.paloteos.length) {
    res.paloteos.sort((a, b) => a.fecha.localeCompare(b.fecha));
    const filas = res.paloteos.map((p) => {
      const venta = p.lineas.reduce((s, l) => s + l.neto, 0);
      const ya = DATOS.ventas.some((l) => l.fecha === p.fecha);
      return `<tr><td>${fmtD(p.fecha)}</td><td class="n">${p.lineas.length}</td><td class="n">${p.mozos.length}</td><td class="n">${money(venta)}</td><td>${ya ? '<span class="tag warn">Reemplaza el día</span>' : '<span class="tag ok">Nuevo</span>'}</td></tr>`;
    });
    partes.push(`<h3>Paloteos de venta: ${res.paloteos.length} día(s)</h3>
      <div class="scroll" style="max-height:260px;margin-top:8px"><table><thead><tr><th>Día</th><th class="n">Líneas</th><th class="n">Mozos</th><th class="n">Venta neta</th><th>Estado</th></tr></thead><tbody>${filas.join("")}</tbody></table></div>`);
  }

  if (res.compras.length) {
    const l = res.compras.flatMap((c) => c.lineas);
    const fechas = l.map((x) => x.fecha).sort();
    partes.push(`<h3 style="margin-top:16px">Compras: ${l.length} líneas del ${fmtD(fechas[0])} al ${fmtD(fechas[fechas.length - 1])}</h3>
      <p class="hint">Total sin IGV ${money(l.reduce((s, x) => s + x.subtotal, 0))}. Reemplaza las compras que ya tengas en esas fechas.</p>`);
  }

  if (res.recetas) {
    const lista = res.recetas.lista;
    const fams = [...new Set(lista.map((r) => r.grupo))].sort();
    const hayCarta = lista.some((r) => r.enCarta !== null);
    const sinCosto = lista.filter((r) => !r.ing.some((g) => g.precio > 0)).length;
    partes.push(`<h3 style="margin-top:16px">Recetas: ${lista.length} en el archivo</h3>
      <p class="hint">${sinCosto} sin costo. Marca qué grupos quieres cargar.</p>
      ${hayCarta ? '<label class="chk" style="min-height:32px"><input type="checkbox" id="soloCarta" checked> Solo recetas que están en la carta</label>' : ""}
      <label class="chk" style="min-height:32px"><input type="checkbox" id="soloCosto"> Solo recetas con costo</label>
      <div class="row" style="gap:4px 14px;margin-top:6px">${fams
        .map((f) => `<label class="chk" style="min-height:30px;font-size:14px"><input type="checkbox" class="recfam" value="${esc(f)}"${/adicional/i.test(f) ? "" : " checked"}> ${esc(f)} (${lista.filter((r) => r.grupo === f).length})</label>`)
        .join("")}</div>
      <label style="font-size:14px;display:block;margin-top:8px">Si la receta ya existe
        <select id="recMode" style="width:auto"><option value="upd">Actualizarla</option><option value="rep">Borrar todas mis recetas y cargar estas</option></select></label>`);
  }

  if (res.desconocidos.length)
    partes.push(`<p class="loss" style="margin-top:12px">No reconocí: ${res.desconocidos.map(esc).join(", ")}. Sube el reporte tal como sale de Inforest.</p>`);

  const hayAlgo = res.paloteos.length || res.compras.length || res.recetas;
  showPrev(`<h2>Archivos de Inforest</h2>${partes.join("")}
    ${hayAlgo ? '<div class="row" style="margin-top:14px"><button class="btn" id="okInf">Confirmar carga</button><button class="btn ghost" id="noInf">Cancelar</button></div>' : ""}`);
}

async function aplicarInforest() {
  const res = pendInf;
  if (!res) return;
  pendInf = null;
  const btn = $("#okInf");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Cargando…";
  }
  const avisos = [];
  try {
    if (res.recetas) avisos.push(aplicarRecetas(res.recetas.lista));

    if (res.paloteos.length) {
      const fechas = new Set(res.paloteos.map((p) => p.fecha));
      const lineas = res.paloteos.flatMap((p) => p.lineas);
      const mozos = res.paloteos.flatMap((p) => p.mozos);
      DATOS.ventas = DATOS.ventas.filter((l) => !fechas.has(l.fecha)).concat(lineas);
      DATOS.mozos = DATOS.mozos.filter((m) => !fechas.has(m.fecha)).concat(mozos);
      const todas = [...new Set(DATOS.ventas.map((l) => l.fecha))].sort();
      S.cfg.ini = todas[0];
      S.cfg.fin = todas[todas.length - 1];
      setSync("Guardando ventas…", "");
      await subirVentas(fechas, lineas, mozos, res.paloteos.map((p) => p.archivo).join(", "));
      avisos.push(`${fechas.size} día(s) de venta`);
    }

    for (const c of res.compras) {
      const f = c.lineas.map((x) => x.fecha).sort();
      if (!f.length) continue;
      const desde = f[0],
        hasta = f[f.length - 1];
      DATOS.compras = DATOS.compras.filter((x) => x.fecha < desde || x.fecha > hasta).concat(c.lineas);
      setSync("Guardando compras…", "");
      await subirCompras(desde, hasta, c.lineas, c.archivo);
      avisos.push(`${c.lineas.length} compras`);
    }
    guardarDatosLocal();
    unidadesDesdeVentas();
    loadCfg();
    renderAll();
    setSync("Guardado en la nube", "ok");
    toast("Cargado: " + avisos.join(", "));
  } catch (e) {
    console.error(e);
    setSync("Error al guardar", "err");
    toast("No se pudo guardar: " + (e.message || e));
  }
  $("#prev").innerHTML = "";
}

/* Convierte las recetas de Inforest en recetas de la app */
function aplicarRecetas(lista) {
  const grupos = new Set($$(".recfam").filter((c) => c.checked).map((c) => c.value));
  const soloCarta = $("#soloCarta")?.checked;
  const soloCosto = $("#soloCosto")?.checked;
  const elegidas = lista.filter(
    (r) => grupos.has(r.grupo) && (!soloCarta || r.enCarta) && (!soloCosto || r.ing.some((g) => g.precio > 0)),
  );
  if ($("#recMode")?.value === "rep") S.rec = [];

  const insPorNombre = new Map(S.ins.map((i) => [norm(i.nom), i]));
  const recPorCod = new Map(S.rec.filter((r) => r.cod).map((r) => [r.cod, r]));
  const recPorNom = new Map(S.rec.map((r) => [norm(r.nom), r]));
  let nuevasIns = 0;

  elegidas.forEach((x) => {
    let r = recPorCod.get(x.cod) || recPorNom.get(norm(x.nom));
    if (!r) {
      r = { id: nid(), ven: 0 };
      S.rec.push(r);
    }
    Object.assign(r, { cod: x.cod, nom: x.nom, cat: x.familia || x.area, precio: x.precio, ing: [] });
    x.ing.forEach((l) => {
      const b = lineaABase(l);
      // Mismo insumo, pero solo si la unidad coincide. Si la receta lo usa en otra
      // unidad (ej. Burrata por unidad y en el maestro por kilo), se usa una ficha aparte.
      let i = insPorNombre.get(norm(l.insumo));
      if (i && i.uni !== b.uni) {
        const nomU = `${l.insumo} (${b.uni})`;
        i = insPorNombre.get(norm(nomU));
        if (!i) {
          i = { id: nid(), cod: "", nom: nomU, cat: "Desde recetas", uni: b.uni, costo: +b.costo.toFixed(4), merma: 0, ii: 0, co: 0, fi: 0 };
          S.ins.push(i);
          insPorNombre.set(norm(nomU), i);
          nuevasIns++;
        }
      }
      if (!i) {
        i = { id: nid(), cod: "", nom: l.insumo, cat: "Desde recetas", uni: b.uni, costo: +b.costo.toFixed(4), merma: 0, ii: 0, co: 0, fi: 0 };
        S.ins.push(i);
        insPorNombre.set(norm(l.insumo), i);
        nuevasIns++;
      } else if (b.costo > 0) i.costo = +b.costo.toFixed(4); // el costeo trae el precio más reciente
      r.ing.push({ id: nid(), iid: i.id, q: +b.cantidad.toFixed(5) });
    });
  });
  S.sel = S.rec[0]?.id;
  return `${elegidas.length} recetas${nuevasIns ? ` y ${nuevasIns} insumos nuevos` : ""}`;
}

/* =====================================================================
   DEMO: Liberty Bistró (cuentas nuevas)
   ===================================================================== */
async function cargarDemo() {
  const r = await fetch("demo/liberty.json");
  if (!r.ok) throw new Error("demo no disponible");
  const d = await r.json();
  S = { cfg: d.cfg, ins: d.ins, rec: d.rec, sel: d.rec[0]?.id };
  DATOS.ventas = d.ventas;
  DATOS.mozos = d.mozos;
  DATOS.compras = d.compras;
  unidadesDesdeVentas();
  guardarDatosLocal();
}

async function subirDemoNube() {
  if (!sb || !user) return;
  await borrarDatosNube();
  await subirVentas(new Set(DATOS.ventas.map((l) => l.fecha)), DATOS.ventas, DATOS.mozos, "demo");
  const f = DATOS.compras.map((x) => x.fecha).sort();
  if (f.length) await subirCompras(f[0], f[f.length - 1], DATOS.compras, "demo");
}

/* ---------- botones de la zona de carga ---------- */
const zonaInf = $("#dropInf");
zonaInf.onclick = () => $("#fInf").click();
zonaInf.ondragover = (e) => {
  e.preventDefault();
  zonaInf.classList.add("over");
};
zonaInf.ondragleave = () => zonaInf.classList.remove("over");
zonaInf.ondrop = (e) => {
  e.preventDefault();
  zonaInf.classList.remove("over");
  leerArchivosInforest([...e.dataTransfer.files]);
};
$("#fInf").onchange = (e) => {
  leerArchivosInforest([...e.target.files]);
  e.target.value = "";
};
$("#prev").addEventListener("click", (e) => {
  if (e.target.id === "okInf") aplicarInforest();
  if (e.target.id === "noInf") {
    pendInf = null;
    $("#prev").innerHTML = "";
  }
});

/* Muestra qué datos hay cargados */
function renderEstadoDatos() {
  const el = $("#estadoDatos");
  if (!el) return;
  const dias = [...new Set(DATOS.ventas.map((l) => l.fecha))].sort();
  const fc = DATOS.compras.map((c) => c.fecha).sort();
  el.innerHTML = [
    ["Insumos", f2(S.ins.length)],
    ["Recetas", f2(S.rec.length)],
    ["Días de venta", dias.length ? `${dias.length} (${fmtD(dias[0])} al ${fmtD(dias[dias.length - 1])})` : "—"],
    ["Compras", fc.length ? `${fc.length} líneas (${fmtD(fc[0])} al ${fmtD(fc[fc.length - 1])})` : "—"],
  ]
    .map(([a, b]) => `<div class="kpi"><small>${a}</small><b style="font-size:18px;white-space:normal">${b}</b></div>`)
    .join("");
}
