/* =====================================================================
   APP PRINCIPAL
   ---------------------------------------------------------------------
   Orden de este archivo (busca cada título con Ctrl + F):
     1. ayudas            formato de soles, porcentajes y nombres
     2. datos de ejemplo  seed() arma un restaurante de prueba
     3. base de datos     leer y guardar en Supabase
     4. cálculos          costo de receta, neto sin impuestos, food cost
     5. pantallas         resumen, insumos, recetas, ventas
     6. excel             leer plantillas propias y exportar
     7. configuración     encabezado verde (meta, IGV, recargo, fechas)
     8. sesión            ingresar, crear cuenta, salir

   La variable S guarda todo lo que ves en pantalla:
     S.cfg  configuración     S.ins  insumos     S.rec  recetas
   ===================================================================== */

/* ---------- 1. ayudas ---------- */
const $ = (s) => document.querySelector(s),
  $$ = (s) => [...document.querySelectorAll(s)];
const esc = (s) =>
  String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
const num = (v) => {
  const n = parseFloat(String(v).replace(",", "."));
  return isFinite(n) ? n : 0;
};
const money = (v) =>
  "S/ " + (v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pct = (v) => (isFinite(v) ? (v * 100).toFixed(1) + "%" : "—");
const f2 = (v) => (v || 0).toLocaleString("en-US", { maximumFractionDigits: 2 });
const norm = (s) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
let uid = 1;
const nid = () =>
  crypto.randomUUID
    ? crypto.randomUUID()
    : "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (c) =>
        (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16),
      );
const isUuid = (v) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(v));
function fixIds(st) {
  const m = new Map();
  const f = (v) => {
    if (isUuid(v)) return v;
    if (!m.has(v)) m.set(v, nid());
    return m.get(v);
  };
  st.ins.forEach((i) => (i.id = f(i.id)));
  st.rec.forEach((r) => {
    r.id = f(r.id);
    r.ing.forEach((g) => {
      g.iid = f(g.iid);
      if (!isUuid(g.id)) g.id = nid();
    });
  });
  if (st.sel != null) st.sel = f(st.sel);
  return st;
}
const KEY = "foodcost-mesacriolla-v1";
const CATS = ["Carnes", "Pescados", "Aves", "Verduras", "Abarrotes", "Lácteos", "Bebidas", "Otros"];
const UNIS = ["kg", "lt", "und"];

/* ---------- 2. datos de ejemplo ---------- */
function seed() {
  uid = 1;
  const I = [
    ["I001", "Lomo fino de res", "Carnes", "kg", 58, 15],
    ["I002", "Filete de corvina", "Pescados", "kg", 38, 10],
    ["I003", "Pechuga de pollo", "Aves", "kg", 16, 12],
    ["I004", "Cebolla roja", "Verduras", "kg", 3.5, 10],
    ["I005", "Tomate", "Verduras", "kg", 3.8, 8],
    ["I006", "Limón", "Verduras", "kg", 6, 45],
    ["I007", "Ají amarillo", "Verduras", "kg", 8, 20],
    ["I008", "Ají limo", "Verduras", "kg", 10, 15],
    ["I009", "Papa amarilla", "Verduras", "kg", 4.5, 18],
    ["I010", "Camote", "Verduras", "kg", 3, 15],
    ["I011", "Choclo", "Verduras", "kg", 4, 50],
    ["I012", "Arroz", "Abarrotes", "kg", 4.2, 0],
    ["I013", "Pan de molde", "Abarrotes", "kg", 9, 5],
    ["I014", "Leche evaporada", "Lácteos", "lt", 6.5, 0],
    ["I015", "Sillao", "Abarrotes", "lt", 12, 0],
    ["I016", "Aceite vegetal", "Abarrotes", "lt", 9, 0],
    ["I017", "Maíz morado", "Abarrotes", "kg", 7, 0],
    ["I018", "Azúcar", "Abarrotes", "kg", 4, 0],
  ].map(([cod, nom, cat, uni, costo, merma]) => ({
    id: nid(),
    cod,
    nom,
    cat,
    uni,
    costo,
    merma,
    ii: 0,
    co: 0,
    fi: 0,
  }));
  const by = (c) => I.find((i) => i.cod === c).id;
  const R = [
    [
      "P001",
      "Lomo saltado",
      "Fondo",
      52,
      320,
      [
        ["I001", 0.18],
        ["I004", 0.08],
        ["I005", 0.08],
        ["I009", 0.2],
        ["I012", 0.12],
        ["I015", 0.02],
        ["I016", 0.04],
      ],
    ],
    [
      "P002",
      "Ceviche clásico",
      "Entrada",
      45,
      280,
      [
        ["I002", 0.2],
        ["I006", 0.12],
        ["I004", 0.06],
        ["I008", 0.01],
        ["I010", 0.1],
        ["I011", 0.1],
      ],
    ],
    [
      "P003",
      "Ají de gallina",
      "Fondo",
      38,
      190,
      [
        ["I003", 0.16],
        ["I007", 0.05],
        ["I013", 0.04],
        ["I014", 0.06],
        ["I009", 0.15],
        ["I012", 0.12],
        ["I016", 0.02],
      ],
    ],
    [
      "P004",
      "Chicha morada (vaso)",
      "Bebida",
      12,
      450,
      [
        ["I017", 0.05],
        ["I018", 0.04],
        ["I006", 0.02],
      ],
    ],
  ].map(([cod, nom, cat, precio, ven, ing]) => ({
    id: nid(),
    cod,
    nom,
    cat,
    precio,
    ven,
    ing: ing.map(([c, q]) => ({ id: nid(), iid: by(c), q })),
  }));
  const st = {
    cfg: {
      nombre: "Mesa Criolla – Cocina peruana",
      logo: null,
      meta: 30,
      igv: true,
      rc: 0,
      ini: "2026-10-01",
      fin: "2026-10-31",
    },
    ins: I,
    rec: R,
    sel: R[0].id,
    uid,
  };
  S = st;
  const extra = { I001: 1.1, I006: 1.15, I002: 1.07 };
  I.forEach((i) => {
    const t = teor(i),
      r = t * (extra[i.cod] || 1.02);
    i.ii = +(t * 0.25).toFixed(1);
    i.co = +Math.ceil(r * 1.05).toFixed(1);
    i.fi = +(i.ii + i.co - r).toFixed(1);
  });
  st.uid = uid;
  return st;
}
let S;
function localLoad(k) {
  try {
    const raw = localStorage.getItem(k);
    if (raw) {
      const d = JSON.parse(raw);
      if (d && d.ins && d.rec) {
        S = fixIds(d);
        return true;
      }
    }
  } catch (e) {}
  return false;
}
let saveT, cloudT;
function save() {
  clearTimeout(saveT);
  saveT = setTimeout(() => {
    try {
      localStorage.setItem(user ? KEY + "-" + user.id : KEY, JSON.stringify(S));
    } catch (e) {}
  }, 300);
  if (sb && user && ready) {
    clearTimeout(cloudT);
    setSync("Cambios sin guardar", "");
    cloudT = setTimeout(cloudSave, 1200);
  }
}

/* ---------- 3. base de datos (Supabase) ---------- */
const sb =
  SUPABASE_URL && SUPABASE_KEY && window.supabase
    ? window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY)
    : null;
let user = null,
  ready = false,
  saving = false,
  again = false;
let known = { ins: new Set(), rec: new Set(), ing: new Set() };
function setSync(t, c) {
  const e = document.getElementById("sync");
  if (e) {
    e.textContent = t;
    e.style.background = c === "err" ? "#C0502E" : c === "ok" ? "rgba(59,167,118,.45)" : "";
  }
}
function snapKnown() {
  known = {
    ins: new Set(S.ins.map((i) => i.id)),
    rec: new Set(S.rec.map((r) => r.id)),
    ing: new Set(S.rec.flatMap((r) => r.ing.map((g) => g.id))),
  };
}
const chk = (r) => {
  if (r.error) throw r.error;
  return r.data;
};
async function cloudLoad() {
  async function all(t, ord) {
    let out = [];
    for (let k = 0; ; k += 1000) {
      const d = chk(
        await sb
          .from(t)
          .select("*")
          .order(ord)
          .order("id")
          .range(k, k + 999),
      );
      out = out.concat(d);
      if (d.length < 1000) break;
    }
    return out;
  }
  const [c, ins, rec, ing] = await Promise.all([
    sb.from("configuracion").select("*").maybeSingle(),
    all("insumos", "nombre"),
    all("recetas", "codigo"),
    all("receta_ingredientes", "orden"),
  ]);
  const cfg = chk(c);
  known = { ins: new Set(), rec: new Set(), ing: new Set() };
  if (!ins.length && !rec.length) {
    // cuenta nueva: se carga la demo de Liberty Bistró
    try {
      await cargarDemo();
      setSync("Preparando la demo…", "");
      await subirDemoNube();
    } catch (e) {
      console.error(e);
      seed();
    }
    saved.clear();
    cfgSaved = "";
    await cloudSave();
    return;
  }
  S = {
    cfg: {
      nombre: cfg?.nombre || "Mi restaurante",
      logo: cfg?.logo || null,
      meta: Number(cfg?.meta_fc ?? 30),
      igv: cfg ? cfg.precios_con_igv : true,
      rc: Number(cfg?.recargo_consumo ?? 0),
      alerta: Number(cfg?.alerta_fc ?? 35),
      ini: cfg?.periodo_ini || "",
      fin: cfg?.periodo_fin || "",
    },
    ins: ins.map((x) => ({
      id: x.id,
      cod: x.codigo || "",
      nom: x.nombre,
      cat: x.categoria || "Otros",
      uni: x.unidad,
      costo: Number(x.costo),
      merma: Number(x.merma),
      ii: Number(x.inv_inicial),
      co: Number(x.compras),
      fi: Number(x.inv_final),
    })),
    rec: rec.map((x) => ({
      id: x.id,
      cod: x.codigo || "",
      nom: x.nombre,
      cat: x.categoria || "Fondo",
      precio: Number(x.precio),
      ven: Number(x.unidades_vendidas),
      ing: ing
        .filter((y) => y.receta_id === x.id)
        .map((y) => ({ id: y.id, iid: y.insumo_id, q: Number(y.cantidad) })),
    })),
  };
  S.sel = S.rec[0]?.id;
  markSaved(buildRows(user.id));
  cfgSaved = sig(cfgRow(user.id));
  await cargarDatosNube();
  unidadesDesdeVentas();
}
const saved = new Map();
let cfgSaved = "";
const sig = (o) => JSON.stringify(o);
function buildRows(u) {
  return {
    ins: S.ins.map((i) => ({
      id: i.id,
      user_id: u,
      codigo: i.cod,
      nombre: i.nom || "Sin nombre",
      categoria: i.cat,
      unidad: i.uni,
      costo: +(+i.costo).toFixed(4),
      merma: i.merma,
      inv_inicial: i.ii,
      compras: i.co,
      inv_final: i.fi,
    })),
    rec: S.rec.map((r) => ({
      id: r.id,
      user_id: u,
      codigo: r.cod,
      nombre: r.nom || "Sin nombre",
      categoria: r.cat,
      precio: r.precio,
      unidades_vendidas: r.ven,
    })),
    ing: S.rec.flatMap((r) =>
      r.ing
        .filter((g) => g.iid && insById(g.iid))
        .map((g, k) => ({
          id: g.id,
          user_id: u,
          receta_id: r.id,
          insumo_id: g.iid,
          cantidad: g.q,
          orden: k,
        })),
    ),
  };
}
const cfgRow = (u) => ({
  user_id: u,
  nombre: S.cfg.nombre,
  logo: S.cfg.logo,
  meta_fc: S.cfg.meta,
  precios_con_igv: S.cfg.igv,
  recargo_consumo: num(S.cfg.rc) || 0,
  alerta_fc: num(S.cfg.alerta) || 35,
  periodo_ini: S.cfg.ini || null,
  periodo_fin: S.cfg.fin || null,
});
function markSaved(b) {
  saved.clear();
  ["ins", "rec", "ing"].forEach((k) => b[k].forEach((x) => saved.set(k + x.id, sig(x))));
  known = {
    ins: new Set(b.ins.map((x) => x.id)),
    rec: new Set(b.rec.map((x) => x.id)),
    ing: new Set(b.ing.map((x) => x.id)),
  };
}
async function inChunks(arr, n, fn) {
  for (let k = 0; k < arr.length; k += n) chk(await fn(arr.slice(k, k + n)));
}
async function cloudSave() {
  if (!sb || !user) return;
  if (saving) {
    again = true;
    return;
  }
  saving = true;
  setSync("Guardando…", "");
  try {
    const u = user.id,
      b = buildRows(u);
    const gone = (set, arr) => {
      const ids = new Set(arr.map((x) => x.id));
      return [...set].filter((id) => !ids.has(id));
    };
    const dIng = gone(known.ing, b.ing),
      dRec = gone(known.rec, b.rec),
      dIns = gone(known.ins, b.ins);
    await inChunks(dIng, 100, (c) => sb.from("receta_ingredientes").delete().in("id", c));
    await inChunks(dRec, 100, (c) => sb.from("recetas").delete().in("id", c));
    await inChunks(dIns, 100, (c) => sb.from("insumos").delete().in("id", c));
    const cr = cfgRow(u);
    if (sig(cr) !== cfgSaved) {
      chk(
        await sb
          .from("configuracion")
          .upsert({ ...cr, updated_at: new Date().toISOString() }, { onConflict: "user_id" }),
      );
      cfgSaved = sig(cr);
    }
    const ch = (k) => b[k].filter((x) => saved.get(k + x.id) !== sig(x));
    const ci = ch("ins"),
      cre = ch("rec"),
      cg = ch("ing");
    if (ci.length > 300) setSync(`Guardando ${ci.length} insumos…`, "");
    await inChunks(ci, 500, (c) => sb.from("insumos").upsert(c));
    await inChunks(cre, 500, (c) => sb.from("recetas").upsert(c));
    await inChunks(cg, 500, (c) => sb.from("receta_ingredientes").upsert(c));
    markSaved(b);
    setSync("Guardado en la nube", "ok");
  } catch (e) {
    console.error(e);
    setSync("Error al guardar", "err");
    toast("No se pudo guardar en la base de datos: " + (e.message || e));
  }
  saving = false;
  if (again) {
    again = false;
    cloudSave();
  }
}
async function cloudVentas(lines, archivo) {
  if (!sb || !user || !lines.length) return;
  const lote = new Date().toISOString();
  const rows = lines.map((l) => ({
    user_id: user.id,
    lote,
    archivo,
    fecha: l.fecha || null,
    plato: l.plato,
    codigo: l.codigo || null,
    unidades: l.u,
    precio: l.precio || null,
    receta_id: l.rid || null,
  }));
  try {
    for (let k = 0; k < rows.length; k += 500) chk(await sb.from("ventas").insert(rows.slice(k, k + 500)));
    toast(`Se guardaron ${rows.length} líneas de venta en la base de datos`);
  } catch (e) {
    console.error(e);
    toast("Las ventas se aplicaron, pero no se guardó el detalle: " + (e.message || e));
  }
}

/* ---------- 4. cálculos ---------- */
/* Busca un insumo por su id. Usa un índice para que sea rápido con miles de insumos. */
let _idx = new Map(),
  _idxArr = null;
function insById(id) {
  if (_idxArr !== S.ins || _idx.size !== S.ins.length) {
    _idx = new Map(S.ins.map((i) => [i.id, i]));
    _idxArr = S.ins;
  }
  return _idx.get(id);
}
const apr = (i) => i.costo / Math.max(0.0001, 1 - i.merma / 100);
const impF = () => 1 + (S.cfg.igv ? 0.18 : 0) + (num(S.cfg.rc) || 0) / 100;
const neto = (p) => p / impF();
const netoTxt = () =>
  S.cfg.igv && num(S.cfg.rc)
    ? " (sin IGV ni recargo)"
    : S.cfg.igv
      ? " (sin IGV)"
      : num(S.cfg.rc)
        ? " (sin recargo)"
        : "";
const costoRec = (r) =>
  r.ing.reduce((s, g) => {
    const i = insById(g.iid);
    return s + (i ? g.q * apr(i) : 0);
  }, 0);
/* Consumo teórico de cada insumo = unidades vendidas x cantidad en la receta (con merma).
   Se calcula una vez para todos los insumos y se guarda hasta el siguiente cambio. */
let _teo = null;
function teorMapa() {
  const m = new Map();
  S.rec.forEach((r) => {
    if (!r.ven) return;
    r.ing.forEach((g) => {
      const i = insById(g.iid);
      if (i) m.set(i.id, (m.get(i.id) || 0) + (r.ven * g.q) / Math.max(0.0001, 1 - i.merma / 100));
    });
  });
  return m;
}
function teor(i) {
  if (!_teo) _teo = teorMapa();
  return _teo.get(i.id) || 0;
}
const real = (i) => i.ii + i.co - i.fi;
const meta = () => S.cfg.meta / 100;
function estado(fc) {
  const m = S.cfg.meta;
  const v = fc * 100;
  return v <= m ? ["ok", "En meta"] : v <= m + 3 ? ["warn", "Alerta"] : ["bad", "Crítico"];
}
function sugerido(c) {
  const g = (c / meta()) * impF();
  return Math.ceil(g * 2) / 2;
}
function totales() {
  let vn = 0,
    ct = 0;
  S.rec.forEach((r) => {
    vn += r.ven * neto(r.precio);
    ct += r.ven * costoRec(r);
  });
  let cr = 0;
  const dev = S.ins.map((i) => {
    const t = teor(i),
      re = real(i);
    cr += re * i.costo;
    return { i, t, re, du: re - t, ds: (re - t) * i.costo };
  });
  return { vn, ct, cr, fcT: vn ? ct / vn : 0, fcR: vn ? cr / vn : 0, dev };
}

/* ---------- 5. pantallas: marca ---------- */
/* Iniciales del restaurante para el logo por defecto: "Liberty Bistró" -> "LB" */
const iniciales = (n) =>
  String(n || "")
    .split(/[\s–-]+/)
    .filter((w) => /^[A-ZÁÉÍÓÚÑ]/.test(w))
    .slice(0, 2)
    .map((w) => w[0])
    .join("") || "FC";
function renderBrand() {
  $("#logo").innerHTML = S.cfg.logo
    ? `<img src="${S.cfg.logo}" alt="Logo">`
    : `<svg viewBox="0 0 56 56" width="50" height="50" aria-label="Logo"><circle cx="28" cy="28" r="26" fill="#1F5B43"/><path d="M14 40c4-12 24-12 28 0" fill="#F2B705"/><text x="28" y="31" text-anchor="middle" font-family="Poppins,sans-serif" font-weight="700" font-size="17" fill="#fff">${esc(iniciales(S.cfg.nombre))}</text></svg>`;
  $("#subName").textContent = S.cfg.nombre;
  document.title = "Food Cost – " + (S.cfg.nombre.split("–")[0].trim() || "Restaurante");
  $("#foot").textContent = `${S.cfg.nombre} · Período ${fmtD(S.cfg.ini)} al ${fmtD(S.cfg.fin)}`;
  $("#perTxt").textContent =
    `Período ${fmtD(S.cfg.ini)} al ${fmtD(S.cfg.fin)}. Puedes editar las unidades o cargarlas desde Excel.`;
}
const fmtD = (d) => (d ? d.split("-").reverse().join("/") : "—");

/* ---------- 5. pantallas: resumen ---------- */
function renderResumen() {
  const T = totales(),
    d = T.fcR - T.fcT,
    mb = T.vn - T.cr;
  const top = T.dev
    .filter((x) => x.ds > 0)
    .sort((a, b) => b.ds - a.ds)
    .slice(0, 3);
  $("#alerta").innerHTML =
    d > 0.01
      ? `<div class="alert bad"><b>El consumo real supera al teórico en ${(d * 100).toFixed(1)} puntos de food cost (${money(T.cr - T.ct)}).</b> Insumos con mayor desvío:<ol>${top.map((x) => `<li>${esc(x.i.nom)}: ${f2(x.du)} ${x.i.uni} de más, ${money(x.ds)}</li>`).join("")}</ol></div>`
      : `<div class="alert ok"><b>Consumo real dentro de lo esperado.</b> La diferencia con el teórico es de ${(d * 100).toFixed(1)} puntos.</div>`;
  const [ct] = estado(T.fcT),
    [cr] = estado(T.fcR);
  $("#kpis").innerHTML = [
    ["Ventas netas", money(T.vn), netoTxt().replace(/[()]/g, "").trim() || "precios netos"],
    ["FC teórico", pct(T.fcT), `<span class="tag ${ct}">meta ${S.cfg.meta}%</span>`],
    ["FC real", pct(T.fcR), `<span class="tag ${cr}">meta ${S.cfg.meta}%</span>`],
    ["Desviación real vs teórico", (d >= 0 ? "+" : "") + (d * 100).toFixed(1) + " pts", money(T.cr - T.ct)],
    ["Margen bruto real", money(mb), pct(T.vn ? mb / T.vn : 0) + " de la venta"],
  ]
    .map(([a, b, c]) => `<div class="kpi"><small>${a}</small><b>${b}</b><span>${c}</span></div>`)
    .join("");
  const vendidas = S.rec.filter((r) => r.ven > 0);
  const base = (vendidas.length ? vendidas : S.rec.slice(0, 30)).sort((a, b) => b.ven - a.ven);
  const rows = base.map((r) => {
    const c = costoRec(r),
      n = neto(r.precio);
    return { r, c, n, fc: n ? c / n : 0 };
  });
  const mx = Math.max(meta() * 1.5, ...rows.map((x) => x.fc * 1.1), 0.05);
  const col = { ok: "var(--ok)", warn: "var(--warn)", bad: "var(--bad)" };
  $("#bars").innerHTML =
    rows
      .slice(0, 15)
      .map((x) => {
        const [e] = estado(x.fc);
        return `<div class="bar"><span>${esc(x.r.nom)}</span><div class="track"><div class="fill" style="width:${Math.min(100, (x.fc / mx) * 100)}%;background:${col[e]}"></div><div class="meta" style="left:${(meta() / mx) * 100}%"></div></div><b class="n">${pct(x.fc)}</b></div>`;
      })
      .join("") || '<p class="hint">Aún no hay recetas.</p>';
  $("#metaLbl").textContent =
    `La línea punteada marca la meta de ${S.cfg.meta}%.` +
    (rows.length > 15 ? ` Se muestran los 15 platos más vendidos de ${rows.length}.` : "");
  $("#tDet").innerHTML = rows
    .map((x) => {
      const [e, l] = estado(x.fc);
      return `<tr><td>${esc(x.r.nom)}</td><td class="n">${f2(x.r.ven)}</td><td class="n">${money(x.r.precio)}</td><td class="n">${money(x.c)}</td><td class="n">${pct(x.fc)}</td><td class="n">${money(x.n - x.c)}</td><td class="n">${money((x.n - x.c) * x.r.ven)}</td><td class="n">${e === "ok" ? "—" : money(sugerido(x.c))}</td><td><span class="tag ${e}">${l}</span></td></tr>`;
    })
    .join("");
}

/* ---------- 5. pantallas: insumos ---------- */
const opts = (arr, v) => arr.map((a) => `<option${a === v ? " selected" : ""}>${a}</option>`).join("");
const LIM = 150;
function catsAll() {
  return [...new Set(S.ins.map((i) => i.cat).filter(Boolean))].sort((a, b) => a.localeCompare(b));
}
function renderIns() {
  const q = norm($("#qIns").value),
    fam = $("#fIns").value,
    cats = catsAll();
  $("#fIns").innerHTML =
    '<option value="">Todas las familias</option>' +
    cats.map((c) => `<option${c === fam ? " selected" : ""}>${esc(c)}</option>`).join("");
  $("#dlCat").innerHTML = [...new Set([...cats, ...CATS])].map((c) => `<option value="${esc(c)}">`).join("");
  const list = S.ins.filter((i) => (!fam || i.cat === fam) && (!q || norm(i.nom + " " + i.cod).includes(q)));
  $("#tIns").innerHTML = list
    .slice(0, LIM)
    .map(
      (i) =>
        `<tr data-id="${i.id}"><td><input type="text" data-f="cod" value="${esc(i.cod)}" style="width:80px"></td><td><input type="text" data-f="nom" value="${esc(i.nom)}" style="min-width:200px"></td><td><input type="text" data-f="cat" list="dlCat" value="${esc(i.cat)}" style="min-width:150px"></td><td><select data-f="uni" style="width:80px">${opts(UNIS, i.uni)}</select></td><td class="n"><input type="number" data-f="costo" step="0.01" min="0" value="${+(+i.costo).toFixed(4)}"></td><td class="n"><input type="number" data-f="merma" step="0.5" min="0" max="99" value="${i.merma}"></td><td class="n" id="ap${i.id}"></td><td><button class="btn del" data-del="${i.id}" aria-label="Eliminar">×</button></td></tr>`,
    )
    .join("");
  $("#insInfo").textContent =
    list.length > LIM
      ? `Mostrando ${LIM} de ${list.length} insumos. Usa el buscador o el filtro de familia para ver el resto.`
      : `${list.length} insumo(s)${q || fam ? " con este filtro" : ""}.`;
  updIns();
}
function updIns() {
  S.ins.forEach((i) => {
    const c = document.getElementById("ap" + i.id);
    if (c) c.textContent = money(apr(i)) + " / " + i.uni;
  });
}
let qT;
function filtIns() {
  clearTimeout(qT);
  qT = setTimeout(renderIns, 200);
}
$("#tIns").addEventListener("input", (e) => {
  const t = e.target,
    tr = t.closest("tr");
  if (!tr || !t.dataset.f) return;
  const i = insById(tr.dataset.id);
  const f = t.dataset.f;
  i[f] = ["costo", "merma"].includes(f)
    ? Math.min(f === "merma" ? 99 : 1e9, Math.max(0, num(t.value)))
    : t.value;
  if (f === "nom" || f === "uni") renderRec();
  refresh();
});
$("#tIns").addEventListener("click", (e) => {
  const b = e.target.closest("[data-del]");
  if (!b) return;
  const id = b.dataset.del,
    i = insById(id);
  let n = 0;
  S.rec.forEach((r) => {
    const l = r.ing.length;
    r.ing = r.ing.filter((g) => g.iid !== id);
    n += l - r.ing.length;
  });
  S.ins = S.ins.filter((x) => x.id !== id);
  renderAll();
  toast(`Se eliminó ${i.nom}` + (n ? ` y se quitó de ${n} receta(s)` : ""));
});
$("#addIns").onclick = () => {
  const i = {
    id: nid(),
    cod: "I" + String(S.ins.length + 1).padStart(3, "0"),
    nom: "Nuevo insumo",
    cat: "Otros",
    uni: "kg",
    costo: 0,
    merma: 0,
    ii: 0,
    co: 0,
    fi: 0,
  };
  S.ins.unshift(i);
  $("#qIns").value = "";
  $("#fIns").value = "";
  renderAll();
  const inp = $(`#tIns tr[data-id="${i.id}"] [data-f=nom]`);
  inp.focus();
  inp.select();
};
$("#qIns").addEventListener("input", filtIns);
$("#fIns").addEventListener("change", renderIns);

/* ---------- 5. pantallas: recetas ---------- */
const recSel = () => S.rec.find((r) => r.id === S.sel) || S.rec[0];
const lbl = (i) => (i ? `${i.nom} · ${i.uni}` : "");
const byLbl = (v) => {
  v = String(v).trim();
  return S.ins.find((i) => lbl(i) === v) || S.ins.find((i) => norm(i.nom) === norm(v));
};
function renderRec() {
  const r = recSel();
  $("#selRec").innerHTML = S.rec
    .map((x) => `<option value="${x.id}"${r && x.id === r.id ? " selected" : ""}>${esc(x.nom)}</option>`)
    .join("");
  if (!r) {
    $("#recBox").innerHTML = '<div class="card"><p class="hint">No hay recetas. Crea una nueva.</p></div>';
    return;
  }
  S.sel = r.id;
  const dl = `<datalist id="dlIns">${S.ins.map((i) => `<option value="${esc(lbl(i))}">`).join("")}</datalist>`;
  $("#recBox").innerHTML = `<div class="grid2"><div class="card">
   <div class="row" style="align-items:end"><label style="flex:2 1 220px;font-size:13px;color:var(--mute)">Nombre del plato<input type="text" data-r="nom" value="${esc(r.nom)}"></label>
   <label style="flex:1 1 100px;font-size:13px;color:var(--mute)">Código<input type="text" data-r="cod" value="${esc(r.cod)}"></label>
   <label style="flex:1 1 120px;font-size:13px;color:var(--mute)">Categoría<select data-r="cat">${opts(["Entrada", "Fondo", "Postre", "Bebida", "Otros"], r.cat)}</select></label>
   <label style="flex:1 1 120px;font-size:13px;color:var(--mute)">Precio de venta S/<input type="number" data-r="precio" step="0.5" min="0" value="${r.precio}"></label></div>
   <h3 style="margin:18px 0 10px">Ingredientes por porción</h3>
   <div class="scroll"><table><thead><tr><th>Insumo</th><th class="n">Cantidad</th><th class="n">Costo</th><th class="n">% del costo</th><th></th></tr></thead><tbody id="tIng">
   ${r.ing.map((g, k) => `<tr data-k="${k}"><td><input type="text" list="dlIns" data-g="iname" value="${esc(lbl(insById(g.iid)))}" placeholder="Escribe para buscar el insumo" style="min-width:240px"></td><td class="n"><input type="number" data-g="q" step="0.005" min="0" value="${g.q}"></td><td class="n" id="gc${k}"></td><td class="n" id="gp${k}"></td><td><button class="btn del" data-gdel="${k}" aria-label="Quitar">×</button></td></tr>`).join("")}
   </tbody></table></div>${dl}
   <button class="btn ghost sm" id="addIng" style="margin-top:12px"${S.ins.length ? "" : " disabled"}>+ Agregar ingrediente</button></div>
   <div class="card"><h3>Resultado por porción</h3><div class="big" id="rFc" style="margin:10px 0"></div><div id="rTag"></div><div class="res" id="rRes" style="margin-top:14px"></div></div></div>`;
  updRec();
}
function updRec() {
  const r = recSel();
  if (!r || !$("#rRes")) return;
  const c = costoRec(r),
    n = neto(r.precio),
    fc = n ? c / n : 0,
    [e, l] = estado(fc);
  r.ing.forEach((g, k) => {
    const i = insById(g.iid),
      gc = i ? g.q * apr(i) : 0;
    const a = document.getElementById("gc" + k),
      b = document.getElementById("gp" + k);
    if (a) a.textContent = money(gc);
    if (b) b.textContent = pct(c ? gc / c : 0);
  });
  $("#rFc").textContent = pct(fc) + " FC";
  $("#rTag").innerHTML = `<span class="tag ${e}">${l}</span>`;
  $("#rRes").innerHTML = [
    ["Costo por porción", money(c)],
    ["Precio de venta", money(r.precio)],
    ["Precio neto" + netoTxt(), money(n)],
    ["Margen", money(n - c) + " · " + pct(n ? (n - c) / n : 0)],
    [`Precio sugerido para ${S.cfg.meta}%`, money(sugerido(c))],
  ]
    .map(([a, b]) => `<div><span>${a}</span><b>${b}</b></div>`)
    .join("");
}
$("#recBox").addEventListener("input", (e) => {
  const t = e.target,
    r = recSel();
  if (t.dataset.r) {
    const f = t.dataset.r;
    r[f] = f === "precio" ? Math.max(0, num(t.value)) : t.value;
    if (f === "nom") {
      const o = $(`#selRec option[value="${r.id}"]`);
      if (o) o.textContent = r.nom;
    }
  } else if (t.dataset.g === "q") {
    r.ing[+t.closest("tr").dataset.k].q = Math.max(0, num(t.value));
  } else return;
  refresh();
});
$("#recBox").addEventListener("change", (e) => {
  const t = e.target;
  if (t.dataset.g !== "iname") return;
  const g = recSel().ing[+t.closest("tr").dataset.k],
    i = byLbl(t.value);
  if (i) {
    g.iid = i.id;
    t.value = lbl(i);
    refresh();
  } else {
    t.value = lbl(insById(g.iid));
    toast("No encontré ese insumo. Elige uno de la lista.");
  }
});
$("#recBox").addEventListener("click", (e) => {
  const r = recSel();
  if (e.target.id === "addIng") {
    r.ing.push({ id: nid(), iid: null, q: 0 });
    renderRec();
    refresh();
    const q = $$("#tIng [data-g=iname]").pop();
    q && q.focus();
  }
  const b = e.target.closest("[data-gdel]");
  if (b) {
    r.ing.splice(+b.dataset.gdel, 1);
    renderRec();
    refresh();
  }
});
$("#selRec").onchange = (e) => {
  S.sel = e.target.value;
  renderRec();
  save();
};
$("#newRec").onclick = () => {
  const r = {
    id: nid(),
    cod: "P" + String(S.rec.length + 1).padStart(3, "0"),
    nom: "Nuevo plato",
    cat: "Fondo",
    precio: 0,
    ven: 0,
    ing: [],
  };
  S.rec.push(r);
  S.sel = r.id;
  renderAll();
  const n = $("#recBox [data-r=nom]");
  n.focus();
  n.select();
};
$("#dupRec").onclick = () => {
  const o = recSel();
  if (!o) return;
  const r = {
    ...JSON.parse(JSON.stringify(o)),
    id: nid(),
    nom: o.nom + " (copia)",
    cod: o.cod + "-C",
    ven: 0,
  };
  r.ing.forEach((g) => (g.id = nid()));
  S.rec.push(r);
  S.sel = r.id;
  renderAll();
  toast("Receta duplicada");
};
$("#delRec").onclick = () => {
  const r = recSel();
  if (!r) return;
  S.rec = S.rec.filter((x) => x.id !== r.id);
  S.sel = S.rec[0]?.id;
  renderAll();
  toast(`Se eliminó ${r.nom}`);
};

/* ---------- 5. pantallas: ventas y consumo ---------- */
let sortLoss = false,
  sinRec = [],
  conAll = false;
function renderVen() {
  const recVen = S.rec.filter((r) => r.ven > 0);
  $("#tVen").innerHTML = (recVen.length ? recVen.sort((a, b) => b.ven - a.ven) : S.rec.slice(0, 60))
    .map(
      (r) =>
        `<tr data-id="${r.id}"><td>${esc(r.nom)}</td><td class="n"><input type="number" data-v="ven" min="0" step="1" value="${r.ven}"></td><td class="n">${money(r.precio)}</td><td class="n" id="vn${r.id}"></td><td class="n" id="vc${r.id}"></td></tr>`,
    )
    .join("");
  $("#sinReceta").innerHTML = sinRec.length
    ? `<div class="alert bad" style="margin:14px 0 0"><b>Platos sin receta en la última carga</b> (no entran al food cost): ${sinRec.map((x) => esc(x.n) + " (" + f2(x.u) + " u)").join(", ")}</div>`
    : "";
  const used = new Set(S.rec.flatMap((r) => r.ing.map((g) => g.iid)));
  const rel = S.ins.filter((i) => used.has(i.id) || i.ii || i.co || i.fi);
  let list = conAll ? [...S.ins] : rel;
  if (sortLoss) list.sort((a, b) => (real(b) - teor(b)) * b.costo - (real(a) - teor(a)) * a.costo);
  $("#conInfo").textContent = conAll
    ? `Mostrando ${Math.min(LIM, list.length)} de ${list.length} insumos.`
    : `${rel.length} insumos usados en recetas o con movimiento de inventario.`;
  $("#allCon").textContent = conAll ? "Solo los usados" : "Mostrar todos";
  $("#tCon").innerHTML = list
    .slice(0, conAll ? LIM : 1e9)
    .map(
      (i) =>
        `<tr data-id="${i.id}"><td>${esc(i.nom)}</td><td>${i.uni}</td>${["ii", "co", "fi"].map((f) => `<td class="n"><input type="number" data-c="${f}" min="0" step="0.1" value="${i[f]}"></td>`).join("")}<td class="n" id="cr${i.id}"></td><td class="n" id="ct${i.id}"></td><td class="n" id="cu${i.id}"></td><td class="n" id="cs${i.id}"></td></tr>`,
    )
    .join("");
  updVen();
}
function updVen() {
  S.rec.forEach((r) => {
    const a = document.getElementById("vn" + r.id),
      b = document.getElementById("vc" + r.id);
    if (a) a.textContent = money(r.ven * neto(r.precio));
    if (b) b.textContent = money(r.ven * costoRec(r));
  });
  S.ins.forEach((i) => {
    const t = teor(i),
      re = real(i),
      du = re - t,
      ds = du * i.costo,
      cls = ds > 0.005 ? "loss" : ds < -0.005 ? "gain" : "";
    const set = (p, v, c) => {
      const el = document.getElementById(p + i.id);
      if (el) {
        el.textContent = v;
        el.className = "n " + (c || "");
      }
    };
    set("cr", f2(re));
    set("ct", f2(t));
    set("cu", (du > 0 ? "+" : "") + f2(du), cls);
    set("cs", (ds > 0 ? "+" : "") + money(ds), cls);
  });
}
$("#tVen").addEventListener("input", (e) => {
  const t = e.target;
  if (t.dataset.v) {
    S.rec.find((r) => r.id === t.closest("tr").dataset.id).ven = Math.max(0, num(t.value));
    refresh();
  }
});
$("#tCon").addEventListener("input", (e) => {
  const t = e.target;
  if (t.dataset.c) {
    insById(t.closest("tr").dataset.id)[t.dataset.c] = Math.max(0, num(t.value));
    refresh();
  }
});
$("#allCon").onclick = () => {
  conAll = !conAll;
  renderVen();
};
$("#sortCon").onclick = () => {
  sortLoss = !sortLoss;
  $("#sortCon").textContent = sortLoss ? "Orden original" : "Ordenar por pérdida";
  renderVen();
};

/* ---------- 6. excel ---------- */
const AL = {
  cod: ["codigo", "cod", "code", "sku", "codigoinsumo"],
  nom: [
    "insumo",
    "nombre",
    "descripcion",
    "producto",
    "item",
    "descripciondeinsumo",
    "descripcioninsumo",
    "articulo",
  ],
  cat: ["categoria", "familia", "grupo"],
  uni: ["unidad", "und", "um", "unidadmedida", "unidadcompra", "unidaddecompra"],
  costo: [
    "costo",
    "costocompra",
    "precio",
    "preciocompra",
    "costounitario",
    "preciopromedio",
    "preciopromedios",
    "costopromedio",
    "ultimocosto",
    "ultcosto",
  ],
  fac: ["factor", "factorconversion", "equivalencia"],
  ucost: ["unidadcosto", "unidaddecosto", "unidadreceta"],
  merma: ["merma", "mermapct", "merma%", "porcentajemerma"],
  ii: ["invinicial", "inventarioinicial", "inicial", "stockinicial"],
  co: ["compras", "compra", "ingresos"],
  fi: ["invfinal", "inventariofinal", "final", "stockfinal"],
};
const ALV = {
  fecha: ["fecha", "dia", "date"],
  plato: ["plato", "producto", "descripcion", "nombre", "item"],
  cod: ["codigo", "cod", "code"],
  u: ["unidades", "cantidad", "cant", "und", "vendidas", "unidadesvendidas"],
  precio: ["precio", "precioventa", "pvp"],
};
function mapHead(h, al) {
  const m = {};
  h.forEach((c, k) => {
    const n = norm(c);
    for (const f in al) if (m[f] == null && al[f].map(norm).includes(n)) m[f] = k;
  });
  return m;
}
const uniN = (u) => {
  const n = norm(u);
  return ["kg", "kilo", "kilos", "kilogramo", "kilogramos", "kgs"].includes(n)
    ? "kg"
    : ["lt", "l", "litro", "litros", "lts"].includes(n)
      ? "lt"
      : [
            "und",
            "un",
            "unidad",
            "unidades",
            "u",
            "uni",
            "porcion",
            "porciones",
            "botella",
            "botellas",
            "pieza",
            "piezas",
          ].includes(n)
        ? "und"
        : null;
};
const UCONV = {
  gramos: ["kg", 1000],
  gramo: ["kg", 1000],
  gr: ["kg", 1000],
  grs: ["kg", 1000],
  g: ["kg", 1000],
  mililitros: ["lt", 1000],
  mililitro: ["lt", 1000],
  ml: ["lt", 1000],
  kilos: ["kg", 1],
  kilo: ["kg", 1],
  kg: ["kg", 1],
  kilogramos: ["kg", 1],
  litros: ["lt", 1],
  litro: ["lt", 1],
  lt: ["lt", 1],
  l: ["lt", 1],
  galon: ["lt", 1 / 3.785],
  galones: ["lt", 1 / 3.785],
  onza: ["kg", 1000 / 28.35],
  onzas: ["kg", 1000 / 28.35],
  oz: ["kg", 1000 / 28.35],
};
[
  "unidad",
  "unidades",
  "und",
  "un",
  "porcion",
  "porciones",
  "botella",
  "botellas",
  "paquete",
  "paquetes",
  "caja",
  "cajas",
  "bolsa",
  "bolsas",
  "lata",
  "latas",
  "pieza",
  "piezas",
  "rollo",
  "rollos",
  "sobre",
  "sobres",
].forEach((k) => (UCONV[k] = ["und", 1]));
let pending = null;
function readFile(file, k) {
  if (!window.XLSX) {
    toast("No se pudo cargar el lector de Excel. Revisa tu conexión.");
    return;
  }
  const fr = new FileReader();
  fr.onload = () => {
    try {
      const wb = XLSX.read(new Uint8Array(fr.result), { type: "array", cellDates: true });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        header: 1,
        defval: "",
        raw: true,
      });
      k === "ins" ? prevIns(rows, file.name) : prevVen(rows, file.name);
    } catch (e) {
      toast("No se pudo leer el archivo: " + e.message);
    }
  };
  fr.readAsArrayBuffer(file);
}
function headRow(rows, al) {
  for (let k = 0; k < Math.min(rows.length, 40); k++) {
    const m = mapHead(rows[k], al);
    if (Object.keys(m).length >= 2) return [k, m];
  }
  return [-1, {}];
}
function prevIns(rows, fn) {
  const [h, m] = headRow(rows, AL);
  if (h < 0 || m.nom == null) {
    showPrev(
      `<p class="loss">No encontré la columna del nombre del insumo en ${esc(fn)}. Usa la plantilla o el reporte de insumos de Inforest.</p>`,
    );
    return;
  }
  const out = [];
  rows.slice(h + 1).forEach((r, k) => {
    if (r.every((c) => String(c).trim() === "")) return;
    const g = (f) => (m[f] == null ? "" : r[m[f]]);
    const err = [];
    const nom = String(g("nom")).trim();
    if (!nom) return;
    let costo = num(g("costo")),
      uni,
      conv = "";
    if (String(g("costo")).trim() === "" || costo < 0) err.push("costo no válido");
    const uc = norm(g("ucost")),
      fac = num(g("fac")) || 1;
    if (m.ucost != null && uc) {
      const cv = UCONV[uc];
      if (cv) {
        uni = cv[0];
        costo = (costo / fac) * cv[1];
      } else uni = null;
      conv = `${g("uni")} ÷ ${f2(fac)} ${g("ucost")}`;
    } else uni = uniN(g("uni") || "kg");
    if (!uni) err.push("unidad no válida");
    let me = null;
    if (m.merma != null && String(g("merma")).trim() !== "") {
      me = num(g("merma"));
      if (me > 0 && me < 1) me *= 100;
      if (me < 0 || me >= 100) err.push("merma fuera de rango");
      me = +me.toFixed(2);
    }
    const fam = String(g("cat")).trim();
    const o = {
      fila: h + k + 2,
      cod: String(g("cod")).trim(),
      nom,
      cat: fam ? CATS.find((c) => norm(c) === norm(fam)) || fam : "",
      uni,
      costo: +costo.toFixed(4),
      merma: me,
      conv,
      err,
      warn: !err.length && costo === 0,
    };
    ["ii", "co", "fi"].forEach((f) => {
      if (m[f] != null && String(g(f)).trim() !== "") {
        const v = num(g(f));
        if (v < 0) err.push("inventario negativo");
        o[f] = v;
      }
    });
    out.push(o);
  });
  const ok = out.filter((o) => !o.err.length);
  const fams = [
    ...ok.reduce(
      (mp, o) => mp.set(o.cat || "Sin familia", (mp.get(o.cat || "Sin familia") || 0) + 1),
      new Map(),
    ),
  ].sort((a, b) => b[1] - a[1]);
  const off = ["suministrosdiversos", "suministros"];
  pending = { k: "ins", ok };
  const show = [...out.filter((o) => o.err.length), ...out.filter((o) => !o.err.length)].slice(0, 200);
  const esEjemplo = S.ins.length === 18 && S.ins.every((i) => /^I0\d\d$/.test(i.cod));
  showPrev(`<h3>Vista previa: ${esc(fn)}</h3><p class="hint">${ok.length} insumos válidos, ${out.length - ok.length} con error, ${ok.filter((o) => o.warn).length} sin precio (entran con costo 0).${m.ucost != null ? " El costo se convierte a S/ por kg, litro o unidad usando el factor de Inforest." : ""}</p>
   ${fams.length > 1 ? `<p style="margin:8px 0 4px"><b>Familias a cargar</b></p><div class="row" style="gap:6px 16px">${fams.map(([f, n]) => `<label class="chk" style="min-height:32px;font-size:14px"><input type="checkbox" class="famchk" value="${esc(f)}"${off.includes(norm(f)) ? "" : " checked"}> ${esc(f)} (${n})</label>`).join("")}</div>` : ""}
   <div class="scroll" style="max-height:340px;margin-top:10px"><table><thead><tr><th>Fila</th><th>Insumo</th><th>Familia</th><th>Unidad</th><th class="n">Costo S/</th><th>Conversión</th><th>Estado</th></tr></thead><tbody>${show.map((o) => `<tr><td>${o.fila}</td><td>${esc(o.nom)}</td><td>${esc(o.cat)}</td><td>${o.uni || ""}</td><td class="n">${f2(o.costo)}</td><td style="font-size:13px;color:var(--mute)">${esc(o.conv)}</td><td>${o.err.length ? `<span class="tag bad">${o.err.join(", ")}</span>` : o.warn ? '<span class="tag warn">Sin precio</span>' : '<span class="tag ok">OK</span>'}</td></tr>`).join("")}</tbody></table></div>
   ${out.length > 200 ? `<p class="hint">Se muestran 200 de ${out.length} filas.</p>` : ""}
   <p class="hint">"Reemplazar todo" borra los insumos actuales y los ingredientes de tus recetas. "Actualizar y agregar" cruza por código o por nombre, actualiza el costo y agrega los nuevos.</p>${modeUI(ok.length, esEjemplo)}`);
}
function prevVen(rows, fn) {
  const [h, m] = headRow(rows, ALV);
  if (h < 0 || (m.plato == null && m.cod == null) || m.u == null) {
    showPrev(
      `<p class="loss">No encontré las columnas "Plato" y "Unidades" en ${esc(fn)}. Usa la plantilla.</p>`,
    );
    return;
  }
  const agg = new Map(),
    bad = [],
    lines = [];
  let dmin = null,
    dmax = null;
  rows.slice(h + 1).forEach((r, k) => {
    if (r.every((c) => String(c).trim() === "")) return;
    const g = (f) => (m[f] == null ? "" : r[m[f]]);
    const u = num(g("u"));
    const nm = String(g("plato")).trim(),
      cd = String(g("cod")).trim();
    if ((!nm && !cd) || u < 0 || String(g("u")).trim() === "") {
      bad.push(h + k + 2);
      return;
    }
    let fd = g("fecha");
    if (fd instanceof Date) fd = fd.toISOString().slice(0, 10);
    else if (typeof fd === "number")
      fd = new Date(Math.round((fd - 25569) * 864e5)).toISOString().slice(0, 10);
    else if (fd) {
      const p = String(fd).match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
      fd = p ? `${p[3]}-${p[2].padStart(2, "0")}-${p[1].padStart(2, "0")}` : "";
    }
    if (fd) {
      if (!dmin || fd < dmin) dmin = fd;
      if (!dmax || fd > dmax) dmax = fd;
    }
    const rec =
      S.rec.find((x) => cd && norm(x.cod) === norm(cd)) || S.rec.find((x) => nm && norm(x.nom) === norm(nm));
    const key = rec ? "r" + rec.id : "n" + norm(nm || cd);
    const a = agg.get(key) || { rec, n: nm || cd, u: 0, pv: 0, pu: 0 };
    a.u += u;
    const pr = num(g("precio"));
    lines.push({
      fecha: fd || null,
      plato: nm || (rec && rec.nom) || cd,
      codigo: cd,
      u,
      precio: pr,
      rid: rec ? rec.id : null,
    });
    if (pr > 0) {
      a.pv += pr * u;
      a.pu += u;
    }
    agg.set(key, a);
  });
  const list = [...agg.values()];
  pending = { k: "ven", list, dmin, dmax, lines, fn };
  showPrev(`<h3>Vista previa: ${esc(fn)}</h3><p class="hint">${list.filter((x) => x.rec).length} platos con receta, ${list.filter((x) => !x.rec).length} sin receta${bad.length ? `, ${bad.length} filas con error (${bad.slice(0, 8).join(", ")}${bad.length > 8 ? "…" : ""})` : ""}.${dmin ? ` Período detectado: ${fmtD(dmin)} al ${fmtD(dmax)}.` : ""}</p>
   <div class="scroll" style="max-height:340px"><table><thead><tr><th>Plato</th><th class="n">Unidades</th><th class="n">Precio prom.</th><th>Estado</th></tr></thead><tbody>${list.map((x) => `<tr><td>${esc(x.rec ? x.rec.nom : x.n)}</td><td class="n">${f2(x.u)}</td><td class="n">${x.pu ? money(x.pv / x.pu) : "—"}</td><td>${x.rec ? '<span class="tag ok">Con receta</span>' : '<span class="tag warn">Sin receta</span>'}</td></tr>`).join("")}</tbody></table></div>
   <label class="chk" style="display:flex;margin-top:10px"><input type="checkbox" id="updPrice"> Actualizar precios de venta con el precio promedio del Excel</label>${modeUI(list.length)}`);
}
const modeUI = (n, r) =>
  n
    ? `<div class="row" style="margin-top:14px"><label style="font-size:14px"><select id="mode" style="width:auto"><option value="upd">Actualizar y agregar</option><option value="rep"${r ? " selected" : ""}>Reemplazar todo</option></select></label><button class="btn" id="okImp">Confirmar carga</button><button class="btn ghost" id="noImp">Cancelar</button></div>`
    : '<p class="loss">No hay filas válidas para cargar.</p>';
function showPrev(h) {
  $("#prev").innerHTML = `<div class="card">${h}</div>`;
  $("#prev").scrollIntoView({ behavior: "smooth", block: "start" });
}
$("#prev").addEventListener("click", (e) => {
  if (e.target.id === "noImp") {
    pending = null;
    $("#prev").innerHTML = "";
  }
  if (e.target.id !== "okImp" || !pending) return;
  const rep = $("#mode").value === "rep";
  if (pending.k === "ins") {
    const sel = $$(".famchk");
    const fam = new Set(sel.filter((c) => c.checked).map((c) => c.value));
    const rows = sel.length ? pending.ok.filter((o) => fam.has(o.cat || "Sin familia")) : pending.ok;
    if (!rows.length) {
      toast("Elige al menos una familia");
      return;
    }
    if (rep) {
      S.ins = [];
      S.rec.forEach((r) => (r.ing = []));
    }
    let a = 0,
      u = 0;
    const byCod = new Map(),
      byNom = new Map();
    S.ins.forEach((i) => {
      if (i.cod) byCod.set(norm(i.cod), i);
      byNom.set(norm(i.nom), i);
    });
    rows.forEach((o) => {
      let i = (o.cod && byCod.get(norm(o.cod))) || byNom.get(norm(o.nom));
      if (!i) {
        i = {
          id: nid(),
          cod: o.cod || "I" + String(S.ins.length + 1).padStart(4, "0"),
          cat: "Otros",
          merma: 0,
          ii: 0,
          co: 0,
          fi: 0,
        };
        S.ins.push(i);
        byNom.set(norm(o.nom), i);
        a++;
      } else u++;
      Object.assign(i, { nom: o.nom, uni: o.uni, costo: o.costo });
      if (o.cat) i.cat = o.cat;
      if (o.merma != null) i.merma = o.merma;
      if (o.cod) i.cod = o.cod;
      ["ii", "co", "fi"].forEach((f) => {
        if (o[f] != null) i[f] = o[f];
      });
    });
    toast(`Insumos cargados: ${a} nuevos, ${u} actualizados` + (rep ? ". Revisa las recetas." : ""));
  } else {
    if (rep) S.rec.forEach((r) => (r.ven = 0));
    const up = $("#updPrice")?.checked;
    pending.list.forEach((x) => {
      if (x.rec) {
        x.rec.ven = x.u;
        if (up && x.pu) x.rec.precio = +(x.pv / x.pu).toFixed(2);
      }
    });
    sinRec = pending.list.filter((x) => !x.rec);
    cloudVentas(pending.lines, pending.fn);
    if (pending.dmin) {
      S.cfg.ini = pending.dmin;
      S.cfg.fin = pending.dmax;
      loadCfg();
    }
    toast("Ventas cargadas" + (sinRec.length ? `. ${sinRec.length} plato(s) sin receta.` : ""));
  }
  pending = null;
  $("#prev").innerHTML = "";
  renderAll();
});
let dropK = "ins";
$$(".drop").forEach((d) => {
  d.onclick = () => {
    dropK = d.dataset.k;
    $("#fXls").click();
  };
  d.ondragover = (e) => {
    e.preventDefault();
    d.classList.add("over");
  };
  d.ondragleave = () => d.classList.remove("over");
  d.ondrop = (e) => {
    e.preventDefault();
    d.classList.remove("over");
    const f = e.dataTransfer.files[0];
    f && readFile(f, d.dataset.k);
  };
});
$("#fXls").onchange = (e) => {
  const f = e.target.files[0];
  f && readFile(f, dropK);
  e.target.value = "";
};

/* ---------- descargas ---------- */
async function saveFile(name, data) {
  try {
    if (window.claude && typeof window.claude.use === "function") {
      const d = await Promise.race([
        window.claude.use("downloads"),
        new Promise((r) => setTimeout(() => r(null), 4000)),
      ]);
      if (d) {
        try {
          await d.save({ filename: name, data });
          toast("Archivo listo");
        } catch (err) {
          if (err && err.code !== "declined") toast("No se pudo descargar en esta vista");
        }
        return;
      }
    }
  } catch (e) {}
  const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data]));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function wbBlob(sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([n, aoa]) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = aoa[0].map((_, k) => ({
      wch: Math.max(10, ...aoa.map((r) => String(r[k] ?? "").length + 2)),
    }));
    XLSX.utils.book_append_sheet(wb, ws, n);
  });
  return new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })]);
}
const HI = [
  "Código",
  "Insumo",
  "Categoría",
  "Unidad",
  "Costo",
  "Merma %",
  "Inv. inicial",
  "Compras",
  "Inv. final",
];
$$("[data-tpl]").forEach(
  (b) =>
    (b.onclick = () => {
      if (!window.XLSX) return toast("Lector de Excel no disponible");
      const t = b.dataset.tpl;
      let aoa;
      if (t === "ins0") aoa = [HI];
      if (t === "ins1")
        aoa = [HI, ...S.ins.map((i) => [i.cod, i.nom, i.cat, i.uni, i.costo, i.merma, i.ii, i.co, i.fi])];
      if (t === "ven0") aoa = [["Fecha", "Plato", "Código", "Unidades", "Precio"]];
      if (t === "ven1") {
        aoa = [["Fecha", "Plato", "Código", "Unidades", "Precio"]];
        const days = 31;
        S.rec.forEach((r) => {
          let rest = r.ven;
          for (let d = 1; d <= days; d++) {
            const u =
              d === days ? rest : Math.round((r.ven / days) * (d % 7 === 5 || d % 7 === 6 ? 1.4 : 0.85));
            const uu = Math.max(0, Math.min(rest, u));
            rest -= uu;
            aoa.push([`2026-10-${String(d).padStart(2, "0")}`, r.nom, r.cod, uu, r.precio]);
          }
        });
      }
      saveFile(
        {
          ins0: "plantilla_insumos.xlsx",
          ins1: "insumos_ejemplo.xlsx",
          ven0: "plantilla_ventas.xlsx",
          ven1: "ventas_ejemplo.xlsx",
        }[t],
        wbBlob([[t.startsWith("ins") ? "Insumos" : "Ventas", aoa]]),
      );
    }),
);
$("#expXls").onclick = () => {
  if (!window.XLSX) return toast("Lector de Excel no disponible");
  const T = totales();
  const res = [
    ["Indicador", "Valor"],
    ["Restaurante", S.cfg.nombre],
    ["Período", `${fmtD(S.cfg.ini)} al ${fmtD(S.cfg.fin)}`],
    ["Meta FC %", S.cfg.meta],
    ["Ventas netas S/", +T.vn.toFixed(2)],
    ["Costo teórico S/", +T.ct.toFixed(2)],
    ["Costo real S/", +T.cr.toFixed(2)],
    ["FC teórico %", +(T.fcT * 100).toFixed(1)],
    ["FC real %", +(T.fcR * 100).toFixed(1)],
    ["Desviación pts", +((T.fcR - T.fcT) * 100).toFixed(1)],
    ["Margen bruto real S/", +(T.vn - T.cr).toFixed(2)],
  ];
  const rec = [
    [
      "Plato",
      "Código",
      "Unidades",
      "Precio",
      "Costo porción",
      "FC %",
      "Margen unit.",
      "Precio sugerido",
      "Estado",
    ],
  ].concat(
    S.rec.map((r) => {
      const c = costoRec(r),
        n = neto(r.precio),
        fc = n ? c / n : 0;
      return [
        r.nom,
        r.cod,
        r.ven,
        r.precio,
        +c.toFixed(2),
        +(fc * 100).toFixed(1),
        +(n - c).toFixed(2),
        sugerido(c),
        estado(fc)[1],
      ];
    }),
  );
  const fic = [["Plato", "Insumo", "Unidad", "Cantidad", "Costo"]];
  S.rec.forEach((r) =>
    r.ing.forEach((g) => {
      const i = insById(g.iid);
      i && fic.push([r.nom, i.nom, i.uni, g.q, +(g.q * apr(i)).toFixed(2)]);
    }),
  );
  const con = [
    ["Insumo", "Unidad", "Inv. inicial", "Compras", "Inv. final", "Real", "Teórico", "Dif. unid.", "Dif. S/"],
  ].concat(
    T.dev.map((x) => [
      x.i.nom,
      x.i.uni,
      x.i.ii,
      x.i.co,
      x.i.fi,
      +x.re.toFixed(3),
      +x.t.toFixed(3),
      +x.du.toFixed(3),
      +x.ds.toFixed(2),
    ]),
  );
  saveFile(
    `food_cost_${S.cfg.ini}_${S.cfg.fin}.xlsx`,
    wbBlob([
      ["Resumen", res],
      ["Platos", rec],
      ["Recetas", fic],
      ["Consumo", con],
    ]),
  );
};
$("#expJson").onclick = () =>
  saveFile(
    "respaldo_food_cost.json",
    new Blob([JSON.stringify({ ...S, uid }, null, 1)], { type: "application/json" }),
  );
$("#impJson").onclick = () => $("#fJson").click();
$("#fJson").onchange = (e) => {
  const f = e.target.files[0];
  if (!f) return;
  f.text()
    .then((t) => {
      const d = JSON.parse(t);
      if (!d.ins || !d.rec) throw 0;
      S = fixIds(d);
      loadCfg();
      renderAll();
      toast("Respaldo cargado");
    })
    .catch(() => toast("El archivo no es un respaldo válido"));
  e.target.value = "";
};
$("#reset").onclick = () => {
  $("#confirmReset").innerHTML =
    `<div class="alert bad" style="margin:14px 0 0">Se reemplazarán todos tus datos por la demo de Liberty Bistró. <span class="row" style="margin-top:8px"><button class="btn sm" id="rsY" style="background:var(--bad)">Sí, restaurar</button><button class="btn ghost sm" id="rsN">Cancelar</button></span></div>`;
};
$("#confirmReset").addEventListener("click", (e) => {
  if (e.target.id === "rsY") {
    (async () => {
      try {
        await cargarDemo();
        await subirDemoNube();
      } catch (err) {
        seed();
      }
      sinRec = [];
      loadCfg();
      renderAll();
      toast("Demo restaurada");
    })();
  }
  if (e.target.id === "rsY" || e.target.id === "rsN") $("#confirmReset").innerHTML = "";
});

/* ---------- 7. configuración ---------- */
function loadCfg() {
  $("#cNombre").value = S.cfg.nombre;
  $("#cMeta").value = S.cfg.meta;
  $("#cIgv").checked = S.cfg.igv;
  $("#cRc").value = num(S.cfg.rc) || 0;
  $("#cIni").value = S.cfg.ini;
  $("#cFin").value = S.cfg.fin;
  $("#cAlerta").value = num(S.cfg.alerta) || 35;
  renderBrand();
}
$("#cNombre").oninput = (e) => {
  S.cfg.nombre = e.target.value;
  renderBrand();
  save();
};
$("#cMeta").oninput = (e) => {
  const v = num(e.target.value);
  if (v > 0 && v < 100) {
    S.cfg.meta = v;
    refresh();
  }
};
$("#cIgv").onchange = (e) => {
  S.cfg.igv = e.target.checked;
  refresh();
};
$("#cRc").oninput = (e) => {
  const v = num(e.target.value);
  if (v >= 0 && v <= 30) {
    S.cfg.rc = v;
    refresh();
  }
};
function cambioPeriodo() {
  unidadesDesdeVentas();
  renderBrand();
  renderVen();
  refresh();
  if (tabActual() === "reportes") renderReportes();
}
$("#cIni").onchange = (e) => {
  S.cfg.ini = e.target.value;
  cambioPeriodo();
};
$("#cFin").onchange = (e) => {
  S.cfg.fin = e.target.value;
  cambioPeriodo();
};
$("#cAlerta").oninput = (e) => {
  const v = num(e.target.value);
  if (v > 0 && v < 100) {
    S.cfg.alerta = v;
    save();
    if (tabActual() === "reportes") renderReportes();
  }
};
$("#bLogo").onclick = () => $("#fLogo").click();
$("#fLogo").onchange = (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const img = new Image(),
    fr = new FileReader();
  fr.onload = () => {
    img.onload = () => {
      const c = document.createElement("canvas");
      c.width = c.height = 160;
      const x = c.getContext("2d"),
        s = Math.min(img.width, img.height);
      x.drawImage(img, (img.width - s) / 2, (img.height - s) / 2, s, s, 0, 0, 160, 160);
      S.cfg.logo = c.toDataURL("image/png");
      renderBrand();
      save();
    };
    img.src = fr.result;
  };
  fr.readAsDataURL(f);
  e.target.value = "";
};

/* ---------- general ---------- */
let tt;
function toast(m) {
  const t = $("#toast");
  t.textContent = m;
  t.classList.add("on");
  clearTimeout(tt);
  tt = setTimeout(() => t.classList.remove("on"), 3200);
}
function refresh() {
  _teo = null;
  renderResumen();
  updIns();
  updRec();
  updVen();
  save();
}
function renderAll() {
  _teo = null;
  renderBrand();
  renderEstadoDatos();
  if (tabActual() === "reportes") renderReportes();
  renderIns();
  renderRec();
  renderVen();
  renderResumen();
  save();
}
const TABS = ["resumen", "reportes", "insumos", "recetas", "ventas", "carga"];
const tabActual = () => $("section[data-tab].on")?.dataset.tab;
function route() {
  const t = (location.hash || "#resumen").slice(1);
  if (t.startsWith("r-")) return;
  const v = TABS.includes(t) ? t : "resumen";
  $$("section[data-tab]").forEach((s) => s.classList.toggle("on", s.dataset.tab === v));
  if (v === "ventas") renderVen();
  if (v === "reportes") renderReportes();
  if (v === "carga") renderEstadoDatos();
  $$("#nav a").forEach((a) => a.classList.toggle("on", a.dataset.t === v));
}
window.addEventListener("hashchange", route);
$$("#nav a").forEach((a) =>
  a.addEventListener("click", (e) => {
    e.preventDefault();
    try {
      history.replaceState(null, "", "#" + a.dataset.t);
    } catch (_) {
      location.hash = a.dataset.t;
    }
    route();
  }),
);
/* ---------- 8. sesión ---------- */
function acctUI() {
  const a = $("#acct");
  if (!sb) {
    a.innerHTML =
      '<span class="sync" title="Completa SUPABASE_URL y SUPABASE_KEY para guardar en la nube">Modo local</span>';
    return;
  }
  a.innerHTML = user
    ? `<span class="sync" id="sync">Conectado</span><span>${esc(user.email)}</span><button class="btn ghost sm" id="bOut">Salir</button>`
    : "";
  const o = $("#bOut");
  if (o)
    o.onclick = async () => {
      clearTimeout(cloudT);
      if (!saving) await cloudSave();
      await sb.auth.signOut();
    };
}
async function enter(u) {
  user = u;
  ready = false;
  $("#auth").classList.add("hide");
  acctUI();
  setSync("Cargando…", "");
  try {
    await cloudLoad();
    ready = true;
    loadCfg();
    renderAll();
    route();
    setSync("Guardado en la nube", "ok");
  } catch (e) {
    console.error(e);
    setSync("Sin conexión", "err");
    toast("No se pudo leer la base de datos: " + (e.message || e) + ". Revisa que creaste las tablas.");
    if (localLoad(KEY + "-" + u.id)) {
      loadCfg();
      renderAll();
      route();
    }
  }
}
function showAuth() {
  user = null;
  ready = false;
  acctUI();
  $("#auth").classList.remove("hide");
  $("#aMsg").textContent = "";
}
$("#authForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  $("#aMsg").textContent = "Ingresando…";
  const { data, error } = await sb.auth.signInWithPassword({
    email: $("#aMail").value.trim(),
    password: $("#aPass").value,
  });
  if (error) {
    $("#aMsg").textContent =
      error.message === "Invalid login credentials" ? "Correo o contraseña incorrectos." : error.message;
    return;
  }
  enter(data.user);
});
$("#aSign").onclick = async () => {
  const email = $("#aMail").value.trim(),
    password = $("#aPass").value;
  if (!email || password.length < 6) {
    $("#aMsg").textContent = "Escribe tu correo y una contraseña de al menos 6 caracteres.";
    return;
  }
  $("#aMsg").textContent = "Creando cuenta…";
  const { data, error } = await sb.auth.signUp({
    email,
    password,
    options: { emailRedirectTo: location.origin + location.pathname },
  });
  if (error) {
    $("#aMsg").textContent = error.message;
    return;
  }
  if (data.session) enter(data.user);
  else $("#aMsg").textContent = "Cuenta creada. Revisa tu correo para confirmarla y luego ingresa.";
};
async function start() {
  leerDatosLocal();
  if (!sb) {
    if (!localLoad(KEY)) {
      try {
        await cargarDemo();
      } catch (e) {
        seed();
      }
    }
    unidadesDesdeVentas();
    acctUI();
    loadCfg();
    renderAll();
    route();
    return;
  }
  if (!localLoad(KEY)) seed();
  loadCfg();
  renderAll();
  route();
  sb.auth.onAuthStateChange((ev, s) => {
    if (ev === "SIGNED_OUT") showAuth();
  });
  const { data } = await sb.auth.getSession();
  if (data.session) enter(data.session.user);
  else showAuth();
}
window.addEventListener("DOMContentLoaded", start);
