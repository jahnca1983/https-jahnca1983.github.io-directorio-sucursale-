<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <title>Consulta de Sucursales</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body { font-family: Arial, sans-serif; margin:0; padding:30px;
      background-image:url('Fondo1.png');
      background-size:cover; background-position:center; color:#fff; }
    h2 { text-shadow:1px 1px 4px #000; }
    input, select, button { margin:5px; padding:10px; font-size:16px; border-radius:5px; border:none; }
    select { width:250px; } input[type="text"] { width:230px; }
    #filtros { display:flex; flex-wrap:wrap; gap:10px; margin-bottom:20px; }
    button { background:#00796b; color:#fff; cursor:pointer; }
    button:hover { background:#004d40; }
    .boton-limpiar { background:#e53935; } .boton-limpiar:hover { background:#b71c1c; }
    .card { background:rgba(255,255,255,0.95); color:#000; border-radius:5px; margin-top:10px; padding:10px; box-shadow:0 2px 5px rgba(0,0,0,0.3); }
    .sucursal-btn { width:100%; text-align:left; background:#4caf50; color:#fff; border:none; padding:10px; font-size:16px; border-radius:4px; cursor:pointer; }
    .sucursal-btn:hover { background:#388e3c; }
    .card-content { display:none; padding-top:10px; border-top:1px solid #ccc; }
    .card-content div { margin-bottom:4px; }
    select:disabled { opacity:0.6; }
    #resultado { margin-top:20px; }
    #resumenGerente table { border-collapse: collapse; margin-top:10px; }
    #resumenGerente th, #resumenGerente td { border:1px solid #333; padding:6px; }
    #resumenGerente th { background:#00796b; color:#fff; }
    #resumenGerente td { background:#fff; color:#000; }
  </style>
</head>
<body>

<h2>Consulta de Sucursales</h2>

<input type="file" id="archivoExcel" accept=".xlsx, .xls"><br><br>

<div id="filtros">
  <input type="text" id="codigo" placeholder="Código Sucursal">
  <input type="text" id="nombre" placeholder="Nombre Sucursal">
  <select id="cluster"><option value="">-- Clúster --</option></select>
  <select id="gerente"><option value="">-- Gerente --</option></select>
  <select id="director"><option value="">-- Director/Subdirector --</option></select>
  <select id="coordinador"><option value="">-- Coordinador/Supervisor --</option></select>
  <select id="region"><option value="">-- Región Económica --</option></select>
  <select id="departamento"><option value="">-- Departamento --</option></select>
  <select id="ciudad"><option value="">-- Ciudad --</option></select>
</div>

<button onclick="filtrar()">Buscar</button>
<button onclick="limpiarBusqueda()" class="boton-limpiar">Limpiar búsqueda</button>

<!-- Resumen del gerente -->
<div id="resumenGerente" style="margin-top:20px;"></div>

<!-- Resultados -->
<div id="resultado"></div>

<script>
/* ---------- Estado global ---------- */
let datos = [];
let KEYS = {};
let HEADERS = [];
let listenersConfigurados = false;

/* ---------- Normalización ---------- */
function n(str) {
  return (str ?? "").toString().trim()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ');
}

/* ---------- Helpers ---------- */
function valoresUnicosDisplay(rows, fieldKeyReal) {
  const mapa = new Map();
  if (!fieldKeyReal) return [];
  rows.forEach(r => {
    const raw = (r[fieldKeyReal] ?? "").toString().trim();
    const k = n(raw);
    if (k && !mapa.has(k)) mapa.set(k, raw);
  });
  return Array.from(mapa.entries())
              .sort((a,b)=> a[1].localeCompare(b[1], undefined, {sensitivity:'base'}));
}

const CANDIDATOS = {
  cod_suc: ["cod. suc","codigo sucursal","codigo","cod suc","id sucursal"],
  nombre_suc: ["nombre sucursal","sucursal","nombre"],
  cluster: ["clúster","cluster"],
  gerente: ["gerente"],
  director: ["director/subdirector","director","subdirector"],
  coordinador: ["coordinador/supervisor","coordinador","supervisor"],
  region: ["región económica","region economica","regional","región"],
  departamento: ["departamento","dpto"],
  ciudad: ["ciudad","municipio","poblacion","población"]
};

function mapearHeaders(muestraObjeto) {
  const llaves = Object.keys(muestraObjeto || {});
  HEADERS = llaves;
  const indiceNorm = {};
  llaves.forEach(k => indiceNorm[n(k)] = k);

  for (const canon in CANDIDATOS) {
    let encontrada = null;
    for (const alias of CANDIDATOS[canon]) {
      if (indiceNorm[n(alias)]) { encontrada = indiceNorm[n(alias)]; break; }
    }
    if (!encontrada) {
      for (const alias of CANDIDATOS[canon]) {
        const aliasNorm = n(alias);
        for (const llave of llaves) {
          const ln = n(llave);
          if (ln.includes(aliasNorm) || aliasNorm.includes(ln)) { encontrada = llave; break; }
        }
        if (encontrada) break;
      }
    }
    if (encontrada) KEYS[canon] = encontrada;
  }
}

const sel = { cluster:"", gerente:"", director:"", coordinador:"", region:"", departamento:"", ciudad:"" };
function valSel(id) { return n(document.getElementById(id).value); }

function filtrarRows(base=datos) {
  let rows = base;
  if (sel.gerente && KEYS.gerente)      rows = rows.filter(r => n(r[KEYS.gerente])      === sel.gerente);
  if (sel.director && KEYS.director)    rows = rows.filter(r => n(r[KEYS.director])     === sel.director);
  if (sel.coordinador && KEYS.coordinador) rows = rows.filter(r => n(r[KEYS.coordinador]) === sel.coordinador);
  if (sel.region && KEYS.region)        rows = rows.filter(r => n(r[KEYS.region])       === sel.region);
  if (sel.departamento && KEYS.departamento) rows = rows.filter(r => n(r[KEYS.departamento]) === sel.departamento);
  if (sel.ciudad && KEYS.ciudad)        rows = rows.filter(r => n(r[KEYS.ciudad])       === sel.ciudad);
  if (sel.cluster && KEYS.cluster)      rows = rows.filter(r => n(r[KEYS.cluster])      === sel.cluster);
  return rows;
}

function setOpciones(id, paresNormDisplay) {
  const selEl = document.getElementById(id);
  const placeholder = selEl.options[0]?.textContent || "-- Seleccione --";
  selEl.innerHTML = `<option value="">${placeholder}</option>`;

  paresNormDisplay.forEach(([kNorm, display]) => {
    // Evita agregar "No asignado" solo en director y coordinador
    if ((id === "director" || id === "coordinador") && n(display) === "no asignado") {
      return; 
    }
    const opt = document.createElement("option");
    opt.value = display;
    opt.textContent = display;
    selEl.appendChild(opt);
  });

  selEl.disabled = paresNormDisplay.length === 0;
  selEl.selectedIndex = 0;
}

function refrescarCombos() {
  if (!datos.length) return;
  const base = filtrarRows(datos);
  const obtenerPares = (canon) => KEYS[canon] ? valoresUnicosDisplay(base, KEYS[canon]) : [];
  setOpciones("gerente",      obtenerPares("gerente"));
  setOpciones("director",     obtenerPares("director"));
  setOpciones("coordinador",  obtenerPares("coordinador"));
  setOpciones("region",       obtenerPares("region"));
  setOpciones("departamento", obtenerPares("departamento"));
  setOpciones("ciudad",       obtenerPares("ciudad"));
  setOpciones("cluster",      obtenerPares("cluster"));
}

function sincronizarEstado() {
  sel.gerente      = valSel("gerente");
  sel.director     = valSel("director");
  sel.coordinador  = valSel("coordinador");
  sel.region       = valSel("region");
  sel.departamento = valSel("departamento");
  sel.ciudad       = valSel("ciudad");
  sel.cluster      = valSel("cluster");
}

function configurarListeners() {
  if (listenersConfigurados) return;
  ["gerente","director","coordinador","region","departamento","ciudad","cluster"]
    .forEach(id => document.getElementById(id).addEventListener("change", () => {
      sincronizarEstado();
      refrescarCombos();
    }));
  listenersConfigurados = true;
}

document.getElementById('archivoExcel').addEventListener('change', function (e) {
  const archivo = e.target.files[0];
  const lector = new FileReader();
  lector.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const hoja = workbook.Sheets["Directorio Consolidado"] || workbook.Sheets[workbook.SheetNames[0]];
    if (!hoja) return alert("No se encontró ninguna hoja.");
    datos = XLSX.utils.sheet_to_json(hoja, { defval: "" });
    if (!datos.length) return alert("La hoja está vacía.");
    mapearHeaders(datos[0]);
    configurarListeners();
    Object.keys(sel).forEach(k => sel[k] = "");
    refrescarCombos();
    alert("Archivo cargado con " + datos.length + " registros.");
  };
  lector.readAsArrayBuffer(archivo);
});

/* ---------- Resumen por Gerente ---------- */
function mostrarResumenGerente(gerenteSeleccionado) {
  const divResumen = document.getElementById("resumenGerente");
  divResumen.innerHTML = "";
  if (!gerenteSeleccionado || !KEYS.gerente) return;

  const registros = datos.filter(r => n(r[KEYS.gerente]) === n(gerenteSeleccionado));
  if (!registros.length) return;

  const directores = [...new Set(registros.map(r => r[KEYS.director]).filter(v => v && n(v) !== "no asignado"))];
  const coordinadores = [...new Set(registros.map(r => r[KEYS.coordinador]).filter(v => v && n(v) !== "no asignado"))];
  const regiones = [...new Set(registros.map(r => r[KEYS.region]).filter(Boolean))];
  const clusters = [...new Set(registros.map(r => r[KEYS.cluster]).filter(Boolean))];
  const cantidadSucursales = registros.length;

  let html = `<h3>Resumen del Gerente: ${gerenteSeleccionado}</h3>`;
  html += `<table><tr><th>Subdirectores</th><td>${directores.join(", ") || "—"}</td></tr>`;
  html += `<tr><th>Coordinadores</th><td>${coordinadores.join(", ") || "—"}</td></tr>`;
  html += `<tr><th>Regiones Económicas</th><td>${regiones.join(", ") || "—"}</td></tr>`;
  html += `<tr><th>Clusters</th><td>${clusters.join(", ") || "—"}</td></tr>`;
  html += `<tr><th>Cantidad de Sucursales</th><td>${cantidadSucursales}</td></tr></table>`;
  html += `<button onclick="exportarResumen('${gerenteSeleccionado}')">Exportar a Excel</button>`;

  divResumen.innerHTML = html;
}

document.getElementById("gerente").addEventListener("change", function() {
  mostrarResumenGerente(this.value);
});

/* ---------- Exportar a Excel ---------- */
function exportarResumen(gerenteSeleccionado) {
  const registros = datos.filter(r => n(r[KEYS.gerente]) === n(gerenteSeleccionado));
  if (!registros.length) return;

  const directores = [...new Set(registros.map(r => r[KEYS.director]).filter(v => v && n(v) !== "no asignado"))];
  const coordinadores = [...new Set(registros.map(r => r[KEYS.coordinador]).filter(v => v && n(v) !== "no asignado"))];
  const regiones = [...new Set(registros.map(r => r[KEYS.region]).filter(Boolean))];
  const clusters = [...new Set(registros.map(r => r[KEYS.cluster]).filter(Boolean))];
  const cantidadSucursales = registros.length;

  const resumen = [
    { Campo: "Gerente", Valor: gerenteSeleccionado },
    { Campo: "Subdirectores", Valor: directores.join(", ") || "—" },
    { Campo: "Coordinadores", Valor: coordinadores.join(", ") || "—" },
    { Campo: "Regiones Económicas", Valor: regiones.join(", ") || "—" },
    { Campo: "Clusters", Valor: clusters.join(", ") || "—" },
    { Campo: "Cantidad de Sucursales", Valor: cantidadSucursales }
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(resumen);
  XLSX.utils.book_append_sheet(wb, ws, "Resumen Gerente");
  XLSX.writeFile(wb, `Resumen_${gerenteSeleccionado}.xlsx`);
}

/* ---------- Búsqueda ---------- */
function filtrar() {
  if (!datos.length) { alert("Primero carga un archivo."); return; }
  const resultadoDiv = document.getElementById("resultado");
  resultadoDiv.innerHTML = "";

  const txtCodigo = n(document.getElementById("codigo").value);
  const txtNombre = n(document.getElementById("nombre").value);

  let res = filtrarRows(datos);

  res = res.filter(s => {
    const okCodigo = !txtCodigo || (KEYS.cod_suc && n(String(s[KEYS.cod_suc] ?? "")) === txtCodigo);
    const okNombre = !txtNombre || (KEYS.nombre_suc && n(String(s[KEYS.nombre_suc] ?? "")).includes(txtNombre));
    return okCodigo && okNombre;
  });

  if (res.length === 0) { resultadoDiv.innerHTML = "<p>No se encontraron resultados.</p>"; return; }

  res.forEach((s, i) => {
    const card = document.createElement("div"); card.className = "card";
    const btn = document.createElement("button"); btn.className = "sucursal-btn";
    const nombre = (KEYS.nombre_suc && s[KEYS.nombre_suc]) || `Sucursal ${i+1}`;
    btn.textContent = nombre;
    btn.onclick = () => {
      const content = card.querySelector(".card-content");
      content.style.display = content.style.display === "none" ? "block" : "none";
    };
    const contentDiv = document.createElement("div"); contentDiv.className = "card-content";
    Object.keys(s).forEach(k => {
      const div = document.createElement("div");
      div.innerHTML = `<strong>${k}:</strong> ${s[k] ?? "—"}`;
      contentDiv.appendChild(div);
    });
    card.appendChild(btn); card.appendChild(contentDiv);
    resultadoDiv.appendChild(card);
    contentDiv.style.display = "none";
  });
}

/* ---------- Limpiar ---------- */
function limpiarBusqueda() {
  document.querySelectorAll('#filtros input').forEach(i => i.value = '');
  Object.keys(sel).forEach(k => sel[k] = "");
  if (datos.length) refrescarCombos();
  document.getElementById('resultado').innerHTML = '';
  document.getElementById('resumenGerente').innerHTML = '';
}
</script>

</body>
</html>
