var SHEET_ID = "13DRivSS2r8rHYVVHekD8zPUlHF3WoW_XGDNm6qovaAU";

function doGet(e) {
  var p = e.parameter;
  var mod = p.mod || "";

  // LOGIN POR GET — resuelve el problema de CORS definitivamente
  if (mod === "auth" && p.action === "login") {
    return login(p.usuario, p.password);
  }

  if (mod === "init") return inicializar();
  if (mod === "usuarios") return listarUsuarios();
  if (mod === "consec") return listarConsecutivos();
  if (mod === "m1") return listarM1();
  if (mod === "m2") {
    return listarHoja("M2_Partes", function(r) {
      return p.rol === "admin" || r.usuario === p.usuario;
    });
  }
  if (mod === "m3a") return listarHoja("M3A_Bitacoras");
  if (mod === "m3b") return listarHoja("M3B_Secuestros");
  if (mod === "m3c") return listarHoja("M3C_Llamados");
  if (mod === "m4") return listarHoja("M4_Visitas");
  if (mod === "m7") return listarHoja("M7_Informes");
  if (mod === "auditoria") return listarHoja("Auditoria");
  if (mod === "docs") return docsGet(p);

  return ok({ error: "modulo no valido: " + mod });
}

function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch(x) {
    try { b = e.parameter; } catch(x2) {}
  }
  return procesarAccion(b);
}

function procesarAccion(b) {
  var mod = b.mod || "";
  var action = b.action || "";

  if (mod === "auth" && action === "login") return login(b.usuario, b.password);
  if (mod === "auth" && action === "cambiar_password") return cambiarPassword(b);

  // Control de Documentación (bloque independiente al final del archivo)
  if (mod === "docs") return docsPost(b);

  // Usuarios se evalúa ANTES que el crear genérico
  if (mod === "usuarios") {
    if (action === "crear") return crearUsuario(b);
    if (action === "permisos") return actualizarPermisos(b);
    if (action === "estado") return actualizarEstado(b);
    if (action === "eliminar") return eliminarUsuario(b);
  }

  if (action === "crear") return crearRegistro(mod, b);

  if (mod === "consec" && action === "set") return actualizarConsecutivos(b);
  if (mod === "reset" && action === "total") return resetTotal();
  if (mod === "reset" && action === "modulo") return resetModulos(b.mods);

  // ── M2 EDITAR ──────────────────────────────────────────────
  if (mod === "m2" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M2_Partes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    // Búsqueda case-insensitive del header id
    for (var i = 0; i < hdrs.length; i++) {
      ci[hdrs[i]] = i;
      ci[String(hdrs[i]).toLowerCase()] = i;
    }
    var colId = ci["id"] !== undefined ? ci["id"] : ci["ID"] !== undefined ? ci["ID"] : ci["parte_id"] !== undefined ? ci["parte_id"] : 0;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][colId] || "") === String(b.id || "")) {
        // Actualizar todos los campos del payload que existan como columnas
        for (var k in b) {
          if (k === "mod" || k === "action") continue;
          if (ci[k] !== undefined) hoja.getRange(r+1, ci[k]+1).setValue(b[k]||"");
        }
        // Mapeo explícito para campos con nombre distinto
        var mapeo = {tipo_lugar:b.tipoLugar, dir_exacta:b.dirExacta, cedula_oficial:b.cedulaOficial};
        for (var mk in mapeo) {
          if (ci[mk] !== undefined) hoja.getRange(r+1, ci[mk]+1).setValue(mapeo[mk]||"");
        }
        var hMods = getOrCreate(ss, "M2_Modificaciones");
        appendRow(hMods, {parte_id:b.id, ts:Utilities.formatDate(new Date(),"America/Costa_Rica","yyyy-MM-dd HH:mm:ss"), usuario:b.u||"", descripcion:"Modificación por "+(b.u||"")});
        return ok({ ok: true });
      }
    }
    return ok({ error: "Parte no encontrado: " + b.id });
  }

  // ── M4 EDITAR ──────────────────────────────────────────────
  if (mod === "m4" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M4_Visitas");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][ci["id"]] || "") === String(b.id || "")) {
        var ts = Utilities.formatDate(new Date(),"America/Costa_Rica","yyyy-MM-dd HH:mm:ss");
        var mapeo4 = {oficial:b.oficial,fecha:b.fecha,programa:b.prog,
          distrito:b.dist,centro:b.centro,nivel:b.nivel,
          matricula:parseInt(b.matricula)||0, asistentes:parseInt(b.est)||0,
          editado_en:ts};
        for (var k in mapeo4) {
          if (ci[k] !== undefined) hoja.getRange(r+1, ci[k]+1).setValue(mapeo4[k]);
        }
        return ok({ ok: true });
      }
    }
    return ok({ error: "Visita no encontrada: " + b.id });
  }

  // ── M7 HISTÓRICO ───────────────────────────────────────────
  if (mod === "m7" && action === "historico") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var anio = parseInt(b.anio) || (new Date().getFullYear() - 1);
    var hoja = getOrCreate(ss, "M7_Historico");
    // Eliminar filas del año
    var vals = hoja.getDataRange().getValues();
    for (var r = vals.length - 1; r >= 1; r--) {
      if (parseInt(vals[r][0]) === anio) hoja.deleteRow(r + 1);
    }
    // Insertar nuevos datos
    var hist = b.hist || {};
    for (var mes in hist) {
      appendRow(hoja, {
        anio: anio,
        mes: parseInt(mes),
        infracciones: parseInt(hist[mes].inf) || 0,
        monto: parseInt(hist[mes].mto) || 0
      });
    }
    return ok({ ok: true });
  }

  // ── M3B ELIMINAR ───────────────────────────────────────────
  if (mod === "m3b" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M3B_Secuestros");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }

  // ── M1 ELIMINAR ──────────────────────────────────────────
  if (mod === "m1" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M1_Reportes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado: " + b.id });
  }

  // ── M3A ELIMINAR ─────────────────────────────────────────
  if (mod === "m3a" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M3A_Bitacoras");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }

  // ── M2 ELIMINAR ──────────────────────────────────────────
  if (mod === "m2" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M2_Partes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }
  // ── M4 ELIMINAR ──────────────────────────────────────────
  if (mod === "m4" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M4_Visitas");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }
  // ── M7 ELIMINAR ──────────────────────────────────────────
  if (mod === "m7" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M7_Informes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        hoja.deleteRow(r + 1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }
  // ── M1 EDITAR ────────────────────────────────────────────
  if (mod === "m1" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M1_Reportes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0].slice(); var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    // Crear columnas faltantes (ap1..ap33 en registros viejos)
    var newCols = Object.keys(b).filter(function(k){
      return ci[k]===undefined && k!=="mod" && k!=="action";
    });
    for (var nc=0; nc<newCols.length; nc++) {
      var newIdx = hdrs.length;
      hoja.getRange(1, newIdx+1).setValue(newCols[nc]);
      ci[newCols[nc]] = newIdx;
      hdrs.push(newCols[nc]);
    }
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        for (var k in b) {
          if (ci[k] !== undefined && k !== "id" && k !== "mod" && k !== "action") {
            hoja.getRange(r+1, ci[k]+1).setValue(b[k]);
          }
        }
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado: " + b.id });
  }
  // ── M7 EDITAR ────────────────────────────────────────────
  if (mod === "m7" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M7_Informes");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][ci["id"]]||"") === String(b.id||"")) {
        var campos = ["mes","anio","mesNom","inf","monto","canc","anul","napel","reg"];
        campos.forEach(function(k){ if(ci[k]!==undefined) hoja.getRange(r+1,ci[k]+1).setValue(b[k]||""); });
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }

  // ── M3B EDITAR ────────────────────────────────────────────
  if (mod === "m3b" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M3B_Secuestros");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) { ci[hdrs[i]]=i; ci[String(hdrs[i]).toLowerCase()]=i; }
    var colId = ci["id"]!==undefined ? ci["id"] : 0;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][colId]||"") === String(b.id||"")) {
        var campos3b = ["inst","op","agente","ced","causa","motivo","fi","ff","estado","u"];
        campos3b.forEach(function(k){ if(ci[k]!==undefined) hoja.getRange(r+1,ci[k]+1).setValue(b[k]||""); });
        return ok({ ok: true });
      }
    }
    return ok({ error: "Secuestro no encontrado: " + b.id });
  }
  // ── M3C EDITAR ────────────────────────────────────────────
  if (mod === "m3c" && action === "editar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M3C_Llamados");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) { ci[hdrs[i]]=i; ci[String(hdrs[i]).toLowerCase()]=i; }
    var colId = ci["id"]!==undefined ? ci["id"] : 0;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][colId]||"") === String(b.id||"")) {
        var campos3c = ["fi","ff","turno","op","rows","u"];
        campos3c.forEach(function(k){
          if(ci[k]!==undefined) hoja.getRange(r+1,ci[k]+1).setValue(
            typeof b[k]==="object"?JSON.stringify(b[k]):b[k]||""
          );
        });
        return ok({ ok: true });
      }
    }
    return ok({ error: "Llamado no encontrado: " + b.id });
  }
  // ── M3C ELIMINAR ──────────────────────────────────────────
  if (mod === "m3c" && action === "eliminar") {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = getOrCreate(ss, "M3C_Llamados");
    var vals = hoja.getDataRange().getValues();
    var hdrs = vals[0]; var ci = {};
    for (var i = 0; i < hdrs.length; i++) { ci[hdrs[i]]=i; ci[String(hdrs[i]).toLowerCase()]=i; }
    var colId = ci["id"]!==undefined ? ci["id"] : 0;
    for (var r = vals.length-1; r >= 1; r--) {
      if (String(vals[r][colId]||"") === String(b.id||"")) {
        hoja.deleteRow(r+1);
        return ok({ ok: true });
      }
    }
    return ok({ error: "No encontrado" });
  }

  return ok({ error: "no procesado: " + mod + "/" + action });
}

// ── LOGIN ─────────────────────────────────────────────────────
function login(usuario, password) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  if (vals.length < 2) return ok({ error: "Sin usuarios registrados" });

  var h = {};
  for (var i = 0; i < vals[0].length; i++) h[vals[0][i]] = i;

  var u = String(usuario || "").trim().toLowerCase();
  var p = String(password || "").trim();

  for (var r = 1; r < vals.length; r++) {
    var ru = String(vals[r][h.usuario] || "").trim().toLowerCase();
    var rp = String(vals[r][h.password] || "").trim();
    var re = String(vals[r][h.estado] || "").trim().toLowerCase();
    if (ru === u && rp === p && re === "activo") {
      var na = [];
      try { na = JSON.parse(vals[r][h.no_acceso] || "[]"); } catch(x) {}
      var user = {
        id: vals[r][h.id], nombre: vals[r][h.nombre],
        n: vals[r][h.nombre], usuario: ru, u: ru,
        rol: vals[r][h.rol], r: vals[r][h.rol],
        estado: "activo", e: "activo",
        cedula: vals[r][h.cedula], c: vals[r][h.cedula],
        noAcceso: na, no_acceso: na
      };
      return ok({ ok: true, user: user });
    }
  }
  return ok({ error: "Usuario o contrasena incorrectos" });
}

// ── LISTAR HOJA ───────────────────────────────────────────────
function listarHoja(nombre, filtro) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, nombre);
  var vals = hoja.getDataRange().getValues();
  if (vals.length < 2) return ok([]);
  var hdrs = vals[0];
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var obj = {};
    for (var j = 0; j < hdrs.length; j++) obj[hdrs[j]] = vals[i][j];
    if (!filtro || filtro(obj)) rows.push(obj);
  }
  return ok(rows.reverse());
}

function listarM1() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "M1_Reportes");
  var vals = hoja.getDataRange().getValues();
  if (vals.length < 2) return ok([]);
  var hdrs = vals[0];
  var rows = [];
  // Columnas JSON: grupo, acciones, decomisos
  // ap1..ap33 son números directos (columnas individuales)
  var jsonCols = ["grupo","acciones","decomisos"];
  for (var i = 1; i < vals.length; i++) {
    var obj = {};
    for (var j = 0; j < hdrs.length; j++) {
      var key = hdrs[j];
      var val = vals[i][j];
      if (jsonCols.indexOf(key) >= 0) {
        try { obj[key] = JSON.parse(String(val) || "[]"); } catch(x) { obj[key] = []; }
      } else {
        obj[key] = val;
      }
    }
    rows.push(obj);
  }
  return ok(rows.reverse());
}

function listarUsuarios() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  if (vals.length < 2) return ok([]);
  var hdrs = vals[0];
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var obj = {};
    for (var j = 0; j < hdrs.length; j++) obj[hdrs[j]] = vals[i][j];
    if (String(obj.estado || "").trim() === "activo") {
      delete obj.password;
      try { obj.no_acceso = JSON.parse(obj.no_acceso || "[]"); } catch(x) { obj.no_acceso = []; }
      obj.noAcceso = obj.no_acceso;
      obj.n = obj.nombre; obj.u = obj.usuario; obj.r = obj.rol;
      rows.push(obj);
    }
  }
  return ok(rows);
}

function listarConsecutivos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Consecutivos");
  var vals = hoja.getDataRange().getValues();
  var c = { m1:1, m2:1, m3a:1, m4:1, m7:1 };
  for (var i = 1; i < vals.length; i++) c[vals[i][0]] = parseInt(vals[i][1]) || 1;
  return ok(c);
}

// ── CREAR REGISTRO ────────────────────────────────────────────
function crearRegistro(mod, b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hojaMap = {
    m1:"M1_Reportes", m2:"M2_Partes", m3a:"M3A_Bitacoras",
    m3b:"M3B_Secuestros", m3c:"M3C_Llamados", m4:"M4_Visitas", m7:"M7_Informes",
    m1_grupo:"M1_Grupo", m1_cuant:"M1_Cuant", m1_acciones:"M1_Acciones",
    m1_decomisos:"M1_Decomisos", m2_imp:"M2_Imputados", m2_of:"M2_Ofendidos",
    m2_asis:"M2_Asistentes", m2_mods:"M2_Modificaciones",
    m3a_inc:"M3A_Incidentes", auditoria:"Auditoria"
  };
  var nombre = hojaMap[mod];
  if (!nombre) return ok({ error: "hoja no mapeada: " + mod });

  var anio = new Date().getFullYear();
  var consec = getConsec(ss, mod);
  var id = generarId(mod, consec, anio);
  b.id = id;
  b.creado_en = Utilities.formatDate(new Date(), "America/Costa_Rica", "yyyy-MM-dd HH:mm:ss");

  var hoja = getOrCreate(ss, nombre);
  appendRow(hoja, b);
  incrementarConsec(ss, mod);

  return ok({ ok: true, id: id, consec: formatConsec(mod, consec + 1, anio) });
}

// ── USUARIOS ──────────────────────────────────────────────────
function crearUsuario(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var n = hoja.getLastRow();
  var ts = Utilities.formatDate(new Date(), "America/Costa_Rica", "yyyy-MM-dd");
  hoja.appendRow([n, b.nombre||"", b.cedula||"", b.usuario||"", b.password||"temporal123", b.rol||"policia", "activo", "[]", ts]);
  return ok({ ok: true });
}

function actualizarPermisos(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  var hdrs = vals[0]; var ci = {};
  for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][ci.id]) === String(b.id)) {
      hoja.getRange(r+1, ci.no_acceso+1).setValue(JSON.stringify(b.no_acceso||[]));
      return ok({ ok: true });
    }
  }
  return ok({ error: "no encontrado" });
}

function actualizarEstado(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  var hdrs = vals[0]; var ci = {};
  for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][ci.id]) === String(b.id)) {
      hoja.getRange(r+1, ci.estado+1).setValue(b.estado||"activo");
      return ok({ ok: true });
    }
  }
  return ok({ error: "no encontrado" });
}

function eliminarUsuario(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  var hdrs = vals[0]; var ci = {};
  for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
  for (var r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][ci.id]) === String(b.id)) {
      hoja.deleteRow(r+1);
      return ok({ ok: true });
    }
  }
  return ok({ error: "no encontrado" });
}

function cambiarPassword(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Usuarios");
  var vals = hoja.getDataRange().getValues();
  var hdrs = vals[0]; var ci = {};
  for (var i = 0; i < hdrs.length; i++) ci[hdrs[i]] = i;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][ci.usuario]||"").trim() === String(b.usuario||"").trim()) {
      if (String(vals[r][ci.password]||"").trim() !== String(b.actual||"").trim())
        return ok({ error: "Contrasena actual incorrecta" });
      hoja.getRange(r+1, ci.password+1).setValue(b.nueva||"");
      return ok({ ok: true });
    }
  }
  return ok({ error: "Usuario no encontrado" });
}

// ── RESET ─────────────────────────────────────────────────────
function resetTotal() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hojas = ["M1_Reportes","M1_Grupo","M1_Cuant","M1_Acciones","M1_Decomisos",
    "M2_Partes","M2_Imputados","M2_Ofendidos","M2_Asistentes","M2_Modificaciones",
    "M3A_Bitacoras","M3A_Incidentes","M3B_Secuestros","M3C_Llamados",
    "M4_Visitas","M7_Informes","M7_Historico","Auditoria"];
  for (var i = 0; i < hojas.length; i++) {
    var h = getOrCreate(ss, hojas[i]);
    if (h.getLastRow() > 1) h.deleteRows(2, h.getLastRow()-1);
  }
  var hc = getOrCreate(ss, "Consecutivos");
  var vc = hc.getDataRange().getValues();
  for (var r = 1; r < vc.length; r++) hc.getRange(r+1,2).setValue(1);
  return ok({ ok: true });
}

function resetModulos(mods) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var map = {
    m1:["M1_Reportes","M1_Grupo","M1_Cuant","M1_Acciones","M1_Decomisos"],
    m2:["M2_Partes","M2_Imputados","M2_Ofendidos","M2_Asistentes","M2_Modificaciones"],
    m3a:["M3A_Bitacoras","M3A_Incidentes"], m3b:["M3B_Secuestros"],
    m3c:["M3C_Llamados"], m4:["M4_Visitas"], m7:["M7_Informes","M7_Historico"],
    auditoria:["Auditoria"]
  };
  for (var i = 0; i < (mods||[]).length; i++) {
    var hs = map[mods[i]] || [];
    for (var j = 0; j < hs.length; j++) {
      var h = getOrCreate(ss, hs[j]);
      if (h.getLastRow() > 1) h.deleteRows(2, h.getLastRow()-1);
    }
  }
  return ok({ ok: true });
}

function actualizarConsecutivos(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = getOrCreate(ss, "Consecutivos");
  var vals = hoja.getDataRange().getValues();
  var keys = ["m1","m2","m3a","m4","m7"];
  for (var i = 0; i < keys.length; i++) {
    if (b[keys[i]]) {
      for (var r = 1; r < vals.length; r++) {
        if (vals[r][0] === keys[i]) {
          hoja.getRange(r+1,2).setValue(parseInt(b[keys[i]])||1);
          break;
        }
      }
    }
  }
  return ok({ ok: true });
}

// ── INICIALIZAR ───────────────────────────────────────────────
function inicializar() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var nombres = ["Usuarios","Consecutivos","Auditoria",
    "M1_Reportes","M1_Grupo","M1_Cuant","M1_Acciones","M1_Decomisos",
    "M2_Partes","M2_Imputados","M2_Ofendidos","M2_Asistentes","M2_Modificaciones",
    "M3A_Bitacoras","M3A_Incidentes","M3B_Secuestros","M3C_Llamados",
    "M4_Visitas","M7_Informes","M7_Historico"];
  for (var i = 0; i < nombres.length; i++) getOrCreate(ss, nombres[i]);

  var hc = getOrCreate(ss, "Consecutivos");
  if (hc.getLastRow() === 0) {
    hc.appendRow(["modulo","valor"]);
    hc.appendRow(["m1",1]); hc.appendRow(["m2",1]);
    hc.appendRow(["m3a",1]); hc.appendRow(["m4",1]); hc.appendRow(["m7",1]);
  }

  var hu = getOrCreate(ss, "Usuarios");
  hu.clearContents();
  var u = [
    ["id","nombre","cedula","usuario","password","rol","estado","no_acceso","creado_en"],
    [1,"Jeison Arce Campos","7-0194-0488","jarce","admin2026","admin","activo","[]","2026-01-01"],
    [2,"Melvin Manuel Benavides Sandoval","3-0337-0284","mbenavides","pml2026","policia","activo","[]","2026-01-01"],
    [3,"Jehudy Alexander Huertas Gutierrez","7-0190-0867","jhuertas","pml2026","preventivo","activo","[]","2026-01-01"],
    [4,"Jacqueline Andrea Guillen Jimenez","7-0191-0931","jguillen","pml2026","policia","activo","[]","2026-01-01"],
    [5,"Jefferson Antonio Picado Matus","1-1486-0038","jpicado","pml2026","policia","activo","[]","2026-01-01"],
    [6,"Karla Johanna Alfaro Campos","7-0144-0657","kalfaro","pml2026","policia","activo","[]","2026-01-01"],
    [7,"Christopher Alvarado Rios","7-0179-0323","calvarado","pml2026","preventivo","activo","[]","2026-01-01"],
    [8,"Klisman Cristobal Mora Diaz","7-0229-0905","kmora","pml2026","policia","activo","[]","2026-01-01"],
    [9,"Kebir Jafeth Blear Araya","7-0221-0923","kblear","pml2026","policia","activo","[]","2026-01-01"],
    [10,"Alberto Gabriel Rubi Arias","7-0204-0627","arubi","pml2026","policia","activo","[]","2026-01-01"],
    [11,"Efrain Jesus Chavarria Bellido","5-0392-0887","echarvarria","pml2026","policia","activo","[]","2026-01-01"],
    [12,"Winier Alberto Escobar Vargas","7-0260-0077","wescobar","pml2026","policia","activo","[]","2026-01-01"],
    [13,"Mirta Hibbert Brown","7-0091-0597","mhibbert","dse2026","jefa_dse","activo","[]","2026-01-01"],
    [14,"Yassir Gomez Oses","7-0284-0932","ygomez","dse2026","operador","activo","[]","2026-01-01"],
    [15,"Guivanny Mina Reid","7-0211-0942","gmina","dse2026","operador","activo","[]","2026-01-01"],
    [16,"Andrey Gutierrez Jaen","7-0251-0679","agutierrez","dse2026","operador","activo","[]","2026-01-01"],
    [17,"Warren Bloonfield Richards","7-0288-0153","wbloonfield","dse2026","operador","activo","[]","2026-01-01"],
    [18,"Johnny Araya Saavedra","7-0112-0698","jaraya","dse2026","operador","activo","[]","2026-01-01"],
    [19,"Wilberth Mora Azofeifa","2-0769-0644","wmora","dse2026","operador","activo","[]","2026-01-01"],
    [20,"Joshua Sterling Maitland","7-0306-0394","jsterling","dse2026","operador","activo","[]","2026-01-01"],
    [21,"Aaron Andre Acuna Arroyo","1-0897-0292","aacuna","esto2026","estacionometro","activo","[]","2026-01-01"],
    [22,"Shaylin Yolieth Harvey Norman","7-0193-0912","sharvey","esto2026","estacionometro","activo","[]","2026-01-01"]
  ];
  hu.getRange(1,1,u.length,u[0].length).setValues(u);
  return ok({ ok: true, msg: "Sistema inicializado con " + (u.length-1) + " usuarios" });
}

// ── UTILIDADES ────────────────────────────────────────────────
function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function appendRow(hoja, obj) {
  var vals = hoja.getDataRange().getValues();
  var hdrs;
  if (vals.length === 0) {
    hdrs = Object.keys(obj);
    hoja.appendRow(hdrs);
  } else {
    hdrs = vals[0];
    var newC = Object.keys(obj).filter(function(k){return hdrs.indexOf(k)<0;});
    for (var i=0;i<newC.length;i++){hoja.getRange(1,hdrs.length+1).setValue(newC[i]);hdrs.push(newC[i]);}
  }
  var row = hdrs.map(function(h){
    var v=obj[h]; if(v===null||v===undefined) return "";
    if(typeof v==="object") return JSON.stringify(v); return v;
  });
  hoja.appendRow(row);
}

function getConsec(ss, mod) {
  var h = getOrCreate(ss,"Consecutivos");
  var v = h.getDataRange().getValues();
  for (var i=1;i<v.length;i++) if(String(v[i][0])===mod) return parseInt(v[i][1])||1;
  return 1;
}

function incrementarConsec(ss, mod) {
  var h = getOrCreate(ss,"Consecutivos");
  var v = h.getDataRange().getValues();
  for (var i=1;i<v.length;i++) if(String(v[i][0])===mod){h.getRange(i+1,2).setValue((parseInt(v[i][1])||1)+1);return;}
}

function generarId(mod, c, anio) {
  var p=("000"+c).slice(-3);
  if(mod==="m1") return "RDL-"+anio+"-"+p;
  if(mod==="m2") return "PP-"+anio+"-"+p;
  if(mod==="m3a") return "BIT-"+anio+"-"+p;
  if(mod==="m4") return "PP-PREV-"+anio+"-"+p;
  if(mod==="m7") return "INF-"+p+"-GSCV-DE-"+anio;
  return mod+"-"+anio+"-"+p;
}

function formatConsec(mod, c, anio) {
  var p=("000"+c).slice(-3);
  if(mod==="m1") return "No RDL-"+anio+"-"+p;
  if(mod==="m2") return "No PP-"+anio+"-"+p;
  if(mod==="m3a") return "No BIT-"+anio+"-"+p;
  if(mod==="m4") return "No PP-PREV-"+anio+"-"+p;
  if(mod==="m7") return "No INF-"+p+"-GSCV-DE-"+anio;
  return p;
}


// ══════════════════════════════════════════════════════════════
// MÓDULO CONTROL DE DOCUMENTACIÓN  (mod = "docs")
// Bloque independiente: no usa ni modifica funciones de otros módulos.
// Hojas propias, creadas automáticamente al primer uso:
//   DOC_Ingresos, DOC_Asignaciones, DOC_Informes, DOC_Bitacora,
//   DOC_Feriados, DOC_Config
// Todas las funciones y variables de este bloque llevan el prefijo "doc"/"DOC_".
// ══════════════════════════════════════════════════════════════
var DOC_TZ = "America/Costa_Rica";
var DOC_SUFIJO = "-INF-PML-";          // Consecutivo: 001-INF-PML-2026
var DOC_MEDIOS = ["Correo", "Ventanilla Única", "Otro"];
var DOC_ROLES_OFICIAL = ["policia", "preventivo", "admin"];
var DOC_HOJAS = {
  ingresos: { nombre: "DOC_Ingresos", cols: [
    "id", "anio", "consec", "fecha_recepcion", "hora_recepcion", "remitente", "asunto",
    "medio", "medio_otro", "plazo_dias", "tipo_dias", "fecha_vencimiento", "estado",
    "registrado_por", "registrado_nombre", "creado_en",
    "fecha_cierre", "cierre_en_plazo", "respuesta", "respuesta_por", "respuesta_nombre"
  ]},
  asignaciones: { nombre: "DOC_Asignaciones", cols: [
    "doc_id", "usuario", "nombre", "cedula", "asignado_ts", "asignado_por",
    "activa", "retirado_ts", "retirado_por"
  ]},
  informes: { nombre: "DOC_Informes", cols: [
    "doc_id", "usuario", "nombre", "cedula", "texto", "estado",
    "creado_ts", "actualizado_ts", "presentado_ts"
  ]},
  bitacora: { nombre: "DOC_Bitacora", cols: ["ts", "usuario", "accion", "doc_id", "detalle"] },
  feriados: { nombre: "DOC_Feriados", cols: ["fecha", "descripcion"] },
  config:   { nombre: "DOC_Config",   cols: ["clave", "valor"] }
};

// ── Despacho ─────────────────────────────────────────────────
function docsGet(p) {
  try {
    docsAsegurar();
    var yo = docsUsuario(p.usuario);
    if (!yo) return ok({ error: "Usuario no valido o inactivo" });
    var accion = p.action || "listar";
    if (accion === "listar") return ok(docsDatos(yo));
    if (accion === "bitacora") {
      if (yo.rol !== "admin") return ok({ error: "Sin permiso" });
      var b = docsLeer("bitacora").filter(function (r) { return r.doc_id === String(p.doc_id || ""); });
      b.forEach(function (r) { delete r._fila; });
      return ok({ bitacora: b.reverse() });
    }
    return ok({ error: "accion no valida: " + accion });
  } catch (e) {
    return ok({ error: "docs: " + e.message });
  }
}

function docsPost(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return ok({ error: "Servidor ocupado, intente de nuevo" }); }
  try {
    docsAsegurar();
    var yo = docsUsuario(b.usuario);
    if (!yo) return ok({ error: "Usuario no valido o inactivo" });
    switch (String(b.action || "")) {
      case "registrar":         return docsRegistrar(yo, b);
      case "asignar":           return docsAsignar(yo, b);
      case "desasignar":        return docsDesasignar(yo, b);
      case "informe_guardar":   return docsInformeGuardar(yo, b);
      case "informe_devolver":  return docsInformeDevolver(yo, b);
      case "cerrar":            return docsCerrar(yo, b);
      case "config_guardar":    return docsConfigGuardar(yo, b);
      case "feriado_agregar":   return docsFeriadoAgregar(yo, b);
      case "feriado_eliminar":  return docsFeriadoEliminar(yo, b);
    }
    return ok({ error: "accion no valida: " + b.action });
  } catch (e) {
    return ok({ error: "docs: " + e.message });
  } finally {
    lock.releaseLock();
  }
}

// Opcional: ejecutar una vez desde el editor para crear las hojas sin abrir el módulo.
function docsInstalar() { docsAsegurar(); }

// ── Lectura para el cliente ──────────────────────────────────
function docsDatos(yo) {
  var admin = yo.rol === "admin";
  var ing = docsLeer("ingresos");
  var asg = docsLeer("asignaciones");
  var inf = docsLeer("informes");

  var mis = {};
  asg.forEach(function (a) { if (a.usuario === yo.usuario && a.activa === "SI") mis[a.doc_id] = true; });

  var docs = ing.filter(function (d) { return admin || mis[d.id]; });
  var ids = {};
  docs.forEach(function (d) { ids[d.id] = true; });

  var asgV = asg.filter(function (a) { return ids[a.doc_id] && (admin || a.usuario === yo.usuario); });
  var infV = inf.filter(function (i) { return ids[i.doc_id] && (admin || i.usuario === yo.usuario); });
  // El administrador ve el texto solo de informes ya presentados (no de borradores).
  if (admin) infV.forEach(function (i) { if (i.estado !== "presentado") i.texto = ""; });

  [docs, asgV, infV].forEach(function (arr) { arr.forEach(function (r) { delete r._fila; }); });

  var cfg = docsConfigMapa();
  var anio = docsHoy().substr(0, 4);
  var fer = docsLeer("feriados").map(function (f) { return { fecha: f.fecha, descripcion: f.descripcion }; });
  return {
    rol: yo.rol, usuario: yo.usuario,
    docs: docs, asignaciones: asgV, informes: infV, feriados: fer,
    config: {
      umbral_amarillo: parseInt(cfg.umbral_amarillo, 10) || 70,
      proximo_consec: parseInt(cfg["consec_" + anio], 10) || 1,
      anio: anio
    }
  };
}

// ── Acciones ─────────────────────────────────────────────────
function docsRegistrar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede registrar ingresos" });
  var rem = docsLimpiar(b.remitente, 200);
  var asu = docsLimpiar(b.asunto, 1500);
  var medio = String(b.medio || "").trim();
  var otro = docsLimpiar(b.medio_otro, 100);
  var tipo = String(b.tipo_dias || "");
  var plazo = parseInt(b.plazo_dias, 10);
  var fecha = String(b.fecha_recepcion || "");
  var hora = String(b.hora_recepcion || "");
  var hoy = docsHoy();

  if (!rem) return ok({ error: "Indique el remitente" });
  if (!asu) return ok({ error: "Indique el asunto" });
  if (DOC_MEDIOS.indexOf(medio) < 0) return ok({ error: "Medio de recepcion no valido" });
  if (medio === "Otro" && !otro) return ok({ error: "Especifique el medio de recepcion" });
  if (medio !== "Otro") otro = "";
  if (!(plazo >= 1 && plazo <= 3650)) return ok({ error: "El plazo debe ser de 1 a 3650 dias" });
  if (tipo !== "habiles" && tipo !== "naturales") return ok({ error: "Indique si el plazo es en dias habiles o naturales" });
  if (!docsParseFecha(fecha)) return ok({ error: "Fecha de recepcion no valida" });
  if (fecha > hoy) return ok({ error: "La fecha de recepcion no puede ser futura" });
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(hora)) return ok({ error: "Hora de recepcion no valida" });

  var venc = docsVencimiento(fecha, plazo, tipo, docsFeriadosSet());
  var anio = parseInt(hoy.substr(0, 4), 10);
  var ing = docsLeer("ingresos");
  var n = docsSiguienteConsec(anio, ing);
  var id = docsId(n, anio);

  docsAgregar("ingresos", {
    id: id, anio: anio, consec: n, fecha_recepcion: fecha, hora_recepcion: hora,
    remitente: rem, asunto: asu, medio: medio, medio_otro: otro,
    plazo_dias: plazo, tipo_dias: tipo, fecha_vencimiento: venc, estado: "Registrado",
    registrado_por: yo.usuario, registrado_nombre: yo.nombre, creado_en: docsAhora(),
    fecha_cierre: "", cierre_en_plazo: "", respuesta: "", respuesta_por: "", respuesta_nombre: ""
  });
  docsLog(yo.usuario, "registrar", id, "Remitente: " + rem + " | Plazo: " + plazo + " " + tipo + " | Vence: " + venc);
  return ok({ ok: true, id: id });
}

function docsAsignar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede asignar" });
  var doc = docsBuscarDoc(b.doc_id);
  if (!doc) return ok({ error: "Gestion no encontrada" });
  if (doc.estado === "Cerrado") return ok({ error: "La gestion ya esta cerrada" });
  var lista = b.oficiales;
  if (!lista || !lista.length) return ok({ error: "Seleccione al menos un oficial" });

  var asg = docsLeer("asignaciones");
  var activos = {};
  asg.forEach(function (a) { if (a.doc_id === doc.id && a.activa === "SI") activos[a.usuario] = true; });

  var nuevos = [], omitidos = [];
  for (var i = 0; i < lista.length; i++) {
    var u = docsUsuario(lista[i]);
    if (!u || DOC_ROLES_OFICIAL.indexOf(u.rol) < 0) { omitidos.push(String(lista[i])); continue; }
    if (activos[u.usuario]) continue;
    activos[u.usuario] = true;
    nuevos.push(u);
  }
  if (!nuevos.length) return ok({ error: "Ningun oficial valido para asignar" + (omitidos.length ? " (sin usuario activo: " + omitidos.join(", ") + ")" : "") });

  var ts = docsAhora();
  nuevos.forEach(function (u) {
    docsAgregar("asignaciones", {
      doc_id: doc.id, usuario: u.usuario, nombre: u.nombre, cedula: u.cedula,
      asignado_ts: ts, asignado_por: yo.usuario, activa: "SI", retirado_ts: "", retirado_por: ""
    });
    docsLog(yo.usuario, "asignar", doc.id, "Asignado a " + u.nombre + " (" + u.usuario + ")");
  });
  docsRecalcEstado(doc.id);
  return ok({ ok: true, asignados: nuevos.length, omitidos: omitidos });
}

function docsDesasignar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede retirar asignaciones" });
  var doc = docsBuscarDoc(b.doc_id);
  if (!doc) return ok({ error: "Gestion no encontrada" });
  if (doc.estado === "Cerrado") return ok({ error: "La gestion ya esta cerrada" });
  var u = String(b.oficial || "").trim().toLowerCase();

  var inf = docsLeer("informes");
  for (var k = 0; k < inf.length; k++) {
    if (inf[k].doc_id === doc.id && inf[k].usuario === u && inf[k].estado === "presentado")
      return ok({ error: "El oficial ya presento su informe; devuelvalo antes de retirar la asignacion" });
  }
  var asg = docsLeer("asignaciones");
  for (var i = 0; i < asg.length; i++) {
    var a = asg[i];
    if (a.doc_id === doc.id && a.usuario === u && a.activa === "SI") {
      docsActualizar("asignaciones", a._fila, { activa: "NO", retirado_ts: docsAhora(), retirado_por: yo.usuario });
      docsLog(yo.usuario, "desasignar", doc.id, "Retirada la asignacion de " + a.nombre + " (" + u + ")");
      docsRecalcEstado(doc.id);
      return ok({ ok: true });
    }
  }
  return ok({ error: "Asignacion no encontrada" });
}

function docsInformeGuardar(yo, b) {
  var doc = docsBuscarDoc(b.doc_id);
  if (!doc) return ok({ error: "Gestion no encontrada" });
  if (doc.estado === "Cerrado") return ok({ error: "La gestion ya esta cerrada" });
  var asg = docsLeer("asignaciones");
  var asignado = false;
  for (var i = 0; i < asg.length; i++)
    if (asg[i].doc_id === doc.id && asg[i].usuario === yo.usuario && asg[i].activa === "SI") asignado = true;
  if (!asignado) return ok({ error: "Usted no esta asignado a esta gestion" });

  var texto = String(b.texto || "").replace(/\r\n/g, "\n").trim();
  if (!texto) return ok({ error: "El informe no puede estar vacio" });
  if (texto.length > 45000) return ok({ error: "El informe excede el largo maximo permitido" });
  var presentar = (b.presentar === true || b.presentar === "true");

  var ts = docsAhora();
  var inf = docsLeer("informes");
  var fila = null, actual = null;
  for (var j = 0; j < inf.length; j++)
    if (inf[j].doc_id === doc.id && inf[j].usuario === yo.usuario) { fila = inf[j]._fila; actual = inf[j]; }

  if (actual && actual.estado === "presentado")
    return ok({ error: "El informe ya fue presentado; solicite su devolucion al administrador" });

  if (actual) {
    docsActualizar("informes", fila, {
      texto: texto, actualizado_ts: ts,
      estado: presentar ? "presentado" : "borrador",
      presentado_ts: presentar ? ts : ""
    });
  } else {
    docsAgregar("informes", {
      doc_id: doc.id, usuario: yo.usuario, nombre: yo.nombre, cedula: yo.cedula, texto: texto,
      estado: presentar ? "presentado" : "borrador",
      creado_ts: ts, actualizado_ts: ts, presentado_ts: presentar ? ts : ""
    });
  }
  docsLog(yo.usuario, presentar ? "informe_presentado" : "informe_borrador", doc.id, yo.nombre);
  docsRecalcEstado(doc.id);
  return ok({ ok: true, estado: presentar ? "presentado" : "borrador" });
}

function docsInformeDevolver(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede devolver informes" });
  var doc = docsBuscarDoc(b.doc_id);
  if (!doc) return ok({ error: "Gestion no encontrada" });
  if (doc.estado === "Cerrado") return ok({ error: "La gestion ya esta cerrada" });
  var u = String(b.oficial || "").trim().toLowerCase();
  var inf = docsLeer("informes");
  for (var i = 0; i < inf.length; i++) {
    if (inf[i].doc_id === doc.id && inf[i].usuario === u && inf[i].estado === "presentado") {
      docsActualizar("informes", inf[i]._fila, { estado: "borrador", presentado_ts: "", actualizado_ts: docsAhora() });
      docsLog(yo.usuario, "informe_devuelto", doc.id, "Devuelto a " + inf[i].nombre + " (" + u + ")");
      docsRecalcEstado(doc.id);
      return ok({ ok: true });
    }
  }
  return ok({ error: "No hay un informe presentado de ese oficial" });
}

function docsCerrar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede cerrar la gestion" });
  var doc = docsBuscarDoc(b.doc_id);
  if (!doc) return ok({ error: "Gestion no encontrada" });
  if (doc.estado === "Cerrado") return ok({ error: "La gestion ya esta cerrada" });

  var resp = String(b.respuesta || "").replace(/\r\n/g, "\n").trim();
  if (!resp) return ok({ error: "La respuesta no puede estar vacia" });
  if (resp.length > 45000) return ok({ error: "La respuesta excede el largo maximo permitido" });

  var asg = docsLeer("asignaciones"), activos = {};
  asg.forEach(function (a) { if (a.doc_id === doc.id && a.activa === "SI") activos[a.usuario] = true; });
  var hayInforme = docsLeer("informes").some(function (i) {
    return i.doc_id === doc.id && i.estado === "presentado" && activos[i.usuario];
  });
  if (!hayInforme) return ok({ error: "Se requiere al menos un informe presentado para cerrar la gestion" });

  var hoy = docsHoy();
  var venc = docsVencimiento(doc.fecha_recepcion, parseInt(doc.plazo_dias, 10), doc.tipo_dias, docsFeriadosSet());
  var enPlazo = hoy <= venc ? "SI" : "NO";
  docsActualizar("ingresos", doc._fila, {
    estado: "Cerrado", fecha_vencimiento: venc, fecha_cierre: docsAhora(), cierre_en_plazo: enPlazo,
    respuesta: resp, respuesta_por: yo.usuario, respuesta_nombre: yo.nombre
  });
  docsLog(yo.usuario, "cerrar", doc.id, "Cierre " + (enPlazo === "SI" ? "dentro" : "fuera") + " del plazo (vencia " + venc + ")");
  return ok({ ok: true, cierre_en_plazo: enPlazo });
}

function docsConfigGuardar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede cambiar la configuracion" });
  var detalle = [];
  if (b.umbral_amarillo !== undefined && b.umbral_amarillo !== "") {
    var u = parseInt(b.umbral_amarillo, 10);
    if (!(u >= 1 && u <= 99)) return ok({ error: "El umbral debe estar entre 1 y 99" });
    docsConfigSet("umbral_amarillo", String(u));
    detalle.push("umbral amarillo = " + u + "%");
  }
  if (b.proximo_consec !== undefined && b.proximo_consec !== "") {
    var c = parseInt(b.proximo_consec, 10);
    if (!(c >= 1 && c <= 999999)) return ok({ error: "Consecutivo no valido" });
    var anio = docsHoy().substr(0, 4);
    docsConfigSet("consec_" + anio, String(c));
    detalle.push("proximo consecutivo " + anio + " = " + c);
  }
  if (!detalle.length) return ok({ error: "Nada que guardar" });
  docsLog(yo.usuario, "config", "", detalle.join(" | "));
  return ok({ ok: true });
}

function docsFeriadoAgregar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede editar feriados" });
  var f = String(b.fecha || "");
  if (!docsParseFecha(f)) return ok({ error: "Fecha no valida" });
  var desc = docsLimpiar(b.descripcion, 100);
  var fer = docsLeer("feriados");
  for (var i = 0; i < fer.length; i++) if (fer[i].fecha === f) return ok({ error: "Ese feriado ya esta registrado" });
  docsAgregar("feriados", { fecha: f, descripcion: desc });
  docsLog(yo.usuario, "feriado_agregar", "", f + " " + desc);
  docsRecalcAbiertos();
  return ok({ ok: true });
}

function docsFeriadoEliminar(yo, b) {
  if (yo.rol !== "admin") return ok({ error: "Solo el administrador puede editar feriados" });
  var f = String(b.fecha || "");
  var fer = docsLeer("feriados");
  for (var i = 0; i < fer.length; i++) {
    if (fer[i].fecha === f) {
      docsHoja("feriados").deleteRow(fer[i]._fila);
      docsLog(yo.usuario, "feriado_eliminar", "", f + " " + fer[i].descripcion);
      docsRecalcAbiertos();
      return ok({ ok: true });
    }
  }
  return ok({ error: "Feriado no encontrado" });
}

// ── Estado y plazos ──────────────────────────────────────────
function docsRecalcEstado(docId) {
  var doc = docsBuscarDoc(docId);
  if (!doc || doc.estado === "Cerrado") return;
  var activos = {}, hayAsg = false;
  docsLeer("asignaciones").forEach(function (a) {
    if (a.doc_id === docId && a.activa === "SI") { activos[a.usuario] = true; hayAsg = true; }
  });
  var hayInf = docsLeer("informes").some(function (i) {
    return i.doc_id === docId && i.estado === "presentado" && activos[i.usuario];
  });
  var nuevo = hayInf ? "Informe presentado" : (hayAsg ? "Asignado" : "Registrado");
  if (nuevo !== doc.estado) docsActualizar("ingresos", doc._fila, { estado: nuevo });
}

// Si cambia la tabla de feriados, se recalcula el vencimiento de las gestiones abiertas.
function docsRecalcAbiertos() {
  var fer = docsFeriadosSet();
  docsLeer("ingresos").forEach(function (d) {
    if (d.estado === "Cerrado") return;
    var v = docsVencimiento(d.fecha_recepcion, parseInt(d.plazo_dias, 10), d.tipo_dias, fer);
    if (v && v !== d.fecha_vencimiento) docsActualizar("ingresos", d._fila, { fecha_vencimiento: v });
  });
}

function docsParseFecha(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ""));
  if (!m) return null;
  var y = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, d = parseInt(m[3], 10);
  var dt = new Date(Date.UTC(y, mo, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo || dt.getUTCDate() !== d) return null;
  return dt;
}
function docsFmtFecha(dt) {
  return dt.getUTCFullYear() + "-" + ("0" + (dt.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + dt.getUTCDate()).slice(-2);
}
function docsSumarDias(dt, n) { return new Date(dt.getTime() + n * 86400000); }
function docsEsHabil(dt, fer) {
  var w = dt.getUTCDay();
  return w !== 0 && w !== 6 && !fer[docsFmtFecha(dt)];
}
// El dia de recepcion no cuenta: el plazo corre desde el dia (habil) siguiente.
function docsVencimiento(rec, n, tipo, fer) {
  var dt = docsParseFecha(rec);
  if (!dt || !(n >= 1)) return "";
  if (tipo === "naturales") return docsFmtFecha(docsSumarDias(dt, n));
  var c = 0;
  while (c < n) { dt = docsSumarDias(dt, 1); if (docsEsHabil(dt, fer)) c++; }
  return docsFmtFecha(dt);
}

// ── Hojas y utilidades propias ───────────────────────────────
function docsAsegurar() {
  for (var k in DOC_HOJAS) docsHoja(k);
}

function docsHoja(key) {
  var def = DOC_HOJAS[key];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var h = ss.getSheetByName(def.nombre);
  if (!h) h = ss.insertSheet(def.nombre);
  if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, def.cols.length).setNumberFormat("@").setValues([def.cols]).setFontWeight("bold");
    h.getRange(2, 1, Math.max(h.getMaxRows() - 1, 1), def.cols.length).setNumberFormat("@");
    h.setFrozenRows(1);
  }
  return h;
}

function docsTxt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, DOC_TZ, "yyyy-MM-dd HH:mm:ss");
  return (v === null || v === undefined) ? "" : String(v);
}

function docsLeer(key) {
  var def = DOC_HOJAS[key], h = docsHoja(key);
  var n = h.getLastRow();
  if (n < 2) return [];
  var vals = h.getRange(2, 1, n - 1, def.cols.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _fila: i + 2 };
    for (var j = 0; j < def.cols.length; j++) o[def.cols[j]] = docsTxt(vals[i][j]);
    out.push(o);
  }
  return out;
}

// Escribe siempre con formato de texto para que Sheets no convierta fechas/numeros.
function docsAgregar(key, obj) {
  var def = DOC_HOJAS[key], h = docsHoja(key);
  var fila = h.getLastRow() + 1;
  var row = def.cols.map(function (c) { return docsTxt(obj[c]); });
  h.getRange(fila, 1, 1, def.cols.length).setNumberFormat("@").setValues([row]);
}

function docsActualizar(key, fila, cambios) {
  var def = DOC_HOJAS[key], h = docsHoja(key);
  var rng = h.getRange(fila, 1, 1, def.cols.length);
  var row = rng.getValues()[0];
  for (var j = 0; j < def.cols.length; j++) {
    if (cambios.hasOwnProperty(def.cols[j])) row[j] = docsTxt(cambios[def.cols[j]]);
    else row[j] = docsTxt(row[j]);
  }
  rng.setNumberFormat("@").setValues([row]);
}

function docsBuscarDoc(id) {
  var s = String(id || "");
  var l = docsLeer("ingresos");
  for (var i = 0; i < l.length; i++) if (l[i].id === s) return l[i];
  return null;
}

function docsUsuario(usuario) {
  var u = String(usuario || "").trim().toLowerCase();
  if (!u) return null;
  var hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Usuarios");
  if (!hoja) return null;
  var vals = hoja.getDataRange().getValues();
  if (vals.length < 2) return null;
  var h = {};
  for (var i = 0; i < vals[0].length; i++) h[vals[0][i]] = i;
  for (var r = 1; r < vals.length; r++) {
    var ru = String(vals[r][h.usuario] || "").trim().toLowerCase();
    var re = String(vals[r][h.estado] || "").trim().toLowerCase();
    if (ru === u && re === "activo") {
      return {
        usuario: ru,
        nombre: String(vals[r][h.nombre] || ""),
        cedula: String(vals[r][h.cedula] || ""),
        rol: String(vals[r][h.rol] || "").trim()
      };
    }
  }
  return null;
}

function docsConfigMapa() {
  var m = {};
  docsLeer("config").forEach(function (r) { m[r.clave] = r.valor; });
  return m;
}
function docsConfigSet(clave, valor) {
  var l = docsLeer("config");
  for (var i = 0; i < l.length; i++) {
    if (l[i].clave === clave) { docsActualizar("config", l[i]._fila, { valor: valor }); return; }
  }
  docsAgregar("config", { clave: clave, valor: valor });
}

function docsFeriadosSet() {
  var s = {};
  docsLeer("feriados").forEach(function (f) { if (docsParseFecha(f.fecha)) s[f.fecha] = true; });
  return s;
}

function docsId(n, anio) {
  var s = String(n);
  while (s.length < 3) s = "0" + s;
  return s + DOC_SUFIJO + anio;
}
// Consecutivo por año, protegido por el LockService de docsPost. Salta ids ya usados.
function docsSiguienteConsec(anio, ingresos) {
  var usados = {};
  ingresos.forEach(function (d) { usados[d.id] = true; });
  var n = parseInt(docsConfigMapa()["consec_" + anio], 10) || 1;
  while (usados[docsId(n, anio)]) n++;
  docsConfigSet("consec_" + anio, String(n + 1));
  return n;
}

function docsLimpiar(v, max) {
  return String(v === null || v === undefined ? "" : v).replace(/\r\n/g, "\n").trim().substr(0, max);
}
function docsHoy() { return Utilities.formatDate(new Date(), DOC_TZ, "yyyy-MM-dd"); }
function docsAhora() { return Utilities.formatDate(new Date(), DOC_TZ, "yyyy-MM-dd HH:mm:ss"); }
function docsLog(usuario, accion, docId, detalle) {
  docsAgregar("bitacora", { ts: docsAhora(), usuario: usuario, accion: accion, doc_id: docId, detalle: detalle });
}
