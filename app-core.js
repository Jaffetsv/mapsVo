(function () {
  "use strict";

  const DATA_FILES = Array.isArray(window.DLSRMAPS_DATA_FILES) ? window.DLSRMAPS_DATA_FILES.slice() : [];
  const DATA_VERSION = window.DLSRMAPS_DATA_VERSION || String(Date.now());

  let workbookPromise = null;
  const sheetCache = new Map();

  const ALIAS = {
    referencia: ["REFERENCIA", "Referencia", "referencia", "PLACA", "Placa", "placa", "CODIGO", "Código", "codigo", "Ref", "REF"],
    contrato: ["Contrato", "CONTRATO", "contrato", "NIS", "nis", "Cuenta", "CUENTA", "No Contrato", "N° Contrato", "Numero Contrato", "Número Contrato"],
    lat: ["Latitud", "LATITUD", "latitud", "Lat", "LAT", "Latitude", "Y"],
    lon: ["Longitud", "LONGITUD", "longitud", "Lon", "LON", "Lng", "LNG", "Long", "LONG", "Longitude", "X"],
    alimentador: ["Alimentador", "ALIMENTADOR", "alimentador", "Feeder", "Circuito", "CIRCUITO"]
  };

  function normalizarTexto(valor) {
    return String(valor ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toLowerCase();
  }

  function normalizarClave(valor) {
    return normalizarTexto(valor).replace(/[^a-z0-9]/g, "");
  }

  function soloDigitos(valor) {
    return String(valor ?? "").replace(/\D/g, "");
  }

  function fechaArchivo(nombre) {
    const m = String(nombre).match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (!m) return 0;
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
  }

  function archivosOrdenados() {
    return DATA_FILES.slice().sort((a, b) => fechaArchivo(b) - fechaArchivo(a) || String(b).localeCompare(String(a)));
  }

  function archivoUrl(nombre) {
    return "./" + encodeURI(nombre) + "?v=" + encodeURIComponent(DATA_VERSION);
  }

  async function cargarWorkbook() {
    if (workbookPromise) return workbookPromise;
    workbookPromise = (async () => {
      for (const archivo of archivosOrdenados()) {
        try {
          const response = await fetch(archivoUrl(archivo), { cache: "no-store" });
          if (!response.ok) continue;
          const arrayBuffer = await response.arrayBuffer();
          return XLSX.read(arrayBuffer, { type: "array", cellDates: false });
        } catch (_) {}
      }
      throw new Error("No fue posible preparar los datos de búsqueda.");
    })();
    return workbookPromise;
  }

  function obtenerHoja(workbook, nombresPosibles) {
    const buscadas = nombresPosibles.map(normalizarClave);
    const nombreHoja = workbook.SheetNames.find(n => buscadas.includes(normalizarClave(n)));
    if (!nombreHoja) throw new Error("No se encontró la hoja requerida para esta búsqueda.");
    return workbook.Sheets[nombreHoja];
  }

  function obtenerCampo(fila, alias) {
    const claves = Object.keys(fila || {});
    const buscadas = alias.map(normalizarClave);
    const clave = claves.find(k => buscadas.includes(normalizarClave(k)));
    return clave ? fila[clave] : "";
  }

  function parseCoord(valor) {
    if (valor === null || valor === undefined || valor === "") return NaN;
    const limpio = String(valor).replace(",", ".").replace(/[^0-9.\-]/g, "");
    return parseFloat(limpio);
  }

  function coordValida(lat, lon) {
    return Number.isFinite(lat) && Number.isFinite(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
  }

  function htmlSeguro(valor) {
    return String(valor ?? "").replace(/[&<>'"]/g, c => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "'": "&#39;",
      '"': "&quot;"
    }[c]));
  }

  function filaPlaca(fila) {
    const lat = parseCoord(obtenerCampo(fila, ALIAS.lat));
    const lon = parseCoord(obtenerCampo(fila, ALIAS.lon));
    const referencia = obtenerCampo(fila, ALIAS.referencia);
    return {
      referencia,
      lat,
      lon,
      alimentador: obtenerCampo(fila, ALIAS.alimentador) || "Sin alimentador"
    };
  }

  function filaContrato(fila) {
    const lat = parseCoord(obtenerCampo(fila, ALIAS.lat));
    const lon = parseCoord(obtenerCampo(fila, ALIAS.lon));
    const contrato = obtenerCampo(fila, ALIAS.contrato);
    const referencia = obtenerCampo(fila, ALIAS.referencia);
    return {
      contrato,
      referencia,
      lat,
      lon,
      alimentador: obtenerCampo(fila, ALIAS.alimentador) || "Sin alimentador"
    };
  }

  function crearIndiceExacto(items, campo) {
    const exacto = new Map();
    for (const item of items) {
      const key = normalizarClave(item[campo]);
      if (key && !exacto.has(key)) exacto.set(key, item);
    }
    return exacto;
  }

  function crearIndiceBusqueda(items, campo) {
    return items
      .filter(item => item && item[campo])
      .map(item => ({
        item,
        label: String(item[campo]),
        key: normalizarClave(item[campo]),
        digits: soloDigitos(item[campo])
      }))
      .filter(row => row.key);
  }

  async function cargarPlacas() {
    if (sheetCache.has("placas")) return sheetCache.get("placas");
    const workbook = await cargarWorkbook();
    const hoja = obtenerHoja(workbook, ["Placas", "Placa", "PLACAS", "PLACA"]);
    const data = XLSX.utils.sheet_to_json(hoja, { defval: "", raw: true })
      .map(filaPlaca)
      .filter(e => e.referencia);
    const payload = {
      data,
      exacto: crearIndiceExacto(data, "referencia"),
      indice: crearIndiceBusqueda(data, "referencia")
    };
    sheetCache.set("placas", payload);
    return payload;
  }

  async function cargarContratos() {
    if (sheetCache.has("contratos")) return sheetCache.get("contratos");
    const workbook = await cargarWorkbook();
    const hoja = obtenerHoja(workbook, ["Contratos", "Contrato", "CONTRATOS", "CONTRATO"]);
    const data = XLSX.utils.sheet_to_json(hoja, { defval: "", raw: true })
      .map(filaContrato)
      .filter(e => e.contrato);
    const payload = {
      data,
      exacto: crearIndiceExacto(data, "contrato"),
      indice: crearIndiceBusqueda(data, "contrato")
    };
    sheetCache.set("contratos", payload);
    return payload;
  }

  function distanciaKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat / 2) ** 2 +
      Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
      Math.sin(dLon / 2) ** 2;
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
  }

  function puntuacion(row, query, modo) {
    const q = normalizarClave(query);
    const qDigits = soloDigitos(query);
    if (!q && !qDigits) return null;

    if (modo === "numero") {
      if (qDigits.length < 2) return null;
      if (row.digits === qDigits) return 0;
      if (row.digits.startsWith(qDigits)) return 1;
      if (row.digits.includes(qDigits)) return 2;
      return null;
    }

    if (row.key === q) return 0;
    if (row.key.startsWith(q)) return 1;
    if (q.length >= 2 && row.key.includes(q)) return 2;
    if (qDigits.length >= 2 && row.digits.includes(qDigits)) return 3;
    return null;
  }

  function sugerencias(indice, valor, limite = 8, modo = "referencia") {
    const q = normalizarClave(valor);
    const qDigits = soloDigitos(valor);
    if ((modo === "numero" && qDigits.length < 2) || (modo !== "numero" && q.length < 2 && qDigits.length < 2)) return [];

    const hallazgos = [];
    const usados = new Set();
    for (const row of indice || []) {
      const score = puntuacion(row, valor, modo);
      if (score === null) continue;
      const unique = row.key;
      if (usados.has(unique)) continue;
      usados.add(unique);
      hallazgos.push({ score, label: row.label, item: row.item, keyLength: row.key.length });
      if (hallazgos.length > limite * 7) break;
    }
    return hallazgos
      .sort((a, b) => a.score - b.score || a.keyLength - b.keyLength || a.label.localeCompare(b.label))
      .slice(0, limite);
  }

  function buscarFlexible(payload, campo, valor, modo = "referencia") {
    const key = normalizarClave(valor);
    const exacto = payload.exacto && payload.exacto.get(key);
    if (exacto) return exacto;
    const top = sugerencias(payload.indice, valor, 1, modo)[0];
    return top ? top.item : null;
  }

  function precargar(fn) {
    if ("requestIdleCallback" in window) {
      window.requestIdleCallback(() => fn().catch(() => {}), { timeout: 1800 });
    } else {
      window.setTimeout(() => fn().catch(() => {}), 650);
    }
  }

  window.DLSRCore = {
    normalizarTexto,
    normalizarClave,
    soloDigitos,
    coordValida,
    htmlSeguro,
    distanciaKm,
    cargarPlacas,
    cargarContratos,
    sugerencias,
    buscarFlexible,
    parseCoord,
    precargar
  };
})();
