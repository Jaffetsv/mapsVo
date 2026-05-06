/* global XLSX, DLSRMAPS_DATA_FILES, DLSRMAPS_DATA_VERSION */
"use strict";

importScripts("https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js", "./data-config.js?v=20260506-6");

const DATA_FILES = Array.isArray(self.DLSRMAPS_DATA_FILES) ? self.DLSRMAPS_DATA_FILES.slice() : [];
const DATA_VERSION = self.DLSRMAPS_DATA_VERSION || "20260506-6";

const ALIAS = {
  referencia: ["REFERENCIA", "Referencia", "referencia", "PLACA", "Placa", "placa", "CODIGO", "Código", "codigo", "Ref", "REF"],
  contrato: ["Contrato", "CONTRATO", "contrato", "NIS", "nis", "Cuenta", "CUENTA", "No Contrato", "N° Contrato", "Numero Contrato", "Número Contrato"],
  lat: ["Latitud", "LATITUD", "latitud", "Lat", "LAT", "Latitude", "Y"],
  lon: ["Longitud", "LONGITUD", "longitud", "Lon", "LON", "Lng", "LNG", "Long", "LONG", "Longitude", "X"],
  alimentador: ["Alimentador", "ALIMENTADOR", "alimentador", "Feeder", "Circuito", "CIRCUITO"]
};

const cache = {
  workbookPromise: null,
  placas: null,
  contratos: null
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
  if (cache.workbookPromise) return cache.workbookPromise;
  cache.workbookPromise = (async () => {
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
  return cache.workbookPromise;
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
  return Number.isFinite(Number(lat)) && Number.isFinite(Number(lon)) &&
    Number(lat) >= -90 && Number(lat) <= 90 && Number(lon) >= -180 && Number(lon) <= 180;
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
  if (cache.placas) return cache.placas;
  const workbook = await cargarWorkbook();
  const hoja = obtenerHoja(workbook, ["Placas", "Placa", "PLACAS", "PLACA"]);
  const data = XLSX.utils.sheet_to_json(hoja, { defval: "", raw: true })
    .map(filaPlaca)
    .filter(e => e.referencia);
  cache.placas = {
    data,
    exacto: crearIndiceExacto(data, "referencia"),
    indice: crearIndiceBusqueda(data, "referencia")
  };
  return cache.placas;
}

async function cargarContratos() {
  if (cache.contratos) return cache.contratos;
  const workbook = await cargarWorkbook();
  const hoja = obtenerHoja(workbook, ["Contratos", "Contrato", "CONTRATOS", "CONTRATO"]);
  const data = XLSX.utils.sheet_to_json(hoja, { defval: "", raw: true })
    .map(filaContrato)
    .filter(e => e.contrato);
  cache.contratos = {
    data,
    exacto: crearIndiceExacto(data, "contrato"),
    indice: crearIndiceBusqueda(data, "contrato")
  };
  return cache.contratos;
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
  const maxPrevio = Math.max(limite * 10, 48);

  for (const row of indice || []) {
    const score = puntuacion(row, valor, modo);
    if (score === null) continue;
    const unique = row.key;
    if (usados.has(unique)) continue;
    usados.add(unique);
    hallazgos.push({ score, label: row.label, item: row.item, keyLength: row.key.length });
    if (hallazgos.length >= maxPrevio) break;
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

function distanciaKm(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function leerCoordenadas(texto) {
  const match = String(texto || "").match(/^\s*(-?\d+(?:[\.,]\d+)?)\s*[,;\s]+\s*(-?\d+(?:[\.,]\d+)?)\s*$/);
  if (!match) return null;
  return { lat: parseCoord(match[1]), lon: parseCoord(match[2]) };
}

function topCercanos(lat, lon, data, limit = 10) {
  const mejores = [];
  for (const item of data) {
    if (!coordValida(item.lat, item.lon)) continue;
    const distancia = distanciaKm(Number(lat), Number(lon), Number(item.lat), Number(item.lon));
    const candidato = { ...item, distancia };
    if (mejores.length < limit) {
      mejores.push(candidato);
      mejores.sort((a, b) => a.distancia - b.distancia);
    } else if (distancia < mejores[mejores.length - 1].distancia) {
      mejores[mejores.length - 1] = candidato;
      mejores.sort((a, b) => a.distancia - b.distancia);
    }
  }
  return mejores;
}

async function resolverCercanosDesdeConsulta(value, limit = 10) {
  const payload = await cargarPlacas();
  const data = payload.data;
  const valor = String(value || "").trim();
  const coords = leerCoordenadas(valor);

  if (coords) {
    if (!coordValida(coords.lat, coords.lon)) throw new Error("Las coordenadas ingresadas no son válidas.");
    return {
      base: { tipo: "coords", titulo: "Coordenadas consultadas", lat: coords.lat, lon: coords.lon, extra: "" },
      cercanos: topCercanos(coords.lat, coords.lon, data, limit)
    };
  }

  const porReferencia = buscarFlexible(payload, "referencia", valor, "referencia");
  if (porReferencia && coordValida(porReferencia.lat, porReferencia.lon)) {
    return {
      base: { tipo: "referencia", titulo: "Referencia consultada", lat: porReferencia.lat, lon: porReferencia.lon, extra: porReferencia.referencia },
      cercanos: topCercanos(porReferencia.lat, porReferencia.lon, data, limit)
    };
  }

  const q = normalizarTexto(valor);
  const coincidenciaAlimentador = data.find(item => normalizarTexto(item.alimentador).includes(q) && coordValida(item.lat, item.lon));
  if (coincidenciaAlimentador) {
    return {
      base: { tipo: "alimentador", titulo: "Alimentador consultado", lat: coincidenciaAlimentador.lat, lon: coincidenciaAlimentador.lon, extra: coincidenciaAlimentador.alimentador },
      cercanos: topCercanos(coincidenciaAlimentador.lat, coincidenciaAlimentador.lon, data, limit)
    };
  }

  return { base: null, cercanos: [] };
}

self.onmessage = async event => {
  const { id, type, payload = {} } = event.data || {};
  try {
    let result;
    if (type === "warmPlacas") result = { ready: true, count: (await cargarPlacas()).data.length };
    else if (type === "warmContratos") result = { ready: true, count: (await cargarContratos()).data.length };
    else if (type === "suggestPlacas") result = sugerencias((await cargarPlacas()).indice, payload.value, payload.limit || 8, payload.mode || "referencia");
    else if (type === "suggestContratos") result = sugerencias((await cargarContratos()).indice, payload.value, payload.limit || 7, "referencia");
    else if (type === "findPlaca") result = buscarFlexible(await cargarPlacas(), "referencia", payload.value, payload.mode || "referencia");
    else if (type === "findContrato") result = buscarFlexible(await cargarContratos(), "contrato", payload.value, "referencia");
    else if (type === "nearestFromPoint") {
      const lat = Number(payload.lat);
      const lon = Number(payload.lon);
      if (!coordValida(lat, lon)) throw new Error("Las coordenadas no son válidas.");
      result = { base: { tipo: "coords", titulo: payload.title || "Punto consultado", lat, lon, extra: payload.extra || "" }, cercanos: topCercanos(lat, lon, (await cargarPlacas()).data, payload.limit || 10) };
    } else if (type === "nearestFromQuery") result = await resolverCercanosDesdeConsulta(payload.value, payload.limit || 10);
    else throw new Error("Solicitud no reconocida.");

    self.postMessage({ id, ok: true, result });
  } catch (err) {
    self.postMessage({ id, ok: false, error: err && err.message ? err.message : "No fue posible completar la consulta." });
  }
};
