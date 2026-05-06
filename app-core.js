(function () {
  "use strict";

  const DATA_FILES = Array.isArray(window.DLSRMAPS_DATA_FILES) && window.DLSRMAPS_DATA_FILES.length
    ? window.DLSRMAPS_DATA_FILES.slice()
    : [];
  const DATA_VERSION = window.DLSRMAPS_DATA_VERSION || "1";

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
          return XLSX.read(arrayBuffer, { type: "array" });
        } catch (_) {
          // Intenta el siguiente archivo declarado.
        }
      }
      throw new Error("No fue posible cargar los datos. Verifique que el archivo Excel esté subido en la raíz del repositorio y que su nombre esté declarado en data-config.js.");
    })();

    return workbookPromise;
  }

  function obtenerHoja(workbook, nombresPosibles) {
    const buscadas = nombresPosibles.map(normalizarClave);
    const nombreHoja = workbook.SheetNames.find(n => buscadas.includes(normalizarClave(n)));
    if (!nombreHoja) throw new Error("No se encontró la hoja requerida dentro del Excel.");
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
    return {
      referencia: obtenerCampo(fila, ALIAS.referencia),
      lat,
      lon,
      alimentador: obtenerCampo(fila, ALIAS.alimentador) || "Sin alimentador"
    };
  }

  function filaContrato(fila) {
    const lat = parseCoord(obtenerCampo(fila, ALIAS.lat));
    const lon = parseCoord(obtenerCampo(fila, ALIAS.lon));
    return {
      contrato: obtenerCampo(fila, ALIAS.contrato),
      referencia: obtenerCampo(fila, ALIAS.referencia),
      lat,
      lon,
      alimentador: obtenerCampo(fila, ALIAS.alimentador) || "Sin alimentador"
    };
  }

  function crearIndice(items, campo) {
    const exacto = new Map();
    for (const item of items) {
      const key = normalizarClave(item[campo]);
      if (key && !exacto.has(key)) exacto.set(key, item);
    }
    return exacto;
  }

  async function cargarPlacas() {
    if (sheetCache.has("placas")) return sheetCache.get("placas");
    const workbook = await cargarWorkbook();
    const hoja = obtenerHoja(workbook, ["Placas", "Placa", "PLACAS", "PLACA"]);
    const data = XLSX.utils.sheet_to_json(hoja, { defval: "" })
      .map(filaPlaca)
      .filter(e => e.referencia);
    const payload = { data, exacto: crearIndice(data, "referencia") };
    sheetCache.set("placas", payload);
    return payload;
  }

  async function cargarContratos() {
    if (sheetCache.has("contratos")) return sheetCache.get("contratos");
    const workbook = await cargarWorkbook();
    const hoja = obtenerHoja(workbook, ["Contratos", "Contrato", "CONTRATOS", "CONTRATO"]);
    const data = XLSX.utils.sheet_to_json(hoja, { defval: "" })
      .map(filaContrato)
      .filter(e => e.contrato);
    const payload = { data, exacto: crearIndice(data, "contrato") };
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

  function buscarParcial(data, campo, valor, minimo) {
    const key = normalizarClave(valor);
    if (!key || key.length < minimo) return null;
    return data.find(item => normalizarClave(item[campo]).includes(key)) || null;
  }

  window.DLSRCore = {
    normalizarTexto,
    normalizarClave,
    coordValida,
    htmlSeguro,
    distanciaKm,
    cargarPlacas,
    cargarContratos,
    buscarParcial,
    parseCoord
  };
})();
