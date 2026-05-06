(function () {
  "use strict";

  const APP_VERSION = "20260506-7";
  const WORKER_URL = `./data-worker.js?v=${APP_VERSION}`;
  const FAV_KEY = "dlsr_maps_favoritos_v7";
  const REC_KEY = "dlsr_maps_recientes_v7";
  const LAST_POS_KEY = "dlsr_maps_last_position_v7";

  function normalizarTexto(valor) {
    return String(valor ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  }
  function normalizarClave(valor) { return normalizarTexto(valor).replace(/[^a-z0-9]/g, ""); }
  function soloDigitos(valor) { return String(valor ?? "").replace(/\D/g, ""); }
  function parseCoord(valor) {
    if (valor === null || valor === undefined || valor === "") return NaN;
    const limpio = String(valor).replace(",", ".").replace(/[^0-9.\-]/g, "");
    return parseFloat(limpio);
  }
  function coordValida(lat, lon) {
    return Number.isFinite(Number(lat)) && Number.isFinite(Number(lon)) && Number(lat) >= -90 && Number(lat) <= 90 && Number(lon) >= -180 && Number(lon) <= 180;
  }
  function htmlSeguro(valor) {
    return String(valor ?? "").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
  }
  function distanciaKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLon / 2) ** 2;
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
  }
  function leerCoordenadas(texto) {
    const s = String(texto || "").trim();
    let m = s.match(/^(-?\d+(?:[\.,]\d+)?)\s*[,;\s]+\s*(-?\d+(?:[\.,]\d+)?)$/);
    if (m) return { lat: parseCoord(m[1]), lon: parseCoord(m[2]) };
    // Formato simple DMS: 13°51'29.4"N 89°17'08.9"W
    m = s.match(/(\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)?\D*(\d+(?:\.\d+)?)?\D*([NS])\s+(\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)?\D*(\d+(?:\.\d+)?)?\D*([EW])/i);
    if (m) {
      const dms = (d, min, sec, hemi) => {
        let v = Number(d) + (Number(min || 0) / 60) + (Number(sec || 0) / 3600);
        if (/S|W/i.test(hemi)) v *= -1;
        return v;
      };
      return { lat: dms(m[1], m[2], m[3], m[4]), lon: dms(m[5], m[6], m[7], m[8]) };
    }
    return null;
  }
  function precargar(fn) {
    const run = () => Promise.resolve().then(fn).catch(() => {});
    if ("requestIdleCallback" in window) window.requestIdleCallback(run, { timeout: 2500 });
    else window.setTimeout(run, 1200);
  }
  function debounce(fn, delay) {
    let t = null;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), delay); };
  }
  function leerJSON(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key) || "") || fallback; } catch (_) { return fallback; }
  }
  function guardarJSON(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
  }
  function itemId(item) {
    const tipo = item.tipo || (item.contrato ? "contrato" : "placa");
    const id = item.referencia || item.contrato || `${item.lat},${item.lon}`;
    return `${tipo}:${normalizarClave(id)}`;
  }
  function itemTitulo(item) { return item.referencia || item.contrato || item.titulo || "Registro"; }
  function itemSubtitulo(item) { return item.alimentador || item.extra || (item.tipo ? item.tipo : ""); }
  function getFavoritos() { return leerJSON(FAV_KEY, []); }
  function setFavoritos(items) { guardarJSON(FAV_KEY, items.slice(0, 200)); }
  function esFavorito(item) { const id = itemId(item); return getFavoritos().some(x => x.id === id); }
  function toggleFavorito(item) {
    const id = itemId(item);
    let favs = getFavoritos();
    const exists = favs.some(x => x.id === id);
    if (exists) favs = favs.filter(x => x.id !== id);
    else favs.unshift({ id, t: Date.now(), item });
    setFavoritos(favs);
    return !exists;
  }
  function getRecientes() { return leerJSON(REC_KEY, []); }
  function addReciente(item) {
    const id = itemId(item);
    let rec = getRecientes().filter(x => x.id !== id);
    rec.unshift({ id, t: Date.now(), item });
    guardarJSON(REC_KEY, rec.slice(0, 30));
  }
  function clearRecientes() { guardarJSON(REC_KEY, []); }
  function clearFavoritos() { guardarJSON(FAV_KEY, []); }
  function mapsUrl(item) { return `https://www.google.com/maps?q=${item.lat},${item.lon}`; }
  function wazeUrl(item) { return `https://waze.com/ul?ll=${item.lat},${item.lon}&navigate=yes`; }
  function textoCompartir(item) {
    const partes = [];
    if (item.referencia) partes.push(`Referencia: ${item.referencia}`);
    if (item.contrato) partes.push(`Contrato: ${item.contrato}`);
    if (item.alimentador) partes.push(`Alimentador: ${item.alimentador}`);
    if (coordValida(item.lat, item.lon)) partes.push(`Ubicación: ${mapsUrl(item)}`);
    return partes.join("\n");
  }
  async function compartir(item) {
    const text = textoCompartir(item);
    if (navigator.share) { await navigator.share({ title: itemTitulo(item), text }); return; }
    await navigator.clipboard.writeText(text);
  }
  async function copiar(text) { await navigator.clipboard.writeText(String(text ?? "")); }
  function guardarUbicacion(lat, lon, accuracy) { guardarJSON(LAST_POS_KEY, { lat, lon, accuracy, t: Date.now() }); }
  function leerUbicacionGuardada(maxAgeMs = 10 * 60 * 1000) {
    const pos = leerJSON(LAST_POS_KEY, null);
    if (!pos || !coordValida(pos.lat, pos.lon)) return null;
    if (Date.now() - Number(pos.t || 0) > maxAgeMs) return null;
    return pos;
  }
  function getPosicion(options) {
    return new Promise((resolve, reject) => navigator.geolocation.getCurrentPosition(resolve, reject, options));
  }
  async function obtenerUbicacion() {
    const fast = { enableHighAccuracy: false, timeout: 6500, maximumAge: 300000 };
    const accurate = { enableHighAccuracy: true, timeout: 15000, maximumAge: 60000 };
    try { return await getPosicion(fast); } catch (_) { return await getPosicion(accurate); }
  }

  let worker = null, callId = 0;
  const pending = new Map();
  function rechazarPendientes(message) { for (const [, req] of pending) req.reject(new Error(message)); pending.clear(); }
  function obtenerWorker() {
    if (worker) return worker;
    if (!("Worker" in window)) throw new Error("Este navegador no soporta búsqueda en segundo plano.");
    worker = new Worker(WORKER_URL);
    worker.onmessage = event => {
      const msg = event.data || {}; const req = pending.get(msg.id); if (!req) return;
      pending.delete(msg.id); msg.ok ? req.resolve(msg.result) : req.reject(new Error(msg.error || "No fue posible completar la consulta."));
    };
    worker.onerror = () => { rechazarPendientes("No fue posible iniciar el motor de búsqueda."); try { worker.terminate(); } catch (_) {} worker = null; };
    return worker;
  }
  function consulta(type, payload = {}) {
    const id = ++callId;
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      try { obtenerWorker().postMessage({ id, type, payload }); }
      catch (err) { pending.delete(id); reject(err); }
    });
  }
  const DLSRData = {
    calentarPlacas: () => consulta("warmPlacas"),
    calentarContratos: () => consulta("warmContratos"),
    calentarTodo: () => consulta("warmAll"),
    sugerirPlacas: (value, mode = "referencia", limit = 8) => consulta("suggestPlacas", { value, mode, limit }),
    sugerirContratos: (value, limit = 8) => consulta("suggestContratos", { value, limit }),
    sugerirTodo: (value, limit = 10) => consulta("suggestAll", { value, limit }),
    buscarPlaca: (value, mode = "referencia") => consulta("findPlaca", { value, mode }),
    buscarContrato: value => consulta("findContrato", { value }),
    buscarTodo: value => consulta("findAll", { value }),
    cercanosDesdePunto: (lat, lon, limit = 10, alimentador = "") => consulta("nearestFromPoint", { lat, lon, limit, alimentador }),
    cercanosDesdeConsulta: (value, limit = 10, alimentador = "") => consulta("nearestFromQuery", { value, limit, alimentador }),
    alimentadores: () => consulta("feeders"),
    estado: () => consulta("status")
  };

  window.DLSRCore = { APP_VERSION, normalizarTexto, normalizarClave, soloDigitos, parseCoord, coordValida, leerCoordenadas, distanciaKm, htmlSeguro, precargar, debounce, mapsUrl, wazeUrl, textoCompartir, compartir, copiar, getFavoritos, toggleFavorito, esFavorito, clearFavoritos, getRecientes, addReciente, clearRecientes, guardarUbicacion, leerUbicacionGuardada, obtenerUbicacion, itemId, itemTitulo, itemSubtitulo };
  window.DLSRData = DLSRData;
})();
