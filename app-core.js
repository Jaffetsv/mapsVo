(function () {
  "use strict";

  const APP_VERSION = "20260506-6";
  const WORKER_URL = `./data-worker.js?v=${APP_VERSION}`;

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

  function parseCoord(valor) {
    if (valor === null || valor === undefined || valor === "") return NaN;
    const limpio = String(valor).replace(",", ".").replace(/[^0-9.\-]/g, "");
    return parseFloat(limpio);
  }

  function coordValida(lat, lon) {
    return Number.isFinite(Number(lat)) && Number.isFinite(Number(lon)) &&
      Number(lat) >= -90 && Number(lat) <= 90 && Number(lon) >= -180 && Number(lon) <= 180;
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

  function distanciaKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat / 2) ** 2 +
      Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
      Math.sin(dLon / 2) ** 2;
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
  }

  function precargar(fn) {
    const run = () => Promise.resolve().then(fn).catch(() => {});
    if ("requestIdleCallback" in window) {
      window.requestIdleCallback(run, { timeout: 1800 });
    } else {
      window.setTimeout(run, 800);
    }
  }

  let worker = null;
  let callId = 0;
  const pending = new Map();

  function rechazarPendientes(message) {
    for (const [, req] of pending) req.reject(new Error(message));
    pending.clear();
  }

  function obtenerWorker() {
    if (worker) return worker;
    if (!("Worker" in window)) {
      throw new Error("Este navegador no soporta el motor rápido de búsqueda.");
    }

    worker = new Worker(WORKER_URL);
    worker.onmessage = event => {
      const msg = event.data || {};
      const req = pending.get(msg.id);
      if (!req) return;
      pending.delete(msg.id);
      if (msg.ok) req.resolve(msg.result);
      else req.reject(new Error(msg.error || "No fue posible completar la consulta."));
    };
    worker.onerror = () => {
      rechazarPendientes("No fue posible iniciar el motor de búsqueda.");
      try { worker.terminate(); } catch (_) {}
      worker = null;
    };
    return worker;
  }

  function llamarWorker(type, payload = {}) {
    return new Promise((resolve, reject) => {
      let w;
      try {
        w = obtenerWorker();
      } catch (err) {
        reject(err);
        return;
      }
      const id = ++callId;
      pending.set(id, { resolve, reject });
      w.postMessage({ id, type, payload });
    });
  }

  window.DLSRCore = {
    normalizarTexto,
    normalizarClave,
    soloDigitos,
    parseCoord,
    coordValida,
    htmlSeguro,
    distanciaKm,
    precargar
  };

  window.DLSRData = {
    calentarPlacas() { return llamarWorker("warmPlacas"); },
    calentarContratos() { return llamarWorker("warmContratos"); },
    sugerirPlacas(value, mode, limit) { return llamarWorker("suggestPlacas", { value, mode, limit }); },
    sugerirContratos(value, limit) { return llamarWorker("suggestContratos", { value, limit }); },
    buscarPlaca(value, mode) { return llamarWorker("findPlaca", { value, mode }); },
    buscarContrato(value) { return llamarWorker("findContrato", { value }); },
    cercanosDesdePunto(lat, lon, limit) { return llamarWorker("nearestFromPoint", { lat, lon, limit }); },
    cercanosDesdeConsulta(value, limit) { return llamarWorker("nearestFromQuery", { value, limit }); }
  };
})();
