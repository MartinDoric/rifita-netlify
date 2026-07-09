// Netlify Function: puente entre la web de Netlify y Apps Script.
// Esta versión incluye diagnóstico en GET para encontrar rápido problemas de URL/env.

const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxy61w8yeIuOCBODrT1mg-V8cXsAkiY-dNvZDBIaqRWp1Tb4L8zlO9b3-5dxzMc8iHXJA/exec";

function json(statusCode, obj) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Allow-Methods": "GET,POST,OPTIONS"
    },
    body: JSON.stringify(obj, null, 2)
  };
}

function cleanAppsScriptUrl(value) {
  let url = String(value || "").trim();

  // Por si en Netlify lo pegaron con comillas.
  url = url.replace(/^['"]+|['"]+$/g, "").trim();

  // Debe quedar solo hasta /exec. Si pegaron parámetros, los sacamos.
  const q = url.indexOf("?");
  if (q !== -1) url = url.slice(0, q);

  return url;
}

function maskUrl(url) {
  const s = String(url || "");
  if (!s) return "";
  if (s.length <= 45) return s;
  return s.slice(0, 38) + "..." + s.slice(-12);
}

function buildAppsScriptCallUrl({ method, args, codigoRifa }) {
  const raw = process.env.APPS_SCRIPT_URL || DEFAULT_APPS_SCRIPT_URL;
  const appsScriptUrl = cleanAppsScriptUrl(raw);

  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(appsScriptUrl)) {
    const err = new Error("APPS_SCRIPT_URL inválida. Tiene que ser el link /exec de Apps Script.");
    err.code = "BAD_APPS_SCRIPT_URL";
    err.appsScriptUrlMasked = maskUrl(appsScriptUrl);
    throw err;
  }

  const url = new URL(appsScriptUrl);
  url.searchParams.set("action", "__call");
  url.searchParams.set("method", String(method || ""));
  url.searchParams.set("codigoRifa", String(codigoRifa || ""));
  url.searchParams.set("args", JSON.stringify(Array.isArray(args) ? args : []));

  return { appsScriptUrl, url };
}

async function callAppsScript({ method, args, codigoRifa }) {
  const { appsScriptUrl, url } = buildAppsScriptCallUrl({ method, args, codigoRifa });

  const response = await fetch(url.toString(), {
    method: "GET",
    redirect: "follow",
    headers: {
      "Accept": "application/json,text/plain,*/*",
      "User-Agent": "Rifita-Netlify-Bridge"
    }
  });

  const text = await response.text();

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (_) {
    const err = new Error("Apps Script no devolvió JSON.");
    err.code = "APPS_SCRIPT_NOT_JSON";
    err.status = response.status;
    err.appsScriptUrlMasked = maskUrl(appsScriptUrl);
    err.calledUrlMasked = maskUrl(url.toString());
    err.preview = text.slice(0, 500);
    throw err;
  }

  return {
    status: response.status,
    appsScriptUrlMasked: maskUrl(appsScriptUrl),
    data: parsed
  };
}

exports.handler = async function(event) {
  if (event.httpMethod === "OPTIONS") {
    return json(204, {});
  }

  try {
    // Diagnóstico: abrir /.netlify/functions/api en el navegador.
    // También permite /.netlify/functions/api?codigoRifa=lechon-julio
    if (event.httpMethod === "GET") {
      const qs = event.queryStringParameters || {};
      const codigoRifa = qs.codigoRifa || "lechon-julio";

      try {
        const result = await callAppsScript({
          method: "apiCheckRifa",
          args: [codigoRifa],
          codigoRifa
        });

        return json(200, {
          ok: true,
          netlifyFunction: "api.js activa",
          usingEnvVar: Boolean(process.env.APPS_SCRIPT_URL),
          appsScriptUrl: result.appsScriptUrlMasked,
          test: result.data
        });
      } catch (e) {
        return json(502, {
          ok: false,
          netlifyFunction: "api.js activa, pero falló Apps Script",
          usingEnvVar: Boolean(process.env.APPS_SCRIPT_URL),
          error: e.message || String(e),
          code: e.code || "ERROR",
          status: e.status || null,
          appsScriptUrl: e.appsScriptUrlMasked || maskUrl(cleanAppsScriptUrl(process.env.APPS_SCRIPT_URL || DEFAULT_APPS_SCRIPT_URL)),
          calledUrl: e.calledUrlMasked || "",
          preview: e.preview || ""
        });
      }
    }

    if (event.httpMethod !== "POST") {
      return json(405, { ok: false, error: "Método no permitido" });
    }

    const body = JSON.parse(event.body || "{}");
    const method = String(body.method || "").trim();
    const args = Array.isArray(body.args) ? body.args : [];
    const codigoRifa = String(body.codigoRifa || "").trim();

    if (!method) {
      return json(400, { ok: false, error: "Falta method" });
    }

    const result = await callAppsScript({ method, args, codigoRifa });

    // Devolvemos exactamente lo que devolvió Apps Script.
    return json(200, result.data);

  } catch (error) {
    return json(500, {
      ok: false,
      error: error.message || String(error),
      code: error.code || "ERROR",
      status: error.status || null,
      appsScriptUrl: error.appsScriptUrlMasked || "",
      calledUrl: error.calledUrlMasked || "",
      preview: error.preview || ""
    });
  }
};
