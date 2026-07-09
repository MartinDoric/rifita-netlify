// Netlify Function: puente entre la web de Netlify y Apps Script.
// Configuración recomendada:
// 1) En Netlify > Site configuration > Environment variables:
//    APPS_SCRIPT_URL = https://script.google.com/macros/s/TU_IMPLEMENTACION/exec
// 2) En Apps Script, pegá apps-script/Code.gs y publicá como Web App:
//    Ejecutar como: Yo
//    Usuarios con acceso: Cualquier persona

const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxy61w8yeIuOCBODrT1mg-V8cXsAkiY-dNvZDBIaqRWp1Tb4L8zlO9b3-5dxzMc8iHXJA/exec";

exports.handler = async function(event) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ ok: false, error: "Método no permitido" })
    };
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const method = String(body.method || "").trim();
    const args = Array.isArray(body.args) ? body.args : [];
    const codigoRifa = String(body.codigoRifa || "").trim();

    if (!method) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ ok: false, error: "Falta method" })
      };
    }

    const appsScriptUrl = process.env.APPS_SCRIPT_URL || DEFAULT_APPS_SCRIPT_URL;

    const response = await fetch(appsScriptUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        action: "__call",
        method,
        args,
        codigoRifa
      })
    });

    const text = await response.text();

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (_) {
      return {
        statusCode: 502,
        headers,
        body: JSON.stringify({
          ok: false,
          error: "Apps Script no devolvió JSON. Revisá la URL /exec y los permisos de implementación.",
          status: response.status,
          preview: text.slice(0, 350)
        })
      };
    }

    return {
      statusCode: response.ok ? 200 : response.status,
      headers,
      body: JSON.stringify(parsed)
    };

  } catch (error) {
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ ok: false, error: error.message || String(error) })
    };
  }
};
