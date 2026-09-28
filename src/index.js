function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { "content-type": "application/json" },
  });
}

let cachedtoken = null;
let cachedexpiry = 0;

async function getgraphtoken(env) {
  if (cachedtoken && Date.now() < cachedexpiry) return cachedtoken;
  const tenant = env.EMAIL_TENANT || "common";
  const params = new URLSearchParams();
  params.set("client_id", env.EMAIL_CLIENT_ID);
  params.set("client_secret", env.EMAIL_CLIENT_SECRET);
  params.set("grant_type", "refresh_token");
  params.set("refresh_token", env.EMAIL_REFRESH_TOKEN);
  params.set("scope", "offline_access https://graph.microsoft.com/Mail.Send");
  const resp = await fetch("https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || !data.access_token) {
    throw new Error("token fetch failed (" + resp.status + "): " + JSON.stringify(data));
  }
  cachedtoken = data.access_token;
  cachedexpiry = Date.now() + ((data.expires_in || 3600) - 60) * 1000;
  return cachedtoken;
}

export default {
  async fetch(request, env) {
    if (request.method !== "POST") return json({ error: "method not allowed" }, 405);

    const auth = request.headers.get("authorization") || "";
    const token = auth.replace(/^Bearer\s+/i, "").trim();
    if (!env.WORKER_SECRET || token !== env.WORKER_SECRET) return json({ error: "unauthorized" }, 401);

    if (!env.EMAIL_CLIENT_ID || !env.EMAIL_CLIENT_SECRET || !env.EMAIL_REFRESH_TOKEN) {
      return json({ error: "worker missing outlook credentials" }, 500);
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return json({ error: "invalid json" }, 400);
    }

    const to = body.to;
    const subject = body.subject;
    if (!to || !subject) return json({ error: "missing to or subject" }, 400);

    let graphtoken;
    try {
      graphtoken = await getgraphtoken(env);
    } catch (e) {
      return json({ error: "auth failed", detail: String(e.message || e) }, 502);
    }

    const message = {
      subject: String(subject),
      body: {
        contentType: body.html ? "HTML" : "Text",
        content: String(body.html || body.text || ""),
      },
      toRecipients: [{ emailAddress: { address: String(to) } }],
    };

    let resp, detail;
    try {
      resp = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: { "authorization": "Bearer " + graphtoken, "content-type": "application/json" },
        body: JSON.stringify({ message, saveToSentItems: true }),
      });
      if (resp.status !== 202) detail = await resp.text();
    } catch (e) {
      return json({ error: "relay request failed", detail: String(e) }, 502);
    }

    if (resp.status !== 202) return json({ error: "send failed", status: resp.status, detail }, 502);
    return json({ ok: true });
  },
};
