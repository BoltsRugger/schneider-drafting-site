// Azure Functions v4 - Node 18+
// Sends mail via Microsoft Graph using client credentials.

import { ClientSecretCredential } from "@azure/identity";

// --- Helpers ---------------------------------------------------------------

// Parse urlencoded or JSON bodies
function parseBody(req) {
    const ctype = (req.headers["content-type"] || "").toLowerCase();
    if (ctype.includes("application/x-www-form-urlencoded")) {
        const params = new URLSearchParams(typeof req.body === "string" ? req.body : "");
        return Object.fromEntries(params.entries());
    }
    if (typeof req.body === "object" && req.body) return req.body;
    try { return JSON.parse(req.rawBody || "{}"); } catch { return {}; }
}

// Basic HTML escape
function esc(s = "") {
    return String(s)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}

// Simple maskers for logs
const mask = v => (v ? "✓" : "✗");
const maskEmail = v => (v ? v.replace(/^(.).+(@.+)$/, "$1***$2") : "(missing)");

// --- Function --------------------------------------------------------------

export default async function (context, req) {
    const rid = context.invocationId || "n/a"; // request id for correlation
    try {
        // 1) Parse + basic validation ------------------------------------------------
        const data = parseBody(req);
        context.log(`[contact] [${rid}] received`, {
            hasBody: !!data,
            ctype: req.headers["content-type"] || "(none)"
        });

        // Honeypot (matches hidden input name)
        if (data.website) {
            context.log(`[contact] [${rid}] honeypot tripped; returning 200 silently`);
            context.res = { status: 200, jsonBody: { ok: true, message: "Thanks!" } };
            return;
        }

        const name = (data.name || "").trim();
        const email = (data.email || "").trim();
        const phone = (data.phone || "").trim();
        const message = (data.message || "").trim();

        if (!name || !email || !message) {
            context.log.warn(`[contact] [${rid}] missing required fields`, {
                name: !!name, email: !!email, message: !!message
            });
            context.res = { status: 400, jsonBody: { ok: false, message: "Please include name, email, and message." } };
            return;
        }
        if (message.length > 5000) {
            context.log.warn(`[contact] [${rid}] message too long`, { len: message.length });
            context.res = { status: 400, jsonBody: { ok: false, message: "Message is too long." } };
            return;
        }

        // 2) Env vars (your current names). Log presence, not values ---------------
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const fromUser = process.env.MAILBOX_ADDRESS; // required: user/shared mailbox UPN
        const toAddress = process.env.MAILBOX_ADDRESS || "jon@schneiderdrafting.com";

        context.log(`[contact] [${rid}] env check`, {
            tenantId: mask(tenantId),
            clientId: mask(clientId),
            clientSecret: mask(clientSecret),
            fromUser: maskEmail(fromUser),
            toAddress: maskEmail(toAddress)
        });

        if (!tenantId || !clientId || !clientSecret || !fromUser) {
            context.log.error(`[contact] [${rid}] missing Graph env vars`);
            context.res = { status: 500, jsonBody: { ok: false, message: "Mail service not configured." } };
            return;
        }

        // 3) Token acquisition -------------------------------------------------------
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        let token;
        try {
            token = await credential.getToken("https://graph.microsoft.com/.default");
            context.log(`[contact] [${rid}] token acquired`);
        } catch (e) {
            context.log.error(`[contact] [${rid}] token error`, e?.message || e);
            throw e;
        }
        if (!token?.token) throw new Error("No Graph token acquired");

        // 4) Build the mail ----------------------------------------------------------
        const html = `
      <div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;">
        <h2 style="margin:0 0 8px;">New Contact — Schneider Drafting Services</h2>
        <p><strong>Name:</strong> ${esc(name)}</p>
        <p><strong>Email:</strong> ${esc(email)}</p>
        ${phone ? `<p><strong>Phone:</strong> ${esc(phone)}</p>` : ""}
        <hr style="border:none;border-top:1px solid #ddd;margin:12px 0;" />
        <p style="white-space:pre-wrap;">${esc(message)}</p>
      </div>
    `;

        const mail = {
            message: {
                subject: `SDS Contact: ${name} <${email}>`,
                body: { contentType: "HTML", content: html },
                toRecipients: [{ emailAddress: { address: toAddress } }],
                replyTo: [{ emailAddress: { address: email, name } }]
            },
            saveToSentItems: true
        };

        // 5) Send via Graph ----------------------------------------------------------
        const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromUser)}/sendMail`;
        context.log(`[contact] [${rid}] sending via Graph`, { url, fromUser: maskEmail(fromUser) });

        const resp = await fetch(url, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token.token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(mail)
        });

        if (!resp.ok) {
            const text = await resp.text().catch(() => "");
            context.log.error(`[contact] [${rid}] Graph sendMail failed`, resp.status, text);
            throw new Error(`Graph sendMail ${resp.status}`);
        }

        context.log(`[contact] [${rid}] sendMail OK`);
        context.res = { status: 200, jsonBody: { ok: true, message: "Thanks! Your message has been sent." } };

    } catch (err) {
        // Final catch-all
        context.log.error(`[contact] [${rid}] unhandled`, err?.message || err);
        context.res = {
            status: 500,
            jsonBody: { ok: false, message: "Send failed. Please email jon@schneiderdrafting.com." }
        };
    }
}
