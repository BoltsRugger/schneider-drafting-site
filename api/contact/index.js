// Azure Functions v4 - Node 18+
// Sends mail via Microsoft Graph using client credentials.

import { ClientSecretCredential } from "@azure/identity";

// In Node 18+, fetch is global. If you’re on older Node, install node-fetch and import it.

// Utility: parse urlencoded or JSON bodies
function parseBody(req) {
    const ctype = (req.headers["content-type"] || "").toLowerCase();
    if (ctype.includes("application/x-www-form-urlencoded")) {
        const params = new URLSearchParams(typeof req.body === "string" ? req.body : "");
        return Object.fromEntries(params.entries());
    }
    if (typeof req.body === "object" && req.body) return req.body;
    try {
        return JSON.parse(req.rawBody || "{}");
    } catch {
        return {};
    }
}

// Basic HTML escape
function esc(s = "") {
    return String(s)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}

export default async function (context, req) {
    try {
        const data = parseBody(req);

        // Honeypot (matches your hidden input name)
        if (data.website) {
            return context.res = { status: 200, jsonBody: { ok: true, message: "Thanks!" } };
        }

        const name = (data.name || "").trim();
        const email = (data.email || "").trim();
        const phone = (data.phone || "").trim();
        const message = (data.message || "").trim();

        if (!name || !email || !message) {
            return context.res = {
                status: 400,
                jsonBody: { ok: false, message: "Please include name, email, and message." }
            };
        }

        // Optional: simple length guard
        if (message.length > 5000) {
            return context.res = {
                status: 400,
                jsonBody: { ok: false, message: "Message is too long." }
            };
        }

        // Env vars (configure in Azure → your SWA → Environment variables)
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const fromUser = process.env.MAILBOX_ADDRESS; // e.g. "jon@schneiderdrafting.com" or a userId (GUID)
        const toAddress = process.env.MAILBOX_ADDRESS || "jon@schneiderdrafting.com";

        if (!tenantId || !clientId || !clientSecret || !fromUser) {
            context.log.error("Missing Graph env vars.");
            return context.res = {
                status: 500,
                jsonBody: { ok: false, message: "Mail service not configured." }
            };
        }

        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const token = await credential.getToken("https://graph.microsoft.com/.default");
        if (!token?.token) throw new Error("Failed to acquire Graph token");

        // Build email (HTML + a text-ish fallback inside)
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

        // Send as the configured user/shared mailbox
        const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromUser)}/sendMail`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token.token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(mail)
        });

        if (!resp.ok) {
            const text = await resp.text();
            context.log.error("Graph sendMail failed:", resp.status, text);
            throw new Error("Email send failed.");
        }

        context.res = {
            status: 200,
            jsonBody: { ok: true, message: "Thanks! Your message has been sent." }
        };
    } catch (err) {
        context.log.error(err);
        context.res = {
            status: 500,
            jsonBody: { ok: false, message: "Send failed. Please email jon@schneiderdrafting.com." }
        };
    }
}
