
import { app } from "@azure/functions";
import * as querystring from "node:querystring";
import msal from "@azure/msal-node";

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAILBOX_ADDRESS } = process.env;

const cca = new msal.ConfidentialClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
});

function sanitize(s){ return String(s||"").replace(/[\r\n]+/g,"\n").trim(); }

function buildMessage({name,email,phone,message}){
  const subject = `New inquiry from ${name}`;
  const body = [`Name: ${name}`, `Email: ${email}`, `Phone: ${phone}`, "", "Message:", message].join("\n");
  return {
    message: {
      subject,
      body: { contentType: "Text", content: body },
      toRecipients: [{ emailAddress: { address: MAILBOX_ADDRESS }}],
      replyTo: email ? [{ emailAddress: { address: email }}] : undefined
    },
    saveToSentItems: "true"
  };
}

async function sendMailViaGraph(msg){
  const result = await cca.acquireTokenByClientCredential({ scopes: ["https://graph.microsoft.com/.default"] });
  const token = result?.accessToken;
  if (!token) throw new Error("No Graph token");
  const endpoint = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MAILBOX_ADDRESS)}/sendMail`;
  const res = await fetch(endpoint, { method:"POST", headers:{ "Authorization":`Bearer ${token}`, "Content-Type":"application/json" }, body: JSON.stringify(msg) });
  if (!res.ok) throw new Error(`Graph sendMail failed (${res.status})`);
}

app.http("contact", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req, ctx) => {
    try {
      const raw = await req.text();
      const data = querystring.parse(raw);
      if (data.website) return { status:200, jsonBody:{ ok:true, message:"Thanks!" } }; // honeypot

      const name = sanitize(data.name);
      const email = sanitize(data.email);
      const phone = sanitize(data.phone);
      const message = sanitize(data.message);
      if (!name || !email) return { status:400, jsonBody:{ ok:false, message:"Please include your name and a valid email." } };

      await sendMailViaGraph(buildMessage({name,email,phone,message}));
      return { status:200, jsonBody:{ ok:true, message:"Thanks! Your message has been sent." } };
    } catch (err) {
      ctx.log("Contact error:", err?.message || err);
      return { status:500, jsonBody:{ ok:false, message:"We couldn’t send your message. Please email info@schneiderdrafting.com." } };
    }
  }
});
