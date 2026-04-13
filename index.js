/**
 * Serveur MCP Outlook pour Claude Code
 * Connecte Claude Code a Microsoft 365 (emails + calendrier)
 *
 * @author Serum & Co
 * @version 1.0.0
 */

require("dotenv").config({ path: require("path").join(__dirname, ".env") });

const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { z } = require("zod");
const msal = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");
const express = require("express");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

// ============================================================
// Configuration depuis .env
// ============================================================
const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI || "http://localhost:3333/callback";
const SCOPES = (process.env.SCOPES || "Mail.Read,Mail.Send,User.Read,Calendars.ReadWrite").split(",");

// Verification des variables obligatoires
if (!CLIENT_ID || !TENANT_ID || !CLIENT_SECRET) {
  console.error("====================================================");
  console.error("ERREUR : Variables manquantes dans le fichier .env");
  console.error("====================================================");
  console.error("Verifiez que votre fichier .env contient :");
  console.error("  CLIENT_ID=...");
  console.error("  TENANT_ID=...");
  console.error("  CLIENT_SECRET=...");
  console.error("");
  console.error("Si vous n'avez pas de fichier .env :");
  console.error("  1. Copiez .env.example en .env");
  console.error("  2. Remplissez les valeurs depuis le portail Azure");
  console.error("====================================================");
  process.exit(1);
}

// ============================================================
// MSAL (authentification Microsoft)
// ============================================================
const CACHE_PATH = path.join(__dirname, ".msal-cache.json");

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  },
};

const beforeCacheAccess = async (cacheContext) => {
  if (fs.existsSync(CACHE_PATH)) {
    cacheContext.tokenCache.deserialize(fs.readFileSync(CACHE_PATH, "utf8"));
  }
};

const afterCacheAccess = async (cacheContext) => {
  if (cacheContext.cacheHasChanged) {
    fs.writeFileSync(CACHE_PATH, cacheContext.tokenCache.serialize());
  }
};

msalConfig.cache = { cachePlugin: { beforeCacheAccess, afterCacheAccess } };
const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
  if (!fs.existsSync(CACHE_PATH)) {
    throw new Error("Non authentifie. Utilisez l'outil 'authenticate' d'abord.");
  }

  try {
    const cache = cca.getTokenCache();
    const accounts = await cache.getAllAccounts();
    if (!accounts.length) {
      throw new Error("Pas de compte en cache.");
    }

    const result = await cca.acquireTokenSilent({
      account: accounts[0],
      scopes: SCOPES.map((s) => `https://graph.microsoft.com/${s}`),
    });
    return result.accessToken;
  } catch (err) {
    if (fs.existsSync(CACHE_PATH)) fs.unlinkSync(CACHE_PATH);
    throw new Error("Token expire. Relancez l'outil 'authenticate'.");
  }
}

function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
}

// ============================================================
// Serveur MCP
// ============================================================
const server = new McpServer({
  name: "outlook-mcp",
  version: "1.0.0",
});

// Tool: authenticate
server.tool("authenticate", "Authentification OAuth2 avec Outlook/Microsoft 365. Ouvre le navigateur pour se connecter.", {}, async () => {
  return new Promise((resolve) => {
    const app = express();
    let serverInstance;

    const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
      `client_id=${CLIENT_ID}` +
      `&response_type=code` +
      `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
      `&scope=${encodeURIComponent(SCOPES.map((s) => `https://graph.microsoft.com/${s}`).join(" ") + " offline_access")}` +
      `&response_mode=query`;

    app.get("/callback", async (req, res) => {
      const code = req.query.code;
      if (!code) {
        res.send("Erreur : pas de code recu.");
        serverInstance.close();
        resolve({ content: [{ type: "text", text: "Erreur d'authentification : pas de code." }] });
        return;
      }

      try {
        await cca.acquireTokenByCode({
          code,
          scopes: SCOPES.map((s) => `https://graph.microsoft.com/${s}`),
          redirectUri: REDIRECT_URI,
        });

        res.send("<html><body style='font-family:Arial;text-align:center;padding:50px'><h1>Connecte !</h1><p>Tu peux fermer cette fenetre.</p></body></html>");
        serverInstance.close();
        resolve({ content: [{ type: "text", text: "Authentification reussie ! Le token a ete sauvegarde." }] });
      } catch (err) {
        res.send("Erreur : " + err.message);
        serverInstance.close();
        resolve({ content: [{ type: "text", text: "Erreur : " + err.message }] });
      }
    });

    serverInstance = app.listen(3333, () => {
      if (process.platform === "win32") {
        exec(`start "" "${authUrl}"`);
      } else {
        const cmd = process.platform === "darwin" ? "open" : "xdg-open";
        exec(`${cmd} "${authUrl}"`);
      }
    });

    setTimeout(() => {
      serverInstance.close();
      resolve({ content: [{ type: "text", text: "Timeout : authentification annulee apres 2 minutes." }] });
    }, 120000);
  });
});

// Tool: list_emails
server.tool(
  "list_emails",
  "Lister les emails Outlook. Peut filtrer par dossier, nombre, et recherche.",
  {
    folder: z.string().optional().default("inbox").describe("Dossier : inbox, sentitems, drafts, junkemail, deleteditems"),
    count: z.number().optional().default(10).describe("Nombre de mails a retourner (max 50)"),
    search: z.string().optional().describe("Recherche dans les mails (sujet, corps, expediteur)"),
    from: z.string().optional().describe("Filtrer par expediteur (email ou nom)"),
    after: z.string().optional().describe("Mails apres cette date (YYYY-MM-DD)"),
    before: z.string().optional().describe("Mails avant cette date (YYYY-MM-DD)"),
    unread_only: z.boolean().optional().default(false).describe("Seulement les mails non lus"),
  },
  async ({ folder, count, search, from, after, before, unread_only }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const folderMap = {
      inbox: "inbox",
      sentitems: "sentitems",
      drafts: "drafts",
      junkemail: "junkemail",
      deleteditems: "deleteditems",
    };

    const folderPath = folderMap[folder.toLowerCase()] || "inbox";
    let url = `/me/mailFolders/${folderPath}/messages`;

    let filters = [];
    if (unread_only) filters.push("isRead eq false");
    if (after) filters.push(`receivedDateTime ge ${after}T00:00:00Z`);
    if (before) filters.push(`receivedDateTime le ${before}T23:59:59Z`);
    if (from) filters.push(`from/emailAddress/address eq '${from}' or contains(from/emailAddress/name, '${from}')`);

    let req = client.api(url)
      .top(Math.min(count, 50))
      .select("id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments")
      .orderby("receivedDateTime desc");

    if (filters.length) req = req.filter(filters.join(" and "));
    if (search) req = req.search(`"${search}"`);

    const result = await req.get();

    const emails = (result.value || []).map((m) => ({
      id: m.id,
      sujet: m.subject,
      de: m.from?.emailAddress?.name + " <" + m.from?.emailAddress?.address + ">",
      date: m.receivedDateTime,
      lu: m.isRead,
      apercu: m.bodyPreview?.substring(0, 200),
      pieces_jointes: m.hasAttachments,
    }));

    return {
      content: [{ type: "text", text: JSON.stringify(emails, null, 2) }],
    };
  }
);

// Tool: read_email
server.tool(
  "read_email",
  "Lire le contenu complet d'un email par son ID.",
  {
    id: z.string().describe("ID du mail (obtenu via list_emails)"),
  },
  async ({ id }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const msg = await client.api(`/me/messages/${id}`)
      .select("subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments")
      .expand("attachments($select=name,contentType,size)")
      .get();

    const email = {
      sujet: msg.subject,
      de: msg.from?.emailAddress?.name + " <" + msg.from?.emailAddress?.address + ">",
      a: (msg.toRecipients || []).map((r) => r.emailAddress?.name + " <" + r.emailAddress?.address + ">"),
      cc: (msg.ccRecipients || []).map((r) => r.emailAddress?.name + " <" + r.emailAddress?.address + ">"),
      date: msg.receivedDateTime,
      corps: msg.body?.content?.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().substring(0, 5000),
      pieces_jointes: (msg.attachments || []).map((a) => ({
        nom: a.name,
        type: a.contentType,
        taille: a.size,
      })),
    };

    return {
      content: [{ type: "text", text: JSON.stringify(email, null, 2) }],
    };
  }
);

// Tool: search_emails
server.tool(
  "search_emails",
  "Rechercher des emails par mots-cles dans tout Outlook.",
  {
    query: z.string().describe("Termes de recherche"),
    count: z.number().optional().default(10).describe("Nombre de resultats (max 50)"),
  },
  async ({ query, count }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const result = await client.api("/me/messages")
      .search(`"${query}"`)
      .top(Math.min(count, 50))
      .select("id,subject,from,receivedDateTime,bodyPreview,hasAttachments")
      .get();

    const emails = (result.value || []).map((m) => ({
      id: m.id,
      sujet: m.subject,
      de: m.from?.emailAddress?.name + " <" + m.from?.emailAddress?.address + ">",
      date: m.receivedDateTime,
      apercu: m.bodyPreview?.substring(0, 200),
    }));

    return {
      content: [{ type: "text", text: JSON.stringify(emails, null, 2) }],
    };
  }
);

// Tool: send_email
server.tool(
  "send_email",
  "Envoyer un email depuis Outlook.",
  {
    to: z.array(z.string()).describe("Destinataires (emails)"),
    subject: z.string().describe("Sujet du mail"),
    body: z.string().describe("Corps du mail (texte ou HTML)"),
    cc: z.array(z.string()).optional().describe("Destinataires en copie"),
    is_html: z.boolean().optional().default(false).describe("true si le corps est en HTML"),
  },
  async ({ to, subject, body, cc, is_html }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const message = {
      subject,
      body: {
        contentType: is_html ? "HTML" : "Text",
        content: body,
      },
      toRecipients: to.map((email) => ({ emailAddress: { address: email } })),
    };

    if (cc?.length) {
      message.ccRecipients = cc.map((email) => ({ emailAddress: { address: email } }));
    }

    await client.api("/me/sendMail").post({ message });

    return {
      content: [{ type: "text", text: `Email envoye a ${to.join(", ")}` }],
    };
  }
);

// Tool: create_event
server.tool(
  "create_event",
  "Creer un evenement dans le calendrier Outlook.",
  {
    subject: z.string().describe("Titre de l'evenement"),
    start: z.string().describe("Date/heure de debut (ISO 8601, ex: 2026-04-15T09:00:00)"),
    end: z.string().describe("Date/heure de fin (ISO 8601, ex: 2026-04-15T09:30:00)"),
    body: z.string().optional().describe("Description / notes de l'evenement (HTML ou texte)"),
    location: z.string().optional().describe("Lieu de l'evenement"),
    is_reminder: z.boolean().optional().default(true).describe("Activer le rappel"),
    reminder_minutes: z.number().optional().default(15).describe("Rappel X minutes avant"),
    attendees: z.array(z.string()).optional().describe("Emails des participants"),
    is_all_day: z.boolean().optional().default(false).describe("Evenement sur toute la journee"),
    is_html: z.boolean().optional().default(false).describe("true si le body est en HTML"),
  },
  async ({ subject, start, end, body, location, is_reminder, reminder_minutes, attendees, is_all_day, is_html }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const event = {
      subject,
      start: { dateTime: start, timeZone: "Europe/Paris" },
      end: { dateTime: end, timeZone: "Europe/Paris" },
      isReminderOn: is_reminder,
      reminderMinutesBeforeStart: reminder_minutes,
      isAllDay: is_all_day,
    };

    if (body) {
      event.body = { contentType: is_html ? "HTML" : "Text", content: body };
    }
    if (location) {
      event.location = { displayName: location };
    }
    if (attendees?.length) {
      event.attendees = attendees.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      }));
    }

    const result = await client.api("/me/events").post(event);

    return {
      content: [{
        type: "text",
        text: `Evenement cree : "${result.subject}" le ${result.start.dateTime} (ID: ${result.id})`,
      }],
    };
  }
);

// Tool: list_events
server.tool(
  "list_events",
  "Lister les evenements du calendrier Outlook.",
  {
    start: z.string().optional().describe("Date de debut (YYYY-MM-DD). Par defaut : aujourd'hui"),
    end: z.string().optional().describe("Date de fin (YYYY-MM-DD). Par defaut : +7 jours"),
    count: z.number().optional().default(20).describe("Nombre max d'evenements (max 50)"),
  },
  async ({ start, end, count }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);

    const now = new Date();
    const startDate = start || now.toISOString().split("T")[0];
    const endDate = end || new Date(now.getTime() + 7 * 86400000).toISOString().split("T")[0];

    const result = await client
      .api("/me/calendarView")
      .query({
        startDateTime: `${startDate}T00:00:00`,
        endDateTime: `${endDate}T23:59:59`,
      })
      .top(Math.min(count, 50))
      .select("id,subject,start,end,location,isAllDay,organizer,attendees,bodyPreview")
      .orderby("start/dateTime")
      .get();

    const events = (result.value || []).map((e) => ({
      id: e.id,
      sujet: e.subject,
      debut: e.start?.dateTime,
      fin: e.end?.dateTime,
      lieu: e.location?.displayName || null,
      journee_entiere: e.isAllDay,
      organisateur: e.organizer?.emailAddress?.name,
      participants: (e.attendees || []).map((a) => a.emailAddress?.address),
      apercu: e.bodyPreview?.substring(0, 200),
    }));

    return {
      content: [{ type: "text", text: JSON.stringify(events, null, 2) }],
    };
  }
);

// Tool: delete_event
server.tool(
  "delete_event",
  "Supprimer un evenement du calendrier Outlook.",
  {
    id: z.string().describe("ID de l'evenement (obtenu via list_events)"),
  },
  async ({ id }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    await client.api(`/me/events/${id}`).delete();
    return {
      content: [{ type: "text", text: "Evenement supprime." }],
    };
  }
);

// ============================================================
// Demarrage
// ============================================================
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(console.error);
