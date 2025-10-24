// src/index.ts
// Power BI MCP server — User-delegated OAuth (Authorization Code + PKCE via localhost)
// + Capability switchboard you can toggle from Claude, persisted to ~/.config/powerbi-mcp/capabilities.json
//
// Extras in this version:
// - Auto-parse stringified JSON for body/query/headers (prevents 400 from string bodies)
// - Improved error surfaces (status + details)
// - New helper tool: datasets_execute_dax
//
// Tools (underscores only):
//   Auth:
//     - powerbi_begin_browser_login  -> returns loginUrl (click to sign in)
//     - powerbi_auth_status          -> shows token status
//   Capability toggles (persisted):
//     - powerbi_capabilities_list
//     - powerbi_capabilities_enable   { name: string | string[] }
//     - powerbi_capabilities_disable  { name: string | string[] }
//     - powerbi_capabilities_set      { enabled: string[] }    // sets all others disabled
//   Generic:
//     - powerbi_request               (gated by capability "raw")
//   Convenience:
//     - datasets_execute_dax          (gated by "datasets"; wraps executeQueries)
//   Typed tools (each gated by a capability, each with mandatory title + description):
//     groups_*     -> "groups"
//     reports_*    -> "reports"
//     datasets_*   -> "datasets"
//     dashboards_* -> "dashboards"
//     capacities_* -> "capacities"
//     push_*       -> "push"
//     admin_*      -> "admin"
//     embed_*      -> "embed"
//
// Required ENV (no secrets):
//   PBI_PUBLIC_CLIENT_ID   = your Entra "Application (client) ID"
//   PBI_TENANT_ID          = your Tenant ID (optional; uses "common" if omitted)
//   PBI_SCOPES             = comma-separated delegated scopes (defaults provided below)
// Optional ENV:
//   PBI_CAPS_DEFAULT       = comma-separated names to enable by default on first run
//   PBI_CAPS_PATH          = persistence file path (default ~/.config/powerbi-mcp/capabilities.json)
//
// App registration must have redirect URIs (Mobile & desktop):
//   http://localhost   and   http://127.0.0.1
//
// NOTE: Loopback URIs allow any port for native apps.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import axios, { type AxiosInstance } from "axios";
import { z } from "zod";
import * as http from "node:http";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import {
  PublicClientApplication,
  CryptoProvider,
  type AuthorizationUrlRequest,
  type AuthorizationCodeRequest,
  type AccountInfo,
} from "@azure/msal-node";

/** ===== Config ===== */
const CLIENT_ID = mustEnv("PBI_PUBLIC_CLIENT_ID");
const TENANT_ID = process.env.PBI_TENANT_ID?.trim();
const AUTHORITY = TENANT_ID
  ? `https://login.microsoftonline.com/${TENANT_ID}`
  : "https://login.microsoftonline.com/common";

const DEFAULT_SCOPES = [
  "openid",
  "offline_access",
  "profile",
  "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
  "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
  "https://analysis.windows.net/powerbi/api/Dashboard.Read.All",
];
const SCOPES =
  process.env.PBI_SCOPES?.split(",").map((s) => s.trim()).filter(Boolean) ||
  DEFAULT_SCOPES;

/** ===== Capabilities (on/off) ===== */
// Coarse feature groups + granular switches you asked for earlier.
const ALL_CAPS = [
  // Coarse
  "groups",
  "reports",
  "datasets",
  "dashboards",
  "capacities",
  "push",
  "admin",
  "embed",
  "raw", // gates powerbi_request

  // Granular switches (independent; off by default unless you enable)
  "App.Read.All",
  "Content.Create",
  "Dashboard.Execute.All",
  "Dashboard.Read.All",
  "Dashboard.ReadWrite.All",
  "Dashboard.Reshare.All",
  "Dataflow.Execute.All",
  "Dataflow.Read.All",
  "Dataflow.ReadWrite.All",
  "Dataflow.Reshare.All",
  "Dataset.Read.All",
  "Gateway.Read.All",
  "Gateway.ReadWrite.All",
  "PaginatedReport.Execute.All",
  "PaginatedReport.Read.All",
  "PaginatedReport.ReadWrite.All",
  "PaginatedReport.Reshare.All",
  "PrincipalDetails.ReadBasic.All",
  "Report.Execute.All",
  "Report.ReadWrite.All",
  "Report.Reshare.All",
  "SemanticModel.Execute.All",
  "SemanticModel.Read.All",
  "SemanticModel.ReadWrite.All",
] as const;
type CapName = typeof ALL_CAPS[number];

function defaultEnabledCaps(): CapName[] {
  const fromEnv = (process.env.PBI_CAPS_DEFAULT || "groups,reports,datasets,dashboards").split(",")
    .map((s) => s.trim())
    .filter(Boolean) as CapName[];
  return normalizeCaps(fromEnv.length ? fromEnv : ["groups","reports","datasets","dashboards"]);
}

function normalizeCaps(names: string[] | undefined): CapName[] {
  const set = new Set<CapName>();
  for (const n of names || []) {
    if ((ALL_CAPS as readonly string[]).includes(n)) set.add(n as CapName);
  }
  return Array.from(set);
}

class CapabilityStore {
  private filePath: string;
  private map: Record<CapName, boolean>;

  constructor() {
    const cfgDir = path.join(os.homedir(), ".config", "powerbi-mcp");
    if (!fs.existsSync(cfgDir)) fs.mkdirSync(cfgDir, { recursive: true });
    this.filePath = process.env.PBI_CAPS_PATH?.trim() || path.join(cfgDir, "capabilities.json");
    this.map = Object.fromEntries(ALL_CAPS.map((c) => [c, false])) as Record<CapName, boolean>;
    this.loadOrInit();
  }

  private loadOrInit() {
    try {
      if (fs.existsSync(this.filePath)) {
        const obj = JSON.parse(fs.readFileSync(this.filePath, "utf8"));
        for (const c of ALL_CAPS) this.map[c] = !!obj[c];
      } else {
        for (const c of defaultEnabledCaps()) this.map[c] = true;
        this.save();
      }
    } catch {
      for (const c of defaultEnabledCaps()) this.map[c] = true;
      this.save();
    }
  }

  private save() {
    const payload = Object.fromEntries(ALL_CAPS.map((c) => [c, this.map[c]]));
    fs.writeFileSync(this.filePath, JSON.stringify(payload, null, 2));
  }

  list() {
    const enabled = ALL_CAPS.filter((c) => this.map[c]);
    const disabled = ALL_CAPS.filter((c) => !this.map[c]);
    return { enabled, disabled, path: this.filePath };
  }

  setAllEnabled(enabled: CapName[]) {
    const set = new Set(enabled);
    for (const c of ALL_CAPS) this.map[c] = set.has(c);
    this.save();
  }

  enable(names: CapName[]) {
    for (const n of names) this.map[n] = true;
    this.save();
  }

  disable(names: CapName[]) {
    for (const n of names) this.map[n] = false;
    this.save();
  }

  isEnabled(name: CapName) {
    return !!this.map[name];
  }
}

/** ===== Auth state (in-memory) ===== */
class PKCEAuth {
  private pca: PublicClientApplication;
  private crypto = new CryptoProvider();

  private accessToken?: string;
  private expiresAtSec = 0;
  private account?: AccountInfo;

  // one login at a time
  private pending?: {
    codeVerifier: string;
    codeChallenge: string;
    state: string;
    redirectUri: string;
    server: http.Server;
  };

  constructor() {
    this.pca = new PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY },
      system: { loggerOptions: { loggerCallback: () => {} } },
    });
  }

  status() {
    const now = Math.floor(Date.now() / 1000);
    const remaining = this.accessToken ? Math.max(0, this.expiresAtSec - now) : null;
    return {
      hasToken: !!this.accessToken,
      expiresAt: this.expiresAtSec || null,
      secondsRemaining: remaining,
      pendingLogin: !!this.pending,
    };
  }

  /** Get a valid token, using silent refresh when possible */
  getToken = async (): Promise<string> => {
    const now = Math.floor(Date.now() / 1000);
    if (this.accessToken && now < this.expiresAtSec - 60) return this.accessToken;

    if (this.account) {
      const res = await this.pca.acquireTokenSilent({
        account: this.account,
        scopes: SCOPES,
        forceRefresh: false,
      });
      if (res?.accessToken) {
        this.accessToken = res.accessToken;
        this.expiresAtSec = res.expiresOn ? Math.floor(res.expiresOn.getTime() / 1000) : now + 3000;
        return this.accessToken;
      }
    }

    throw new Error("Not authenticated or token expired. Run powerbi_begin_browser_login.");
  };

  /** Starts a localhost listener and returns a login URL */
  async beginBrowserLogin(): Promise<{ loginUrl: string; redirectUri: string }> {
    if (this.pending) {
      const { redirectUri, state, codeChallenge } = this.pending;
      const loginUrl = await this.makeAuthUrl(redirectUri, state, codeChallenge);
      return { loginUrl, redirectUri };
    }

    const { verifier: codeVerifier, challenge: codeChallenge } = await this.crypto.generatePkceCodes();
    const state = randomId();

    const { server, redirectUri } = await startLoopbackServer(async (fullUrl) => {
      const url = new URL(fullUrl);
      if (url.pathname !== "/callback") return;
      const code = url.searchParams.get("code");
      const gotState = url.searchParams.get("state");
      if (!code || gotState !== state) throw new Error("Invalid OAuth response");

      const tokenReq: AuthorizationCodeRequest = {
        scopes: SCOPES,
        code,
        redirectUri,
        codeVerifier,
      };
      const res = await this.pca.acquireTokenByCode(tokenReq);
      if (!res?.accessToken) throw new Error("No access token in response");

      this.accessToken = res.accessToken;
      this.expiresAtSec = res.expiresOn ? Math.floor(res.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3000;
      this.account = res.account ?? this.account;

      try { server.close(); } catch {}
      this.pending = undefined;
    });

    this.pending = { codeVerifier, codeChallenge, state, redirectUri, server };

    const loginUrl = await this.makeAuthUrl(redirectUri, state, codeChallenge);
    return { loginUrl, redirectUri };
  }

  private async makeAuthUrl(redirectUri: string, state: string, codeChallenge: string): Promise<string> {
    const urlReq: AuthorizationUrlRequest = {
      scopes: SCOPES,
      redirectUri,
      state,
      codeChallenge,
      codeChallengeMethod: "S256",
      prompt: "select_account",
    };
    return await this.pca.getAuthCodeUrl(urlReq);
  }
}

/** Start a localhost HTTP server on an ephemeral port and return redirectUri */
async function startLoopbackServer(onCallback: (fullUrl: string) => Promise<void>) {
  const srv = http.createServer(async (req, res) => {
    try {
      if (!req.url) throw new Error("No URL");
      const port = (srv.address() as any).port;
      const fullUrl = `http://127.0.0.1:${port}${req.url}`;
      if (req.url.startsWith("/callback")) {
        await onCallback(fullUrl);
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(`<html><body><h2>Login complete</h2><p>You can close this window and return to Claude.</p></body></html>`);
      } else {
        res.writeHead(404);
        res.end("Not found");
      }
    } catch (e: any) {
      res.writeHead(500, { "Content-Type": "text/plain" });
      res.end(`Auth error: ${e?.message || String(e)}`);
    }
  });

  await new Promise<void>((resolve) => srv.listen(0, "127.0.0.1", () => resolve()));
  const port = (srv.address() as any).port as number;
  const redirectUri = `http://127.0.0.1:${port}/callback`;
  return { server: srv, redirectUri };
}

/** Utilities */
function randomId() {
  return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}
function mustEnv(k: string) {
  const v = process.env[k];
  if (!v) throw new Error(`Missing env ${k}`);
  return v.trim();
}

/** ===== HTTP client (injects token) ===== */
function makeClient(getToken: () => Promise<string>): AxiosInstance {
  const client = axios.create({ baseURL: "https://api.powerbi.com", timeout: 60000 });
  client.interceptors.request.use(async (cfg) => {
    const t = await getToken();
    cfg.headers = { ...(cfg.headers as any), Authorization: `Bearer ${t}` } as any;
    return cfg;
  });
  client.interceptors.response.use(undefined, async (err) => {
    const status = err?.response?.status;
    if (status === 429 || status === 503) {
      const retryAfter = Number(err?.response?.headers?.["retry-after"] ?? 2);
      await new Promise((r) => setTimeout(r, Math.max(1, retryAfter) * 1000));
      return client.request(err.config);
    }
    throw err;
  });
  return client;
}

/** ===== Tool catalog (underscores only, with mandatory title + description) ===== */
type Op = {
  name: string;
  title: string;         // mandatory
  description: string;   // mandatory
  method: "GET" | "POST" | "PATCH" | "DELETE";
  path: string;
  pathParams?: string[];
  hasBody?: boolean;
  cap: CapName; // coarse capability gate
};

const ops: Op[] = [
  // Groups (Workspaces)
  {
    name: "groups_get_groups",
    title: "Groups: List workspaces",
    description: "GET /v1.0/myorg/groups — Lists workspaces the signed-in user can access.",
    method: "GET", path: "/v1.0/myorg/groups", cap: "groups",
  },
  {
    name: "groups_get_group_users",
    title: "Groups: List users in a workspace",
    description: "GET /v1.0/myorg/groups/{groupId}/users — Lists principals (users/service principals) in a workspace.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/users", pathParams: ["groupId"], cap: "groups",
  },
  {
    name: "groups_add_user",
    title: "Groups: Add user to workspace",
    description: "POST /v1.0/myorg/groups/{groupId}/users — Adds a principal to a workspace with a specified access right.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/users", pathParams: ["groupId"], hasBody: true, cap: "groups",
  },
  {
    name: "groups_delete_user_in_group",
    title: "Groups: Remove user from workspace",
    description: "DELETE /v1.0/myorg/groups/{groupId}/users/{user} — Removes a principal from a workspace.",
    method: "DELETE", path: "/v1.0/myorg/groups/{groupId}/users/{user}", pathParams: ["groupId","user"], cap: "groups",
  },

  // Reports
  {
    name: "reports_get_reports",
    title: "Reports: List all reports (My Org)",
    description: "GET /v1.0/myorg/reports — Lists reports available to the signed-in user in My Org.",
    method: "GET", path: "/v1.0/myorg/reports", cap: "reports",
  },
  {
    name: "reports_get_reports_in_group",
    title: "Reports: List reports in workspace",
    description: "GET /v1.0/myorg/groups/{groupId}/reports — Lists reports within a specific workspace.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/reports", pathParams: ["groupId"], cap: "reports",
  },
  {
    name: "reports_get_report_in_group",
    title: "Reports: Get report (workspace)",
    description: "GET /v1.0/myorg/groups/{groupId}/reports/{reportId} — Gets metadata for a specific report in a workspace.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}", pathParams: ["groupId","reportId"], cap: "reports",
  },
  {
    name: "reports_clone_in_group",
    title: "Reports: Clone report",
    description: "POST /v1.0/myorg/groups/{groupId}/reports/{reportId}/Clone — Clones a report to a target workspace or dataset.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}/Clone", pathParams: ["groupId","reportId"], hasBody: true, cap: "reports",
  },
  {
    name: "reports_rebind_in_group",
    title: "Reports: Rebind report to dataset",
    description: "POST /v1.0/myorg/groups/{groupId}/reports/{reportId}/Rebind — Rebinds a report to a different dataset.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}/Rebind", pathParams: ["groupId","reportId"], hasBody: true, cap: "reports",
  },
  {
    name: "reports_export_to_file_in_group",
    title: "Reports: Export to file",
    description: "POST /v1.0/myorg/groups/{groupId}/reports/{reportId}/ExportTo — Exports a report to a file format (e.g., PDF, PPTX).",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}/ExportTo", pathParams: ["groupId","reportId"], hasBody: true, cap: "reports",
  },
  {
    name: "reports_get_export_to_file_status_in_group",
    title: "Reports: Get export status",
    description: "GET /v1.0/myorg/groups/{groupId}/reports/{reportId}/exports/{exportId} — Gets export job status.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}/exports/{exportId}", pathParams: ["groupId","reportId","exportId"], cap: "reports",
  },
  {
    name: "reports_get_file_of_export_to_file_in_group",
    title: "Reports: Download exported file",
    description: "GET /v1.0/myorg/groups/{groupId}/reports/{reportId}/exports/{exportId}/file — Downloads the exported file bytes.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/reports/{reportId}/exports/{exportId}/file", pathParams: ["groupId","reportId","exportId"], cap: "reports",
  },

  // Datasets
  {
    name: "datasets_get_datasets_in_group",
    title: "Datasets: List datasets in workspace",
    description: "GET /v1.0/myorg/groups/{groupId}/datasets — Lists datasets in a workspace.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/datasets", pathParams: ["groupId"], cap: "datasets",
  },
  {
    name: "datasets_execute_queries_in_group",
    title: "Datasets: Execute DAX queries",
    description: "POST /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/executeQueries — Executes a single DAX query that returns a table.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/executeQueries", pathParams: ["groupId","datasetId"], hasBody: true, cap: "datasets",
  },
  {
    name: "datasets_refresh_in_group",
    title: "Datasets: Trigger refresh",
    description: "POST /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes — Triggers a dataset refresh.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes", pathParams: ["groupId","datasetId"], hasBody: true, cap: "datasets",
  },
  {
    name: "datasets_get_refresh_history_in_group",
    title: "Datasets: Get refresh history",
    description: "GET /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes — Gets refresh executions for a dataset.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshes", pathParams: ["groupId","datasetId"], cap: "datasets",
  },
  {
    name: "datasets_get_refresh_schedule_in_group",
    title: "Datasets: Get refresh schedule",
    description: "GET /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule — Gets the scheduled refresh settings.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule", pathParams: ["groupId","datasetId"], cap: "datasets",
  },
  {
    name: "datasets_update_refresh_schedule_in_group",
    title: "Datasets: Update refresh schedule",
    description: "PATCH /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule — Updates the dataset’s scheduled refresh.",
    method: "PATCH", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/refreshSchedule", pathParams: ["groupId","datasetId"], hasBody: true, cap: "datasets",
  },

  // Push datasets
  {
    name: "push_post_dataset_in_group",
    title: "Push: Create dataset (push)",
    description: "POST /v1.0/myorg/groups/{groupId}/datasets — Creates a push dataset in a workspace.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/datasets", pathParams: ["groupId"], hasBody: true, cap: "push",
  },
  {
    name: "push_post_rows_in_group",
    title: "Push: Add rows to table",
    description: "POST /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/tables/{tableName}/rows — Adds rows to a table in a push dataset.",
    method: "POST", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/tables/{tableName}/rows", pathParams: ["groupId","datasetId","tableName"], hasBody: true, cap: "push",
  },
  {
    name: "push_delete_rows_in_group",
    title: "Push: Delete all rows in table",
    description: "DELETE /v1.0/myorg/groups/{groupId}/datasets/{datasetId}/tables/{tableName}/rows — Deletes all rows from a push dataset table.",
    method: "DELETE", path: "/v1.0/myorg/groups/{groupId}/datasets/{datasetId}/tables/{tableName}/rows", pathParams: ["groupId","datasetId","tableName"], cap: "push",
  },

  // Dashboards
  {
    name: "dashboards_get_dashboards_in_group",
    title: "Dashboards: List dashboards in workspace",
    description: "GET /v1.0/myorg/groups/{groupId}/dashboards — Lists dashboards in a workspace.",
    method: "GET", path: "/v1.0/myorg/groups/{groupId}/dashboards", pathParams: ["groupId"], cap: "dashboards",
  },

  // Capacities
  {
    name: "capacities_get_capacities",
    title: "Capacities: List capacities",
    description: "GET /v1.0/myorg/capacities — Lists capacities available to the tenant/user (where applicable).",
    method: "GET", path: "/v1.0/myorg/capacities", cap: "capacities",
  },
  {
    name: "capacities_get_refreshables_for_capacity",
    title: "Capacities: List refreshables",
    description: "GET /v1.0/myorg/capacities/{capacityId}/refreshables — Lists refreshable items for a capacity.",
    method: "GET", path: "/v1.0/myorg/capacities/{capacityId}/refreshables", pathParams: ["capacityId"], cap: "capacities",
  },

  // Admin
  {
    name: "admin_get_activity_events",
    title: "Admin: Get activity events",
    description: "GET /v1.0/myorg/admin/activityevents — Retrieves activity events (requires admin privileges).",
    method: "GET", path: "/v1.0/myorg/admin/activityevents", cap: "admin",
  },

  // Embed tokens
  {
    name: "embed_generate_token",
    title: "Embed: Generate token",
    description: "POST /v1.0/myorg/GenerateToken — Generates an embed token for supported artifacts.",
    method: "POST", path: "/v1.0/myorg/GenerateToken", hasBody: true, cap: "embed",
  },
];

/** ===== MCP server ===== */
const mcpServer = new McpServer({ name: "powerbi-mcp", version: "4.4.0" });
const auth = new PKCEAuth();
const httpClient = makeClient(() => auth.getToken());
const caps = new CapabilityStore();
const tools: string[] = [];

/** Helpers */
function gate<T extends Record<string, unknown>>(cap: CapName, handler: (args: T) => Promise<any>) {
  return async (args: T) => {
    if (!caps.isEnabled(cap)) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: `Capability '${cap}' is disabled. Enable it with powerbi_capabilities_enable.`,
              capability: cap,
            }),
          },
        ],
      };
    }
    return handler(args);
  };
}
function tryParseMaybe(v: any) {
  if (v == null) return v;
  if (typeof v === "string") {
    try { return JSON.parse(v); } catch { return v; }
  }
  return v;
}

/** Capability tools */
mcpServer.registerTool(
  "powerbi_capabilities_list",
  {
    title: "Power BI Capabilities: List",
    description: "Lists which capabilities are enabled/disabled and the persistence file location.",
    inputSchema: {},
  },
  async () => ({ content: [{ type: "text", text: JSON.stringify(caps.list()) }] })
);
tools.push("powerbi_capabilities_list");

mcpServer.registerTool(
  "powerbi_capabilities_enable",
  {
    title: "Power BI Capabilities: Enable",
    description: `Enable one or more capabilities. Valid: ${ALL_CAPS.join(", ")}.`,
    inputSchema: { name: z.union([z.string(), z.array(z.string())]) },
  },
  async ({ name }) => {
    const names = Array.isArray(name) ? name : [name];
    const valid = normalizeCaps(names);
    if (!valid.length) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "No valid capability names provided." }) }] };
    }
    caps.enable(valid);
    return { content: [{ type: "text", text: JSON.stringify({ ok: true, ...caps.list() }) }] };
  }
);
tools.push("powerbi_capabilities_enable");

mcpServer.registerTool(
  "powerbi_capabilities_disable",
  {
    title: "Power BI Capabilities: Disable",
    description: `Disable one or more capabilities. Valid: ${ALL_CAPS.join(", ")}.`,
    inputSchema: { name: z.union([z.string(), z.array(z.string())]) },
  },
  async ({ name }) => {
    const names = Array.isArray(name) ? name : [name];
    const valid = normalizeCaps(names);
    if (!valid.length) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "No valid capability names provided." }) }] };
    }
    caps.disable(valid);
    return { content: [{ type: "text", text: JSON.stringify({ ok: true, ...caps.list() }) }] };
  }
);
tools.push("powerbi_capabilities_disable");

mcpServer.registerTool(
  "powerbi_capabilities_set",
  {
    title: "Power BI Capabilities: Set",
    description: `Enable exactly the provided list (all others disabled). Valid: ${ALL_CAPS.join(", ")}.`,
    inputSchema: { enabled: z.array(z.string()) },
  },
  async ({ enabled }) => {
    const valid = normalizeCaps(enabled);
    caps.setAllEnabled(valid);
    return { content: [{ type: "text", text: JSON.stringify({ ok: true, ...caps.list() }) }] };
  }
);
tools.push("powerbi_capabilities_set");

/** Auth tools */
mcpServer.registerTool(
  "powerbi_begin_browser_login",
  {
    title: "Power BI Auth: Begin browser login (PKCE)",
    description: "Returns a loginUrl. Click it, sign in, then verify with powerbi_auth_status.",
    inputSchema: {},
  },
  async () => {
    try {
      const { loginUrl, redirectUri } = await auth.beginBrowserLogin();
      return { content: [{ type: "text", text: JSON.stringify({ ok: true, loginUrl, redirectUri, scopes: SCOPES }) }] };
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: e?.message || String(e) }) }] };
    }
  }
);
tools.push("powerbi_begin_browser_login");

mcpServer.registerTool(
  "powerbi_auth_status",
  {
    title: "Power BI Auth: Status",
    description: "Shows whether you’re logged in and when the access token expires.",
    inputSchema: {},
  },
  async () => ({ content: [{ type: "text", text: JSON.stringify(auth.status()) }] })
);
tools.push("powerbi_auth_status");

/** Convenience: execute DAX safely (prevents stringified-body issues) */
mcpServer.registerTool(
  "datasets_execute_dax",
  {
    title: "Datasets: Execute DAX (helper)",
    description: "Runs a single DAX query against a dataset (wraps POST /executeQueries). Your DAX must return a table (use EVALUATE).",
    inputSchema: {
      groupId: z.string(),
      datasetId: z.string(),
      dax: z.string(),
      includeNulls: z.boolean().optional(),
    },
  },
  gate("datasets", async ({ groupId, datasetId, dax, includeNulls }) => {
    try {
      const url = `/v1.0/myorg/groups/${encodeURIComponent(groupId)}/datasets/${encodeURIComponent(datasetId)}/executeQueries`;
      const body = {
        queries: [{ query: dax }],
        serializerSettings: includeNulls ? { includeNulls: true } : undefined,
      };
      const res = await httpClient.request({ method: "POST", url, data: body });
      return { content: [{ type: "text", text: JSON.stringify(res.data) }] };
    } catch (e: any) {
      const status = e?.response?.status;
      const details = e?.response?.data;
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, status, error: e?.message || String(e), details }) }] };
    }
  }))
);
tools.push("datasets_execute_dax");

/** Generic request (gated by 'raw') with auto-parse + better errors */
mcpServer.registerTool(
  "powerbi_request",
  {
    title: "Power BI Request (raw)",
    description: "Calls a documented Power BI REST endpoint under /v1.0/myorg. (Gated by capability 'raw'.) Auto-parses stringified JSON for body/query/headers.",
    inputSchema: {
      method: z.enum(["GET", "POST", "PATCH", "DELETE"]),
      path: z.string().regex(/^\/v1\.0\/myorg(\/.*)?$/),
      query: z.any().optional(),
      body: z.any().optional(),
      headers: z.any().optional(),
    },
  },
  gate("raw", async ({ method, path, query, body, headers }) => {
    try {
      const res = await httpClient.request({
        method,
        url: path,
        params: tryParseMaybe(query),
        data: tryParseMaybe(body),
        headers: tryParseMaybe(headers),
        responseType: path.endsWith("/file") ? "arraybuffer" : "json",
      });

      if (path.endsWith("/file")) {
        const ct = res.headers["content-type"] || "application/octet-stream";
        return { content: [{ type: "text", text: JSON.stringify({ contentType: ct, base64: Buffer.from(res.data).toString("base64") }) }] };
      }
      return { content: [{ type: "text", text: JSON.stringify(res.data) }] };
    } catch (e: any) {
      const status = e?.response?.status;
      const details = e?.response?.data;
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, status, error: e?.message || String(e), details }) }] };
    }
  }))
);
tools.push("powerbi_request");

/** Typed wrappers (each gated by its coarse capability) — with auto-parse + better errors */
for (const op of ops) {
  const shape: Record<string, z.ZodTypeAny> = { query: z.any().optional() };
  for (const p of op.pathParams ?? []) shape[p] = z.string();
  if (op.hasBody) shape.body = z.any();

  mcpServer.registerTool(
    op.name,
    { title: op.title, description: op.description, inputSchema: shape },
    gate(op.cap, async (args: Record<string, unknown>) => {
      try {
        const url = op.path.replace(/\{(\w+)\}/g, (_, k) => {
          const v = (args as any)[k];
          if (!v) throw new Error(`Missing path param ${k}`);
          return encodeURIComponent(String(v));
        });

        const res = await httpClient.request({
          method: op.method,
          url,
          params: tryParseMaybe((args as any).query),
          data: op.hasBody ? tryParseMaybe((args as any).body) : undefined,
          responseType: url.endsWith("/file") ? "arraybuffer" : "json",
        });

        if (url.endsWith("/file")) {
          const ct = res.headers["content-type"] || "application/octet-stream";
          return { content: [{ type: "text", text: JSON.stringify({ contentType: ct, base64: Buffer.from(res.data).toString("base64") }) }] };
        }
        return { content: [{ type: "text", text: JSON.stringify(res.data) }] };
      } catch (e: any) {
        const status = e?.response?.status;
        const details = e?.response?.data;
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, status, error: e?.message || String(e), details }) }] };
      }
    })
  );
  tools.push(op.name);
}

/** Start */
(async () => {
  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);
  console.error(`Power BI MCP server ready (PKCE + caps). Authority=${AUTHORITY}`);
  console.error(`ClientId=${CLIENT_ID}`);
  console.error(`Scopes=${SCOPES.join(", ")}`);
  console.error(`Capabilities file: ${caps.list().path}`);
  console.error(`Enabled: ${caps.list().enabled.join(", ") || "(none)"}`);
  console.error(`Registered tools (${tools.length}): ${tools.join(", ")}`);
})();