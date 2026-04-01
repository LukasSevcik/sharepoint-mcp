import {
  PublicClientApplication,
  type ICachePlugin,
  type TokenCacheContext,
  type AccountInfo,
} from "@azure/msal-node";
import fs from "fs";
import path from "path";
import os from "os";

// PnP Management Shell — public client designed for SharePoint delegated access
const CLIENT_ID = "31359c7f-bd7e-475c-86db-fdb8c937548e";
const AUTHORITY = "https://login.microsoftonline.com/organizations";
const CACHE_DIR = path.join(os.homedir(), ".sharepoint-mcp");
const CACHE_FILE = path.join(CACHE_DIR, "token-cache.json");

// ---------------------------------------------------------------------------
// File-based token cache (persists logins between server restarts)
// ---------------------------------------------------------------------------
class FileCachePlugin implements ICachePlugin {
  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (fs.existsSync(CACHE_FILE)) {
      ctx.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (ctx.cacheHasChanged) {
      fs.mkdirSync(CACHE_DIR, { recursive: true });
      fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize());
    }
  }
}

// ---------------------------------------------------------------------------
// Singleton PCA
// ---------------------------------------------------------------------------
let _pca: PublicClientApplication | null = null;

function getPca(): PublicClientApplication {
  if (!_pca) {
    _pca = new PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY },
      cache: { cachePlugin: new FileCachePlugin() },
    });
  }
  return _pca;
}

// ---------------------------------------------------------------------------
// Derive SharePoint scope from site URL
// e.g. https://contoso.sharepoint.com/sites/x  →  https://contoso.sharepoint.com/.default
// ---------------------------------------------------------------------------
export function scopeFromSiteUrl(siteUrl: string): string {
  const url = new URL(siteUrl);
  return `https://${url.hostname}/.default`;
}

// ---------------------------------------------------------------------------
// Main export: get a valid access token (silent first, browser fallback)
// ---------------------------------------------------------------------------
export async function getAccessToken(siteUrl: string): Promise<string> {
  const scope = scopeFromSiteUrl(siteUrl);
  const pca = getPca();
  const cache = pca.getTokenCache();

  // 1. Try silent with every cached account
  const accounts: AccountInfo[] = await cache.getAllAccounts();
  for (const account of accounts) {
    try {
      const res = await pca.acquireTokenSilent({ account, scopes: [scope] });
      if (res?.accessToken) return res.accessToken;
    } catch {
      // expired / wrong tenant — try next
    }
  }

  // 2. Interactive browser login
  process.stderr.write(
    "\n[sharepoint-mcp] Opening browser for authentication…\n"
  );

  const { default: open } = await import("open");

  const res = await pca.acquireTokenInteractive({
    scopes: [scope],
    prompt: "select_account",
    openBrowser: async (url: string) => {
      await open(url);
    },
    successTemplate: `
      <html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;padding:40px;text-align:center">
        <h2>&#10003; Authentication successful!</h2>
        <p>You can close this tab and return to the terminal.</p>
      </body></html>`,
    errorTemplate: `
      <html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;padding:40px;text-align:center">
        <h2>&#10007; Authentication failed</h2><p>{errorMessage}</p>
      </body></html>`,
  });

  if (!res?.accessToken) throw new Error("Nepodarilo sa získať access token.");
  return res.accessToken;
}

// ---------------------------------------------------------------------------
// Expose for logout / cache clearing
// ---------------------------------------------------------------------------
export async function clearTokenCache(): Promise<void> {
  if (fs.existsSync(CACHE_FILE)) fs.rmSync(CACHE_FILE);
  _pca = null; // force re-init
}
