import { getAccessToken } from "./auth.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

class SpError extends Error {
  constructor(public status: number, body: string) {
    let msg = `SharePoint API ${status}`;
    try {
      const j = JSON.parse(body) as { error?: { message?: { value?: string } | string } };
      const m = j?.error?.message;
      const detail = typeof m === "object" ? (m as { value?: string })?.value : m;
      msg += ": " + (detail ?? body);
    } catch {
      msg += ": " + body;
    }
    super(msg);
  }
}

export function escapeOData(s: string): string {
  return s.replace(/'/g, "''");
}

// ---------------------------------------------------------------------------
// SharePoint REST client
// ---------------------------------------------------------------------------

export class SharePointClient {
  private siteUrl = "";

  setSiteUrl(url: string): void {
    this.siteUrl = url.replace(/\/$/, "");
  }

  getSiteUrl(): string {
    return this.siteUrl;
  }

  /** /sites/mysite  (no trailing slash) */
  getSitePath(): string {
    return new URL(this.siteUrl).pathname.replace(/\/$/, "");
  }

  // -------------------------------------------------------------------------
  // Private request kernel
  // -------------------------------------------------------------------------
  private async request<T>(
    method: string,
    apiPath: string,
    body?: unknown,
    extraHeaders?: Record<string, string>,
    rawBody?: Buffer
  ): Promise<T> {
    if (!this.siteUrl) {
      throw new Error(
        "Site URL nie je nastavený. Zavolaj najprv nástroj set_site."
      );
    }

    const token = await getAccessToken(this.siteUrl);
    const url = `${this.siteUrl}/_api/${apiPath.replace(/^\//, "")}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      ...extraHeaders,
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let fetchBody: any;
    if (rawBody) {
      fetchBody = rawBody; // Buffer is a valid Node.js 18+ fetch body
    } else if (body !== undefined) {
      headers["Content-Type"] = "application/json;odata=verbose";
      fetchBody = JSON.stringify(body);
    }

    const res = await fetch(url, { method, headers, body: fetchBody });

    if (!res.ok) {
      throw new SpError(res.status, await res.text());
    }

    if (res.status === 204 || res.headers.get("content-length") === "0") {
      return null as T;
    }

    const text = await res.text();
    if (!text) return null as T;
    return JSON.parse(text) as T;
  }

  // -------------------------------------------------------------------------
  // Public CRUD helpers
  // -------------------------------------------------------------------------

  async get<T = Record<string, unknown>>(
    path: string,
    qs?: Record<string, string>
  ): Promise<T> {
    let p = path;
    if (qs && Object.keys(qs).length) {
      p += "?" + new URLSearchParams(qs).toString();
    }
    return this.request<T>("GET", p);
  }

  async post<T = Record<string, unknown>>(
    path: string,
    body: unknown
  ): Promise<T> {
    return this.request<T>("POST", path, body);
  }

  async patch<T = Record<string, unknown>>(
    path: string,
    body: unknown
  ): Promise<T> {
    return this.request<T>("PATCH", path, body, { "IF-MATCH": "*" });
  }

  async delete(path: string): Promise<void> {
    await this.request<null>("DELETE", path, undefined, { "IF-MATCH": "*" });
  }

  /** Binary upload (file content as Buffer) */
  async uploadBinary<T = Record<string, unknown>>(
    path: string,
    buffer: Buffer
  ): Promise<T> {
    return this.request<T>("POST", path, undefined, {}, buffer);
  }

  // -------------------------------------------------------------------------
  // Convenience: fetch server-relative URL of a library's root folder
  // -------------------------------------------------------------------------
  async getLibraryRootUrl(libraryName: string): Promise<string> {
    const data = await this.get<{ ServerRelativeUrl: string }>(
      `web/lists/GetByTitle('${escapeOData(libraryName)}')/RootFolder`,
      { $select: "ServerRelativeUrl" }
    );
    return data.ServerRelativeUrl;
  }

  // -------------------------------------------------------------------------
  // Convenience: get ListItemEntityTypeFullName (needed for create/update items)
  // -------------------------------------------------------------------------
  private _entityTypeCache = new Map<string, string>();

  async getEntityTypeName(listName: string): Promise<string> {
    // Key includes siteUrl so switching sites doesn't return stale values
    const key = `${this.siteUrl}::${listName.toLowerCase()}`;
    if (!this._entityTypeCache.has(key)) {
      const data = await this.get<{ ListItemEntityTypeFullName: string }>(
        `web/lists/GetByTitle('${escapeOData(listName)}')`,
        { $select: "ListItemEntityTypeFullName" }
      );
      this._entityTypeCache.set(key, data.ListItemEntityTypeFullName);
    }
    return this._entityTypeCache.get(key)!;
  }
}

// Singleton used by all tool modules
export const sp = new SharePointClient();
