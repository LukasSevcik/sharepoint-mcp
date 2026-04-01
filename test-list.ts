import { getAccessToken } from "./src/auth.js";

const SITE_URL = process.argv[2] ?? "https://mill.sharepoint.com/sites/demo";
const LIST_NAME = process.argv[3] ?? "DRAiReviewers";

console.log(`Site:  ${SITE_URL}`);
console.log(`List:  ${LIST_NAME}\n`);

try {
  const token = await getAccessToken(SITE_URL);

  // Decode token claims (no signature verification, just inspect)
  const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
  console.log("Token claims:");
  console.log("  aud:", payload.aud);
  console.log("  iss:", payload.iss);
  console.log("  tid:", payload.tid);
  console.log("  scp:", payload.scp);
  console.log("  appid:", payload.appid);
  console.log("  exp:", new Date(payload.exp * 1000).toISOString());
  console.log()

  // Discover what tenant SharePoint expects
  const discovery = await fetch(`${SITE_URL}/_vti_bin/client.svc/`, { method: "GET" });
  const wwwAuth = discovery.headers.get("www-authenticate") ?? "";
  const realmMatch = wwwAuth.match(/realm="([^"]+)"/);
  const spTenantId = realmMatch?.[1] ?? "unknown";
  console.log("SharePoint expects realm (tenant):", spTenantId);
  console.log("Token tenant (tid):               ", payload.tid);
  console.log("Match:", spTenantId === payload.tid ? "YES ✓" : "NO ✗ — MISMATCH!");
  console.log();

  console.log("Calling SharePoint REST...\n");

  // Headers from unauthenticated request (to see what realm SharePoint expects)
  const unauth = await fetch(`${SITE_URL}/_api/web`);
  console.log("Unauthenticated /_api/web status:", unauth.status);
  console.log("  www-authenticate:", unauth.headers.get("www-authenticate"));
  console.log("  x-ms-diagnostics:", unauth.headers.get("x-ms-diagnostics"));
  console.log();

  // Test 1: basic web endpoint
  const test1 = await fetch(`${SITE_URL}/_api/web`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  console.log(`Test 1 /_api/web: ${test1.status}`);
  console.log("  www-authenticate:", test1.headers.get("www-authenticate"));
  console.log("  x-ms-diagnostics:", test1.headers.get("x-ms-diagnostics"));

  // Test 2: same but with odata=nometadata
  const test2 = await fetch(`${SITE_URL}/_api/web`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
  });
  console.log(`Test 2 /_api/web (odata=nometadata):       ${test2.status}`);

  // Test 3: list items
  const test3url = `${SITE_URL}/_api/web/lists/GetByTitle('${LIST_NAME}')/items?$top=5`;
  const test3 = await fetch(test3url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
  });
  const test3body = await test3.text();
  console.log(`Test 3 list items:                         ${test3.status}`);
  if (!test3.ok) console.log("  body:", test3body);
  else {
    const data = JSON.parse(test3body) as { value: unknown[] };
    console.log(`  Got ${data.value.length} item(s)`);
  }
} catch (err) {
  console.error("ERROR:", err);
  process.exit(1);
}

// Usage:
//   npx tsx test-list.ts
//   npx tsx test-list.ts https://mill.sharepoint.com/sites/demo DRAiReviewers
