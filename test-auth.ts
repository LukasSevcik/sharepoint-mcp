import { getAccessToken } from "./src/auth.js";

const SITE_URL = process.argv[2] ?? "https://mill.sharepoint.com/sites/demo";

console.log(`Testing auth for: ${SITE_URL}`);
console.log("Browser will open for login...\n");

try {
  const token = await getAccessToken(SITE_URL);
  console.log("SUCCESS! Token received.");
  console.log(`Token (first 80 chars): ${token.slice(0, 80)}...`);
} catch (err) {
  console.error("FAILED:", err);
}

// z terminalu zavolat
//npx tsx test-auth.ts
//npx tsx test-auth.ts https://contoso.sharepoint.com/sites/x