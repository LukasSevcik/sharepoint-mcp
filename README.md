# sharepoint-mcp

MCP server pre SharePoint Online — správa listov, knižníc, stĺpcov, pohľadov, položiek a súborov priamo z AI agenta.

---

## Štruktúra projektu

```
sharepoint-mcp/
├── .env.example          # Šablóna konfigurácie (skopíruj do .env)
├── package.json
├── tsconfig.json
└── src/
    ├── index.ts          # MCP server + registrácia všetkých nástrojov
    ├── auth.ts           # MSAL browser login + file token cache
    ├── sharepoint.ts     # SharePoint REST klient (GET / POST / PATCH / DELETE / upload)
    ├── types.ts          # Zdieľané TypeScript typy (ColumnSpec, ViewSpec)
    ├── column-types.ts   # Mapovanie typov stĺpcov + buildFieldBody
    └── tools/
        ├── lists.ts      # Nástroje pre listy a knižnice
        ├── columns.ts    # Nástroje pre stĺpce
        ├── items.ts      # Nástroje pre položky
        ├── views.ts      # Nástroje pre pohľady
        └── files.ts      # Nástroje pre súbory
```

---

## Dostupné nástroje (22)

| Nástroj | Popis |
|---|---|
| `set_site` | Nastaví SharePoint site URL pre aktuálnu reláciu |
| `list_lists` | Zoznam všetkých listov a knižníc na site |
| `get_list_schema` | Stĺpce, pohľady a metadáta listu / knižnice |
| `get_list_schema_xml` | Export kompletného SchemaXml (backup, klonovanie) |
| `create_list` | Vytvorí list alebo knižnicu so stĺpcami a pohľadmi |
| `create_list_from_xml` | Vytvorí list zo SchemaXml reťazca |
| `delete_list` | Vymaže list alebo knižnicu |
| `get_items` | Načíta položky (OData filter, select, orderby, paging) |
| `create_item` | Vytvorí novú položku |
| `update_item` | Aktualizuje položku podľa ID |
| `delete_item` | Vymaže položku podľa ID |
| `add_column` | Pridá stĺpec do existujúceho listu |
| `update_column` | Zmení nastavenia stĺpca (názov, required, default, choices…) |
| `delete_column` | Vymaže stĺpec |
| `get_views` | Zoznam pohľadov listu |
| `create_view` | Vytvorí nový pohľad so stĺpcami a filtrom |
| `update_view` | Aktualizuje pohľad (stĺpce, filter, row limit) |
| `list_files` | Zoznam súborov a priečinkov v knižnici |
| `upload_file` | Nahrá súbor (base64 obsah) |
| `download_file` | Stiahne súbor (vráti base64 obsah) |
| `delete_file` | Vymaže súbor |

---

## Autentifikácia

Nevyžaduje app registration ani admin súhlas. Používa **PnP Management Shell** — well-known public client (`31359c7f-bd7e-475c-86db-fdb8c937548e`).

- Pri prvom spustení sa otvorí prehliadač → prihlásenie pod vlastným účtom → jednorazový consent
- Token je uložený v `~/.sharepoint-mcp/token-cache.json` — ďalšie spustenia nevyžadujú login
- Jeden token pokrýva **všetky site collections** na tom istom tenante
- Prepínanie medzi site collections: zavolaj `set_site` s novou URL

---

## Inštalácia a spustenie

```bash
# 1. Nainštaluj závislosti
npm install

# 2. Vytvor .env zo šablóny
cp .env.example .env
# Uprav SITE_URL v .env

# 3. Spusti vývojový server (voliteľné — Claude Code štartuje server sám)
npm run dev
```

### .env

```env
SITE_URL=https://contoso.sharepoint.com/sites/mysite
```

---

## Konfigurácia MCP v Claude Desktop

Pridaj do `claude_desktop_config.json` (zvyčajne `%APPDATA%\Claude\claude_desktop_config.json` alebo na Windows `%LOCALAPPDATA%\Packages\Claude_...\LocalCache\Roaming\Claude\claude_desktop_config.json`):

### Možnosť 1 — skompilovaný build (odporúčané)

Najprv sprav build projektu:

```bash
npm run build
```

Potom pridaj do konfigurácie (uprav cestu podľa svojho umiestnenia projektu):

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["C:/Users/Lukas/Documents/AI Projects/sharepoint-mcp/dist/index.js"]
    }
  }
}
```

Po každej zmene kódu treba znovu spustiť `npm run build`.

### Možnosť 2 — priamo TypeScript cez tsx (len pre vývoj)

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "npx",
      "args": ["tsx", "C:/Users/Lukas/Documents/AI Projects/sharepoint-mcp/src/index.ts"]
    }
  }
}
```

> **Pozor:** `cwd` v konfigurácii Claude Desktop nefunguje spoľahlivo — vždy používaj absolútnu cestu v `args`.

Po uložení reštartuj Claude Desktop — server sa spustí automaticky pri prvom volaní nástroja.

> **Tip:** Ak pracuješ na viacerých site collections, stačí na začiatku konverzácie zavolať `set_site` s príslušnou URL. Token sa znovu použije, prehliadač sa neotvorí.

---

## Typy stĺpcov

| Typ | Popis | Voliteľné parametre |
|---|---|---|
| `text` | Jednoriadkový text | — |
| `note` | Viacriadkový text | `richText` |
| `number` | Číslo | `min`, `max`, `decimals` |
| `currency` | Mena | `decimals` |
| `date` | Dátum a čas | `dateOnly` |
| `boolean` | Áno / Nie | `defaultValue` |
| `choice` | Výber jednej hodnoty | `choices[]` |
| `multichoice` | Výber viacerých hodnôt | `choices[]` |
| `person` | People picker | `multiple` |
| `url` | Odkaz | — |
| `calculated` | Vypočítaný stĺpec | `formula`, `outputType` |
