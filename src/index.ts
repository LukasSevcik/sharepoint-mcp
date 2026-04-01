import "dotenv/config";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";

import { sp } from "./sharepoint.js";
import * as lists from "./tools/lists.js";
import * as columns from "./tools/columns.js";
import * as items from "./tools/items.js";
import * as views from "./tools/views.js";
import * as files from "./tools/files.js";

// ---------------------------------------------------------------------------
// Bootstrap: pre-set site URL from .env if provided
// ---------------------------------------------------------------------------
if (process.env["SITE_URL"]) {
  sp.setSiteUrl(process.env["SITE_URL"]);
}

// ---------------------------------------------------------------------------
// JSON Schema fragment reused in multiple tools
// ---------------------------------------------------------------------------
const COLUMN_SPEC_SCHEMA = {
  type: "object",
  properties: {
    name: {
      type: "string",
      description:
        "Internal name (no spaces, use PascalCase). SharePoint will use this as the field's StaticName.",
    },
    displayName: {
      type: "string",
      description: "Display name shown in the SharePoint UI (optional, defaults to name).",
    },
    type: {
      type: "string",
      enum: [
        "text", "note", "number", "currency", "date",
        "boolean", "choice", "multichoice", "person", "url", "calculated",
      ],
      description:
        "Column type. text=single line, note=multi-line, date=DateTime, " +
        "choice=single choice, multichoice=multi-select, person=People picker.",
    },
    required: { type: "boolean", description: "Whether the column is required." },
    description: { type: "string", description: "Column description / help text." },
    defaultValue: { description: "Default value (string, number, or boolean)." },
    indexed: { type: "boolean", description: "Index this column for faster queries." },
    choices: {
      type: "array",
      items: { type: "string" },
      description: "Allowed choices (for choice / multichoice columns).",
    },
    richText: { type: "boolean", description: "Enable rich text (note columns)." },
    dateOnly: { type: "boolean", description: "Date only without time (date columns)." },
    multiple: { type: "boolean", description: "Allow multiple people (person columns)." },
    min: { type: "number", description: "Minimum value (number columns)." },
    max: { type: "number", description: "Maximum value (number columns)." },
    decimals: { type: "number", description: "Decimal places (number / currency columns)." },
    formula: { type: "string", description: "Formula string (calculated columns), e.g. '=[Price]*1.2'." },
    outputType: {
      type: "string",
      enum: ["text", "number", "boolean", "date", "currency"],
      description: "Output type for calculated columns.",
    },
  },
  required: ["name", "type"],
} as const;

// ---------------------------------------------------------------------------
// Tool definitions
// ---------------------------------------------------------------------------
const TOOLS: Tool[] = [
  // --- Site ---
  {
    name: "set_site",
    description:
      "Set the SharePoint site URL to work with. Must be called before any other tool if SITE_URL is not set in .env.",
    inputSchema: {
      type: "object",
      properties: {
        siteUrl: {
          type: "string",
          description: "Full SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/mysite",
        },
      },
      required: ["siteUrl"],
    },
  },

  // --- Lists & Libraries ---
  {
    name: "list_lists",
    description:
      "List all lists and document libraries in the SharePoint site. Returns name, id, type (list / documentLibrary), item count, etc.",
    inputSchema: {
      type: "object",
      properties: {
        includeHidden: {
          type: "boolean",
          description: "Include hidden / system lists (default: false).",
        },
      },
    },
  },
  {
    name: "get_list_schema",
    description:
      "Get the full schema of a list or library: all columns (type, required, default, indexed), all views, and list metadata. Works on both lists (baseTemplate 100) and document libraries (101).",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string", description: "Display name of the list / library." },
      },
      required: ["listName"],
    },
  },
  {
    name: "get_list_schema_xml",
    description:
      "Export the full SharePoint SchemaXml of a list or library. The XML contains every field, view, content type, and JSON formatting definition. Useful for backup or cloning.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
      },
      required: ["listName"],
    },
  },
  {
    name: "create_list",
    description:
      "Create a new SharePoint list or document library with optional columns and views. " +
      "Columns are described using structured specs (type, choices, required, indexed, etc.). " +
      "New columns are automatically added to the default view.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "List / library display name." },
        type: {
          type: "string",
          enum: ["list", "library"],
          description: "list = generic list (baseTemplate 100), library = document library (101). Default: list.",
        },
        description: { type: "string" },
        columns: {
          type: "array",
          items: COLUMN_SPEC_SCHEMA,
          description: "Columns to create (leave empty for a title-only list).",
        },
        views: {
          type: "array",
          description: "Additional views to create (beyond the default All Items view).",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              isDefault: { type: "boolean" },
              columns: { type: "array", items: { type: "string" }, description: "Internal column names." },
              rowLimit: { type: "number" },
              query: { type: "string", description: "CAML query for filtering." },
            },
            required: ["name", "columns"],
          },
        },
      },
      required: ["name"],
    },
  },
  {
    name: "create_list_from_xml",
    description:
      "Create a list from a full SharePoint SchemaXml string (e.g. obtained from get_list_schema_xml). Useful for cloning or restoring a list.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name for the new list." },
        schemaXml: { type: "string", description: "Full SchemaXml string." },
        description: { type: "string" },
      },
      required: ["name", "schemaXml"],
    },
  },
  {
    name: "delete_list",
    description: "Permanently delete a list or document library.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
      },
      required: ["listName"],
    },
  },

  // --- Items ---
  {
    name: "get_items",
    description:
      "Query items in a list. Supports OData $filter, $select, $orderby, $top, $skip, $expand.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        filter: { type: "string", description: "OData filter, e.g. \"Status eq 'Active' and DueDate lt '2025-01-01'\"" },
        select: { type: "array", items: { type: "string" }, description: "Field internal names to return." },
        orderBy: { type: "string", description: "e.g. 'Title asc' or 'Modified desc'" },
        top: { type: "number", description: "Max items to return (default: 100)." },
        skip: { type: "number", description: "Skip N items (for paging)." },
        expand: { type: "array", items: { type: "string" }, description: "Navigation properties to expand, e.g. ['Author', 'AssignedTo']" },
      },
      required: ["listName"],
    },
  },
  {
    name: "create_item",
    description:
      "Create a new item in a list. Pass field values as key-value pairs using internal column names. " +
      "For person fields use the format {FieldNameId: userId}. " +
      "For date fields use ISO 8601 strings.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        fields: {
          type: "object",
          description: "Key-value map of internal field name → value.",
          additionalProperties: true,
        },
      },
      required: ["listName", "fields"],
    },
  },
  {
    name: "update_item",
    description: "Update fields of an existing list item by its numeric ID.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        id: { type: "number", description: "Item ID." },
        fields: {
          type: "object",
          description: "Fields to update (only include fields that need to change).",
          additionalProperties: true,
        },
      },
      required: ["listName", "id", "fields"],
    },
  },
  {
    name: "delete_item",
    description: "Delete a list item by its numeric ID.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        id: { type: "number" },
      },
      required: ["listName", "id"],
    },
  },

  // --- Columns ---
  {
    name: "add_column",
    description:
      "Add a new column to an existing list or library. The column is automatically added to the default view unless addToDefaultView is false.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        column: COLUMN_SPEC_SCHEMA,
        addToDefaultView: { type: "boolean", description: "Default: true." },
      },
      required: ["listName", "column"],
    },
  },
  {
    name: "update_column",
    description: "Update settings of an existing column (display name, description, required, default value, indexed, choices).",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        internalName: { type: "string", description: "Internal name of the column." },
        updates: {
          type: "object",
          properties: {
            displayName: { type: "string" },
            description: { type: "string" },
            required: { type: "boolean" },
            defaultValue: { type: "string" },
            indexed: { type: "boolean" },
            choices: { type: "array", items: { type: "string" } },
          },
        },
      },
      required: ["listName", "internalName", "updates"],
    },
  },
  {
    name: "delete_column",
    description: "Delete a column from a list (by internal name). Cannot delete read-only or sealed columns.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        internalName: { type: "string" },
      },
      required: ["listName", "internalName"],
    },
  },

  // --- Views ---
  {
    name: "get_views",
    description: "List all views of a list or library with their columns, filters, and row limits.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
      },
      required: ["listName"],
    },
  },
  {
    name: "create_view",
    description: "Create a new view for a list or library with specified columns, row limit, and optional CAML filter.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        name: { type: "string", description: "View display name." },
        columns: {
          type: "array",
          items: { type: "string" },
          description: "Ordered list of internal column names to show in the view.",
        },
        isDefault: { type: "boolean", description: "Make this the default view. Default: false." },
        rowLimit: { type: "number", description: "Items per page. Default: 30." },
        query: { type: "string", description: "CAML query string for filtering / ordering." },
      },
      required: ["listName", "name", "columns"],
    },
  },
  {
    name: "update_view",
    description: "Update an existing view: change columns, CAML query, row limit, or make it the default.",
    inputSchema: {
      type: "object",
      properties: {
        listName: { type: "string" },
        viewName: { type: "string" },
        columns: {
          type: "array",
          items: { type: "string" },
          description: "Replace view columns with this list (optional).",
        },
        query: { type: "string" },
        rowLimit: { type: "number" },
        isDefault: { type: "boolean" },
      },
      required: ["listName", "viewName"],
    },
  },

  // --- Files ---
  {
    name: "list_files",
    description:
      "List files and sub-folders inside a document library or a specific folder within it.",
    inputSchema: {
      type: "object",
      properties: {
        libraryName: { type: "string", description: "Display name of the document library." },
        folderPath: {
          type: "string",
          description: "Relative path within the library, e.g. 'Reports/2025'. Omit for the root.",
        },
      },
      required: ["libraryName"],
    },
  },
  {
    name: "upload_file",
    description: "Upload a file to a document library. Content must be base64-encoded.",
    inputSchema: {
      type: "object",
      properties: {
        libraryName: { type: "string" },
        folderPath: { type: "string", description: "Target sub-folder (optional)." },
        fileName: { type: "string", description: "File name including extension." },
        content: { type: "string", description: "Base64-encoded file content." },
        overwrite: { type: "boolean", description: "Overwrite if file exists. Default: true." },
      },
      required: ["libraryName", "fileName", "content"],
    },
  },
  {
    name: "download_file",
    description: "Download a file from SharePoint. Returns base64-encoded content.",
    inputSchema: {
      type: "object",
      properties: {
        serverRelativeUrl: {
          type: "string",
          description: "Server-relative URL of the file, e.g. /sites/mysite/Documents/report.pdf",
        },
      },
      required: ["serverRelativeUrl"],
    },
  },
  {
    name: "delete_file",
    description: "Permanently delete a file from a document library.",
    inputSchema: {
      type: "object",
      properties: {
        serverRelativeUrl: {
          type: "string",
          description: "Server-relative URL of the file to delete.",
        },
      },
      required: ["serverRelativeUrl"],
    },
  },
];

// ---------------------------------------------------------------------------
// Tool dispatcher
// ---------------------------------------------------------------------------
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function dispatch(name: string, args: any): Promise<unknown> {
  switch (name) {
    // site
    case "set_site":
      sp.setSiteUrl(args.siteUrl as string);
      return { success: true, message: `Site set to ${args.siteUrl}` };

    // lists
    case "list_lists":         return lists.listLists(args);
    case "get_list_schema":    return lists.getListSchema(args);
    case "get_list_schema_xml":return lists.getListSchemaXml(args);
    case "create_list":        return lists.createList(args);
    case "create_list_from_xml":return lists.createListFromXml(args);
    case "delete_list":        return lists.deleteList(args);

    // items
    case "get_items":   return items.getItems(args);
    case "create_item": return items.createItem(args);
    case "update_item": return items.updateItem(args);
    case "delete_item": return items.deleteItem(args);

    // columns
    case "add_column":    return columns.addColumn(args);
    case "update_column": return columns.updateColumn(args);
    case "delete_column": return columns.deleteColumn(args);

    // views
    case "get_views":    return views.getViews(args);
    case "create_view":  return views.createView(args);
    case "update_view":  return views.updateView(args);

    // files
    case "list_files":     return files.listFiles(args);
    case "upload_file":    return files.uploadFile(args);
    case "download_file":  return files.downloadFile(args);
    case "delete_file":    return files.deleteFile(args);

    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// ---------------------------------------------------------------------------
// MCP Server
// ---------------------------------------------------------------------------
const server = new Server(
  { name: "sharepoint-mcp", version: "0.1.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  try {
    const result = await dispatch(name, args ?? {});
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: "text", text: `Error: ${msg}` }],
      isError: true,
    };
  }
});

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------
const transport = new StdioServerTransport();
await server.connect(transport);
