import { sp, escapeOData } from "../sharepoint.js";
import { buildFieldBody, fieldTypeKindToLabel } from "../column-types.js";
import type { ColumnSpec, ViewSpec } from "../types.js";

// ---------------------------------------------------------------------------
// list_lists
// ---------------------------------------------------------------------------
export async function listLists(args: { includeHidden?: boolean }) {
  const data = await sp.get<{ value: Record<string, unknown>[] }>(
    "web/lists",
    {
      $select:
        "Title,Id,BaseTemplate,Description,ItemCount,Hidden,IsPrivate,LastItemModifiedDate,Created",
      $orderby: "Title",
    }
  );

  return (data.value ?? [])
    .filter(
      (l) => args.includeHidden || (!l["Hidden"] && !l["IsPrivate"])
    )
    .map((l) => ({
      title: l["Title"],
      id: l["Id"],
      baseTemplate: l["BaseTemplate"],
      type:
        l["BaseTemplate"] === 101
          ? "documentLibrary"
          : l["BaseTemplate"] === 100
          ? "list"
          : `template${l["BaseTemplate"]}`,
      description: l["Description"],
      itemCount: l["ItemCount"],
      created: l["Created"],
      lastModified: l["LastItemModifiedDate"],
    }));
}

// ---------------------------------------------------------------------------
// get_list_schema
// ---------------------------------------------------------------------------
export async function getListSchema(args: { listName: string }) {
  const base = `web/lists/GetByTitle('${escapeOData(args.listName)}')`;

  const [listData, fieldsData, viewsData] = await Promise.all([
    sp.get<Record<string, unknown>>(base, {
      $select:
        "Title,Id,BaseTemplate,Description,ItemCount,Created,LastItemModifiedDate",
    }),
    sp.get<{ value: Record<string, unknown>[] }>(`${base}/fields`, {
      $select:
        "Title,InternalName,FieldTypeKind,TypeAsString,Required,DefaultValue," +
        "Description,EnforceUniqueValues,Indexed,Hidden,ReadOnlyField,Choices,AllowMultipleValues",
      $filter:
        "Hidden eq false and ReadOnlyField eq false and FieldTypeKind ne 0",
      $orderby: "InternalName",
    }),
    sp.get<{ value: Record<string, unknown>[] }>(`${base}/views`, {
      $select: "Title,Id,DefaultView,RowLimit,ViewQuery",
      $expand: "ViewFields",
    }),
  ]);

  return {
    list: {
      title: listData["Title"],
      id: listData["Id"],
      baseTemplate: listData["BaseTemplate"],
      type: listData["BaseTemplate"] === 101 ? "documentLibrary" : "list",
      description: listData["Description"],
      itemCount: listData["ItemCount"],
      created: listData["Created"],
      lastModified: listData["LastItemModifiedDate"],
    },
    columns: (fieldsData.value ?? []).map((f) => ({
      internalName: f["InternalName"],
      displayName: f["Title"],
      type: fieldTypeKindToLabel(f["FieldTypeKind"] as number),
      fieldTypeKind: f["FieldTypeKind"],
      required: f["Required"],
      defaultValue: f["DefaultValue"],
      description: f["Description"],
      indexed: f["Indexed"] || f["EnforceUniqueValues"],
      choices: (f["Choices"] as { results?: string[] } | null)?.results,
      allowMultiple: f["AllowMultipleValues"],
    })),
    views: (viewsData.value ?? []).map((v) => {
      const vf = v["ViewFields"] as
        | { Items?: { results?: string[] }; results?: string[] }
        | null;
      return {
        id: v["Id"],
        title: v["Title"],
        isDefault: v["DefaultView"],
        rowLimit: v["RowLimit"],
        query: v["ViewQuery"],
        fields: vf?.Items?.results ?? vf?.results ?? [],
      };
    }),
  };
}

// ---------------------------------------------------------------------------
// get_list_schema_xml
// ---------------------------------------------------------------------------
export async function getListSchemaXml(args: { listName: string }) {
  const data = await sp.get<{ value?: string }>(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/SchemaXml`
  );
  return { schemaXml: data?.value ?? data };
}

// ---------------------------------------------------------------------------
// create_list
// ---------------------------------------------------------------------------
export async function createList(args: {
  name: string;
  type?: "list" | "library";
  description?: string;
  columns?: ColumnSpec[];
  views?: ViewSpec[];
}) {
  const baseTemplate = args.type === "library" ? 101 : 100;
  const escaped = escapeOData(args.name);

  // 1. Create list
  await sp.post("web/lists", {
    __metadata: { type: "SP.List" },
    AllowContentTypes: true,
    BaseTemplate: baseTemplate,
    ContentTypesEnabled: false,
    Description: args.description ?? "",
    Title: args.name,
  });

  // 2. Add columns
  if (args.columns?.length) {
    for (const col of args.columns) {
      await sp.post(
        `web/lists/GetByTitle('${escaped}')/fields`,
        buildFieldBody(col)
      );
    }

    // Add new columns to the default view
    try {
      const vData = await sp.get<{ value: { Id: string }[] }>(
        `web/lists/GetByTitle('${escaped}')/views`,
        { $filter: "DefaultView eq true", $select: "Id" }
      );
      if (vData.value?.length) {
        const viewId = vData.value[0].Id;
        for (const col of args.columns) {
          await sp
            .post(
              `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/addviewfield('${col.name}')`,
              {}
            )
            .catch(() => {});
        }
      }
    } catch {
      // non-fatal
    }
  }

  // 3. Create custom views
  if (args.views?.length) {
    for (const view of args.views) {
      const vResult = await sp.post<{ Id?: string }>(
        `web/lists/GetByTitle('${escaped}')/views`,
        {
          __metadata: { type: "SP.View" },
          Title: view.name,
          DefaultView: view.isDefault ?? false,
          RowLimit: view.rowLimit ?? 30,
          ViewQuery: view.query ?? "",
        }
      );
      const viewId = vResult?.Id;
      if (viewId) {
        for (const col of view.columns) {
          await sp
            .post(
              `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/addviewfield('${col}')`,
              {}
            )
            .catch(() => {});
        }
      }
    }
  }

  return {
    success: true,
    message: `${args.type === "library" ? "Document library" : "List"} '${args.name}' created successfully`,
    columnsCreated: args.columns?.length ?? 0,
    viewsCreated: args.views?.length ?? 0,
  };
}

// ---------------------------------------------------------------------------
// create_list_from_xml
// ---------------------------------------------------------------------------
export async function createListFromXml(args: {
  name: string;
  schemaXml: string;
  description?: string;
}) {
  const result = await sp.post<{ Id?: string }>("web/lists", {
    __metadata: { type: "SP.List" },
    Title: args.name,
    Description: args.description ?? "",
    CustomSchemaXml: args.schemaXml,
  });

  return {
    success: true,
    message: `List '${args.name}' created from XML schema`,
    id: result?.Id,
  };
}

// ---------------------------------------------------------------------------
// delete_list
// ---------------------------------------------------------------------------
export async function deleteList(args: { listName: string }) {
  await sp.delete(`web/lists/GetByTitle('${escapeOData(args.listName)}')`);
  return { success: true, message: `'${args.listName}' deleted` };
}
