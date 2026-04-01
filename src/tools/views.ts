import { sp, escapeOData } from "../sharepoint.js";

// ---------------------------------------------------------------------------
// get_views
// ---------------------------------------------------------------------------
export async function getViews(args: { listName: string }) {
  const data = await sp.get<{ value: Record<string, unknown>[] }>(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/views`,
    {
      $select: "Title,Id,DefaultView,RowLimit,ViewQuery",
      $expand: "ViewFields",
    }
  );

  return (data.value ?? []).map((v) => {
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
  });
}

// ---------------------------------------------------------------------------
// create_view
// ---------------------------------------------------------------------------
export async function createView(args: {
  listName: string;
  name: string;
  columns: string[];
  isDefault?: boolean;
  rowLimit?: number;
  query?: string;
}) {
  const escaped = escapeOData(args.listName);

  // Create view (without columns — add them separately)
  const vResult = await sp.post<{ Id?: string }>(
    `web/lists/GetByTitle('${escaped}')/views`,
    {
      __metadata: { type: "SP.View" },
      Title: args.name,
      DefaultView: args.isDefault ?? false,
      RowLimit: args.rowLimit ?? 30,
      ViewQuery: args.query ?? "",
    }
  );

  const viewId = vResult?.Id;

  if (viewId && args.columns.length) {
    // Remove the default fields that SharePoint adds automatically
    await sp
      .post(
        `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/removeallviewfields`,
        {}
      )
      .catch(() => {});

    for (const col of args.columns) {
      await sp
        .post(
          `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/addviewfield('${col}')`,
          {}
        )
        .catch(() => {});
    }
  }

  return {
    success: true,
    id: viewId,
    message: `View '${args.name}' created in '${args.listName}'`,
  };
}

// ---------------------------------------------------------------------------
// update_view
// ---------------------------------------------------------------------------
export async function updateView(args: {
  listName: string;
  viewName: string;
  columns?: string[];
  query?: string;
  rowLimit?: number;
  isDefault?: boolean;
}) {
  const escaped = escapeOData(args.listName);

  // Resolve view ID
  const vData = await sp.get<{ value: { Id: string }[] }>(
    `web/lists/GetByTitle('${escaped}')/views`,
    {
      $filter: `Title eq '${escapeOData(args.viewName)}'`,
      $select: "Id",
    }
  );

  if (!vData.value?.length) {
    throw new Error(`View '${args.viewName}' not found in '${args.listName}'`);
  }
  const viewId = vData.value[0].Id;

  // Update metadata
  const patchBody: Record<string, unknown> = {
    __metadata: { type: "SP.View" },
  };
  if (args.query !== undefined) patchBody["ViewQuery"] = args.query;
  if (args.rowLimit !== undefined) patchBody["RowLimit"] = args.rowLimit;
  if (args.isDefault !== undefined) patchBody["DefaultView"] = args.isDefault;

  if (Object.keys(patchBody).length > 1) {
    await sp.patch(
      `web/lists/GetByTitle('${escaped}')/views('${viewId}')`,
      patchBody
    );
  }

  // Replace columns if provided
  if (args.columns) {
    await sp
      .post(
        `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/removeallviewfields`,
        {}
      )
      .catch(() => {});

    for (const col of args.columns) {
      await sp
        .post(
          `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/addviewfield('${col}')`,
          {}
        )
        .catch(() => {});
    }
  }

  return { success: true, message: `View '${args.viewName}' updated` };
}
