import { sp, escapeOData } from "../sharepoint.js";
import { buildFieldBody } from "../column-types.js";
import type { ColumnSpec } from "../types.js";

// ---------------------------------------------------------------------------
// add_column
// ---------------------------------------------------------------------------
export async function addColumn(args: {
  listName: string;
  column: ColumnSpec;
  addToDefaultView?: boolean;
}) {
  const escaped = escapeOData(args.listName);

  await sp.post(
    `web/lists/GetByTitle('${escaped}')/fields`,
    buildFieldBody(args.column)
  );

  // Add to default view unless explicitly disabled
  if (args.addToDefaultView !== false) {
    try {
      const vData = await sp.get<{ value: { Id: string }[] }>(
        `web/lists/GetByTitle('${escaped}')/views`,
        { $filter: "DefaultView eq true", $select: "Id" }
      );
      if (vData.value?.length) {
        const viewId = vData.value[0].Id;
        await sp
          .post(
            `web/lists/GetByTitle('${escaped}')/views('${viewId}')/viewfields/addviewfield('${args.column.name}')`,
            {}
          )
          .catch(() => {});
      }
    } catch {
      // non-fatal
    }
  }

  return {
    success: true,
    message: `Column '${args.column.name}' added to '${args.listName}'`,
  };
}

// ---------------------------------------------------------------------------
// update_column
// ---------------------------------------------------------------------------
export async function updateColumn(args: {
  listName: string;
  internalName: string;
  updates: {
    displayName?: string;
    description?: string;
    required?: boolean;
    defaultValue?: string;
    indexed?: boolean;
    choices?: string[];
  };
}) {
  const escaped = escapeOData(args.listName);
  const fieldPath = `web/lists/GetByTitle('${escaped}')/fields/GetByInternalNameOrTitle('${escapeOData(args.internalName)}')`;

  const body: Record<string, unknown> = {
    __metadata: { type: "SP.Field" },
  };

  if (args.updates.displayName !== undefined) body["Title"] = args.updates.displayName;
  if (args.updates.description !== undefined) body["Description"] = args.updates.description;
  if (args.updates.required !== undefined) body["Required"] = args.updates.required;
  if (args.updates.defaultValue !== undefined) body["DefaultValue"] = args.updates.defaultValue;
  if (args.updates.indexed !== undefined) {
    body["Indexed"] = args.updates.indexed;
    body["EnforceUniqueValues"] = false;
  }
  if (args.updates.choices !== undefined) {
    body["Choices"] = {
      __metadata: { type: "Collection(Edm.String)" },
      results: args.updates.choices,
    };
  }

  await sp.patch(fieldPath, body);

  return { success: true, message: `Column '${args.internalName}' updated` };
}

// ---------------------------------------------------------------------------
// delete_column
// ---------------------------------------------------------------------------
export async function deleteColumn(args: {
  listName: string;
  internalName: string;
}) {
  await sp.delete(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/fields/GetByInternalNameOrTitle('${escapeOData(args.internalName)}')`
  );
  return {
    success: true,
    message: `Column '${args.internalName}' deleted from '${args.listName}'`,
  };
}
