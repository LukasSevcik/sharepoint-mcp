import { sp, escapeOData } from "../sharepoint.js";

// ---------------------------------------------------------------------------
// get_items
// ---------------------------------------------------------------------------
export async function getItems(args: {
  listName: string;
  filter?: string;
  select?: string[];
  orderBy?: string;
  top?: number;
  skip?: number;
  expand?: string[];
}) {
  const qs: Record<string, string> = {};
  if (args.filter) qs["$filter"] = args.filter;
  if (args.select?.length) qs["$select"] = args.select.join(",");
  if (args.orderBy) qs["$orderby"] = args.orderBy;
  if (args.top !== undefined) qs["$top"] = String(args.top);
  if (args.skip !== undefined) qs["$skip"] = String(args.skip);
  if (args.expand?.length) qs["$expand"] = args.expand.join(",");

  const data = await sp.get<{ value: unknown[] }>(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/items`,
    qs
  );

  return {
    items: data.value ?? [],
    count: data.value?.length ?? 0,
  };
}

// ---------------------------------------------------------------------------
// create_item
// ---------------------------------------------------------------------------
export async function createItem(args: {
  listName: string;
  fields: Record<string, unknown>;
}) {
  const entityType = await sp.getEntityTypeName(args.listName);

  const data = await sp.post<Record<string, unknown>>(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/items`,
    { __metadata: { type: entityType }, ...args.fields }
  );

  return {
    success: true,
    id: data?.["Id"],
    item: data,
  };
}

// ---------------------------------------------------------------------------
// update_item
// ---------------------------------------------------------------------------
export async function updateItem(args: {
  listName: string;
  id: number;
  fields: Record<string, unknown>;
}) {
  const entityType = await sp.getEntityTypeName(args.listName);

  await sp.patch(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/items(${args.id})`,
    { __metadata: { type: entityType }, ...args.fields }
  );

  return { success: true, message: `Item ${args.id} updated` };
}

// ---------------------------------------------------------------------------
// delete_item
// ---------------------------------------------------------------------------
export async function deleteItem(args: {
  listName: string;
  id: number;
}) {
  await sp.delete(
    `web/lists/GetByTitle('${escapeOData(args.listName)}')/items(${args.id})`
  );
  return { success: true, message: `Item ${args.id} deleted` };
}
