import { sp, escapeOData } from "../sharepoint.js";
import { getAccessToken } from "../auth.js";

// ---------------------------------------------------------------------------
// list_files
// ---------------------------------------------------------------------------
export async function listFiles(args: {
  libraryName: string;
  folderPath?: string;
}) {
  const rootUrl = await sp.getLibraryRootUrl(args.libraryName);
  const folderUrl = args.folderPath
    ? `${rootUrl}/${args.folderPath.replace(/^\//, "")}`
    : rootUrl;

  const enc = encodeURIComponent(folderUrl);

  const [filesData, foldersData] = await Promise.all([
    sp.get<{ value: Record<string, unknown>[] }>(
      `web/GetFolderByServerRelativeUrl('${enc}')/Files`,
      {
        $select: "Name,ServerRelativeUrl,Length,TimeLastModified",
        $orderby: "Name",
      }
    ),
    sp.get<{ value: Record<string, unknown>[] }>(
      `web/GetFolderByServerRelativeUrl('${enc}')/Folders`,
      {
        $select: "Name,ServerRelativeUrl,ItemCount",
        $orderby: "Name",
      }
    ),
  ]);

  return {
    path: folderUrl,
    files: (filesData.value ?? []).map((f) => ({
      name: f["Name"],
      serverRelativeUrl: f["ServerRelativeUrl"],
      sizeBytes: f["Length"],
      modified: f["TimeLastModified"],
    })),
    folders: (foldersData.value ?? [])
      .filter((f) => f["Name"] !== "Forms") // hide the internal Forms folder
      .map((f) => ({
        name: f["Name"],
        serverRelativeUrl: f["ServerRelativeUrl"],
        itemCount: f["ItemCount"],
      })),
  };
}

// ---------------------------------------------------------------------------
// upload_file
// ---------------------------------------------------------------------------
export async function uploadFile(args: {
  libraryName: string;
  folderPath?: string;
  fileName: string;
  content: string; // base64
  overwrite?: boolean;
}) {
  const rootUrl = await sp.getLibraryRootUrl(args.libraryName);
  const folderUrl = args.folderPath
    ? `${rootUrl}/${args.folderPath.replace(/^\//, "")}`
    : rootUrl;

  const overwrite = args.overwrite !== false;
  const enc = encodeURIComponent(folderUrl);
  const name = encodeURIComponent(args.fileName);

  const buffer = Buffer.from(args.content, "base64");

  const data = await sp.uploadBinary<Record<string, unknown>>(
    `web/GetFolderByServerRelativeUrl('${enc}')/Files/Add(overwrite=${overwrite},url='${name}')`,
    buffer
  );

  return {
    success: true,
    message: `File '${args.fileName}' uploaded to '${folderUrl}'`,
    serverRelativeUrl: data?.["ServerRelativeUrl"],
  };
}

// ---------------------------------------------------------------------------
// download_file
// ---------------------------------------------------------------------------
export async function downloadFile(args: { serverRelativeUrl: string }) {
  const siteUrl = sp.getSiteUrl();
  if (!siteUrl) throw new Error("Site URL not set. Call set_site first.");

  const token = await getAccessToken(siteUrl);
  const enc = encodeURIComponent(args.serverRelativeUrl);

  const res = await fetch(
    `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${enc}')/$value`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!res.ok) {
    throw new Error(`Download failed (${res.status}): ${await res.text()}`);
  }

  const buffer = Buffer.from(await res.arrayBuffer());
  const fileName = args.serverRelativeUrl.split("/").pop() ?? "file";

  return {
    fileName,
    content: buffer.toString("base64"),
    encoding: "base64",
    sizeBytes: buffer.length,
  };
}

// ---------------------------------------------------------------------------
// delete_file
// ---------------------------------------------------------------------------
export async function deleteFile(args: { serverRelativeUrl: string }) {
  const enc = escapeOData(args.serverRelativeUrl);
  await sp.delete(`web/GetFileByServerRelativeUrl('${enc}')`);
  return {
    success: true,
    message: `File deleted: ${args.serverRelativeUrl}`,
  };
}
