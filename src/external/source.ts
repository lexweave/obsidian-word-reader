import { App, Notice, TFile } from "obsidian";
import { getElectronModule, getFileSystemPromises, getPathModule, getUrlModule } from "../desktopRuntime";
import type { PptxSource } from "../pptx/pptxSource";
import { buildPresentationSummaryNote } from "../pptx/pptxSummaryNote";
import type { ExternalOfficeAccess } from "./access";
import type { BinaryOfficeSource, OfficeSourceFile } from "../reader/officeSource";
import { externalOfficeFormat } from "./links";
import { readOfficeDocument, revision, validateStat } from "./fileAccess";

export interface ExternalOfficeSource extends BinaryOfficeSource {
  changed(): Promise<boolean>;
  createSummaryMarkdown(content: string, createdMessage: string): Promise<void>;
}

export async function externalOfficeSource(app: App, uri: string, access: ExternalOfficeAccess): Promise<ExternalOfficeSource> {
  const io = getFileSystemPromises();
  const path = await access.require(uri, false);
  const format = externalOfficeFormat(uri);
  if (!path || !format || getPathModule().extname(path).toLowerCase() !== `.${format}`) throw new Error("External document access is not allowed.");
  const canonicalUri = getUrlModule().pathToFileURL(path).href;
  let expected = "";
  const file: OfficeSourceFile = {
    path: canonicalUri, name: getPathModule().basename(path),
    basename: getPathModule().basename(path, getPathModule().extname(path)),
    extension: format, stat: { mtime: 0, ctime: 0, size: 0 },
  };
  const checkAccess = async () => {
    if (await io.realpath(path) !== path || !await access.allowed(path, format)) throw new Error("External document access changed.");
  };
  const refresh = async () => {
    await checkAccess();
    const stat = await io.stat(path);
    validateStat(stat); expected = revision(stat);
    file.stat = { mtime: stat.mtimeMs, ctime: stat.ctimeMs, size: stat.size };
  };
  await refresh();
  return {
    file, refresh,
    read: async () => { await checkAccess(); return readOfficeDocument(io, path, expected, io.constants.O_RDONLY | (io.constants.O_NOFOLLOW ?? 0)); },
    changed: async () => { await checkAccess(); return revision(await io.stat(path)) !== expected; },
    openExternal: async () => {
      await checkAccess();
      if (await getElectronModule().shell.openPath(path)) throw new Error("Could not open the document.");
    },
    createSummaryMarkdown: async (markdown, createdMessage) => {
      // Only a deliberate summary action writes a Markdown note, always inside the vault.
      const id = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(canonicalUri));
      const suffix = [...new Uint8Array(id)].map((n) => n.toString(16).padStart(2, "0")).join("").slice(0, 16);
      const folder = "Office Reader Notes";
      if (!app.vault.getAbstractFileByPath(folder)) await app.vault.createFolder(folder);
      const basename = file.basename.replace(/[\\/:*?"<>|#^[\]]/g, "_").slice(0, 50);
      const notePath = `${folder}/${basename}-${suffix}.md`;
      const existing = app.vault.getAbstractFileByPath(notePath);
      if (existing instanceof TFile) { await app.workspace.getLeaf(true).openFile(existing); return; }
      const content = markdown.replace(`[[${canonicalUri}]]`, `[${file.basename.replace(/[[\]]/g, "")}](<${canonicalUri}>)`);
      const note = await app.vault.create(notePath, content);
      await app.workspace.getLeaf(true).openFile(note);
      new Notice(createdMessage);
    },
  };
}

export async function externalPptxSource(app: App, uri: string, access: ExternalOfficeAccess): Promise<PptxSource & { changed(): Promise<boolean> }> {
  if (externalOfficeFormat(uri) !== "pptx") throw new Error("Expected a PPTX source.");
  const source = await externalOfficeSource(app, uri, access);
  return { ...source, createSummary: (slides, page, text) => source.createSummaryMarkdown(
    buildPresentationSummaryNote(source.file, slides, page, text), text.notices.createdSummaryNote,
  ) };
}
