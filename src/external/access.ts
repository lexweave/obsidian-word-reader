import { App, Modal, Setting } from "obsidian";
import { getFileSystemPromises, getPathModule, getUrlModule } from "../desktopRuntime";
import { isWithinRoot } from "./fileAccess";

import { externalOfficeFormat, parseExternalOfficeLink, type ExternalOfficeFormat } from "./links";

const key = (format: ExternalOfficeFormat) => `word-reader-external-${format}-folders`;

export class ExternalOfficeAccess {
  constructor(private readonly app: App) {}

  private roots(format: ExternalOfficeFormat): string[] {
    const value: unknown = this.app.loadLocalStorage(key(format));
    return Array.isArray(value) ? value.filter((p): p is string => typeof p === "string") : [];
  }
  clear(): void { for (const format of ["pptx", "docx", "xlsx"] as const) this.app.saveLocalStorage(key(format), null); }

  async resolve(uri: string): Promise<string> {
    const io = getFileSystemPromises();
    return io.realpath(getUrlModule().fileURLToPath(uri));
  }
  async allowed(path: string, format: ExternalOfficeFormat = "pptx"): Promise<boolean> {
    const io = getFileSystemPromises();
    const paths = getPathModule();
    for (const root of this.roots(format)) {
      try {
        const realRoot = await io.realpath(root);
        if (isWithinRoot(paths.relative(realRoot, path), paths.sep)) return true;
      } catch { /* An unavailable root grants no access. */ }
    }
    return false;
  }
  async require(uri: string, prompt: boolean): Promise<string | null> {
    const format = externalOfficeFormat(uri);
    if (!format || !parseExternalOfficeLink(uri)) throw new Error("Unsupported external Office link.");
    const path = await this.resolve(uri);
    if (getPathModule().extname(path).toLowerCase() !== `.${format}`) throw new Error("Source format changed.");
    if (await this.allowed(path, format)) return path;
    if (!prompt) return null;
    const root = await new FolderAccessModal(this.app, getPathModule().dirname(path), format).ask();
    if (!root) return null;
    const realRoot = await getFileSystemPromises().realpath(root);
    if (!(await getFileSystemPromises().stat(realRoot)).isDirectory() ||
        !isWithinRoot(getPathModule().relative(realRoot, path), getPathModule().sep)) {
      throw new Error("Choose a folder containing this document.");
    }
    this.app.saveLocalStorage(key(format), [...new Set([...this.roots(format), realRoot])]);
    return path;
  }
}

class FolderAccessModal extends Modal {
  private result: ((root: string | null) => void) | null = null;
  private root: string;
  constructor(app: App, root: string, private readonly format: ExternalOfficeFormat) { super(app); this.root = root; }
  ask(): Promise<string | null> {
    return new Promise((resolve) => { this.result = resolve; this.open(); });
  }
  onOpen(): void {
    this.setTitle(`Allow external ${this.format.toUpperCase()} files`);
    this.contentEl.createEl("p", { text: `Allow read-only access to .${this.format} files in this folder on this device. Files stay in place. You may choose a containing parent folder. Other formats require their own permission.` });
    new Setting(this.contentEl).setName("Folder / 文件夹").addText((field) => field.setValue(this.root).onChange((value) => { this.root = value; }));
    new Setting(this.contentEl)
      .addButton((button) => button.setButtonText("Cancel / 取消").onClick(() => this.close()))
      .addButton((button) => button.setButtonText("Allow / 允许").setCta().onClick(() => {
        const resolve = this.result; this.result = null; resolve?.(this.root); this.close();
      }));
  }
  onClose(): void { this.result?.(null); this.result = null; this.contentEl.empty(); }
}
