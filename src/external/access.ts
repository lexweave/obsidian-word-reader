import { App, Modal, Setting } from "obsidian";
import { getFileSystemPromises, getPathModule, getUrlModule } from "../desktopRuntime";
import { isWithinRoot } from "./fileAccess";

const KEY = "word-reader-external-pptx-folders";

export class ExternalPptxAccess {
  constructor(private readonly app: App) {}

  private roots(): string[] {
    const value: unknown = this.app.loadLocalStorage(KEY);
    return Array.isArray(value) ? value.filter((p): p is string => typeof p === "string") : [];
  }
  clear(): void { this.app.saveLocalStorage(KEY, null); }

  async resolve(uri: string): Promise<string> {
    const io = getFileSystemPromises();
    return io.realpath(getUrlModule().fileURLToPath(uri));
  }
  async allowed(path: string): Promise<boolean> {
    const io = getFileSystemPromises();
    const paths = getPathModule();
    for (const root of this.roots()) {
      try {
        const realRoot = await io.realpath(root);
        if (isWithinRoot(paths.relative(realRoot, path), paths.sep)) return true;
      } catch { /* An unavailable root grants no access. */ }
    }
    return false;
  }
  async require(uri: string, prompt: boolean): Promise<string | null> {
    const path = await this.resolve(uri);
    if (await this.allowed(path)) return path;
    if (!prompt) return null;
    const root = await new FolderAccessModal(this.app, getPathModule().dirname(path)).ask();
    if (!root) return null;
    const realRoot = await getFileSystemPromises().realpath(root);
    if (!(await getFileSystemPromises().stat(realRoot)).isDirectory() ||
        !isWithinRoot(getPathModule().relative(realRoot, path), getPathModule().sep)) {
      throw new Error("Choose a folder containing this presentation.");
    }
    this.app.saveLocalStorage(KEY, [...new Set([...this.roots(), realRoot])]);
    return path;
  }
}

class FolderAccessModal extends Modal {
  private result: ((root: string | null) => void) | null = null;
  private root: string;
  constructor(app: App, root: string) { super(app); this.root = root; }
  ask(): Promise<string | null> {
    return new Promise((resolve) => { this.result = resolve; this.open(); });
  }
  onOpen(): void {
    this.setTitle("Allow external presentations / 允许读取外部课件");
    this.contentEl.createEl("p", { text: "This reader will read presentation files in this folder on this device. Files stay in place. You may choose a containing parent folder. / 只读访问此目录中的演示文稿，不复制文件；可改为上级目录。" });
    new Setting(this.contentEl).setName("Folder / 文件夹").addText((field) => field.setValue(this.root).onChange((value) => { this.root = value; }));
    new Setting(this.contentEl)
      .addButton((button) => button.setButtonText("Cancel / 取消").onClick(() => this.close()))
      .addButton((button) => button.setButtonText("Allow / 允许").setCta().onClick(() => {
        const resolve = this.result; this.result = null; resolve?.(this.root); this.close();
      }));
  }
  onClose(): void { this.result?.(null); this.result = null; this.contentEl.empty(); }
}
