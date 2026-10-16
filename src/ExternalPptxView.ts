import { ItemView, type TFile, type ViewStateResult, type WorkspaceLeaf } from "obsidian";
import type WordReaderPlugin from "./main";
import { createPptxReader } from "./pptx/PptxReaderView";
import { externalPptxSource } from "./external/source";
import { parseExternalPptxLink } from "./external/links";

export const EXTERNAL_PPTX_VIEW_TYPE = "external-pptx-reader-view";

export class ExternalPptxView extends ItemView {
  private readonly reader;
  private readonly plugin: WordReaderPlugin;
  private uri: string | null = null;
  private generation = 0;
  private source: Awaited<ReturnType<typeof externalPptxSource>> | null = null;
  private checking = false;
  private invalidated = false;
  private unavailableEl: HTMLElement | null = null;
  readonly managesOwnSource = true;
  readonly file: TFile | null = null;

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf); this.plugin = plugin; this.reader = createPptxReader(this, plugin);
  }
  get adapter() { return this.reader.adapter; }
  get capabilities() { return this.reader.capabilities; }
  hasSource(): boolean { return this.source !== null; }
  getViewType(): string { return EXTERNAL_PPTX_VIEW_TYPE; }
  getDisplayText(): string { return this.reader.getDisplayText(); }
  getIcon(): string { return "presentation"; }
  getState(): Record<string, unknown> { return { uri: this.uri }; }
  async setState(state: unknown, result: ViewStateResult): Promise<void> {
    const previousUri = this.uri;
    const generation = ++this.generation;
    this.source = null;
    this.invalidated = false;
    this.unavailableEl?.remove();
    this.unavailableEl = null;
    await this.reader.onUnloadFile();
    this.uri = typeof state === "object" && state !== null && "uri" in state && typeof state.uri === "string"
      ? parseExternalPptxLink(state.uri) : null;
    result.history = previousUri !== this.uri;
    if (!this.uri) { this.showUnavailable(); return; }
    try {
      const source = await externalPptxSource(this.app, this.uri, this.plugin.externalPptxAccess);
      if (generation !== this.generation) return;
      this.source = source;
      await this.reader.onLoadSource(source);
    } catch {
      if (generation === this.generation) this.showUnavailable();
    }
    if (generation === this.generation) await super.setState(state, result);
  }
  async onOpen(): Promise<void> {
    await this.reader.onOpen();
    this.registerInterval(this.contentEl.win.setInterval(() => { void this.checkSource(); }, 1500));
  }
  private async checkSource(): Promise<void> {
    const source = this.source;
    const generation = this.generation;
    if (!source || this.checking) return;
    this.checking = true;
    try {
      const changed = await source.changed();
      if (!changed) this.invalidated = false;
      if (changed && !this.invalidated && generation === this.generation) {
        this.reader.handleFileChange({ kind: "modified", path: source.file.path });
        this.invalidated = true;
      }
    } catch {
      if (generation === this.generation && !this.invalidated) {
        this.reader.handleWorkspaceRecovery("missing-file");
        this.invalidated = true;
      }
    } finally { this.checking = false; }
  }
  private showUnavailable(): void {
    this.unavailableEl?.remove();
    this.unavailableEl = this.contentEl.createDiv();
    this.unavailableEl.createEl("p", { text: "Presentation unavailable or folder access not granted. / 文件不可用或尚未授权目录。" });
    if (this.uri) {
      this.unavailableEl.createEl("button", { text: "Open presentation / 打开课件" }).addEventListener("click", () => {
        void this.plugin.openExternalPptx(this.uri ?? "", this.leaf);
      });
    }
  }
  async onClose(): Promise<void> { ++this.generation; this.source = null; await this.reader.onClose(); }
  onLoadFile(): Promise<void> { return Promise.resolve(); }
  onUnloadFile() { return this.reader.onUnloadFile(); }
  reload() { return this.reader.reload(); }
  refreshInterfaceLanguage() { this.reader.refreshInterfaceLanguage(); }
  copyText() { return this.reader.copyText(); }
  copyDiagnostics() { return this.reader.copyDiagnostics(); }
  createSummaryNote() { return this.reader.createSummaryNote(); }
  openExternal() { return this.reader.openExternal(); }
  previousPage() { return this.reader.previousPage(); }
  nextPage() { return this.reader.nextPage(); }
  toggleNotes() { this.reader.toggleNotes(); }
  toggleFullscreen() { return this.reader.toggleFullscreen(); }
  focusSearch() { this.reader.focusSearch(); }
}
