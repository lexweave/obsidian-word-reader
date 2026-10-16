import { ItemView, type TFile, type ViewStateResult, type WorkspaceLeaf } from "obsidian";
import type WordReaderPlugin from "./main";
import { createPptxReader } from "./pptx/PptxReaderView";
import { externalOfficeSource, type ExternalOfficeSource } from "./external/source";
import { DocxSession } from "./docx/DocxSession";
import { XlsxSession } from "./xlsx/XlsxSession";
import { PptxSession } from "./pptx/PptxSession";
import { buildSummaryNote } from "./docx/docxSummaryNote";
import { buildXlsxSummaryNote } from "./xlsx/xlsxSummaryNote";
import { buildPresentationSummaryNote } from "./pptx/pptxSummaryNote";
import { parseExternalOfficeLink, externalOfficeFormat, type ExternalOfficeFormat } from "./external/links";

export const EXTERNAL_PPTX_VIEW_TYPE = "external-pptx-reader-view";

export const EXTERNAL_DOCX_VIEW_TYPE = "external-docx-reader-view";
export const EXTERNAL_XLSX_VIEW_TYPE = "external-xlsx-reader-view";
export const externalOfficeViewType = (format: ExternalOfficeFormat): string => `external-${format}-reader-view`;

export class ExternalOfficeView extends ItemView {
  private readonly reader: PptxSession | DocxSession | XlsxSession;
  private readonly loadSource: (source: ExternalOfficeSource) => Promise<void>;
  private readonly plugin: WordReaderPlugin;
  private uri: string | null = null;
  private generation = 0;
  private source: ExternalOfficeSource | null = null;
  private checking = false;
  private invalidated = false;
  private unavailableEl: HTMLElement | null = null;
  readonly managesOwnSource = true;
  readonly file: TFile | null = null;

  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin, private readonly format: ExternalOfficeFormat) {
    super(leaf); this.plugin = plugin;
    if (format === "docx") {
      const reader = new DocxSession(this, plugin); this.reader = reader;
      this.loadSource = (source) => reader.onLoadSource({ ...source,
        createSummary: (text) => source.createSummaryMarkdown(buildSummaryNote(source.file, text), text.notices.createdSummaryNote),
      });
    } else if (format === "xlsx") {
      const reader = new XlsxSession(this, plugin); this.reader = reader;
      this.loadSource = (source) => reader.onLoadSource({ ...source,
        createSummary: (workbook, sheet, range, markdown, text) => source.createSummaryMarkdown(
          buildXlsxSummaryNote(source.file, workbook, sheet, range, markdown, text), text.notices.createdSummaryNote),
      });
    } else {
      const reader = createPptxReader(this, plugin); this.reader = reader;
      this.loadSource = (source) => reader.onLoadSource({ ...source,
        createSummary: (slides, page, text) => source.createSummaryMarkdown(
          buildPresentationSummaryNote(source.file, slides, page, text), text.notices.createdSummaryNote),
      });
    }
  }
  get adapter() { return this.reader.adapter; }
  get capabilities() { return this.reader.capabilities; }
  hasSource(): boolean { return this.source !== null; }
  getViewType(): string { return externalOfficeViewType(this.format); }
  getDisplayText(): string { return this.reader.getDisplayText(); }
  getIcon(): string { return this.reader.getIcon(); }
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
      ? parseExternalOfficeLink(state.uri) : null;
    result.history = previousUri !== this.uri;
    if (!this.uri || externalOfficeFormat(this.uri) !== this.format) { this.uri = null; this.showUnavailable(); return; }
    try {
      const source = await externalOfficeSource(this.app, this.uri, this.plugin.externalOfficeAccess);
      if (generation !== this.generation) return;
      this.source = source;
      await this.loadSource(source);
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
    this.unavailableEl.createEl("p", { text: "Document unavailable or folder access not granted." });
    if (this.uri) {
      this.unavailableEl.createEl("button", { text: "Open document" }).addEventListener("click", () => {
        void this.plugin.openExternalOffice(this.uri ?? "", this.leaf);
      });
    }
  }
  async onClose(): Promise<void> { ++this.generation; this.source = null; await this.reader.onClose(); }
  onLoadFile(): Promise<void> { return Promise.resolve(); }
  onUnloadFile() { return this.reader.onUnloadFile(); }
  reload() { return this.reader.reload(); }
  refreshInterfaceLanguage() { this.reader.refreshInterfaceLanguage(); }
  copyText() { return this.reader.copyText(); }
  copyMarkdown() { return this.reader instanceof PptxSession ? Promise.resolve() : this.reader.copyMarkdown(); }
  copyFormulas() { return this.reader instanceof XlsxSession ? this.reader.copyFormulas() : Promise.resolve(); }
  copyDiagnostics() { return this.reader.copyDiagnostics(); }
  createSummaryNote() { return this.reader.createSummaryNote(); }
  openExternal() { return this.reader.openExternal(); }
  previousPage() { return this.reader instanceof PptxSession ? this.reader.previousPage() : Promise.resolve(); }
  nextPage() { return this.reader instanceof PptxSession ? this.reader.nextPage() : Promise.resolve(); }
  toggleNotes() { if (this.reader instanceof PptxSession) this.reader.toggleNotes(); }
  toggleFullscreen() { return this.reader instanceof PptxSession ? this.reader.toggleFullscreen() : Promise.resolve(); }
  focusSearch() { if (this.reader instanceof PptxSession || this.reader instanceof XlsxSession) this.reader.focusSearch(); }
  focusNameBox() { if (this.reader instanceof XlsxSession) this.reader.focusNameBox(); }
}

export class ExternalPptxView extends ExternalOfficeView {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) { super(leaf, plugin, "pptx"); }
}
export class ExternalDocxView extends ExternalOfficeView {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) { super(leaf, plugin, "docx"); }
}
export class ExternalXlsxView extends ExternalOfficeView {
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) { super(leaf, plugin, "xlsx"); }
}
