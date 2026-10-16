import { FileView, type TFile, type WorkspaceLeaf } from "obsidian";
import type WordReaderPlugin from "./main";
import { XlsxSession } from "./xlsx/XlsxSession";
import { XLSX_VIEW_TYPE } from "./xlsx/XlsxAdapter";
import { createNoteFromXlsx } from "./commands/createNoteFromXlsx";
import { openExternalFile } from "./commands/openExternal";
import type { ReaderFileChange } from "./reader/fileRevision";
import type { ReaderWorkspaceRecoveryReason } from "./reader/workspaceRecovery";

export const VIEW_TYPE_XLSX_READER = XLSX_VIEW_TYPE;

export class XlsxView extends FileView {
  protected readonly reader: XlsxSession;
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf); this.reader = new XlsxSession(this, plugin);
  }
  get adapter() { return this.reader.adapter; }
  get capabilities() { return this.reader.capabilities; }
  getViewType(): string { return XLSX_VIEW_TYPE; }
  getDisplayText(): string { return this.reader.getDisplayText(); }
  getIcon(): string { return this.reader.getIcon(); }
  onOpen() { return this.reader.onOpen(); }
  onClose() { return this.reader.onClose(); }
  onLoadFile(file: TFile) {
    return this.reader.onLoadSource({
      file, vaultFile: file,
      read: () => this.app.vault.readBinary(file),
      refresh: () => Promise.resolve(),
      openExternal: () => openExternalFile(this.app, file, {
        desktopOnly: "Only available on desktop", localVaultOnly: "Requires a local vault",
      }),
      createSummary: (workbook, sheet, range, markdown, text) => createNoteFromXlsx(this.app, file, workbook, sheet, range, markdown, text),
    });
  }
  onUnloadFile() { return this.reader.onUnloadFile(); }
  refreshInterfaceLanguage() { this.reader.refreshInterfaceLanguage(); }
  handleFileChange(change: ReaderFileChange) { this.reader.handleFileChange(change); }
  handleWorkspaceRecovery(reason: ReaderWorkspaceRecoveryReason) { this.reader.handleWorkspaceRecovery(reason); }
  reload() { return this.reader.reload(); }
  copyText() { return this.reader.copyText(); }
  copyMarkdown() { return this.reader.copyMarkdown(); }
  copyDiagnostics() { return this.reader.copyDiagnostics(); }
  createSummaryNote() { return this.reader.createSummaryNote(); }
  openExternal() { return this.reader.openExternal(); }
  focusSearch() { this.reader.focusSearch(); }
  focusNameBox() { this.reader.focusNameBox(); }
  copyFormulas() { return this.reader.copyFormulas(); }
}
