import { FileView, type TFile, type WorkspaceLeaf } from "obsidian";
import type WordReaderPlugin from "./main";
import { PPTX_VIEW_TYPE } from "./pptx/PptxAdapter";
import { createNoteFromPptx } from "./commands/createNoteFromPptx";
import { openExternalFile } from "./commands/openExternal";
import { createPptxReader } from "./pptx/PptxReaderView";
import type { ReaderFileChange } from "./reader/fileRevision";
import type { ReaderWorkspaceRecoveryReason } from "./reader/workspaceRecovery";

export const VIEW_TYPE_PPTX_READER = PPTX_VIEW_TYPE;

export class PptxView extends FileView {
  protected readonly reader;
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf);
    this.reader = createPptxReader(this, plugin);
  }
  get adapter() { return this.reader.adapter; }
  get capabilities() { return this.reader.capabilities; }
  getViewType(): string { return PPTX_VIEW_TYPE; }
  getDisplayText(): string { return this.reader.getDisplayText(); }
  getIcon(): string { return "presentation"; }
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
      createSummary: (slides, page, text) => createNoteFromPptx(this.app, file, slides, page, text),
    });
  }
  onUnloadFile() { return this.reader.onUnloadFile(); }
  reload() { return this.reader.reload(); }
  refreshInterfaceLanguage() { this.reader.refreshInterfaceLanguage(); }
  handleFileChange(change: ReaderFileChange) { this.reader.handleFileChange(change); }
  handleWorkspaceRecovery(reason: ReaderWorkspaceRecoveryReason) { this.reader.handleWorkspaceRecovery(reason); }
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
