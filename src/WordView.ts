import { FileView, type TFile, type WorkspaceLeaf } from "obsidian";
import type WordReaderPlugin from "./main";
import { DocxSession } from "./docx/DocxSession";
import { DOCX_VIEW_TYPE } from "./docx/DocxAdapter";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import { openExternalDocx } from "./commands/openExternal";
import type { ReaderFileChange } from "./reader/fileRevision";
import type { ReaderWorkspaceRecoveryReason } from "./reader/workspaceRecovery";

export const VIEW_TYPE_WORD_READER = DOCX_VIEW_TYPE;

export class WordView extends FileView {
  protected readonly reader: DocxSession;
  private readonly plugin: WordReaderPlugin;
  constructor(leaf: WorkspaceLeaf, plugin: WordReaderPlugin) {
    super(leaf); this.plugin = plugin; this.reader = new DocxSession(this, plugin);
  }
  get adapter() { return this.reader.adapter; }
  get capabilities() { return this.reader.capabilities; }
  getViewType(): string { return DOCX_VIEW_TYPE; }
  getDisplayText(): string { return this.reader.getDisplayText(); }
  getIcon(): string { return this.reader.getIcon(); }
  onOpen() { return this.reader.onOpen(); }
  onClose() { return this.reader.onClose(); }
  onLoadFile(file: TFile) {
    return this.reader.onLoadSource({
      file, vaultFile: file,
      read: () => this.app.vault.readBinary(file),
      refresh: () => Promise.resolve(),
      openExternal: () => openExternalDocx(this.app, file, this.plugin.text),
      createSummary: (text) => createNoteFromDocx(this.app, file, text),
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
}
