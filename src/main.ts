import { Plugin, TFile, Platform, Notice, type WorkspaceLeaf } from "obsidian";
import { ExternalPptxView, ExternalDocxView, ExternalXlsxView, EXTERNAL_DOCX_VIEW_TYPE, EXTERNAL_XLSX_VIEW_TYPE, EXTERNAL_PPTX_VIEW_TYPE, externalOfficeViewType } from "./ExternalOfficeView";
import { ExternalOfficeAccess } from "./external/access";
import { parseExternalOfficeLink, externalOfficeAtPosition, externalOfficeFormat } from "./external/links";
import { EditorView } from "@codemirror/view";

import {
  PptxView,
  VIEW_TYPE_PPTX_READER,
} from "./PptxView";
import {
  WordView,
  VIEW_TYPE_WORD_READER,
} from "./WordView";
import {
  XlsxView,
  VIEW_TYPE_XLSX_READER,
} from "./XlsxView";
import { DOCX_ADAPTER } from "./docx/DocxAdapter";
import { getWordReaderText, type WordReaderText } from "./i18n";
import { PPTX_ADAPTER } from "./pptx/PptxAdapter";
import { getPptxReaderText } from "./pptx/pptxI18n";
import {
  XLSX_ADAPTER,
} from "./xlsx/XlsxAdapter";
import { getXlsxReaderText } from "./xlsx/xlsxI18n";
import {
  ReadingStateStore,
  type ReaderFileIdentity,
  type ReaderFormat,
  type ReaderViewState,
} from "./reader/readingState";
import {
  type ReaderFileChange,
  type ReaderFileDescriptor,
  type ReaderFileRevision,
  ReaderFileRevisionTracker,
} from "./reader/fileRevision";
import { recoverPluginData } from "./reader/pluginData";
import {
  isReaderSession,
  type ReaderSession,
} from "./reader/session";
import {
  planReaderWorkspaceRecovery,
  type ReaderWorkspaceEntry,
} from "./reader/workspaceRecovery";
import {
  DEFAULT_OFFICE_READER_SETTINGS,
  WordReaderSettingTab,
  type OfficeReaderSettings,
  normalizeOfficeReaderSettings,
} from "./settings";

const DATA_SAVE_DEBOUNCE_MS = 500;

export default class WordReaderPlugin extends Plugin {
  settings: OfficeReaderSettings = normalizeOfficeReaderSettings(
    DEFAULT_OFFICE_READER_SETTINGS,
  );
  externalOfficeAccess = new ExternalOfficeAccess(this.app);
  private openingExternal = false;
  private readonly externalDocuments = new Set<Document>();
  private readingStates = new ReadingStateStore();
  private readonly fileRevisions = new ReaderFileRevisionTracker();
  private dataSaveTimer: number | null = null;
  private dataSavePromise: Promise<void> = Promise.resolve();

  async onload(): Promise<void> {
    await this.loadSettings();
    if (Platform.isDesktopApp) {
      this.registerView(EXTERNAL_PPTX_VIEW_TYPE, (leaf) => new ExternalPptxView(leaf, this));
      this.registerView(EXTERNAL_DOCX_VIEW_TYPE, (leaf) => new ExternalDocxView(leaf, this));
      this.registerView(EXTERNAL_XLSX_VIEW_TYPE, (leaf) => new ExternalXlsxView(leaf, this));
      this.registerExternalOfficeLinks(activeDocument);
      this.app.workspace.iterateAllLeaves((leaf) => this.registerExternalOfficeLinks(leaf.view.containerEl.doc));
      this.registerEvent(this.app.workspace.on("window-open", (_workspace, win) => {
        this.registerExternalOfficeLinks(win.document);
      }));
      this.addCommand({ id: "reset-external-pptx-access", name: "Forget external office folder permissions", callback: () => {
        this.externalOfficeAccess.clear();
        new Notice("External office folder permissions cleared.");
      } });
    }
    this.addSettingTab(new WordReaderSettingTab(this.app, this));
    const text = this.text;
    const pptxText = getPptxReaderText(this.settings.common.language);
    const xlsxText = getXlsxReaderText(this.settings.common.language);

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_PPTX_READER,
      (leaf) => new PptxView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_XLSX_READER,
      (leaf) => new XlsxView(leaf, this),
    );
    this.registerExtensions(
      [...DOCX_ADAPTER.extensions],
      DOCX_ADAPTER.viewType,
    );
    this.registerExtensions(
      [...PPTX_ADAPTER.extensions],
      PPTX_ADAPTER.viewType,
    );
    this.registerExtensions(
      [...XLSX_ADAPTER.extensions],
      XLSX_ADAPTER.viewType,
    );
    this.registerEvent(this.app.vault.on("modify", (file) => {
      if (file instanceof TFile) {
        this.notifyReaderFileChange({ kind: "modified", path: file.path });
      }
    }));
    this.registerEvent(this.app.vault.on("delete", (file) => {
      if (file instanceof TFile) {
        this.notifyReaderFileChange({ kind: "deleted", path: file.path });
      }
    }));
    this.registerEvent(this.app.vault.on("rename", (file, oldPath) => {
      if (file instanceof TFile) {
        this.fileRevisions.markRenamed(oldPath, file.path);
        this.notifyReaderFileChange({
          kind: "renamed",
          path: oldPath,
          newPath: file.path,
          newExtension: file.extension,
        }, false);
      }
    }));
    this.app.workspace.onLayoutReady(() => {
      try {
        this.recoverWorkspaceViews();
      } catch (error) {
        console.error("Office Reader could not inspect restored views", error);
      }
    });

    this.registerReaderCommand(
      "reload",
      text.commands.reload,
      (session) => session.capabilities.reload,
      (session) => session.reload(),
    );
    this.registerReaderCommand(
      "copy-text",
      text.commands.copyText,
      (session) => session.capabilities.copyText && Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-markdown",
      text.commands.copyMarkdown,
      (session) =>
        session.adapter.format === "docx" &&
        session.capabilities.copyMarkdown && Boolean(session.copyMarkdown),
      (session) => session.copyMarkdown?.(),
    );
    this.registerReaderCommand(
      "create-note",
      text.commands.createNote,
      (session) =>
        session.adapter.format === "docx" &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "open-external",
      text.commands.openExternal,
      (session) =>
        session.capabilities.openExternal && Boolean(session.openExternal),
      (session) => session.openExternal?.(),
    );

    this.registerReaderCommand(
      "previous-slide",
      pptxText.commands.previousSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.previousPage),
      (session) => session.previousPage?.(),
    );
    this.registerReaderCommand(
      "next-slide",
      pptxText.commands.nextSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.nextPage),
      (session) => session.nextPage?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-fullscreen",
      pptxText.commands.toggleFullscreen,
      (session) =>
        session.capabilities.fullscreen && Boolean(session.toggleFullscreen),
      (session) => session.toggleFullscreen?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-text",
      pptxText.commands.copySlideText,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.copyText &&
        Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-render-diagnostics",
      pptxText.commands.copyRenderDiagnostics,
      (session) =>
        session.capabilities.diagnostics && Boolean(session.copyDiagnostics),
      (session) => session.copyDiagnostics?.(),
    );
    this.registerReaderCommand(
      "create-presentation-note",
      pptxText.commands.createSummaryNote,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-notes",
      pptxText.commands.toggleNotes,
      (session) =>
        session.capabilities.notes && Boolean(session.toggleNotes),
      (session) => session.toggleNotes?.(),
    );
    this.registerReaderCommand(
      "search-presentation",
      pptxText.commands.focusSearch,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.search &&
        Boolean(session.focusSearch),
      (session) => session.focusSearch?.(),
    );
    this.registerReaderCommand(
      "search-spreadsheet",
      xlsxText.commands.focusSearch,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.search &&
        Boolean(session.focusSearch),
      (session) => session.focusSearch?.(),
    );
    this.registerReaderCommand(
      "go-to-spreadsheet-cell",
      xlsxText.commands.goToCell,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.navigation &&
        Boolean(session.focusNameBox),
      (session) => session.focusNameBox?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-values",
      xlsxText.commands.copyValues,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyText &&
        Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-formulas",
      xlsxText.commands.copyFormulas,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyText &&
        Boolean(session.copyFormulas),
      (session) => session.copyFormulas?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-markdown",
      xlsxText.commands.copyMarkdown,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyMarkdown &&
        Boolean(session.copyMarkdown),
      (session) => session.copyMarkdown?.(),
    );
    this.registerReaderCommand(
      "create-spreadsheet-note",
      xlsxText.commands.createSummaryNote,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
  }

  onunload(): void {
    this.externalDocuments.clear();
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }
  }

  private registerExternalOfficeLinks(doc: Document): void {
    if (this.externalDocuments.has(doc)) return;
    this.externalDocuments.add(doc);
    // Obsidian Live Preview uses CodeMirror spans, not <a> elements. Capture
    // mousedown as well: its editor handler can open URLs before click fires.
    this.registerDomEvent(doc, "mousedown", (event: MouseEvent) => {
      const target = event.targetNode;
      if (event.button !== 0 || !target?.instanceOf(Element)) return;
      if (!this.externalOfficeFromEditorEvent(event, target)) return;
      if (!this.externalLinkLeaf(target)) return;
      event.preventDefault(); event.stopImmediatePropagation();
    }, true);
    this.registerDomEvent(doc, "click", (event: MouseEvent) => {
      const target = event.targetNode;
      if (event.button !== 0 || !target?.instanceOf(Element)) return;
      const anchor = target.closest("a[href]");
      const href = anchor?.getAttribute("href");
      const uri = href ? parseExternalOfficeLink(href) : this.externalOfficeFromEditorEvent(event, target);
      if (!uri) return;
      const sourceLeaf = this.externalLinkLeaf(anchor ?? target);
      if (!sourceLeaf) return;
      event.preventDefault(); event.stopImmediatePropagation();
      void this.openExternalOffice(uri, sourceLeaf);
    }, true);
  }

  private externalOfficeFromEditorEvent(event: MouseEvent, target: Element): string | null {
    const editor = target.closest(".cm-editor");
    if (!editor) return null;
    const renderedLink = target.closest(".cm-link, .cm-url, .external-link");
    const livePreview = !!target.closest(".is-live-preview");
    if (!(event.ctrlKey || event.metaKey) && !(livePreview && renderedLink)) return null;
    try {
      const view = EditorView.findFromDOM(editor as HTMLElement);
      if (!view) return null;
      // DOM positions also work for hidden URL/icon widgets in Live Preview.
      const position = renderedLink ? view.posAtDOM(target) : view.posAtCoords({ x: event.clientX, y: event.clientY });
      if (position === null) return null;
      const line = view.state.doc.lineAt(position);
      return externalOfficeAtPosition(line.text, position - line.from);
    } catch { return null; }
  }

  private externalLinkLeaf(element: Element): WorkspaceLeaf | null {
    const pane = element.closest(".workspace-leaf");
    let result: WorkspaceLeaf | null = null;
    this.app.workspace.iterateAllLeaves((leaf) => {
      if (pane?.contains(leaf.view.containerEl)) result = leaf;
    });
    return result;
  }

  async openExternalOffice(uri: string, leaf: WorkspaceLeaf): Promise<void> {
    if (!Platform.isDesktopApp || !parseExternalOfficeLink(uri) || this.openingExternal) return;
    this.openingExternal = true;
    try {
      const path = await this.externalOfficeAccess.require(uri, true);
      if (!path) return;
      const format = externalOfficeFormat(uri);
      if (!format) return;
      await leaf.setViewState({ type: externalOfficeViewType(format), state: { uri }, active: true });
      await this.app.workspace.revealLeaf(leaf);
    } catch {
      new Notice("Could not open the external office document. Check the file and folder permission.");
    } finally { this.openingExternal = false; }
  }

  get text(): WordReaderText {
    return getWordReaderText(this.settings.common.language);
  }

  refreshWordReaderViews(): void {
    for (const viewType of [
      VIEW_TYPE_WORD_READER,
      VIEW_TYPE_PPTX_READER,
      VIEW_TYPE_XLSX_READER,
      EXTERNAL_PPTX_VIEW_TYPE,
      EXTERNAL_DOCX_VIEW_TYPE,
      EXTERNAL_XLSX_VIEW_TYPE,
    ]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        if (isReaderSession(leaf.view)) {
          leaf.view.refreshInterfaceLanguage();
        }
      }
    }
  }

  async loadSettings(): Promise<void> {
    let loadedData: unknown;
    try {
      loadedData = await this.loadData();
    } catch (error) {
      console.error("Office Reader could not load plugin data", error);
      const recovered = recoverPluginData(undefined);
      this.settings = recovered.settings;
      this.readingStates = recovered.readingStates;
      return;
    }
    const recovered = recoverPluginData(loadedData);
    this.settings = recovered.settings;
    this.readingStates = recovered.readingStates;
    if (recovered.requiresRepair) {
      await this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not repair plugin data", error);
      });
    }
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeOfficeReaderSettings(this.settings);
    await this.flushData();
  }

  getReadingState(
    file: TFile | ReaderFileIdentity | string,
  ): ReaderViewState | undefined {
    return typeof file === "string"
      ? this.readingStates.get(file)
      : this.readingStates.get(toReaderIdentity(file));
  }

  updateReadingState(
    file: TFile | ReaderFileIdentity | string,
    state: ReaderViewState,
  ): void {
    if (typeof file === "string") {
      this.readingStates.set(file, state);
    } else {
      this.readingStates.set(toReaderIdentity(file), state);
    }
    this.scheduleDataSave();
  }

  captureFileRevision(
    file: ReaderFileDescriptor,
    format: ReaderFormat,
  ): ReaderFileRevision {
    return this.fileRevisions.capture(file, format);
  }

  isFileRevisionCurrent(
    revision: ReaderFileRevision | null,
    file: ReaderFileDescriptor | null,
    format: ReaderFormat,
  ): boolean {
    return this.fileRevisions.isCurrent(revision, file, format);
  }

  async flushData(): Promise<void> {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
    }
    await this.persistData().catch((error: unknown) => {
      console.error("Office Reader could not save plugin data", error);
    });
  }

  private registerReaderCommand(
    id: string,
    name: string,
    isAvailable: (session: ReaderSession) => boolean,
    run: (session: ReaderSession) => void | Promise<void>,
  ): void {
    this.addCommand({
      id,
      name,
      checkCallback: (checking) => {
        const session = this.getActiveReaderSession();
        if (!session || (!session.file && !session.hasSource?.()) || !isAvailable(session)) {
          return false;
        }
        if (!checking) {
          void Promise.resolve(run(session)).catch((error: unknown) => {
            console.error("Office Reader command failed", error);
          });
        }
        return true;
      },
    });
  }

  private getActiveReaderSession(): ReaderSession | null {
    const view = this.app.workspace.getMostRecentLeaf()?.view;
    return isReaderSession(view) ? view : null;
  }

  private scheduleDataSave(): void {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
    }
    this.dataSaveTimer = window.setTimeout(() => {
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }, DATA_SAVE_DEBOUNCE_MS);
  }

  private persistData(): Promise<void> {
    const data = {
      ...this.settings,
      readingStates: this.readingStates.serialize(),
    };
    this.dataSavePromise = this.dataSavePromise
      .catch(() => undefined)
      .then(async () => {
        await this.saveData(data);
      });
    return this.dataSavePromise;
  }

  private notifyReaderFileChange(
    change: ReaderFileChange,
    markChanged = true,
  ): void {
    if (markChanged) {
      this.fileRevisions.markChanged(change.path);
    }
    if (this.readingStates.invalidate(change.path)) {
      this.scheduleDataSave();
    }
    for (const session of this.getReaderSessions()) {
      try {
        session.handleFileChange?.(change);
      } catch (error) {
        console.error("Office Reader could not invalidate a changed file", error);
      }
    }
  }

  private recoverWorkspaceViews(): void {
    const sessions = this.getReaderSessions().filter((session) => !session.managesOwnSource);
    const entries: ReaderWorkspaceEntry[] = sessions.map((session, index) => {
      const file = session.file;
      const currentFile = file
        ? this.app.vault.getAbstractFileByPath(file.path)
        : null;
      const availableFile = currentFile instanceof TFile ? currentFile : null;
      return {
        id: String(index),
        format: session.adapter.format,
        path: availableFile?.path ?? null,
        extension: availableFile?.extension ?? file?.extension ?? null,
      };
    });
    for (const action of planReaderWorkspaceRecovery(entries)) {
      if (!action.reason) {
        continue;
      }
      try {
        sessions[Number(action.id)]?.handleWorkspaceRecovery?.(action.reason);
      } catch (error) {
        console.error("Office Reader could not recover a workspace view", error);
      }
    }
  }

  private getReaderSessions(): ReaderSession[] {
    const sessions: ReaderSession[] = [];
    for (const viewType of [
      VIEW_TYPE_WORD_READER,
      VIEW_TYPE_PPTX_READER,
      VIEW_TYPE_XLSX_READER,
      EXTERNAL_PPTX_VIEW_TYPE,
      EXTERNAL_DOCX_VIEW_TYPE,
      EXTERNAL_XLSX_VIEW_TYPE,
    ]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        if (isReaderSession(leaf.view)) {
          sessions.push(leaf.view);
        }
      }
    }
    return sessions;
  }
}

function toReaderIdentity(
  file: TFile | ReaderFileIdentity,
): ReaderFileIdentity {
  if ("format" in file) {
    return file;
  }
  return {
    path: file.path,
    mtime: file.stat.mtime,
    format: getReaderFormat(file.extension),
  };
}

function getReaderFormat(extension: string): ReaderFormat {
  switch (extension.toLowerCase()) {
    case "pptx":
      return "pptx";
    case "xlsx":
      return "xlsx";
    default:
      return "docx";
  }
}
