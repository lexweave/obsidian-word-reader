/* eslint-disable obsidianmd/no-global-this, obsidianmd/prefer-instanceof -- This test host defines the browser globals and the instanceOf primitive itself. */
import { Window } from "happy-dom";

export const dom = new Window({ url: "http://localhost/" });
const doc = dom.document;
Object.assign(globalThis, { window: dom, document: doc, activeWindow: dom, activeDocument: doc,
  HTMLElement: dom.HTMLElement, Element: dom.Element, Node: dom.Node,
  MouseEvent: dom.MouseEvent, KeyboardEvent: dom.KeyboardEvent, DOMParser: dom.DOMParser,
  HTMLImageElement: dom.HTMLImageElement, HTMLStyleElement: dom.HTMLStyleElement,
  ResizeObserver: class { observe() {} disconnect() {} unobserve() {} },
});
Object.defineProperty(globalThis, "navigator", { configurable: true, value: dom.navigator });
Object.defineProperty(dom.Event.prototype, "targetNode", { get: function(this: Event) { return this.target; } });
Object.defineProperty(dom.Node.prototype, "instanceOf", { value(this: Node, type: new (...args: never[]) => unknown) { return this instanceof type; } });
Object.defineProperty(dom.Node.prototype, "doc", { get: function(this: Node) { return this.ownerDocument ?? doc; } });
Object.defineProperty(dom.Node.prototype, "win", { get() { return dom; } });
Object.defineProperty(dom.Node.prototype, "empty", { value(this: HTMLElement) { this.replaceChildren(); } });
Object.defineProperty(dom.HTMLElement.prototype, "setText", { value(this: HTMLElement, text: string) { this.textContent = text; } });
Object.defineProperty(dom.HTMLElement.prototype, "addClass", { value(this: HTMLElement, ...names: string[]) { this.classList.add(...names); } });
Object.defineProperty(dom.HTMLElement.prototype, "removeClass", { value(this: HTMLElement, ...names: string[]) { this.classList.remove(...names); } });
Object.defineProperty(dom.HTMLElement.prototype, "toggleClass", { value(this: HTMLElement, name: string, on: boolean) { this.classList.toggle(name, on); } });
Object.defineProperty(dom.HTMLElement.prototype, "setCssProps", { value(this: HTMLElement, props: Record<string, string>) { for (const [key, value] of Object.entries(props)) this.style.setProperty(key, value); } });
interface Options { cls?: string; text?: string; attr?: Record<string, string>; type?: string; }
function create(parent: HTMLElement, tag: string, options: Options = {}): HTMLElement {
  const el = doc.createElement(tag);
  if (options.cls) el.className = options.cls;
  if (options.text) el.textContent = options.text;
  for (const [key, value] of Object.entries(options.attr ?? {})) el.setAttribute(key, value);
  parent.appendChild(el as unknown as Node);
  return el as unknown as HTMLElement;
}
Object.defineProperty(dom.HTMLElement.prototype, "createEl", { value(this: HTMLElement, tag: string, options: Options) { return create(this, tag, options); } });
Object.defineProperty(dom.HTMLElement.prototype, "createDiv", { value(this: HTMLElement, options: Options) { return create(this, "div", options); } });
Object.defineProperty(dom.HTMLElement.prototype, "createSpan", { value(this: HTMLElement, options: Options) { return create(this, "span", options); } });
export const Platform = { isDesktopApp: true };
export function setIcon() {}
export function normalizePath(path: string): string { return path.replace(/\\/g, "/"); }
export class Notice { constructor(_message: string) {} }
export class TFile {
  stat = { mtime: Date.now(), ctime: Date.now(), size: 0 };
  extension = "pptx";
  basename: string;
  name: string;
  constructor(public path: string) { this.name = path.split("/").pop() ?? path; this.basename = this.name.replace(/\.pptx$/, ""); }
}
export class FileSystemAdapter { getFullPath(path: string) { return path; } }
export class TestApp {
  storage = new Map<string, unknown>();
  views = new Map<string, (leaf: TestLeaf) => ItemView>();
  leaves: TestLeaf[] = [];
  files = new Map<string, TFile>();
  data = new Map<string, ArrayBuffer | string>();
  active: TestLeaf | null = null;
  vault = {
    adapter: new FileSystemAdapter(),
    on: () => ({}),
    readBinary: (file: TFile) => Promise.resolve(this.data.get(file.path) as ArrayBuffer),
    getAbstractFileByPath: (path: string) => this.files.get(path) ?? null,
    createFolder: (path: string) => { this.files.set(path, new TFile(path)); return Promise.resolve(); },
    create: (path: string, content: string) => {
      const file = new TFile(path); this.files.set(path, file); this.data.set(path, content); return Promise.resolve(file);
    },
  };
  workspace = {
    on: () => ({}),
    onLayoutReady: (_callback: () => void) => {},
    getMostRecentLeaf: () => this.active,
    getLeavesOfType: (type: string) => this.leaves.filter((leaf) => leaf.view?.getViewType() === type),
    iterateAllLeaves: (callback: (leaf: TestLeaf) => void) => { this.leaves.forEach(callback); },
    getLeaf: (_newLeaf?: boolean) => new TestLeaf(this),
    revealLeaf: (leaf: TestLeaf) => { this.active = leaf; return Promise.resolve(); },
  };
  loadLocalStorage(key: string): unknown { return this.storage.get(key) ?? null; }
  saveLocalStorage(key: string, data: unknown) { this.storage.set(key, data); }
}
export class TestLeaf {
  app: TestApp;
  view: ItemView;
  element = doc.createElement("div");
  changes = 0;
  constructor(app: TestApp) {
    this.app = app; app.leaves.push(this); app.active = this;
    this.element.className = "workspace-leaf"; doc.body.appendChild(this.element);
    this.view = new ItemView(this);
  }
  async setViewState(state: { type: string; state?: unknown }) {
    this.changes++;
    await this.view.onClose(); this.view.unload(); this.view.containerEl.remove();
    const factory = this.app.views.get(state.type);
    if (!factory) throw new Error("Unknown test view");
    this.view = factory(this); await this.view.onOpen(); await this.view.setState(state.state, {});
  }
  async openFile(file: TFile) {
    this.changes++;
    if (this.view instanceof FileView) { this.view.file = file; await this.view.onLoadFile(file); }
  }
}
export class ItemView {
  app: TestApp;
  containerEl = doc.createElement("div") as unknown as HTMLElement;
  contentEl: HTMLElement;
  private cleanups: Array<() => void> = [];
  constructor(public leaf: TestLeaf) {
    this.app = leaf.app; leaf.element.appendChild(this.containerEl as unknown as typeof leaf.element);
    this.contentEl = create(this.containerEl, "div");
  }
  getViewType(): string { return "markdown"; }
  async onOpen() {}
  async onClose() {}
  async setState(_state: unknown, _result: unknown) {}
  registerInterval(id: number) { this.cleanups.push(() => dom.clearInterval(id as unknown as ReturnType<typeof dom.setInterval>)); }
  unload() { this.cleanups.forEach((fn) => fn()); }
}
export class FileView extends ItemView {
  file: TFile | null = null;
  async onLoadFile(_file: TFile) {}
}
export class Plugin {
  commands = new Map<string, { callback?: () => void; checkCallback?: (checking: boolean) => boolean }>();
  private cleanups: Array<() => void> = [];
  constructor(public app: TestApp, _manifest?: unknown) {}
  async loadData() { return null; }
  async saveData(_data: unknown) {}
  addSettingTab(_tab: unknown) {}
  registerView(type: string, factory: (leaf: TestLeaf) => ItemView) { this.app.views.set(type, factory); }
  registerExtensions(_ext: string[], _type: string) {}
  registerEvent(_event: unknown) {}
  registerEditorExtension(_extension: unknown) {}
  registerDomEvent(target: Document, name: string, fn: EventListener, options?: boolean) {
    target.addEventListener(name, fn, options); this.cleanups.push(() => target.removeEventListener(name, fn, options));
  }
  addCommand(command: { id: string; callback?: () => void; checkCallback?: (checking: boolean) => boolean }) { this.commands.set(command.id, command); }
  unload() { this.cleanups.forEach((fn) => fn()); }
}
export class Modal {
  contentEl = doc.createElement("div") as unknown as HTMLElement;
  constructor(_app: TestApp) { this.contentEl.className = "modal"; }
  setTitle(_title: string) {}
  open() { doc.body.appendChild(this.contentEl as unknown as typeof doc.body); this.onOpen(); }
  close() { this.onClose(); this.contentEl.remove(); }
  onOpen() {}
  onClose() {}
}
export class Setting {
  constructor(private container: HTMLElement) {}
  setName(_name: string) { return this; }
  addText(callback: (field: { setValue: (value: string) => unknown; onChange: (cb: (value: string) => void) => unknown }) => void) {
    const input = create(this.container, "input") as unknown as HTMLInputElement;
    const field = { setValue: (value: string) => { input.value = value; return field; }, onChange: (cb: (value: string) => void) => { input.addEventListener("input", () => cb(input.value)); return field; } };
    callback(field); return this;
  }
  addButton(callback: (button: { setButtonText: (value: string) => unknown; onClick: (cb: () => void) => unknown; setCta: () => unknown }) => void) {
    const el = create(this.container, "button");
    const button = { setButtonText: (value: string) => { el.textContent = value; return button; }, setCta: () => button, onClick: (cb: () => void) => { el.addEventListener("click", cb); return button; } };
    callback(button); return this;
  }
}
export class PluginSettingTab { constructor(_app: TestApp, _plugin: Plugin) {} }
