import assert from "node:assert/strict";
import test, { mock } from "node:test";
import { EditorView } from "@codemirror/view";
import * as fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import type { App, PluginManifest, TFile as VaultFile, WorkspaceLeaf } from "obsidian";
import WordReaderPlugin from "../src/main";
import { ExternalPptxView, EXTERNAL_PPTX_VIEW_TYPE } from "../src/ExternalPptxView";
import { PptxView } from "../src/PptxView";
import { TestApp, TestLeaf, TFile, FileView, dom } from "./helpers/obsidianMock";
import { createMinimalPptx } from "./pptxFixture";

async function until(check: () => boolean): Promise<void> {
  for (let i = 0; i < 300; i++) {
    if (check()) return;
    await new Promise((resolve) => window.setTimeout(resolve, 10));
  }
  assert.ok(check(), "View did not reach the expected state");
}

void test("Live Preview spans intercept ordinary clicks and early editor mousedown only for PPTX", async () => {
  const app = new TestApp();
  const plugin = new WordReaderPlugin(app as unknown as App, {} as PluginManifest);
  await plugin.onload();
  const leaf = new TestLeaf(app);
  const source = dom.document.createElement("div"); source.className = "is-live-preview";
  const editorEl = dom.document.createElement("div"); editorEl.className = "cm-editor";
  const link = dom.document.createElement("span"); link.className = "cm-link";
  const label = dom.document.createElement("span"); label.className = "cm-underline"; label.textContent = "Slides";
  link.appendChild(label); editorEl.appendChild(link); source.appendChild(editorEl);
  leaf.view.contentEl.appendChild(source as unknown as Node);
  const target = source.querySelector(".cm-underline")!;
  let text = "[Slides](<file:///tmp/slides.pptx>) · [Canvas](https://example.org)";
  let position = 3;
  const editor = mock.method(EditorView, "findFromDOM", () => ({
    posAtDOM: () => position,
    state: { doc: { lineAt: () => ({ text, from: 0 }) } },
  }));
  const open = mock.method(plugin, "openExternalPptx", async () => {});
  let externalCalls = 0;
  target.addEventListener("mousedown", () => { externalCalls++; });
  target.addEventListener("click", () => { externalCalls++; });
  const dispatch = (type: string) => target.dispatchEvent(new dom.MouseEvent(type, { bubbles: true, cancelable: true, button: 0 }));
  try {
    assert.equal(dispatch("mousedown"), false);
    assert.equal(open.mock.callCount(), 0, "Open only once, on click");
    assert.equal(dispatch("click"), false);
    assert.equal(externalCalls, 0, "Stop Obsidian's external opener before it runs");
    assert.equal(open.mock.callCount(), 1);
    assert.equal(open.mock.calls[0]?.arguments[0], "file:///tmp/slides.pptx");
    assert.ok(open.mock.calls[0]?.arguments[1] === leaf as unknown as WorkspaceLeaf);
    position = text.indexOf("Canvas") + 1;
    dispatch("mousedown"); dispatch("click");
    assert.equal(externalCalls, 2, "Leave other links to Obsidian");
    position = 3; text = "[Slides](<file:///tmp/slides.docx>)";
    dispatch("mousedown"); dispatch("click");
    assert.equal(open.mock.callCount(), 1, "Do not intercept other file formats");
    text = "[Slides](<file:///tmp/slides.pptx>)"; source.className = "";
    dispatch("mousedown"); dispatch("click");
    assert.equal(open.mock.callCount(), 1, "Plain source-mode clicks keep editing behavior");
  } finally { editor.mock.restore(); open.mock.restore(); plugin.unload(); source.remove(); }
});

void test("actual PPTX views: click routing, permission cancellation, rendering, summaries, replacement and close", async () => {
  const folder = await fs.mkdtemp(path.join(os.tmpdir(), "office-reader-view-"));
  const app = new TestApp();
  const plugin = new WordReaderPlugin(app as unknown as App, {} as PluginManifest);
  const bytes = await createMinimalPptx();
  const filePath = path.join(folder, "课件 (one).pptx");
  await fs.writeFile(filePath, Buffer.from(bytes));
  const before = await fs.readFile(filePath);
  const uri = pathToFileURL(filePath).href;
  try {
    await plugin.onload();
    const clicked = new TestLeaf(app);
    const other = new TestLeaf(app); // Deliberately make a different pane the active pane.
    const anchor = dom.document.createElement("a");
    anchor.href = uri; anchor.textContent = "Slides";
    clicked.view.contentEl.appendChild(anchor as unknown as Node);
    const click = () => anchor.dispatchEvent(new dom.MouseEvent("click", { bubbles: true, cancelable: true, button: 0 }));
    click();
    await until(() => dom.document.querySelectorAll(".modal button").length === 2);
    (dom.document.querySelector(".modal button") as unknown as HTMLButtonElement).click();
    await new Promise((resolve) => window.setTimeout(resolve, 20));
    assert.equal(clicked.changes, 0, "Cancel must preserve the navigation pane");
    assert.equal(other.changes, 0);
    click();
    await until(() => dom.document.querySelectorAll(".modal button").length === 2);
    (dom.document.querySelectorAll(".modal button")[1] as unknown as HTMLButtonElement).click();
    await until(() => clicked.view.getViewType() === EXTERNAL_PPTX_VIEW_TYPE && clicked.view.contentEl.querySelector("svg") !== null && app.active === clicked);
    const view = clicked.view as unknown as ExternalPptxView;
    assert.equal(other.changes, 0, "Must open in the clicked pane, not the most recent pane");
    assert.ok(app.active === clicked);
    assert.ok(view.contentEl.querySelector("svg"), "Use the real PPTX renderer");
    assert.equal(view.hasSource(), true);
    assert.equal(view.file, null, "External files must never masquerade as vault TFiles");
    assert.deepEqual(await fs.readFile(filePath), before);
    assert.deepEqual(await fs.readdir(folder), [path.basename(filePath)]);
    assert.equal(app.files.size, 0, "Opening must not create a vault copy");
    await view.createSummaryNote();
    const summary = [...app.data.entries()].find(([name]) => name.endsWith(".md"));
    assert.ok(summary?.[0].startsWith("Office Reader Notes/"));
    assert.ok(typeof summary?.[1] === "string");
    assert.match(summary[1], /file:\/\/\/.*\.pptx/);
    assert.deepEqual(await fs.readdir(folder), [path.basename(filePath)], "Summary belongs in the vault");
    const state = view.getState();
    await view.onClose(); view.unload();
    const restoredLeaf = new TestLeaf(app);
    const restored = new ExternalPptxView(restoredLeaf as unknown as WorkspaceLeaf, plugin);
    restoredLeaf.view = restored as unknown as typeof restoredLeaf.view;
    await restored.onOpen();
    const result = { history: false };
    await restored.setState(state, result);
    assert.equal(result.history, true, "External URI changes belong in navigation history");
    assert.ok(restored.contentEl.textContent?.includes("Hello PPTX"));
    // Replace the source at the same path, as a Canvas update would.
    const replacement = path.join(folder, "updated.pptx");
    await fs.writeFile(replacement, Buffer.from(bytes));
    await fs.rename(replacement, filePath);
    await until(() => restored.contentEl.querySelector(".pptx-reader-error") !== null);
    await restored.reload();
    assert.ok(restored.contentEl.textContent?.includes("Hello PPTX"));
    await fs.unlink(filePath);
    await restored.reload();
    assert.ok(restored.contentEl.querySelector(".pptx-reader-error"));
    await restored.onClose(); restored.unload();

    // Original vault source still uses FileView, rendering and existing summary behavior.
    const nativeLeaf = new TestLeaf(app);
    const native = new PptxView(nativeLeaf as unknown as WorkspaceLeaf, plugin);
    nativeLeaf.view = native as unknown as typeof nativeLeaf.view;
    const file = new TFile("lesson.pptx"); file.stat.size = bytes.byteLength;
    app.files.set(file.path, file); app.data.set(file.path, bytes);
    native.file = file as unknown as VaultFile;
    await native.onOpen(); await native.onLoadFile(file as unknown as VaultFile);
    assert.ok(native instanceof FileView);
    assert.ok(native.contentEl.querySelector("svg"));
    await native.reload();
    await native.createSummaryNote();
    assert.ok(app.files.has("lesson.md"), "Preserve same-name summaries for vault files");
    await native.onUnloadFile(); await native.onClose(); native.unload();
  } finally {
    plugin.onunload(); (plugin as unknown as { unload(): void }).unload();
    for (const leaf of app.leaves) { await leaf.view.onClose(); leaf.view.unload(); leaf.element.remove(); }
    await fs.rm(folder, { recursive: true, force: true });
    await dom.happyDOM.abort();
  }
});

void test("supplied external PPTX renders through the view without changing its bytes", {
  skip: !process.env.OFFICE_READER_TEST_PPTX,
}, async () => {
  const filePath = process.env.OFFICE_READER_TEST_PPTX;
  assert.ok(filePath);
  const original = await fs.readFile(filePath);
  const app = new TestApp();
  const plugin = new WordReaderPlugin(app as unknown as App, {} as PluginManifest);
  try {
    await plugin.onload();
    const realPath = await fs.realpath(filePath);
    app.saveLocalStorage("word-reader-external-pptx-folders", [path.dirname(realPath)]);
    const leaf = new TestLeaf(app);
    await leaf.setViewState({ type: EXTERNAL_PPTX_VIEW_TYPE, state: { uri: pathToFileURL(realPath).href } });
    assert.ok(leaf.view.contentEl.querySelector("svg"), "Actual presentation must render");
    assert.deepEqual(await fs.readFile(filePath), original);
    assert.equal(app.files.size, 0, "No vault copies or summary notes during opening");
  } finally {
    plugin.onunload(); (plugin as unknown as { unload(): void }).unload();
    for (const leaf of app.leaves) { await leaf.view.onClose(); leaf.view.unload(); leaf.element.remove(); }
    await dom.happyDOM.abort();
  }
});
