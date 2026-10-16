import assert from "node:assert/strict";
import test from "node:test";
import * as fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { pathToFileURL } from "node:url";
import type { App, PluginManifest, WorkspaceLeaf, TFile as VaultFile } from "obsidian";
import WordReaderPlugin from "../src/main";
import { ExternalDocxView, ExternalXlsxView } from "../src/ExternalPptxView";
import { WordView } from "../src/WordView";
import { XlsxView } from "../src/XlsxView";
import { TestApp, TestLeaf, FileView, TFile, dom } from "./helpers/obsidianMock";
import { createLargeDocx } from "./performanceFixtures";
import { createRichXlsx } from "./xlsxFixture";
import { parseExternalOfficeLink, externalOfficeAtPosition } from "../src/external/links";

async function until(check: () => boolean): Promise<void> {
  for (let i = 0; i < 500; i++) {
    if (check()) return;
    await new Promise((resolve) => window.setTimeout(resolve, 10));
  }
  assert.ok(check(), "Office view did not reach expected state");
}

void test("external Office links accept DOCX/XLSX and reject legacy, macro and remote files", () => {
  for (const extension of ["docx", "xlsx", "PPTX"]) {
    const uri = `file:///tmp/Study%20%28one%29.${extension}`;
    assert.ok(parseExternalOfficeLink(uri));
    const line = `[Material](<${uri}>)`;
    assert.equal(externalOfficeAtPosition(line, 3), uri);
  }
  for (const uri of ["file:///tmp/a.doc", "file:///tmp/a.xls", "file:///tmp/a.xlsm", "https://example.org/a.docx", "file://server/a.xlsx", "file:///tmp/a.xlsx?key=1", "file:///tmp/a.docx#part", "file:///tmp/a%00.docx"]) assert.equal(parseExternalOfficeLink(uri), null);
});

for (const format of ["docx", "xlsx"] as const) {
  void test(`external ${format}: click, separate permission, rendering, summary, replacement, restore and vault regression`, async () => {
    const folder = await fs.mkdtemp(path.join(os.tmpdir(), `office-reader-${format}-`));
    const bytes = format === "docx" ? await createLargeDocx(1) : await createRichXlsx();
    const filePath = path.join(folder, `课件 (one).${format}`);
    await fs.writeFile(filePath, Buffer.from(bytes));
    const app = new TestApp();
    const plugin = new WordReaderPlugin(app as unknown as App, {} as PluginManifest);
    const views: Array<ExternalDocxView | ExternalXlsxView | WordView | XlsxView> = [];
    const sourceText = format === "docx" ? "Performance page 1" : "Merged title";
    try {
      await plugin.onload();
      // A prior PPTX grant must not silently authorize Word or Excel.
      app.storage.set("word-reader-external-pptx-folders", [folder]);
      const uri = pathToFileURL(filePath).href;
      assert.equal(await plugin.externalOfficeAccess.require(uri, false), null);
      const leaf = new TestLeaf(app);
      const other = new TestLeaf(app);
      const anchor = dom.document.createElement("a"); anchor.href = uri;
      leaf.view.contentEl.appendChild(anchor as unknown as Node);
      const click = () => anchor.dispatchEvent(new dom.MouseEvent("click", { bubbles: true, cancelable: true, button: 0 }));
      click(); await until(() => dom.document.querySelectorAll(".modal button").length === 2);
      (dom.document.querySelector(".modal button") as unknown as HTMLButtonElement).click();
      await until(() => !dom.document.querySelector(".modal"));
      await new Promise((resolve) => window.setTimeout(resolve, 20));
      assert.equal(leaf.changes, 0);
      click(); await until(() => dom.document.querySelectorAll(".modal button").length === 2);
      (dom.document.querySelectorAll(".modal button")[1] as unknown as HTMLButtonElement).click();
      await until(() => leaf.view.getViewType() === `external-${format}-reader-view` && leaf.view.contentEl.textContent?.includes(sourceText) === true && app.active === leaf).catch(() => { throw new Error(`Fixture render failed: ${leaf.view.contentEl.textContent}`); });
      const view = leaf.view as unknown as ExternalDocxView | ExternalXlsxView; views.push(view);
      assert.equal(other.changes, 0); assert.ok(app.active === leaf);
      assert.equal(view.file, null); assert.equal(view.hasSource(), true);
      assert.equal(app.files.size, 0);
      assert.equal(await plugin.externalOfficeAccess.allowed(filePath, format === "docx" ? "xlsx" : "docx"), false);
      await view.createSummaryNote();
      const summary = [...app.data.entries()].find(([name]) => name.endsWith(".md"));
      assert.ok(summary?.[0].startsWith("Office Reader Notes/"));
      assert.ok(typeof summary?.[1] === "string");
      assert.ok(summary[1].includes(pathToFileURL(await fs.realpath(filePath)).href));
      assert.deepEqual(await fs.readFile(filePath), Buffer.from(bytes));
      assert.deepEqual(await fs.readdir(folder), [path.basename(filePath)]);
      const state = view.getState(); await view.onClose(); view.unload(); views.pop();
      const restoredLeaf = new TestLeaf(app);
      const restored = format === "docx" ? new ExternalDocxView(restoredLeaf as unknown as WorkspaceLeaf, plugin) : new ExternalXlsxView(restoredLeaf as unknown as WorkspaceLeaf, plugin);
      restoredLeaf.view = restored as unknown as typeof restoredLeaf.view; views.push(restored);
      await restored.onOpen(); await restored.setState(state, { history: false });
      assert.ok(restored.contentEl.textContent?.includes(sourceText));
      const replacement = path.join(folder, `replacement.${format}`);
      await fs.writeFile(replacement, Buffer.from(bytes)); await fs.rename(replacement, filePath);
      await until(() => restored.contentEl.querySelector(".office-reader-message") !== null);
      await restored.reload();
      assert.ok(restored.contentEl.textContent?.includes(sourceText));
      await fs.unlink(filePath); await restored.reload();
      assert.equal(restored.contentEl.textContent?.includes(sourceText), false, "Missing source must clear stale content");
      await fs.writeFile(filePath, Buffer.from(bytes)); await restored.reload();
      assert.ok(restored.contentEl.textContent?.includes(sourceText));
      // Revoking format permission must also clear stale data on reload.
      app.storage.set(`word-reader-external-${format}-folders`, null); await restored.reload();
      assert.equal(restored.contentEl.textContent?.includes(sourceText), false);
      await restored.onClose(); restored.unload(); views.pop();
      const nativeLeaf = new TestLeaf(app);
      const native = format === "docx" ? new WordView(nativeLeaf as unknown as WorkspaceLeaf, plugin) : new XlsxView(nativeLeaf as unknown as WorkspaceLeaf, plugin);
      nativeLeaf.view = native as unknown as typeof nativeLeaf.view; views.push(native);
      const file = new TFile(`vault-${format}.${format}`); file.stat.size = bytes.byteLength;
      app.files.set(file.path, file); app.data.set(file.path, bytes);
      native.file = file as unknown as VaultFile;
      await native.onOpen(); await native.onLoadFile(file as unknown as VaultFile);
      assert.ok(native instanceof FileView);
      assert.ok(native.contentEl.textContent?.includes(sourceText), `Vault fixture render: ${native.contentEl.textContent}`);
      await native.createSummaryNote();
      assert.ok(app.files.has(`vault-${format}.md`));
      await native.reload();
      await until(() => native.contentEl.textContent?.includes(sourceText) === true);
      await native.onUnloadFile();
    } finally {
      for (const view of views) { await view.onClose(); view.unload(); }
      plugin.unload(); for (const leaf of app.leaves) leaf.element.remove();
      await fs.rm(folder, { recursive: true, force: true });
    }
  });
}
