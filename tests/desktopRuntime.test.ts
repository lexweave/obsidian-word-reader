import assert from "node:assert/strict";
import test from "node:test";
import { build } from "esbuild";
import { runInNewContext } from "node:vm";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as url from "node:url";

void test("desktop helpers work with Obsidian's module.exports-only loader and lazy require", async () => {
  const bundle = await build({ entryPoints: ["src/desktopRuntime.ts"], bundle: true,
    platform: "node", format: "cjs", external: ["obsidian", "electron"], write: false });
  const platform = { isDesktopApp: true };
  const calls: string[] = [];
  const electron = { shell: {} };
  const modules: Record<string, unknown> = { obsidian: { Platform: platform }, "node:fs/promises": fs,
    "node:path": path, "node:url": url, electron };
  const pluginModule = { exports: {} as Record<string, () => unknown> };
  // Actual Obsidian plugins receive a require argument; module.require is absent.
  runInNewContext(bundle.outputFiles[0].text, { module: pluginModule, exports: pluginModule.exports,
    require: (id: string) => { calls.push(id); if (!(id in modules)) throw new Error("Unexpected module"); return modules[id]; } });
  assert.deepEqual(calls, ["obsidian"], "Desktop modules must be loaded lazily");
  assert.equal(pluginModule.exports.getFileSystemPromises(), fs);
  assert.equal(pluginModule.exports.getPathModule(), path);
  assert.equal(pluginModule.exports.getUrlModule(), url);
  assert.equal(pluginModule.exports.getElectronModule(), electron);
  platform.isDesktopApp = false;
  const before = calls.length;
  assert.throws(() => pluginModule.exports.getFileSystemPromises(), /only available on desktop/);
  assert.equal(calls.length, before, "Mobile guard must precede require");
});
