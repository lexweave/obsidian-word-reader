import assert from "node:assert/strict";
import test from "node:test";
import * as fs from "node:fs/promises";
import { constants, type PathLike } from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { parseExternalPptxLink, externalPptxAtPosition } from "../src/external/links";
import { isWithinRoot, readPresentation, revision } from "../src/external/fileAccess";
import { PptxPackage } from "../src/pptx/pptxPackage";
import { createMinimalPptx } from "./pptxFixture";

void test("external links accept encoded local PPTX only, including Windows drive URLs", () => {
  for (const uri of ["file:///tmp/a%20%28test%29%20%E8%AF%BE%E4%BB%B6.PPTX", "file:///C:/Study/a.pptx", "file://localhost/tmp/a.pptx"]) {
    assert.ok(parseExternalPptxLink(uri));
  }
  for (const uri of ["https://example.org/a.pptx", "file://server/share/a.pptx", "file:///tmp/a.docx", "file:///tmp/a.pptx?url=remote", "file:///tmp/a.pptx#other", "file:///tmp/a%00.pptx", "file:///tmp/%QQ.pptx", "javascript:alert(1)"]) {
    assert.equal(parseExternalPptxLink(uri), null, uri);
  }
});

void test("source-mode links resolve only the clicked span and preserve URI escaping", () => {
  const line = "prefix [Slides](<file:///tmp/a%20b.pptx>) and [Web](https://example.org/a.pptx)";
  assert.equal(externalPptxAtPosition(line, 10), "file:///tmp/a%20b.pptx");
  assert.equal(externalPptxAtPosition(line, 1), null);
  assert.equal(externalPptxAtPosition(line, line.length - 4), null);
});

void test("folder permissions require a real directory boundary on POSIX and Windows", () => {
  assert.equal(isWithinRoot("slides/a.pptx", "/"), true);
  assert.equal(isWithinRoot("../other/a.pptx", "/"), false);
  assert.equal(isWithinRoot("..", "/"), false);
  assert.equal(isWithinRoot("/absolute/a.pptx", "/"), false);
  assert.equal(isWithinRoot("..\\other\\a.pptx", "\\"), false);
  assert.equal(isWithinRoot("D:\\other\\a.pptx", "\\"), false);
});

void test("external file read parses with the existing PPTX engine and leaves no copy or modification", async () => {
  const folder = await fs.mkdtemp(path.join(os.tmpdir(), "office-reader-external-"));
  try {
    const file = path.join(folder, "课件 (one).pptx");
    const content = Buffer.from(await createMinimalPptx());
    await fs.writeFile(file, content);
    const before = await fs.stat(file);
    const data = await readPresentation(fs, file, revision(before));
    const presentation = await PptxPackage.load(data);
    assert.equal(presentation.slideCount, 1);
    assert.equal((await presentation.getSlideMetadata(0)).title, "Hello PPTX");
    presentation.clearCaches();
    assert.deepEqual(await fs.readFile(file), content);
    assert.equal(revision(await fs.stat(file)), revision(before));
    assert.deepEqual(await fs.readdir(folder), [path.basename(file)]);
    assert.ok(parseExternalPptxLink(pathToFileURL(file).href));
    await fs.writeFile(file, "replacement");
    await assert.rejects(readPresentation(fs, file, revision(before)), /changed/);
    await fs.unlink(file);
    await assert.rejects(readPresentation(fs, file, revision(before)));
  } finally { await fs.rm(folder, { recursive: true, force: true }); }
});

void test("replacements during a read are rejected and descriptors are closed", async () => {
  const folder = await fs.mkdtemp(path.join(os.tmpdir(), "office-reader-race-"));
  try {
    const file = path.join(folder, "slides.pptx");
    await fs.writeFile(file, "original");
    const expected = revision(await fs.stat(file));
    const io = { open: fs.open, stat: async (p: PathLike) => {
      await fs.writeFile(file, "replacement changed size");
      return fs.stat(p);
    } };
    // Structural overloads on fs.stat need the exact library interface.
    await assert.rejects(readPresentation(io as Pick<typeof fs, "open" | "stat">, file, expected), /changed/);
    const link = path.join(folder, "link.pptx");
    await fs.symlink(file, link);
    if (constants.O_NOFOLLOW) {
      await assert.rejects(readPresentation(fs, link, revision(await fs.stat(file)), constants.O_RDONLY | constants.O_NOFOLLOW));
    }
  } finally { await fs.rm(folder, { recursive: true, force: true }); }
});
