# External PPTX preview (development build)

The external-source feature now also supports DOCX and XLSX. See the current
[external Office document guide](external-office.md) for all formats, separate
permissions, and updated validation coverage. The notes below describe the
original PPTX implementation and its initial validation.

This fork can open a local `file:///...pptx` link from a Markdown note in the pane
where it was clicked. It reads the original file in place. It does not import the
presentation into the vault, create a presentation copy, create a symbolic link,
or fetch a remote URL. Existing vault PPTX/DOCX/XLSX reading is retained.

## Use

1. Install the development build's three files (`main.js`, `manifest.json`,
   `styles.css`) in `.obsidian/plugins/word-reader/` after backing up the existing
   plugin. Reload the plugin or restart Obsidian. The upstream plugin ID and
   version are retained; this is a fork development build, not an upstream release.
2. Click a normal Markdown link, for example:

   ```markdown
   [Slides](<file:///Users/you/Documents/Lectures/Synesthesia%20slides.pptx>)
   ```

3. The first click requests permission to read PPTX files from the containing
   folder. You may enter a parent folder to cover a collection of presentations.
   Permission is stored in Obsidian's device-local app storage, outside synced
   plugin settings. Canceling keeps the current note open.
4. Reading mode/rendered Live Preview links open directly. For source text in the
   editor, use Ctrl/Cmd-click on the Markdown link, following Obsidian's editing
   convention. Only local PPTX links are intercepted; other formats and web links
   keep their original behavior.

The presentation uses the existing slide renderer, thumbnails, search, zoom,
notes, fullscreen, copy, and reload functionality. The same view host tracks the
clicked pane, rather than blindly opening in the most recently active pane.
A valid source retains its page and view preferences across restore. External
source state is separate from vault file identities.

If a source changes or disappears, the view clears stale content and presents the
existing recovery UI. The reload action checks the current source again. The
view polls only its opened source every 1.5 seconds; its timer belongs to the
view lifecycle and is disposed when the view closes. A restored view without
folder permission displays an explicit open action rather than silently reading.

The **Create presentation note** action is deliberate and writes only Markdown,
under `Office Reader Notes/` in the vault. Its name includes a source identity
suffix to distinguish same-named external files. Existing summary notes are opened,
not overwritten. This action does not write next to the external presentation.

Use **Forget external presentation folder permissions** in the command palette to
clear grants on this device. Other devices need their own files and permission;
this feature does not make external material available through iCloud.

## Boundaries

- Desktop only, `.pptx` only. No additional support for legacy `.ppt` is implied.
- No HTTP/HTTPS downloads, remote file hosts/UNC URLs, URL query strings or fragments.
- Folder boundaries use canonical paths, not string prefixes. Each read rechecks
  the grant. Opening uses a read-only descriptor with no-follow protection where
  available, bounds the byte allocation at 512 MiB, and compares identity/size/
  timestamps before and after the read to reject in-progress replacements.
- It retains the upstream Office package validation and rendering limits.
- Workspace/reading-state files can contain the local file URI. Diagnostics omit
  paths and document contents; no telemetry or network requests are added.
- This is document opening in the current pane, not an inline slide embed inside
  the Markdown document. Native desktop layout/popout behavior still needs manual
  acceptance testing on the user's installed Obsidian build.

## Validation

```sh
npm ci --ignore-scripts
npm run check
npm test -- externalPptxView
```

An optional read-only integration test can render a real local presentation:

```sh
OFFICE_READER_TEST_PPTX='/absolute/path/to/slides.pptx' npm run check
```

The view integration host runs the actual plugin link handler, PPTX views, source
reader, ZIP parser and SVG renderer with a DOM and simulated Obsidian host APIs.
It verifies the clicked pane even when a different pane is active, cancellation,
permissions, no vault copy, source preservation, workspace restore, summary
placement, replacement/deletion recovery, and the original vault FileView path.
It is not a real installed-Obsidian UI test.

The minified bundle allowance is 512 KiB (previously 500 KiB) to accommodate the
bounded external-source and editor-link paths. Runtime rendering, memory, and
resource-residue budgets are unchanged.

## Manual desktop acceptance

1. Back up the installed plugin and enable this development build.
2. In two panes, open the navigation note in one and a different note in the other.
   Click an external PPTX in the navigation pane: only that pane should change.
3. Check reading mode and Live Preview/source Ctrl/Cmd-click, and repeat in a
   popout window. Cancel the initial permission once before granting it.
4. Test search, slide navigation, zoom, reload, notes, fullscreen and summary creation.
5. Close/reopen the view and restart Obsidian. Verify the file URI/page restore.
6. Replace or move a disposable external test PPTX. Verify recovery/reload, and
   ensure neither its original bytes nor the vault contents were changed by opening.
7. Check a vault PPTX, DOCX and XLSX to confirm normal file opening remains intact.

On 2026-10-15, the installed development build was tested in the user's actual
Obsidian 1.14.4 session through the CLI. A normal Live Preview click on Synesthesia
in the existing Perception navigation note prevented both mousedown and click
from reaching the system opener. After granting the existing course directory,
the same leaf became `external-pptx-reader-view` and rendered the real slide text
and SVG content. The external source was present, with no reader error. The
previous plugin files were backed up before replacement. This verifies that
specific desktop path; the remaining acceptance steps above are not all claimed.

This test caught two gaps in the earlier simulated host: Live Preview renders
links as CodeMirror spans, and Obsidian supplies a `require` argument while its
`module` object has no `require` method. Regression tests now cover both cases.
Plain source-mode clicks keep editing behavior; Ctrl/Cmd-click opens the PPTX.


### Local result

The PPTX-stage full `npm run check` passed: **131 passed, 0 failed, 1 skipped**
(the optional real-file test needs `OFFICE_READER_TEST_PPTX`). The earlier run
included the supplied `Synesthesia slides.pptx`, whose original bytes were
unchanged; the actual desktop test above now also verifies that presentation. Read-only ZIP, performance, 400 lifecycle stress
cycles, dependency inventory, and source/bundle security scans passed. The final
bundle is approximately 503 KiB within the 512 KiB bound. The upstream lockfile
also has npm advisory findings (8 reported during setup); the repository's
`dependencies:audit` is an inventory/license/artifact check, not a claim that npm
advisories are absent. This feature does not update those upstream dependencies.
