# External Office document preview (development build)

Open a local `file://` link to a Word (`.docx`), PowerPoint (`.pptx`), or Excel
(`.xlsx`) document in the Obsidian pane where the link was clicked. The source is
read in place, without a vault import, duplicate document, or symlink. Existing
in-vault readers remain registered under their original view types.

## Installation and use

1. Back up your installed plugin, then place `main.js`, `manifest.json`, and
   `styles.css` from this development build in `.obsidian/plugins/word-reader/`.
   Reload the plugin or restart Obsidian. The upstream plugin ID and version are
   retained, so this replaces the upstream installation rather than installing
   alongside it. An upstream update may overwrite these development files.
2. Add ordinary Markdown links with absolute local file URIs:

   ```markdown
   [Lecture slides](<file:///Users/you/Documents/Courses/Lecture%20slides.pptx>)
   [Reading handout](<file:///Users/you/Documents/Courses/Reading%20handout.docx>)
   [Course workbook](<file:///Users/you/Documents/Courses/Course%20workbook.xlsx>)
   ```

3. Click a link in reading mode or Live Preview. Use Ctrl/Cmd-click when editing
   raw Markdown source. The clicked pane becomes the appropriate Office reader.
4. On first use, allow access to the containing directory or choose an ancestor
   directory covering your materials. Canceling preserves the navigation note.

Permissions are **separate for DOCX, PPTX, and XLSX**, and are stored on the current
device. An existing PPTX grant does not silently allow Word or Excel. The existing
PPTX storage key is retained for compatibility. Use **Forget external office
folder permissions** to clear grants for all three formats on this device.

## Retained reading features

| Format | Reader behavior |
| --- | --- |
| DOCX | Existing local document rendering, headings/outline, search, zoom, reading position, text/Markdown copying, and summary creation. |
| PPTX | Existing slide rendering, thumbnails, slide navigation, search, speaker notes, zoom, fullscreen, copying, and summary creation. |
| XLSX | Existing bounded virtual grid, worksheet navigation, search, named ranges, selection, text/formula copying, Markdown/TSV export, cached formula display, and summary creation. |

Rendering compatibility remains that of the upstream engines. This is a read-only
viewer: it neither edits Office documents nor recalculates workbook formulas.
Legacy `.doc`, `.xls`, and `.ppt`, macro-enabled formats, and other extensions
are not added to the external-link feature. Original vault legacy-format fallback
behavior is retained.

## Summaries and source recovery

An explicit summary action creates Markdown under `Office Reader Notes/` inside
the vault. A suffix derived from the source URI distinguishes same-named files.
Existing summary notes are opened without overwriting them. Nothing is written
beside the external Office document. Original vault summaries keep their existing
same-name behavior.

Each opened external source is checked every 1.5 seconds. Replacement, deletion,
or lost folder access clears stale content and offers recovery; reload checks
access and refreshes source metadata before reading. Closing a view disposes its
timer and reader resources. Workspace/reading-state files retain local source
identities, so moved files may need links to be updated. External files and folder
grants must exist separately on each device; this feature does not sync them.

## Local access boundaries

- Desktop-only, explicit local `file://` links. Web URLs, remote/UNC hosts, URL
  queries/fragments, credentials, and malformed paths are rejected.
- Folder boundaries use canonical paths and directory boundaries. File extension
  and per-format grants are checked again for the canonical source.
- Reads use bounded read-only descriptors, no-follow protection where available,
  and before/after file identity, size, and timestamp comparisons. The shared
  external input ceiling is 512 MiB; package-specific limits remain in force.
- All formats retain upstream OOXML package checks and rendering limits. There
  are no document uploads, network downloads, telemetry, or new remote services.
- Device workspace/reading-state storage may contain the source URI. Diagnostics
  continue to omit absolute paths and document contents.

## Validation

Run `npm run check` for lint, typecheck, unit/integration tests, ZIP production
smoke, build, performance, lifecycle stress, dependency inventory, and source/bundle
security checks. The current complete run passed **134 tests, with zero failures
and one optional real-PPTX test skipped**.

The new DOCX and XLSX integration tests use the actual sessions, document parsers,
and renderers with a simulated Obsidian host and DOM. They verify click routing,
canceled grants, independent format permissions, no vault import, unchanged source
bytes, summaries, workspace restore, source replacement/deletion, permission
revocation, and original vault rendering/summary/reload behavior. The Word reload
path also handles a vault reader returning the same ArrayBuffer again.

The source-link tests cover Live Preview spans and early mouse events for all
three formats. A separate regression test models Obsidian's actual CommonJS
loader, where `require` is supplied but `module.require` is absent.

The build stays within the existing 512 KiB bundle allowance. The repository's
dependency audit is an inventory/license/artifact check; existing upstream npm
advisory findings are not claimed to be fixed by this feature.

## Native desktop result (2026-10-16)

The installed development build was tested in actual macOS Obsidian 1.14.4 using
ordinary links in a temporary navigation note:

| Format | Reading mode | Live Preview |
| --- | --- | --- |
| DOCX | Passed: same pane, rendered document text | Passed: same pane, rendered document text |
| XLSX | Passed: same pane, rendered worksheet cells | Passed: same pane, rendered worksheet cells |
| PPTX | Passed: same pane, rendered slide content | Passed: same pane, rendered slide content |

All six clicks were intercepted and the expected external view and source were
present, with no reader error. The three original document hashes were unchanged;
no extra files appeared beside them. The temporary navigation note, test view,
format-specific test-directory grants, and source fixtures were removed afterward.
The user's prior plugin build was backed up before installation. Existing grants
and the original workspace panes were retained.

## Remaining acceptance work

Further compatibility testing should cover Windows/Linux, popout windows,
multiple Obsidian versions, large real documents, restart/restore, complex Word
layouts, and spreadsheet drawings/charts. Automated tests and a local macOS test
do not establish universal Office layout fidelity or cross-platform acceptance.
