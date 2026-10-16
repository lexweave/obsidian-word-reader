# Why this fork exists

This is a development fork of [Office Reader by qianwei4712](https://github.com/qianwei4712/obsidian-word-reader).
The upstream project supplies the existing document parsers, rendering engines,
reader features, and read-only design. This fork extends how a document reaches
those readers; it does not replace the upstream project or claim its work as new.
The upstream license and attribution remain in place.

## The problem

In this study workflow, an Obsidian vault contains course notes and weekly
navigation pages and is synced through iCloud. Large lecture presentations and
other teaching materials are stored in a separate local directory. Putting a
second copy in the vault would add storage and synchronization work. Symlinks
would introduce filesystem and cross-device management concerns.

The original Office Reader uses Obsidian vault files (`TFile` and
`vault.readBinary`). A Markdown `file:///...pptx` link to an external presentation
therefore launches the system application instead of the reader. The goal is to
click a course material link and read it in the same Obsidian pane while leaving
the source exactly where it is.

Canvas-to-Obsidian synchronization motivated the workflow, but this plugin feature
is independent of Canvas. It requires no Canvas account, API token, or downloader.

## What changed and why

| Change | Reason |
| --- | --- |
| Source interfaces and shared DOCX, PPTX, and XLSX reading sessions | Reuse existing renderers for vault and external files without pretending an external file is a vault `TFile`. |
| External Office `ItemView` hosts | Open the file in the clicked pane and retain its URI in navigation and workspace state. |
| Link handling for reading mode, Live Preview, and source Ctrl/Cmd-click | Keep the reading workflow inside Obsidian while preserving normal source editing and unrelated links. |
| Capture of Live Preview spans and early mouse events | Live Preview uses CodeMirror spans rather than ordinary anchors, and its native handler can open links before a later click handler. |
| Lazy desktop module loading through Obsidian's supplied `require` | The actual plugin loader does not provide `module.require`; filesystem access must work in the real host as well as tests. |
| Explicit format-specific folder grants stored on each device | Limit the feature to a user-selected local directory without syncing that machine's grants through plugin settings. |
| Canonical path checks, bounded read-only access, and source revision checks | Respect folder boundaries, limit resource use, and reject files replaced during a read. |
| Source change detection and recoverable reload | Avoid stale presentation content when a local synchronization tool replaces or moves a file. |
| Explicit summary creation inside `Office Reader Notes/` | Allow useful Markdown notes without writing beside the external presentation; distinguish sources with the same filename. |
| Integration, filesystem, URL, and loader regression tests | Cover both the original vault workflow and the new external source path. |
| A 512 KiB bundle budget instead of 500 KiB | Accommodate the added code; rendering, memory, and lifecycle budgets remain unchanged. |

## Current scope

- Desktop-only external `.docx`, `.pptx`, and `.xlsx` preview from explicit local `file://` links.
- Read the source in place: no presentation import, duplicate copy, or symlink.
- Retain the existing Word, presentation, and workbook renderers, their reading
  features, and the original vault file views.
- No remote downloads, telemetry, document upload, or Office editing.
- Folder grants and source availability are local to each device. This feature
  does not synchronize external documents or make local links portable.
- External DOCX and XLSX are implemented alongside PPTX. Each format requires its
  own directory grant; existing PPTX grants remain valid for PPTX only. Legacy
  DOC, XLS, and PPT are not added by this change.

Installation details, permission behavior, file-access boundaries, and the
acceptance checklist are documented in [external Office usage](docs/external-office.md).
The plugin retains upstream ID `word-reader` and version `3.5.0`, so this build
replaces the upstream plugin rather than installing beside it. Back up an existing
installation; an upstream update may replace the fork's installed files.

## Validation and its limits

The latest full `npm run check` passed with **134 tests passed, zero failures,
and one optional real-file test skipped**. That command includes lint, typecheck,
ZIP production smoke checks, build, performance, lifecycle stress, dependency
inventory, and source/bundle security scanning. An earlier run also passed the
optional real-file integration test with an external course presentation and
verified that its bytes were unchanged.

The installed development build was tested in actual macOS Obsidian 1.14.4.
On 2026-10-16, all six combinations of DOCX/PPTX/XLSX and reading mode/Live Preview
opened in the clicked pane and rendered the expected content. All three source
hashes were unchanged. Temporary test notes, directory grants, and files were
cleaned up. Earlier, a real course PPTX also rendered successfully, and the user
confirmed successful opening. Windows, Linux, popouts, and the entire manual
acceptance matrix are not claimed as fully validated.

The repository's dependency audit checks inventory, licenses, and artifact size.
Existing upstream npm advisory findings were not remediated by this feature;
passing that check does not mean the dependencies have no reported advisories.

## Word and Excel extension

External Word (`.docx`) and Excel (`.xlsx`) opening are now implemented, rather
than only proposed. Both sessions use a binary document source instead of relying
on a vault `TFile`. Their original `FileView` wrappers preserve vault reads,
commands, summaries, and legacy fallback behavior. Shared external access and
view lifecycle handling cover all three formats.

Word retains its outline, search, zoom, copy, and reading state. Excel retains its
worksheet navigation, bounded grid, search, selection/export, cached formulas,
and workbook summaries. Explicit external summaries always go into the vault.
The Word buffer registration order was also corrected so reloading the same
ArrayBuffer cannot accidentally clear the replacement reader buffer.

## Next steps

These are future validation and maintenance directions, not release commitments:

- Broaden native compatibility testing to Windows, Linux, popouts, and supported
  Obsidian versions.
- Exercise more real DOCX layouts and XLSX drawings, charts, and large workbooks.
- Extend native restart, permission revocation, source replacement, and long-session
  regression coverage.
- Review upstream dependency advisory updates independently of the external-source
  feature.

Legacy formats, macro-enabled files, editing, and automatic cloud transfer remain
outside this change. The development manifest still retains the upstream version;
a separately versioned release is a later packaging decision.

## Development status

Work is on [`feat/external-pptx`](https://github.com/lexweave/obsidian-word-reader/tree/feat/external-pptx)
and reviewed through [draft PR #1](https://github.com/lexweave/obsidian-word-reader/pull/1)
in this fork. No upstream pull request or new public release has been published.
