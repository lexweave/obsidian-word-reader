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
| A source interface and shared PPTX reading session | Reuse the existing renderer for vault and external files without pretending an external file is a vault `TFile`. |
| A dedicated external PPTX `ItemView` | Open the file in the clicked pane and retain its URI in navigation and workspace state. |
| Link handling for reading mode, Live Preview, and source Ctrl/Cmd-click | Keep the reading workflow inside Obsidian while preserving normal source editing and unrelated links. |
| Capture of Live Preview spans and early mouse events | Live Preview uses CodeMirror spans rather than ordinary anchors, and its native handler can open links before a later click handler. |
| Lazy desktop module loading through Obsidian's supplied `require` | The actual plugin loader does not provide `module.require`; filesystem access must work in the real host as well as tests. |
| Explicit folder grants stored on each device | Limit the feature to a user-selected local directory without syncing that machine's grants through plugin settings. |
| Canonical path checks, bounded read-only access, and source revision checks | Respect folder boundaries, limit resource use, and reject files replaced during a read. |
| Source change detection and recoverable reload | Avoid stale presentation content when a local synchronization tool replaces or moves a file. |
| Explicit summary creation inside `Office Reader Notes/` | Allow useful Markdown notes without writing beside the external presentation; distinguish sources with the same filename. |
| Integration, filesystem, URL, and loader regression tests | Cover both the original vault workflow and the new external source path. |
| A 512 KiB bundle budget instead of 500 KiB | Accommodate the added code; rendering, memory, and lifecycle budgets remain unchanged. |

## Current scope

- Desktop-only external `.pptx` preview from explicit local `file://` links.
- Read the source in place: no presentation import, duplicate copy, or symlink.
- Retain the existing presentation renderer, navigation, search, speaker notes,
  zoom, fullscreen, copy, reload, and original vault file views.
- No remote downloads, telemetry, document upload, or Office editing.
- Folder grants and source availability are local to each device. This feature
  does not synchronize external documents or make local links portable.
- External DOCX and XLSX are not yet supported. Existing in-vault DOCX and XLSX
  support remains available. Legacy DOC, XLS, and PPT are not added by this change.

Installation details, permission behavior, file-access boundaries, and the
acceptance checklist are documented in [external PPTX usage](docs/external-pptx.md).
The plugin retains upstream ID `word-reader` and version `3.5.0`, so this build
replaces the upstream plugin rather than installing beside it. Back up an existing
installation; an upstream update may replace the fork's installed files.

## Validation and its limits

The latest full `npm run check` passed with **131 tests passed, zero failures,
and one optional real-file test skipped**. That command includes lint, typecheck,
ZIP production smoke checks, build, performance, lifecycle stress, dependency
inventory, and source/bundle security scanning. An earlier run also passed the
optional real-file integration test with an external course presentation and
verified that its bytes were unchanged.

The installed development build was then tested in actual macOS Obsidian 1.14.4:
a normal Live Preview click on an external presentation changed the clicked leaf
to the external reader and displayed slide text and SVG content. The user also
confirmed successful PPTX opening. This is evidence for that desktop workflow;
Windows, Linux, popout behavior, and the entire manual acceptance matrix are not
claimed as fully validated.

The repository's dependency audit checks inventory, licenses, and artifact size.
Existing upstream npm advisory findings were not remediated by this feature;
passing that check does not mean the dependencies have no reported advisories.

## Possible next steps: external Word and Excel

These are proposals for later work, not features delivered by the current branch
or commitments to a release date.

| Stage | Proposed work | Completion criteria |
| --- | --- | --- |
| Shared external source handling | Generalize PPTX path validation, folder grants, file identity, and lifecycle behavior where useful. | Keep existing vault readers working and add format-specific tests before expanding link routing. |
| External Word (`.docx`) | Adapt the current Word reader to an external source and open local DOCX links in the clicked pane. | Preserve outline, search, zoom, reading position, copy, and explicit vault-only summary creation; no writes to the source. |
| External Excel (`.xlsx`) | Adapt the workbook reader to an external source and open local XLSX links in the clicked pane. | Preserve worksheet navigation, bounded virtualization, search, selection/export, cached formula display, and source replacement recovery. |
| Desktop compatibility | Exercise each implemented format in supported Obsidian versions on macOS, Windows, and Linux. | Test reading mode, Live Preview, source modifiers, multiple panes, popouts, restart, missing files, and canceled or revoked folder grants. |

Word and Excel should reuse the established renderers and read-only package
validation. A file link alone must not silently expand folder access. Legacy
formats, macro-enabled documents, editing, and automatic cloud transfer require
separate decisions and are outside this plan.

## Development status

Work is on [`feat/external-pptx`](https://github.com/lexweave/obsidian-word-reader/tree/feat/external-pptx)
and reviewed through [draft PR #1](https://github.com/lexweave/obsidian-word-reader/pull/1)
in this fork. No upstream pull request or new public release has been published.
