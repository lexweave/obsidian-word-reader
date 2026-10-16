/** Recognize only explicit local Office links; never treat web/UNC URLs as local. */
export function parseExternalOfficeLink(href: string): string | null {
  try {
    const url = new URL(href);
    const path = decodeURIComponent(url.pathname);
    if (url.protocol !== "file:" || (url.hostname && url.hostname !== "localhost") ||
        url.username || url.password || url.search || url.hash ||
        [...path].some((c) => c.charCodeAt(0) < 32) || !externalOfficeFormat(url.href)) return null;
    return url.href;
  } catch { return null; }
}

/** Source-mode Ctrl/Cmd-click, when CodeMirror has no rendered anchor element. */
export function externalOfficeAtPosition(line: string, position: number): string | null {
  const links = /\[[^\]\n]*\]\(\s*(?:<(file:\/\/[^>\n]+)>|(file:\/\/[^\s)]+))\s*\)/gi;
  for (const match of line.matchAll(links)) {
    const start = match.index ?? 0;
    if (position >= start && position < start + match[0].length) {
      return parseExternalOfficeLink(match[1] ?? match[2]);
    }
  }
  return null;
}

export type ExternalOfficeFormat = "pptx" | "docx" | "xlsx";
export function externalOfficeFormat(uri: string): ExternalOfficeFormat | null {
  try {
    const extension = decodeURIComponent(new URL(uri).pathname).split(".").pop()?.toLowerCase();
    return extension === "pptx" || extension === "docx" || extension === "xlsx" ? extension : null;
  } catch { return null; }
}
// Keep PPTX helpers for existing callers without broadening their contract.
export function parseExternalPptxLink(href: string): string | null {
  const uri = parseExternalOfficeLink(href);
  return uri && externalOfficeFormat(uri) === "pptx" ? uri : null;
}
export function externalPptxAtPosition(line: string, position: number): string | null {
  const uri = externalOfficeAtPosition(line, position);
  return uri && externalOfficeFormat(uri) === "pptx" ? uri : null;
}
