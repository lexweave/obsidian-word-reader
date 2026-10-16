/** Recognize only explicit local PPTX links; never treat web/UNC URLs as local. */
export function parseExternalPptxLink(href: string): string | null {
  try {
    const url = new URL(href);
    const path = decodeURIComponent(url.pathname);
    if (url.protocol !== "file:" || (url.hostname && url.hostname !== "localhost") ||
        url.username || url.password || url.search || url.hash ||
        [...path].some((c) => c.charCodeAt(0) < 32) || !path.toLowerCase().endsWith(".pptx")) return null;
    return url.href;
  } catch { return null; }
}

/** Source-mode Ctrl/Cmd-click, when CodeMirror has no rendered anchor element. */
export function externalPptxAtPosition(line: string, position: number): string | null {
  const links = /\[[^\]\n]*\]\(\s*(?:<(file:\/\/[^>\n]+)>|(file:\/\/[^\s)]+))\s*\)/gi;
  for (const match of line.matchAll(links)) {
    const start = match.index ?? 0;
    if (position >= start && position < start + match[0].length) {
      return parseExternalPptxLink(match[1] ?? match[2]);
    }
  }
  return null;
}
