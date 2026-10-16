import type { ItemView } from "obsidian";
import type WordReaderPlugin from "../main";
import { PptxSession } from "./PptxSession";

/** Shared delegation keeps vault and external hosts on the same reader implementation. */
export function createPptxReader(host: ItemView, plugin: WordReaderPlugin): PptxSession {
  return new PptxSession(host, plugin);
}
