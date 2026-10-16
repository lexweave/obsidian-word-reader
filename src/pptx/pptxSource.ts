import type { TFile } from "obsidian";
import type { PptxReaderText } from "./pptxI18n";
import type { PptxSlideMetadata } from "./pptxMetadata";
import type { ReaderFileDescriptor } from "../reader/fileRevision";

export interface PptxSourceFile extends ReaderFileDescriptor {
  name: string;
  basename: string;
  stat: { mtime: number; ctime: number; size: number };
}

/** A document source, without pretending an external file is a vault TFile. */
export interface PptxSource {
  readonly file: PptxSourceFile;
  readonly vaultFile?: TFile;
  read(): Promise<ArrayBuffer>;
  refresh(): Promise<void>;
  openExternal(): Promise<void>;
  createSummary(slides: readonly PptxSlideMetadata[], page: number, text: PptxReaderText): Promise<void>;
}
