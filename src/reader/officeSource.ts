import type { TFile } from "obsidian";
import type { ReaderFileDescriptor } from "./fileRevision";
import type { WordReaderText } from "../i18n";
import type { XlsxReaderText } from "../xlsx/xlsxI18n";
import type { XlsxWorkbookSummary } from "../xlsx/xlsxSummaryNote";

export interface OfficeSourceFile extends ReaderFileDescriptor {
  name: string;
  basename: string;
  stat: { mtime: number; ctime: number; size: number };
}
export interface BinaryOfficeSource {
  readonly file: OfficeSourceFile;
  readonly vaultFile?: TFile;
  read(): Promise<ArrayBuffer>;
  refresh(): Promise<void>;
  openExternal(): Promise<void>;
}
export interface DocxSource extends BinaryOfficeSource {
  createSummary(text: WordReaderText): Promise<void>;
}
export interface XlsxSource extends BinaryOfficeSource {
  createSummary(workbook: XlsxWorkbookSummary, sheet: number, range: string,
    markdown: string, text: XlsxReaderText): Promise<void>;
}
