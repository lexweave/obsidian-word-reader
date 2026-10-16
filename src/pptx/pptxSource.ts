import type { PptxReaderText } from "./pptxI18n";
import type { PptxSlideMetadata } from "./pptxMetadata";
import type { BinaryOfficeSource, OfficeSourceFile } from "../reader/officeSource";

export type PptxSourceFile = OfficeSourceFile;

/** A document source, without pretending an external file is a vault TFile. */
export interface PptxSource extends BinaryOfficeSource {
  createSummary(slides: readonly PptxSlideMetadata[], page: number, text: PptxReaderText): Promise<void>;
}
