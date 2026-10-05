import { analyzeSlides } from "./analyze";
import { readPowerPoint } from "./pptx-reader";
import type {
  Finding,
  PresentationAnalysis,
  SlideRecord,
  UserNote,
} from "./types";

export type { Finding, PresentationAnalysis, SlideRecord, UserNote } from "./types";

export async function parsePresentation(
  file: File,
): Promise<PresentationAnalysis> {
  const slides = await readPowerPoint(file);
  return analyzeSlides(file.name, slides);
}
