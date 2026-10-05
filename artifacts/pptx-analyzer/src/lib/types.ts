export type FindingKind = "مشكلة" | "ملاحظة" | "توصية";
export type FindingSeverity = "مرتفع" | "متوسط" | "منخفض";

export interface SlideRecord {
  number: number;
  title: string;
  text: string;
}

export interface Finding {
  id: string;
  slideNumber: number;
  section: string;
  kind: FindingKind;
  severity: FindingSeverity;
  text: string;
  evidence: string;
}

export interface UserNote {
  id: string;
  text: string;
  slideNumber?: number;
  quote?: string;
  imageDataUrl?: string;
  createdAt: string;
}

export interface PresentationAnalysis {
  id: string;
  fileName: string;
  createdAt: string;
  slideCount: number;
  slides: SlideRecord[];
  findings: Finding[];
  notes: UserNote[];
}

export interface ExtractedSlide {
  number: number;
  title: string;
  text: string;
}
