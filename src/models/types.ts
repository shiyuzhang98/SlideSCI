export interface ShapeInfo {
  id: string;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  type: string;
}

export interface FontProperties {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  underline: boolean;
}

export interface CopiedPosition {
  centerX: number;
  centerY: number;
}

export interface CopiedDimensions {
  width: number;
  height: number;
}

export interface CropSettings {
  cropLeft: number;
  cropRight: number;
  cropTop: number;
  cropBottom: number;
  originalHeight: number;
  croppedHeight: number;
}

export interface MarkdownSegment {
  content: string;
  isCodeBlock: boolean;
  isTable: boolean;
  isMathBlock: boolean;
  isBlockQuote: boolean;
  language?: string;
}

export interface ImageGroup {
  shapes: ShapeInfo[];
  minTop: number;
  maxBottom: number;
}

export type LabelTemplate =
  | "A" | "a" | "A)" | "a)" | "(A)" | "(a)"
  | "1" | "1)"
  | "Ⅰ" | "Ⅰ)"
  | "①" | "①)"
  | "一" | "一)";

export type ArrangeMode = "maxColumnWidth" | "uniformHeight" | "waterfall";

export type SortMode = "byPosition" | "bySelectionOrder";

export interface CodeHighlightToken {
  start: number;
  length: number;
  type: string;
  color: string;
}
