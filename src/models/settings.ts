export interface TitleSettings {
  fontName: string;
  fontSize: number;
  distanceFromBottom: number;
  titleText: string;
  autoGroup: boolean;
  centerTitle: boolean;
  excludeTextShapes: boolean;
}

export interface LabelSettings {
  offsetX: number;
  offsetY: number;
  template: string;
  fontName: string;
  fontSize: number;
  bold: boolean;
  startIndex: number;
  autoUpdateIndex: boolean;
}

export interface ArrangeSettings {
  sortType: number; // 0 = by position, 1 = by selection order
  alignType: number; // 0 = max column width, 1 = uniform height, 2 = waterfall
  colNum: number;
  colSpace: number;
  rowSpace: number;
  imgWidth: string;
  imgHeight: string;
  excludeTextShapes: boolean;
}

export interface CodeBlockSettings {
  language: string;
  darkTheme: boolean;
}

export interface ExportSettings {
  format: string;
  dpi: number;
}

export interface AppSettings {
  title: TitleSettings;
  label: LabelSettings;
  arrange: ArrangeSettings;
  codeBlock: CodeBlockSettings;
  export: ExportSettings;
}

export const DEFAULT_SETTINGS: AppSettings = {
  title: {
    fontName: "微软雅黑",
    fontSize: 14,
    distanceFromBottom: 0,
    titleText: "图片标题",
    autoGroup: false,
    centerTitle: true,
    excludeTextShapes: true,
  },
  label: {
    offsetX: -20,
    offsetY: -7,
    template: "A",
    fontName: "Arial",
    fontSize: 12,
    bold: true,
    startIndex: 1,
    autoUpdateIndex: true,
  },
  arrange: {
    sortType: 0,
    alignType: 0,
    colNum: 3,
    colSpace: 10,
    rowSpace: 25,
    imgWidth: "",
    imgHeight: "",
    excludeTextShapes: true,
  },
  codeBlock: {
    language: "python",
    darkTheme: true,
  },
  export: {
    format: "PNG",
    dpi: 300,
  },
};
