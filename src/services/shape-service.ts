import { ShapeInfo } from "../models/types";

export class ShapeService {
  /**
   * Get selected shapes from the active slide.
   * Filters out text boxes and auto shapes if excludeText is true.
   */
  async getSelectedShapes(excludeText: boolean = false): Promise<ShapeInfo[]> {
    return PowerPoint.run(async (context) => {
      // Use getSelectedShapes() to only get user-selected shapes (PowerPointApi 1.5+)
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type");
      await context.sync();

      let result = shapes.items.map((shape) => ({
        id: shape.id,
        name: shape.name,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        type: shape.type as string,
      }));

      if (excludeText) {
        // PowerPoint API ShapeType: Image, GeometricShape, Group, Line, Table, etc.
        // Keep only Image and Group shapes when excluding text
        result = result.filter(
          (s) => s.type === "Image" || s.type === "Group"
        );
      }

      return result;
    });
  }

  /**
   * Add a text box to the current slide and return its shape ID.
   */
  async addTextBox(
    left: number,
    top: number,
    width: number,
    height: number,
    text: string,
    options?: {
      fontName?: string;
      fontSize?: number;
      fontColor?: string;
      bold?: boolean;
      italic?: boolean;
      alignment?: string;
      wordWrap?: boolean;
      autoSize?: boolean;
    }
  ): Promise<string> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(text, {
        left,
        top,
        width,
        height,
      });
      textBox.load("id,name");

      if (options) {
        const font = textBox.textFrame.textRange.font;
        if (options.fontName) font.name = options.fontName;
        if (options.fontSize) font.size = options.fontSize;
        if (options.fontColor) font.color = options.fontColor;
        if (options.bold !== undefined) font.bold = options.bold;
        if (options.italic !== undefined) font.italic = options.italic;
        if (options.alignment) {
          textBox.textFrame.textRange.paragraphFormat.horizontalAlignment =
            options.alignment as PowerPoint.ParagraphHorizontalAlignment;
        }
        if (options.autoSize) {
          textBox.textFrame.autoSizeSetting =
            "AutoSizeShapeToFitText" as PowerPoint.ShapeAutoSize;
        }
        if (options.wordWrap === false) {
          textBox.textFrame.wordWrap = false;
        }
      }

      await context.sync();
      return textBox.id;
    });
  }

  /**
   * Add a title text box for a given shape.
   * isBottom=true: title below image; false: title above image.
   */
  async addTitleForShape(
    shape: ShapeInfo,
    titleText: string,
    isBottom: boolean,
    options: {
      fontName: string;
      fontSize: number;
      distance: number;
      centerTitle: boolean;
    }
  ): Promise<string> {
    const { fontName, fontSize, distance, centerTitle } = options;
    const titleHeight = fontSize * 2;
    const titleTop = isBottom
      ? shape.top + shape.height + distance
      : shape.top - titleHeight - distance;

    return this.addTextBox(
      shape.left,
      titleTop,
      shape.width,
      titleHeight,
      titleText,
      {
        fontName,
        fontSize,
        alignment: centerTitle ? "Center" : "Left",
        autoSize: true,
      }
    );
  }

  /**
   * Add a label text box at the specified offset from a shape.
   */
  async addLabelForShape(
    shape: ShapeInfo,
    label: string,
    options: {
      fontName: string;
      fontSize: number;
      offsetX: number;
      offsetY: number;
      bold: boolean;
    }
  ): Promise<string> {
    const { fontName, fontSize, offsetX, offsetY, bold } = options;

    return this.addTextBox(
      shape.left + offsetX,
      shape.top + offsetY,
      0, // Initial width - auto-sized
      fontSize * 2,
      label,
      {
        fontName,
        fontSize,
        bold,
        wordWrap: false,
        autoSize: true,
      }
    );
  }

  /**
   * Set shape position.
   */
  async setShapePosition(
    shapeId: string,
    left: number,
    top: number
  ): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      shape.left = left;
      shape.top = top;
      await context.sync();
    });
  }

  /**
   * Set shape dimensions.
   */
  async setShapeDimensions(
    shapeId: string,
    width: number,
    height: number
  ): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      shape.width = width;
      shape.height = height;
      await context.sync();
    });
  }

  /**
   * Get shape font properties (for format copy).
   */
  async getShapeFont(shapeId: string): Promise<{
    name: string;
    size: number;
    bold: boolean;
    italic: boolean;
    color: string;
    underline: boolean;
  } | null> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      const font = shape.textFrame.textRange.font;
      font.load("name,size,bold,italic,color,underline");
      await context.sync();
      return {
        name: font.name ?? "",
        size: font.size ?? 12,
        bold: font.bold ?? false,
        italic: font.italic ?? false,
        color: font.color ?? "#000000",
        underline: font.underline === "Single",
      };
    });
  }

  /**
   * Set shape font properties (for format paste).
   */
  async setShapeFont(
    shapeId: string,
    fontProps: {
      name?: string;
      size?: number;
      bold?: boolean;
      italic?: boolean;
      color?: string;
      underline?: boolean;
    }
  ): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      const font = shape.textFrame.textRange.font;
      if (fontProps.name) font.name = fontProps.name;
      if (fontProps.size) font.size = fontProps.size;
      if (fontProps.bold !== undefined) font.bold = fontProps.bold;
      if (fontProps.italic !== undefined) font.italic = fontProps.italic;
      if (fontProps.color) font.color = fontProps.color;
      if (fontProps.underline !== undefined) {
        font.underline = fontProps.underline
          ? ("Single" as PowerPoint.ShapeFontUnderlineStyle)
          : ("None" as PowerPoint.ShapeFontUnderlineStyle);
      }
      await context.sync();
    });
  }

  /**
   * Delete a shape by ID.
   */
  async deleteShape(shapeId: string): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.shapes.getItem(shapeId).delete();
      await context.sync();
    });
  }
}

export const shapeService = new ShapeService();
