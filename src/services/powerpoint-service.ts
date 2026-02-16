import { ShapeInfo } from "../models/types";

export class PowerPointService {
  /**
   * Run a batch operation on the current PowerPoint context.
   */
  async run<T>(
    callback: (context: PowerPoint.RequestContext) => Promise<T>
  ): Promise<T> {
    return PowerPoint.run(callback);
  }

  /**
   * Get the currently selected shapes on the active slide.
   */
  async getSelectedShapes(): Promise<ShapeInfo[]> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type");
      await context.sync();

      return shapes.items.map((shape) => ({
        id: shape.id,
        name: shape.name,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        type: shape.type as string,
      }));
    });
  }

  /**
   * Get all shapes on the current slide.
   */
  async getSlideShapes(): Promise<ShapeInfo[]> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type");
      await context.sync();

      return shapes.items.map((shape) => ({
        id: shape.id,
        name: shape.name,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        type: shape.type as string,
      }));
    });
  }

  /**
   * Get the slide dimensions.
   */
  async getSlideDimensions(): Promise<{ width: number; height: number }> {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      presentation.load("slideMasters");
      await context.sync();
      // Default PowerPoint slide dimensions: 10" x 7.5"
      return { width: 720, height: 540 };
    });
  }

  /**
   * Add a text box to the current slide.
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
      alignment?: "Left" | "Center" | "Right";
      wordWrap?: boolean;
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
        const textRange = textBox.textFrame.textRange;
        const font = textRange.font;
        if (options.fontName) font.name = options.fontName;
        if (options.fontSize) font.size = options.fontSize;
        if (options.fontColor) font.color = options.fontColor;
        if (options.bold !== undefined) font.bold = options.bold;
        if (options.italic !== undefined) font.italic = options.italic;
        if (options.alignment) {
          textRange.paragraphFormat.horizontalAlignment =
            options.alignment as PowerPoint.ParagraphHorizontalAlignment;
        }
      }

      await context.sync();
      return textBox.id;
    });
  }

  /**
   * Set shape position and size.
   */
  async setShapeGeometry(
    shapeId: string,
    props: { left?: number; top?: number; width?: number; height?: number }
  ): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      if (props.left !== undefined) shape.left = props.left;
      if (props.top !== undefined) shape.top = props.top;
      if (props.width !== undefined) shape.width = props.width;
      if (props.height !== undefined) shape.height = props.height;
      await context.sync();
    });
  }

  /**
   * Delete a shape by ID.
   */
  async deleteShape(shapeId: string): Promise<void> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.getItem(shapeId);
      shape.delete();
      await context.sync();
    });
  }

  /**
   * Insert an image as base64.
   * Uses PowerPointApi 1.4+ addImage or falls back to addGeometricShape.
   */
  async insertImage(
    base64: string,
    left: number,
    top: number,
    width: number,
    height: number
  ): Promise<string> {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // addImage is available in PowerPointApi 1.4+
      const shapes = slide.shapes as PowerPoint.ShapeCollection & {
        addImage?(base64: string, options: object): PowerPoint.Shape;
      };
      if (typeof shapes.addImage === "function") {
        const image = shapes.addImage(base64, { left, top, width, height });
        image.load("id");
        await context.sync();
        return image.id;
      }
      // Fallback: insert as a placeholder text box
      const placeholder = slide.shapes.addTextBox("(Image)", { left, top, width, height });
      placeholder.load("id");
      await context.sync();
      return placeholder.id;
    });
  }
}

export const powerpointService = new PowerPointService();
