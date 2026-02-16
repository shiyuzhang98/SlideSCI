import { ShapeInfo } from "../models/types";
import { cmToPoints } from "../utils/unit-conversion";

export interface ArrangeOptions {
  colNum: number;
  colSpace: number; // points
  rowSpace: number; // points
  customWidth: number; // cm, 0 = no custom
  customHeight: number; // cm, 0 = no custom
}

interface ShapeUpdate {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Mode 0: Max column width arrangement.
 * Each column has fixed width = max width of shapes in that column.
 * Reference: Ribbon1.cs lines 1007-1092
 */
export function arrangeMaxColumnWidth(
  shapes: ShapeInfo[],
  options: ArrangeOptions
): ShapeUpdate[] {
  const { colNum, colSpace, rowSpace, customWidth, customHeight } = options;
  const useCustomWidth = customWidth > 0;
  const useCustomHeight = customHeight > 0;
  const cwPt = cmToPoints(customWidth);
  const chPt = cmToPoints(customHeight);

  const updates: ShapeUpdate[] = shapes.map((s) => ({
    id: s.id,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  }));

  // 1. Pre-assign shapes to columns
  const columns: ShapeUpdate[][] = [];
  for (let i = 0; i < colNum; i++) {
    columns.push([]);
  }
  for (let i = 0; i < updates.length; i++) {
    columns[i % colNum].push(updates[i]);
  }

  // 2. Apply custom dimensions and calculate max width per column
  const columnWidths: number[] = [];
  for (let i = 0; i < colNum; i++) {
    let columnMaxWidth = 0;
    for (const shape of columns[i]) {
      const aspectRatio = shape.width / shape.height;
      if (useCustomWidth && !useCustomHeight) {
        shape.width = cwPt;
        shape.height = cwPt / aspectRatio;
      } else if (!useCustomWidth && useCustomHeight) {
        shape.height = chPt;
        shape.width = chPt * aspectRatio;
      } else if (useCustomWidth && useCustomHeight) {
        shape.width = cwPt;
        shape.height = chPt;
      }
      columnMaxWidth = Math.max(columnMaxWidth, shape.width);
    }
    columnWidths.push(columnMaxWidth);
  }

  // 3. Arrange by rows
  const startX = updates[0].left;
  let currentY = updates[0].top;
  let currentX = startX;
  let rowMaxHeight = 0;
  let colCount = 0;

  for (const shape of updates) {
    const aspectRatio = shape.width / shape.height;
    if (useCustomWidth && !useCustomHeight) {
      shape.width = cwPt;
      shape.height = cwPt / aspectRatio;
    } else if (!useCustomWidth && useCustomHeight) {
      shape.height = chPt;
      shape.width = chPt * aspectRatio;
    } else if (useCustomWidth && useCustomHeight) {
      shape.width = cwPt;
      shape.height = chPt;
    }

    if (colCount >= colNum) {
      colCount = 0;
      currentX = startX;
      currentY += rowMaxHeight + rowSpace;
      rowMaxHeight = 0;
    }

    shape.left = currentX;
    shape.top = currentY;
    rowMaxHeight = Math.max(rowMaxHeight, shape.height);
    currentX += columnWidths[colCount] + colSpace;
    colCount++;
  }

  return updates;
}

/**
 * Mode 1: Uniform height arrangement.
 * All shapes in a row have the same height (reference height from first shape or custom).
 * Reference: Ribbon1.cs lines 1093-1158
 */
export function arrangeUniformHeight(
  shapes: ShapeInfo[],
  options: ArrangeOptions
): ShapeUpdate[] {
  const { colNum, colSpace, rowSpace, customWidth, customHeight } = options;
  const useCustomWidth = customWidth > 0;
  const useCustomHeight = customHeight > 0;
  const cwPt = cmToPoints(customWidth);
  const chPt = cmToPoints(customHeight);

  const updates: ShapeUpdate[] = shapes.map((s) => ({
    id: s.id,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  }));

  let referenceHeight = updates[0].height;
  if (useCustomWidth && !useCustomHeight) {
    referenceHeight = 0;
  }

  const startX = updates[0].left;
  let currentY = updates[0].top;
  let currentX = startX;
  let colCount = 0;

  for (const shape of updates) {
    const aspectRatio = shape.width / shape.height;

    if (!useCustomWidth && !useCustomHeight) {
      shape.height = referenceHeight;
      shape.width = referenceHeight * aspectRatio;
    } else if (useCustomWidth && !useCustomHeight) {
      shape.width = cwPt;
      shape.height = cwPt / aspectRatio;
    } else if (!useCustomWidth && useCustomHeight) {
      shape.height = chPt;
      shape.width = chPt * aspectRatio;
      referenceHeight = chPt;
    } else {
      shape.width = cwPt;
      shape.height = chPt;
    }

    if (colCount >= colNum) {
      colCount = 0;
      currentX = startX;
      currentY += referenceHeight + rowSpace;
      if (useCustomWidth && !useCustomHeight) {
        referenceHeight = 0;
      }
    }

    shape.left = currentX;
    shape.top = currentY;
    currentX += shape.width + colSpace;
    colCount++;

    if (useCustomWidth && !useCustomHeight) {
      referenceHeight = Math.max(referenceHeight, shape.height);
    }
  }

  return updates;
}

/**
 * Mode 2: Waterfall / uniform width arrangement.
 * All shapes have uniform width, placed in the shortest column.
 * Reference: Ribbon1.cs lines 1159-1201
 */
export function arrangeWaterfall(
  shapes: ShapeInfo[],
  options: ArrangeOptions
): ShapeUpdate[] {
  const { colNum, colSpace, rowSpace, customWidth } = options;
  const cwPt = customWidth > 0 ? cmToPoints(customWidth) : shapes[0].width;

  const updates: ShapeUpdate[] = shapes.map((s) => ({
    id: s.id,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  }));

  const startX = updates[0].left;
  const startY = updates[0].top;

  // Initialize column positions
  const columnTops: number[] = new Array(colNum).fill(startY);
  const columnLefts: number[] = [];
  for (let i = 0; i < colNum; i++) {
    columnLefts.push(startX + i * (cwPt + colSpace));
  }

  for (const shape of updates) {
    // Uniform width, maintain aspect ratio
    const aspectRatio = shape.width / shape.height;
    shape.width = cwPt;
    shape.height = cwPt / aspectRatio;

    // Find the shortest column
    let minColumn = 0;
    let minHeight = columnTops[0];
    for (let i = 1; i < colNum; i++) {
      if (columnTops[i] < minHeight) {
        minHeight = columnTops[i];
        minColumn = i;
      }
    }

    // Place shape
    shape.left = columnLefts[minColumn];
    shape.top = columnTops[minColumn];

    // Update column height
    columnTops[minColumn] += shape.height + rowSpace;
  }

  return updates;
}

/**
 * Main arrangement function that dispatches to the correct mode.
 * @param shapes - Pre-sorted shapes to arrange
 * @param alignType - 0: max column width, 1: uniform height, 2: waterfall
 * @param options - Arrangement options
 */
export function arrangeShapes(
  shapes: ShapeInfo[],
  alignType: number,
  options: ArrangeOptions
): ShapeUpdate[] {
  if (shapes.length === 0) return [];

  switch (alignType) {
    case 1:
      return arrangeUniformHeight(shapes, options);
    case 2:
      return arrangeWaterfall(shapes, options);
    default:
      return arrangeMaxColumnWidth(shapes, options);
  }
}
