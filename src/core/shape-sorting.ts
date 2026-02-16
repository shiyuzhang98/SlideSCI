import { ShapeInfo, ImageGroup } from "../models/types";

/**
 * Sort shapes by position: primarily by top (Y), then by left (X).
 * Reference: Ribbon1.cs GetSortedSelection() lines 620-653
 */
export function sortShapesByPosition(shapes: ShapeInfo[]): ShapeInfo[] {
  return [...shapes].sort((a, b) => {
    if (a.top !== b.top) {
      return a.top - b.top;
    }
    return a.left - b.left;
  });
}

/**
 * Check if a shape vertically overlaps with an image group.
 * Overlap threshold: 50% of shape height.
 * Reference: Ribbon1.cs ImageGroup.OverlapsWith() lines 1829-1841
 */
function overlapsWith(group: ImageGroup, shape: ShapeInfo): boolean {
  const shapeHeight = shape.height;
  const threshold = shapeHeight * 0.5;
  const shapeBottom = shape.top + shapeHeight;

  const overlapStart = Math.max(group.minTop, shape.top);
  const overlapEnd = Math.min(group.maxBottom, shapeBottom);
  const overlapHeight = overlapEnd - overlapStart;

  return overlapHeight >= threshold;
}

/**
 * Add a shape to an image group, updating bounds.
 * Reference: Ribbon1.cs ImageGroup.AddShape() lines 1843-1856
 */
function addShapeToGroup(group: ImageGroup, shape: ShapeInfo): void {
  if (group.shapes.length === 0) {
    group.minTop = shape.top;
    group.maxBottom = shape.top + shape.height;
  } else {
    group.minTop = Math.min(group.minTop, shape.top);
    group.maxBottom = Math.max(group.maxBottom, shape.top + shape.height);
  }
  group.shapes.push(shape);
}

/**
 * Group shapes by vertical overlap, then sort within groups by X position,
 * and sort groups by their minimum Y position.
 * Returns a flattened list of shapes in row-then-column order.
 *
 * This is used for both image arrangement and label assignment.
 * Reference: Ribbon1.cs lines 949-993 and 2461-2518
 */
export function groupAndSortShapes(shapes: ShapeInfo[]): ShapeInfo[] {
  const groups: ImageGroup[] = [];

  // Group shapes based on vertical overlap
  for (const shape of shapes) {
    let addedToExistingGroup = false;
    for (const group of groups) {
      if (overlapsWith(group, shape)) {
        addShapeToGroup(group, shape);
        addedToExistingGroup = true;
        break;
      }
    }

    if (!addedToExistingGroup) {
      const newGroup: ImageGroup = {
        shapes: [],
        minTop: 0,
        maxBottom: 0,
      };
      addShapeToGroup(newGroup, shape);
      groups.push(newGroup);
    }
  }

  // Sort shapes within each group by X position
  for (const group of groups) {
    group.shapes.sort((a, b) => a.left - b.left);
  }

  // Sort groups by their minimum Y position
  groups.sort((a, b) => a.minTop - b.minTop);

  // Flatten all groups into a single sorted list
  const result: ShapeInfo[] = [];
  for (const group of groups) {
    result.push(...group.shapes);
  }
  return result;
}
