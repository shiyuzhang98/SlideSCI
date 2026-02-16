const CM_TO_POINTS = 28.34646;

export function cmToPoints(cm: number): number {
  return cm * CM_TO_POINTS;
}

export function pointsToCm(points: number): number {
  return points / CM_TO_POINTS;
}

export function pxToPoints(px: number, dpi: number = 96): number {
  return (px / dpi) * 72;
}

export function pointsToPx(points: number, dpi: number = 96): number {
  return (points / 72) * dpi;
}

export function inchesToPoints(inches: number): number {
  return inches * 72;
}

export function pointsToInches(points: number): number {
  return points / 72;
}
