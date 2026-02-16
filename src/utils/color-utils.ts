export function rgbToHex(r: number, g: number, b: number): string {
  return "#" + [r, g, b].map((x) => x.toString(16).padStart(2, "0")).join("");
}

export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  if (!result) {
    return { r: 0, g: 0, b: 0 };
  }
  return {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16),
  };
}

/** Convert OLE color integer (BGR) to hex string */
export function oleToHex(oleColor: number): string {
  const b = (oleColor >> 16) & 0xff;
  const g = (oleColor >> 8) & 0xff;
  const r = oleColor & 0xff;
  return rgbToHex(r, g, b);
}

/** Convert hex to OLE color integer (BGR) */
export function hexToOle(hex: string): number {
  const { r, g, b } = hexToRgb(hex);
  return (b << 16) | (g << 8) | r;
}
