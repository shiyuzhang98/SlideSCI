/**
 * Label template definitions.
 * Maps template key to the ordered set of label characters.
 */
const LABEL_CHARS: Record<string, string> = {
  "A": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
  "a": "abcdefghijklmnopqrstuvwxyz",
  "A)": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
  "a)": "abcdefghijklmnopqrstuvwxyz",
  "(A)": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
  "(a)": "abcdefghijklmnopqrstuvwxyz",
  "1": "123456789",
  "1)": "123456789",
  "\u2160": "\u2160\u2161\u2162\u2163\u2164\u2165\u2166\u2167\u2168\u2169",
  "\u2160)": "\u2160\u2161\u2162\u2163\u2164\u2165\u2166\u2167\u2168\u2169",
  "\u2460": "\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469",
  "\u2460)": "\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469",
  "\u4e00": "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341",
  "\u4e00)": "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341",
};

/**
 * Generate a label string for the given index using the specified template.
 * @param template - Label template key (e.g., "A", "a)", "(A)", "1", etc.)
 * @param index - 0-based index of the shape
 * @param startIndex - 1-based starting number
 */
export function generateLabel(
  template: string,
  index: number,
  startIndex: number = 1
): string {
  const chars = LABEL_CHARS[template] || LABEL_CHARS["A"];
  const isNumeric = template.startsWith("1");

  let label: string;
  if (isNumeric) {
    label = (startIndex + index).toString();
  } else {
    const charIndex = (startIndex - 1 + index) % chars.length;
    label = chars[charIndex];
  }

  // Apply formatting suffix/prefix based on template
  if (template.endsWith(")") && !template.startsWith("(")) {
    label += ")";
  } else if (template.startsWith("(") && template.endsWith(")")) {
    label = "(" + label + ")";
  }

  return label;
}

/**
 * Generate an array of labels for a given count.
 * @param template - Label template key
 * @param count - Number of labels to generate
 * @param startIndex - 1-based starting number
 */
export function generateLabels(
  template: string,
  count: number,
  startIndex: number = 1
): string[] {
  const labels: string[] = [];
  for (let i = 0; i < count; i++) {
    labels.push(generateLabel(template, i, startIndex));
  }
  return labels;
}
