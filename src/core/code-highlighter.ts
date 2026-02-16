export interface HighlightToken {
  start: number;
  length: number;
  type: string;
  color: string;
}

interface ThemeColors {
  keyword: string;
  comment: string;
  string: string;
  number: string;
  property: string;
  selector: string;
}

const DARK_THEME: ThemeColors = {
  keyword: "#569CD6",   // blue
  comment: "#57A64A",   // green
  string: "#D69D85",    // brown
  number: "#B5CEA8",    // teal
  property: "#9CDCFE",  // light blue
  selector: "#D7BA7D",  // gold
};

const LIGHT_THEME: ThemeColors = {
  keyword: "#0000FF",
  comment: "#008000",
  string: "#A31515",
  number: "#098658",
  property: "#0086D1",
  selector: "#A86300",
};

type PatternDef = [string, string, string]; // [pattern, flags, type]

const LANGUAGE_PATTERNS: Record<string, PatternDef[]> = {
  csharp: [
    ["/\\*[\\s\\S]*?\\*/", "", "comment"],
    ["//[^\\n]*", "", "comment"],
    ['@"([^"]|"")*"|"([^"\\n\\\\]|\\\\.)*"', "", "string"],
    ["\\b\\d*\\.?\\d+([eE][-+]?\\d+)?\\b", "", "number"],
    [
      "\\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|void|volatile|while)\\b",
      "",
      "keyword",
    ],
  ],
  python: [
    ["#.*?$", "m", "comment"],
    [
      "(?:[ruf]|rf|fr)?(?:'''[\\s\\S]*?'''|\"\"\"[\\s\\S]*?\"\"\"|'[^'\\n\\\\]*(?:\\\\.[^'\\n\\\\]*)*'|\"[^\"\\n\\\\]*(?:\\\\.[^\"\\n\\\\]*)*\")",
      "",
      "string",
    ],
    ["\\b\\d*\\.?\\d+([eE][-+]?\\d+)?\\b", "", "number"],
    [
      "\\b(and|as|assert|async|await|break|class|continue|def|del|elif|else|except|False|finally|for|from|global|if|import|in|is|lambda|None|nonlocal|not|or|pass|raise|return|True|try|while|with|yield)\\b",
      "",
      "keyword",
    ],
  ],
  javascript: [
    ["/\\*[\\s\\S]*?\\*/|//.*?$", "m", "comment"],
    ["`(?:[^`\\\\]|\\\\.)*`|'(?:[^'\\\\]|\\\\.)*'|\"(?:[^\"\\\\]|\\\\.)*\"", "", "string"],
    ["\\b\\d*\\.?\\d+([eE][-+]?\\d+)?\\b", "", "number"],
    [
      "\\b(async|await|break|case|catch|class|const|continue|debugger|default|delete|do|else|export|extends|finally|for|function|if|import|in|instanceof|new|return|super|switch|this|throw|try|typeof|var|void|while|with|yield|let)\\b",
      "",
      "keyword",
    ],
  ],
  matlab: [
    ["%.*?$", "m", "comment"],
    ['(?:"[^"\\n]*"|\'[^\'\\n]*\')', "", "string"],
    ["\\b\\d*\\.?\\d+([eE][-+]?\\d+)?\\b", "", "number"],
    [
      "\\b(break|case|catch|classdef|continue|else|elseif|end|for|function|global|if|otherwise|parfor|persistent|return|switch|try|while|clear|close|load|save|figure|plot|xlabel|ylabel|title|grid|hold|zeros|ones|rand|eye|disp|input|fprintf|strcmp|length|size|max|min|sum|mean|std|find|sort|reshape)\\b",
      "",
      "keyword",
    ],
  ],
  css: [
    ["/\\*[\\s\\S]*?\\*/", "", "comment"],
    ["[^{}]+(?=\\s*\\{)", "", "selector"],
    ["[\\w-]+(?=\\s*:)", "", "property"],
    ["'[^']*'|\"[^\"]*\"", "", "string"],
    ["\\b\\d*\\.?\\d+(?:px|em|rem|vh|vw|%|s|ms|deg|rad|turn)?\\b", "", "number"],
    [
      "\\b(important|inherit|initial|unset|none|auto|hidden|visible|block|inline|flex|grid|absolute|relative|fixed|static|left|right|top|bottom|center|justify|stretch|wrap|nowrap|solid|dashed|dotted)\\b",
      "",
      "keyword",
    ],
  ],
  html: [
    ["<!--[\\s\\S]*?-->", "", "comment"],
    ["'[^']*'|\"[^\"]*\"", "", "string"],
    ["</?\\w+(?:\\s+\\w+(?:\\s*=\\s*(?:\".*?\"|'.*?'|[^'\"\\s]\\S+))?)*\\s*/?>", "", "keyword"],
    ["\\b\\d+\\b", "", "number"],
  ],
  r: [
    ["#.*?$", "m", "comment"],
    ['(?:"[^"\\n\\\\]*(?:\\\\.[^"\\n\\\\]*)*"|\'[^\'\\n\\\\]*(?:\\\\.[^\'\\n\\\\]*)*\')', "", "string"],
    ["\\b\\d*\\.?\\d+([eE][-+]?\\d+)?L?\\b", "", "number"],
    [
      "\\b(if|else|for|in|while|function|repeat|next|break|TRUE|FALSE|NULL|Inf|NaN|NA|NA_integer_|NA_real_|NA_complex_|NA_character_|c|list|data\\.frame|matrix|array|factor|length|names|dim|class|str|summary|head|tail|print|cat|paste|paste0|substr|nchar|grep|gsub|which|is\\.na|is\\.null|is\\.numeric|is\\.character|is\\.logical|as\\.numeric|as\\.character|as\\.logical|mean|median|sd|var|min|max|sum|apply|lapply|sapply|mapply|tapply|aggregate|merge|rbind|cbind|subset|sort|order|unique|duplicated|table|plot|hist|boxplot|barplot|pie|lines|points|abline|legend|title|xlabel|ylabel|par|dev\\.new|dev\\.off|png|pdf|jpeg|library|require|install\\.packages|source|load|save|write\\.csv|read\\.csv|read\\.table|write\\.table)\\b",
      "",
      "keyword",
    ],
  ],
  fortran: [
    ["!.*?$", "mi", "comment"],
    ['(?:"[^"\\n\\\\]*(?:\\\\.[^"\\n\\\\]*)*"|\'[^\'\\n\\\\]*(?:\\\\.[^\'\\n\\\\]*)*\')', "", "string"],
    ["\\b\\d*\\.?\\d+([eEdD][-+]?\\d+)?\\b", "", "number"],
    ["\\b(?:implicit\\s+none)\\b", "i", "keyword"],
    ["\\b(?:double\\s+precision)\\b", "i", "keyword"],
    ["\\b(?:end\\s+(?:program|subroutine|function|module|type|interface|do|if|select))\\b", "i", "keyword"],
    [
      "\\b(program|end|subroutine|function|module|use|implicit|none|integer|real|complex|logical|character|parameter|dimension|allocatable|pointer|target|intent|in|out|inout|optional|public|private|save|data|common|equivalence|external|intrinsic|interface|contains|procedure|abstract|extends|class|type|select|case|default|where|elsewhere|forall|pure|elemental|recursive|result|only|operator|assignment|generic|sequence|bind|value|volatile|asynchronous|protected|enum|enumerator|associate|block|critical|sync|all|images|memory|lock|unlock|event|post|wait|if|then|else|elseif|endif|do|while|enddo|continue|exit|cycle|stop|pause|return|call|goto|assign|to|format|open|close|read|write|print|rewind|backspace|endfile|inquire|namelist|include|import)\\b",
      "i",
      "keyword",
    ],
    [
      "\\b(abs|acos|aimag|aint|alog|alog10|amax0|amax1|amin0|amin1|amod|anint|asin|atan|atan2|cabs|ccos|cexp|char|clog|cmplx|conjg|cos|cosh|csin|csqrt|dabs|dacos|dasin|datan|datan2|dble|dcos|dcosh|ddim|dexp|dim|dint|dlog|dlog10|dmax1|dmin1|dmod|dnint|dprod|dsign|dsin|dsinh|dsqrt|dtan|dtanh|exp|float|iabs|ichar|idim|idint|idnint|ifix|index|int|isign|len|lge|lgt|lle|llt|log|log10|max|max0|max1|min|min0|min1|mod|nint|real|sign|sin|sinh|sngl|sqrt|tan|tanh|trim|adjustl|adjustr|all|any|count|maxval|minval|product|sum|matmul|dot_product|pack|unpack|reshape|spread|merge|eoshift|cshift|transpose|lbound|ubound|size|shape|allocated|associated|present|kind|selected_int_kind|selected_real_kind|huge|tiny|epsilon|precision|radix|range|digits|minexponent|maxexponent|fraction|exponent|spacing|rrspacing|scale|set_exponent|nearest|ceiling|floor|modulo|sign|verify|scan|null|transfer|bit_size|btest|iand|ibclr|ibits|ibset|ieor|ior|ishft|ishftc|not|mvbits|random_number|random_seed|system_clock|date_and_time|cpu_time)\\b",
      "i",
      "property",
    ],
  ],
};

const LANGUAGE_ALIASES: Record<string, string> = {
  js: "javascript",
  "c#": "csharp",
  cs: "csharp",
  py: "python",
  m: "matlab",
  htm: "html",
  R: "r",
  f90: "fortran",
  f95: "fortran",
  f03: "fortran",
  f08: "fortran",
  f77: "fortran",
  for: "fortran",
  f: "fortran",
};

/**
 * Resolve language aliases to canonical name.
 */
export function resolveLanguage(language: string): string {
  const lower = language.toLowerCase();
  return LANGUAGE_ALIASES[lower] || LANGUAGE_ALIASES[language] || lower;
}

/**
 * Get theme colors for the specified theme.
 */
export function getThemeColors(isDark: boolean): ThemeColors {
  return isDark ? { ...DARK_THEME } : { ...LIGHT_THEME };
}

/**
 * Get background and default text colors for a theme.
 */
export function getThemeBackground(isDark: boolean): {
  background: string;
  foreground: string;
  border: string;
} {
  return isDark
    ? { background: "#1E1E1E", foreground: "#FFFFFF", border: "#C8C8C8" }
    : { background: "#FFFFFF", foreground: "#000000", border: "#C8C8C8" };
}

/**
 * Compute syntax highlighting tokens for the given code and language.
 * Returns an array of tokens with position, length, type, and color.
 *
 * Ported from CodeHighlighter.cs ApplyHighlighting() method.
 * Priority: comments (0) > strings (1) > everything else (2).
 * Overlapping lower-priority matches are skipped.
 * Comments inside string literals are filtered out.
 */
export function highlightCode(
  code: string,
  language: string,
  isDark: boolean
): HighlightToken[] {
  // Normalize \r\n to \n to avoid position mismatches
  const normalizedCode = code.replace(/\r\n/g, "\n");
  const lang = resolveLanguage(language);
  const patterns = LANGUAGE_PATTERNS[lang];
  if (!patterns) return [];

  const colors = getThemeColors(isDark);

  // Collect all matches with priority
  interface MatchEntry {
    start: number;
    end: number;
    type: string;
    priority: number;
  }

  const allMatches: MatchEntry[] = [];

  for (const [pattern, flags, type] of patterns) {
    const priority = type === "comment" ? 0 : type === "string" ? 1 : 2;
    const regex = new RegExp(pattern, "g" + flags);
    let match: RegExpExecArray | null;
    while ((match = regex.exec(normalizedCode)) !== null) {
      if (match[0].length === 0) {
        regex.lastIndex++;
        continue;
      }
      allMatches.push({
        start: match.index,
        end: match.index + match[0].length,
        type,
        priority,
      });
    }
  }

  // Filter out comments that fall inside string literals
  const stringIntervals = allMatches
    .filter((m) => m.type === "string")
    .map((m) => ({ start: m.start, end: m.end }));

  const filtered =
    stringIntervals.length > 0
      ? allMatches.filter(
          (m) =>
            m.type !== "comment" ||
            !stringIntervals.some((si) => m.start >= si.start && m.start < si.end)
        )
      : allMatches;

  // Sort by priority (lower number = higher priority)
  filtered.sort((a, b) => a.priority - b.priority);

  // Process matches, skipping overlaps with already-processed ranges
  const processedIntervals: { start: number; end: number }[] = [];
  const tokens: HighlightToken[] = [];

  for (const entry of filtered) {
    // Check overlap with already-processed ranges
    const hasOverlap = processedIntervals.some(
      (pi) => entry.start < pi.end && entry.end > pi.start
    );
    if (hasOverlap) continue;

    tokens.push({
      start: entry.start,
      length: entry.end - entry.start,
      type: entry.type,
      color: colors[entry.type as keyof ThemeColors] || colors.keyword,
    });

    processedIntervals.push({ start: entry.start, end: entry.end });
  }

  return tokens;
}
