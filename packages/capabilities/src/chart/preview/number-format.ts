type Token =
  | { readonly kind: "literal"; readonly text: string }
  | { readonly kind: "digit"; readonly ch: "0" | "#" | "?" }
  | { readonly kind: "dot" }
  | { readonly kind: "comma" }
  | { readonly kind: "percent" }
  | { readonly kind: "exp"; readonly sign: "+" | "-" }
  | { readonly kind: "general" };

function splitSections(formatCode: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuotes = false;
  let inBracket = false;
  for (let i = 0; i < formatCode.length; i += 1) {
    const ch = formatCode[i]!;
    if (inQuotes) {
      current += ch;
      if (ch === '"') inQuotes = false;
      continue;
    }
    if (ch === "\\" || ch === "_") {
      current += ch + (formatCode[i + 1] ?? "");
      i += 1;
      continue;
    }
    if (ch === "[") {
      inBracket = true;
      current += ch;
      continue;
    }
    if (inBracket) {
      current += ch;
      if (ch === "]") inBracket = false;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      current += ch;
      continue;
    }
    if (ch === ";") {
      sections.push(current);
      current = "";
      if (sections.length === 3) {
        sections.push(formatCode.slice(i + 1));
        return sections;
      }
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
}

function tokenizeSection(section: string): Token[] {
  const tokens: Token[] = [];
  for (let i = 0; i < section.length; i += 1) {
    const ch = section[i]!;
    if (ch === '"') {
      let literal = "";
      for (i += 1; i < section.length; i += 1) {
        if (section[i] === '"') break;
        literal += section[i]!;
      }
      if (literal !== "") tokens.push({ kind: "literal", text: literal });
      continue;
    }
    if (ch === "\\" || ch === "*") {
      const next = section[i + 1];
      if (next !== undefined) {
        tokens.push({ kind: "literal", text: next });
        i += 1;
      }
      continue;
    }
    if (ch === "_") {
      i += 1;
      continue;
    }
    if (ch === "[") {
      const end = section.indexOf("]", i + 1);
      if (end >= 0) {
        const directive = section.slice(i + 1, end);
        if (directive.startsWith("$")) {
          const rest = directive.slice(1);
          const dash = rest.indexOf("-");
          const currency = dash < 0 ? rest : rest.slice(0, dash);
          if (currency !== "") {
            tokens.push({ kind: "literal", text: currency });
          }
        }
        i = end;
        continue;
      }
    }
    if (ch === "0" || ch === "#" || ch === "?") {
      tokens.push({ kind: "digit", ch });
      continue;
    }
    if (ch === ".") {
      tokens.push({ kind: "dot" });
      continue;
    }
    if (ch === ",") {
      tokens.push({ kind: "comma" });
      continue;
    }
    if (ch === "%") {
      tokens.push({ kind: "percent" });
      continue;
    }
    if (
      (ch === "E" || ch === "e") &&
      (section[i + 1] === "+" || section[i + 1] === "-")
    ) {
      tokens.push({ kind: "exp", sign: section[i + 1] as "+" | "-" });
      i += 1;
      continue;
    }
    if (
      (ch === "G" || ch === "g") &&
      section.slice(i, i + 7).toLowerCase() === "general" &&
      (section[i + 7] === undefined || !/[a-z]/iu.test(section[i + 7]!))
    ) {
      tokens.push({ kind: "general" });
      i += 6;
      continue;
    }
    tokens.push({ kind: "literal", text: ch });
  }
  return tokens;
}

function formatGeneral(value: number): string {
  if (Object.is(value, -0) || value === 0) return "0";
  const clean = Number(value.toPrecision(12));
  return String(clean);
}

function addGrouping(intStr: string): string {
  if (intStr.length <= 3) return intStr;
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function formatFixedParts(
  absValue: number,
  minInt: number,
  minFrac: number,
  maxFrac: number,
  useGrouping: boolean,
): string {
  const clampedFrac = Math.min(Math.max(0, maxFrac), 20);
  const fixed = absValue.toFixed(clampedFrac);
  const dotIdx = fixed.indexOf(".");
  let intPart = dotIdx < 0 ? fixed : fixed.slice(0, dotIdx);
  let fracPart = dotIdx < 0 ? "" : fixed.slice(dotIdx + 1);
  if (intPart === "0" && minInt === 0) {
    intPart = "";
  } else if (intPart.length < minInt) {
    intPart = intPart.padStart(minInt, "0");
  }
  if (useGrouping && intPart.length > 3) {
    intPart = addGrouping(intPart);
  }
  while (fracPart.length > minFrac && fracPart.endsWith("0")) {
    fracPart = fracPart.slice(0, -1);
  }
  if (intPart === "" && fracPart === "" && minFrac === 0) {
    return "";
  }
  return fracPart.length > 0 ? `${intPart}.${fracPart}` : intPart;
}

function formatSection(absValue: number, section: string): string {
  if (section === "") return "";
  const tokens = tokenizeSection(section);
  if (tokens.some((token) => token.kind === "general")) {
    return tokens
      .map((token) => {
        if (token.kind === "general") return formatGeneral(absValue);
        if (token.kind === "literal") return token.text;
        if (token.kind === "percent") return "%";
        return "";
      })
      .join("");
  }

  let firstNumIdx = -1;
  let lastNumIdx = -1;
  let lastDigitIdx = -1;
  let expIdx = -1;
  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]!;
    if (token.kind === "digit" || token.kind === "dot") {
      if (firstNumIdx < 0) firstNumIdx = i;
      lastNumIdx = i;
    }
    if (token.kind === "digit") {
      lastDigitIdx = i;
    }
    if (token.kind === "exp" && expIdx < 0) {
      expIdx = i;
    }
  }

  let scaled = absValue;
  let percentCount = 0;
  for (const token of tokens) {
    if (token.kind === "percent") percentCount += 1;
  }
  if (percentCount > 0) {
    scaled *= 100 ** percentCount;
  }

  // Trailing commas immediately after the last digit placeholder scale down by 1000 each.
  let trailingCommas = 0;
  if (lastDigitIdx >= 0) {
    for (let i = lastDigitIdx + 1; i < tokens.length; i += 1) {
      const token = tokens[i]!;
      if (token.kind === "comma") {
        trailingCommas += 1;
        lastNumIdx = Math.max(lastNumIdx, i);
      } else if (token.kind === "literal" || token.kind === "percent") {
        break;
      }
    }
  }
  if (trailingCommas > 0 && expIdx < 0) {
    scaled /= 1000 ** trailingCommas;
  }

  if (firstNumIdx < 0) {
    return tokens
      .map((token) =>
        token.kind === "literal"
          ? token.text
          : token.kind === "percent"
            ? "%"
            : "",
      )
      .join("");
  }

  const prefix = tokens
    .slice(0, firstNumIdx)
    .map((token) =>
      token.kind === "literal"
        ? token.text
        : token.kind === "percent"
          ? "%"
          : "",
    )
    .join("");
  const suffix = tokens
    .slice(lastNumIdx + 1)
    .map((token) =>
      token.kind === "literal"
        ? token.text
        : token.kind === "percent"
          ? "%"
          : "",
    )
    .join("");

  if (expIdx >= 0 && expIdx > firstNumIdx) {
    const expToken = tokens[expIdx] as { kind: "exp"; sign: "+" | "-" };
    const mantissaTokens = tokens.slice(firstNumIdx, expIdx);
    const expTokens = tokens.slice(expIdx + 1, lastNumIdx + 1);
    const dotPos = mantissaTokens.findIndex((token) => token.kind === "dot");
    const intDigits = (
      dotPos < 0 ? mantissaTokens : mantissaTokens.slice(0, dotPos)
    ).filter((token) => token.kind === "digit");
    const fracDigits = (
      dotPos < 0 ? [] : mantissaTokens.slice(dotPos + 1)
    ).filter((token) => token.kind === "digit");
    const expMinDigits = expTokens.filter(
      (token) => token.kind === "digit" && token.ch === "0",
    ).length;
    const minInt = intDigits.filter((token) => token.ch === "0").length;
    const minFrac = fracDigits.filter((token) => token.ch === "0").length;
    const maxFrac = fracDigits.length;

    let expVal = 0;
    let mantissa = 0;
    if (scaled > 0) {
      const leadDigits = Math.max(1, intDigits.length);
      const rawExp = Math.floor(Math.log10(scaled));
      expVal = rawExp - (leadDigits - 1);
      mantissa = scaled / 10 ** expVal;
    }
    const mantissaStr = formatFixedParts(
      mantissa,
      Math.max(1, minInt),
      minFrac,
      maxFrac,
      false,
    );
    const expSign = expVal < 0 ? "-" : expToken.sign === "+" ? "+" : "";
    const expDigits = String(Math.abs(expVal)).padStart(
      Math.max(1, expMinDigits),
      "0",
    );
    return `${prefix}${mantissaStr}E${expSign}${expDigits}${suffix}`;
  }

  const numTokens = tokens.slice(firstNumIdx, lastNumIdx + 1);
  const dotPos = numTokens.findIndex((token) => token.kind === "dot");
  const intTokens = dotPos < 0 ? numTokens : numTokens.slice(0, dotPos);
  const fracTokens = dotPos < 0 ? [] : numTokens.slice(dotPos + 1);

  const minInt = intTokens.filter(
    (token) => token.kind === "digit" && token.ch === "0",
  ).length;
  const useGrouping = intTokens.some((token) => token.kind === "comma");
  const fracDigits = fracTokens.filter((token) => token.kind === "digit");
  const minFrac = fracDigits.filter(
    (token) => token.kind === "digit" && token.ch === "0",
  ).length;
  const maxFrac = fracDigits.length;

  const formatted = formatFixedParts(
    scaled,
    minInt,
    minFrac,
    maxFrac,
    useGrouping,
  );
  // Preserve percent/literal tokens that appeared inside the number span (e.g. `0%`).
  const innerSuffix = numTokens
    .filter((token) => token.kind === "percent")
    .map(() => "%")
    .join("");
  return `${prefix}${formatted}${innerSuffix}${suffix}`;
}

/**
 * Formats a numeric chart value using an Excel/OOXML number format string.
 */
export function formatChartNumber(value: number, formatCode?: string): string {
  if (!Number.isFinite(value)) return String(value);
  if (
    formatCode === undefined ||
    formatCode.trim().toLowerCase() === "general"
  ) {
    return formatGeneral(value);
  }
  const sections = splitSections(formatCode);
  const normalized = Object.is(value, -0) ? 0 : value;
  if (normalized === 0 && sections.length >= 3) {
    return formatSection(0, sections[2]!);
  }
  if (normalized < 0) {
    if (sections.length >= 2) {
      return formatSection(Math.abs(normalized), sections[1]!);
    }
    return `-${formatSection(Math.abs(normalized), sections[0]!)}`;
  }
  return formatSection(normalized, sections[0]!);
}
