/** Format the supported Excel-style numeric subset, or return undefined for unsupported syntax.
 * Separators are deterministic (comma grouping, decimal point), independent of host locale.
 * Currency labels are literals, not currency conversion or locale inference.
 */
export function formatNumber(
  value: number,
  format: string,
): string | undefined {
  if (!Number.isFinite(value) || format.length > 4096) return undefined;
  // Tokenize before splitting: quoted/escaped semicolons belong to literals.
  const sections: string[][] = [[]];
  let section = sections[0]!;
  for (let index = 0; index < format.length; index++) {
    const char = format[index]!;
    if (char === ";") {
      section = [];
      sections.push(section);
    } else if (char === '"') {
      const end = format.indexOf('"', index + 1);
      if (end < 0) return undefined;
      section.push(JSON.stringify(format.slice(index + 1, end)));
      index = end;
    } else if (char === "\\" || char === "_") {
      if (++index >= format.length) return undefined;
      // Excel underscore reserves glyph width; plain text output omits that padding.
      if (char === "\\") section.push(JSON.stringify(format[index]));
    } else if (char === "[") {
      const end = format.indexOf("]", index + 1);
      if (end < 0) return undefined;
      const token = format.slice(index + 1, end);
      const currency = /^\$([^\[\]-]+)(?:-[0-9a-f]+)?$/i.exec(token);
      if (currency) section.push(JSON.stringify(currency[1]));
      else if (
        !/^(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)$/i.test(token)
      )
        return undefined;
      index = end;
    } else section.push(char);
  }
  if (sections.length > 4) return undefined;
  const parsed = sections.slice(0, 3).map(parseSection);
  if (parsed.some((item) => item === undefined)) return undefined;
  const selected =
    parsed[
      value < 0 && parsed.length > 1
        ? 1
        : value === 0 && parsed.length > 2
          ? 2
          : 0
    ];
  if (!selected) return undefined;
  if (!selected.pattern) return selected.prefix;
  const decimalPlaces =
    selected.pattern.split(".")[1]?.replace("%", "").length ?? 0;
  const magnitude =
    Math.abs(value) * (selected.pattern.endsWith("%") ? 100 : 1);
  if (!Number.isFinite(magnitude)) return undefined;
  const number = new Intl.NumberFormat("en-US", {
    useGrouping: selected.pattern.includes(","),
    minimumFractionDigits: decimalPlaces,
    maximumFractionDigits: decimalPlaces,
  }).format(magnitude);
  return (
    (value < 0 && parsed.length === 1 ? "-" : "") +
    selected.prefix +
    number +
    (selected.pattern.endsWith("%") ? "%" : "") +
    selected.suffix
  );
}

function parseSection(
  tokens: readonly string[],
): { prefix: string; pattern: string; suffix: string } | undefined {
  let prefix = "",
    pattern = "",
    suffix = "",
    ended = false;
  for (const token of tokens) {
    if (/^[0#,.%]$/.test(token)) {
      if (ended) return undefined;
      pattern += token;
    } else {
      let literal: string;
      if (token.startsWith('"')) literal = JSON.parse(token) as string;
      else if (/^[\p{Sc} ()+\-]$/u.test(token)) literal = token;
      else return undefined;
      if (pattern) {
        ended = true;
        suffix += literal;
      } else prefix += literal;
    }
  }
  if (pattern && !/^(?:0|#,##0)(?:\.0{1,10})?%?$/.test(pattern))
    return undefined;
  return { prefix, pattern, suffix };
}
