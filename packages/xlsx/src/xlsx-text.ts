/** Decodes SpreadsheetML escaped UTF-16 once, so escaped literal underscores stay literal. */
export function decodeXlsxText(text: string): string {
  return text.replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

/** Encodes SpreadsheetML strings before XML escaping, including control characters and literal escapes. */
export function encodeXlsxText(text: string, attribute = false): string {
  const literal = text.replace(
    /_x[0-9a-fA-F]{4}_/g,
    (match) => `_x005F_${match.slice(1)}`,
  );
  return literal.replace(/[\u0000-\u001f\ufffe\uffff]/g, (character) => {
    const code = character.charCodeAt(0);
    if (!attribute && (code === 9 || code === 10)) return character;
    return `_x${code.toString(16).toUpperCase().padStart(4, "0")}_`;
  });
}
