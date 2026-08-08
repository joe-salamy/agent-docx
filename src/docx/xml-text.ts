/** Removes code points forbidden by XML 1.0 from emitted text. */
export const sanitizeXmlText = (text: string): string => {
  let sanitized = "";
  for (const character of text) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x09 ||
      codePoint === 0x0a ||
      codePoint === 0x0d ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff)
    )
      sanitized += character;
  }
  return sanitized;
};
