// Lexical evidence only. PDF extraction may split a word over lines. Neither
// match policy establishes geometry, legibility, editability or native playback.
export function pdfTextEvidence(text, markers) {
  if (typeof text !== "string" || !Array.isArray(markers) || markers.some(m => typeof m !== "string" || !m.trim().length)) {
    throw new TypeError("Expected extracted PDF text and nonempty string markers.");
  }
  const words = value => value.replace(/\s+/gu, " ").trim();
  const compact = value => value.replace(/\s+/gu, "");
  const normalized = words(text), stream = compact(text);
  const missingTextMarkers = markers.filter(marker => !normalized.includes(words(marker)));
  const missingWhitespaceInsensitiveMarkers = markers.filter(marker => !stream.includes(compact(marker)));
  return {
    missingTextMarkers, allTextMarkersInPdf: missingTextMarkers.length === 0,
    missingWhitespaceInsensitiveMarkers,
    allWhitespaceInsensitiveMarkersInPdf: missingWhitespaceInsensitiveMarkers.length === 0,
    textMatchStatus: missingTextMarkers.length === 0 ? "word-spacing-match"
      : missingWhitespaceInsensitiveMarkers.length === 0 ? "whitespace-fragmented-match" : "unmatched-markers",
  };
}
