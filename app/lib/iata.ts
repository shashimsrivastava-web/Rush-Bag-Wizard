// Configurable lookup table of IATA airline numeric owner codes to two-letter prefixes.
export const DEFAULT_IATA_AIRLINE_MAP: Record<string, string> = {
  '0220': 'LH', // Lufthansa
  '0724': 'LX', // SWISS
  '0257': 'OS', // Austrian Airlines
  '0082': 'SN', // Brussels Airlines
  '0016': 'UA', // United Airlines
  '0098': 'AI', // Air India
  '0125': 'BA', // British Airways
  '0555': 'SU', // Aeroflot
  '0074': 'KL', // KLM
  '0057': 'AF', // Air France
  '0001': 'AA', // American Airlines
  '0160': 'CX', // Cathay Pacific
  '0618': 'SQ', // Singapore Airlines
  '0006': 'DL', // Delta Air Lines
  '0176': 'EK', // Emirates
  '0607': 'EY', // Etihad Airways
  '0738': 'VN', // Vietnam Airlines
};

/**
 * Convert any tag to canonical airline prefix format (e.g., LH123456)
 * Supports:
 * - 10-digit plate format (e.g., 0220123456)
 * - 9-digit plate format (e.g., 220123456)
 * - Prefix format (e.g., LH123456)
 */
export function getCanonicalTag(tag: string, map: Record<string, string>): string {
  if (!tag) return '';
  const trimmed = tag.trim().toUpperCase();

  // If standard 10-digit license plate format (e.g. 0220123456)
  if (/^\d{10}$/.test(trimmed)) {
    const numericCode = trimmed.slice(0, 4);
    const serial = trimmed.slice(4);
    const prefix = map[numericCode];
    if (prefix) {
      return `${prefix}${serial}`;
    }
  }

  // If 9-digit plate where leading zero might be omitted (e.g. 220123456)
  if (/^\d{9}$/.test(trimmed)) {
    const numericCode = '0' + trimmed.slice(0, 3);
    const serial = trimmed.slice(3);
    const prefix = map[numericCode];
    if (prefix) {
      return `${prefix}${serial}`;
    }
  }

  return trimmed;
}

/**
 * Convert any tag to 10-digit standard plate format (e.g., 0220123456)
 */
export function get10DigitTag(tag: string, map: Record<string, string>): string | null {
  if (!tag) return null;
  const clean = tag.trim().toUpperCase();

  // If already 10 digits
  if (/^\d{10}$/.test(clean)) return clean;

  // If 9 digits, pad with a leading zero
  if (/^\d{9}$/.test(clean)) return '0' + clean;

  // If prefix format (e.g. LH123456)
  const match = clean.match(/^([A-Z]{2})(\d{6})$/);
  if (match) {
    const prefix = match[1];
    const serial = match[2];
    // Find numeric code for this prefix
    const numericCode = Object.keys(map).find(k => map[k] === prefix);
    if (numericCode) {
      return `${numericCode}${serial}`;
    }
  }

  return null;
}

/**
 * Intelligent Tag Matching
 * Checks if two tags match regardless of representing format
 */
export function matchTag(tagA: string, tagB: string, map: Record<string, string>): boolean {
  if (!tagA || !tagB) return false;
  const aClean = tagA.trim().toUpperCase();
  const bClean = tagB.trim().toUpperCase();

  if (aClean === bClean) return true;

  const aCanon = getCanonicalTag(aClean, map);
  const bCanon = getCanonicalTag(bClean, map);

  if (aCanon && bCanon && aCanon === bCanon) return true;

  const a10 = get10DigitTag(aClean, map);
  const b10 = get10DigitTag(bClean, map);

  if (a10 && b10 && a10 === b10) return true;

  // Partial match fallback
  if (aClean.includes(bClean) || bClean.includes(aClean)) return true;
  if (aCanon && aCanon.includes(bClean)) return true;
  if (bCanon && bCanon.includes(aClean)) return true;

  return false;
}
