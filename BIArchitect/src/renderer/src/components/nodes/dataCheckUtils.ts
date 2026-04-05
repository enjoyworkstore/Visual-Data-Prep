export type MatchType = 'includes' | 'exact' | 'not' | 'gt' | 'lt';

const FULL_WIDTH_CHARS = /[Ａ-Ｚａ-ｚ０-９．，－ー：／　]/g;
const FULL_WIDTH_OFFSET = 0xfee0;
const DATE_LIKE_PATTERN = /^\d{4}[/-]\d{1,2}[/-]\d{1,2}(?:[ tT]\d{1,2}:\d{2}(?::\d{2})?)?$/;

const toHalfWidthChar = (char: string): string => {
  if (char === '　') return ' ';
  if (char === 'ー') return '-';
  const code = char.charCodeAt(0);
  if (code >= 0xff01 && code <= 0xff5e) {
    return String.fromCharCode(code - FULL_WIDTH_OFFSET);
  }
  return char;
};

export const normalizeCompareText = (value: unknown): string => {
  if (value == null) return '';
  return String(value)
    .replace(FULL_WIDTH_CHARS, toHalfWidthChar)
    .trim()
    .toLowerCase();
};

const tryParseNumber = (value: unknown): number | null => {
  const normalized = normalizeCompareText(value).replace(/,/g, '');
  if (!normalized) return null;
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(normalized)) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
};

const tryParseDate = (value: unknown): number | null => {
  if (value instanceof Date) {
    const timestamp = value.getTime();
    return Number.isNaN(timestamp) ? null : timestamp;
  }

  const normalized = normalizeCompareText(value);
  if (!normalized || !DATE_LIKE_PATTERN.test(normalized)) return null;

  const timestamp = new Date(normalized.replace(/\//g, '-')).getTime();
  return Number.isNaN(timestamp) ? null : timestamp;
};

const getComparablePair = (left: unknown, right: unknown) => {
  const leftNumber = tryParseNumber(left);
  const rightNumber = tryParseNumber(right);
  if (leftNumber !== null && rightNumber !== null) {
    return { kind: 'number' as const, left: leftNumber, right: rightNumber };
  }

  const leftDate = tryParseDate(left);
  const rightDate = tryParseDate(right);
  if (leftDate !== null && rightDate !== null) {
    return { kind: 'date' as const, left: leftDate, right: rightDate };
  }

  return {
    kind: 'string' as const,
    left: normalizeCompareText(left),
    right: normalizeCompareText(right),
  };
};

export const getCheckOperatorLabel = (matchType?: string): string => {
  switch (matchType) {
    case 'gt':
      return '>';
    case 'lt':
      return '<';
    case 'exact':
      return '=';
    case 'not':
      return '≠';
    default:
      return 'inc';
  }
};

export const matchesCondition = (
  row: Record<string, unknown>,
  col: string,
  value: unknown,
  matchType: MatchType | string = 'includes'
): boolean => {
  const sourceValue = row?.[col];
  const comparable = getComparablePair(sourceValue, value);

  switch (matchType) {
    case 'exact':
      return comparable.left === comparable.right;
    case 'not':
      return comparable.left !== comparable.right;
    case 'gt':
      return comparable.left > comparable.right;
    case 'lt':
      return comparable.left < comparable.right;
    case 'includes':
    default:
      return normalizeCompareText(sourceValue).includes(normalizeCompareText(value));
  }
};
