import React, { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import {
  FileText, CheckCircle, AlertTriangle,
  Info, XCircle, Search, Download, Users, FileSpreadsheet,
  Eye, Filter, RefreshCw, Settings, AlertCircle, ChevronLeft, ChevronRight
} from 'lucide-react';

/* ------------------------------------------------------------------ */
/* الأنواع                                                             */
/* ------------------------------------------------------------------ */

declare global {
  interface Window {
    XLSX?: any;
  }
}

type CellValue = string | number | boolean | Date;
type Row = Record<string, CellValue | undefined>;
type ColumnMap = Record<string, string | null>;
type MatchStatus = 'EXACT' | 'MATCH_DIFF' | 'ONLY_F1' | 'ONLY_F2' | 'REVIEW';
type MatchType = 'NAT_ID' | 'ID' | 'NAME_BIRTH' | 'NAME' | 'FUZZY_HIGH' | 'REVIEW' | null;
type FilterKey = MatchStatus | 'ALL';
type SortKey = 'default' | 'name_asc' | 'name_desc';
type DiffType = 'DIFFERENT' | 'FORMAL';

interface LoadedFile {
  name: string;
  sheetName: string;
  rows: Row[];
  columns: ColumnInfo[];
  colMap: ColumnMap;
}

interface ColumnInfo {
  key: string;
  header: string;
  group: string;
  label: string;
}

interface ComparisonColumn {
  field: string;
  key1: string | null;
  key2: string | null;
}

interface FieldDiff {
  field: string;
  val1: string;
  val2: string;
  type: DiffType;
}

interface CompareRecord {
  status: MatchStatus;
  r1: Row | null;
  r2: Row | null;
  i1: number;
  i2: number;
  diffs: FieldDiff[];
  matchType: MatchType;
  score: number;
  search: string;
}

interface Stats {
  totalF1: number;
  totalF2: number;
  exact: number;
  diff: number;
  onlyF1: number;
  onlyF2: number;
  review: number;
}

interface StatCardConfig {
  key: MatchStatus;
  label: string;
  value: number;
  cls: string;
}

/* ------------------------------------------------------------------ */
/* ثوابت عامة                                                          */
/* ------------------------------------------------------------------ */

const XLSX_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const XLSX_SCRIPT_ID = 'xlsx-cdn-script';
const FUZZY_HIGH_CONFIDENCE = 98;
const XLSX_LOAD_TIMEOUT = 20000;
const MAX_FUZZY_CANDIDATES = 4000;

// الأعمدة القياسية التي يتم البحث عنها والمقارنة بينها
const STANDARD_COLUMNS: string[] = [
  "المحافظة", "موقع الوحدة الإدارية", "السوية التنظيمية الأولى", "السوية التنظيمية الثانية",
  "السوية التنظيمية الثالثة", "السوية التنظيمية الرابعة", "الرقم الوطني", "اسم الموظف الثلاثي",
  "التولد", "الصفة الوظيفية", "الراتب لآخر شهر", "ملاحظات", "رقم ID", "نمط الدوام"
];

// بدائل مقبولة لكل عمود قياسي (تتم مطابقتها بعد التطبيع النصي)
const COLUMN_ALIASES: Record<string, string[]> = {
  "المحافظة": ["المحافظة", "اسم المحافظة", "المحافه", "محافظة الموقع"],
  "موقع الوحدة الإدارية": ["موقع الوحدة الإدارية", "الوحدة الإدارية", "موقع الوحدة", "موقع الوحده"],
  "السوية التنظيمية الأولى": ["السوية التنظيمية الأولى", "السوية التنظيمية 1", "السوية الأولى", "السوية التنظيميه 1"],
  "السوية التنظيمية الثانية": ["السوية التنظيمية الثانية", "السوية التنظيمية 2", "السوية الثانية", "السوية التنظيميه 2"],
  "السوية التنظيمية الثالثة": ["السوية التنظيمية الثالثة", "السوية التنظيمية 3", "السوية الثالثة", "السوية التنظيميه 3"],
  "السوية التنظيمية الرابعة": ["السوية التنظيمية الرابعة", "السوية التنظيمية 4", "السوية الرابعة", "السوية التنظيميه 4"],
  "الرقم الوطني": ["الرقم الوطني", "الرقم الوطني إن وجد", "الرقم الوطني للتثبت", "الرقم الوطني للموظف"],
  "اسم الموظف الثلاثي": ["اسم الموظف الثلاثي", "اسم الموظف", "الاسم الثلاثي", "الاسم الرباعي", "الاسم", "اسم الموظف الرباعي"],
  "التولد": ["التولد", "تاريخ التولد", "تاريخ الميلاد", "الميلاد"],
  "الصفة الوظيفية": ["الصفة الوظيفية", "الصفة", "الصفة الوظيفيه"],
  "الراتب لآخر شهر": ["الراتب لآخر شهر", "الراتب لاخر شهر", "الراتب الشهري لآخر شهر", "الراتب"],
  "ملاحظات": ["ملاحظات", "الملاحظات", "ملاحظة"],
  "رقم ID": ["رقم ID", "رقم المعرف", "الرقم التعريفي", "ID", "رقم الهوية"],
  "نمط الدوام": ["نمط الدوام", "نمط دوام", "نظام الدوام", "الدوام", "نمط العمل"]
};

const STATUS_LABELS: Record<MatchStatus, string> = {
  EXACT: 'متطابق',
  MATCH_DIFF: 'يوجد اختلافات',
  ONLY_F1: 'في الأول فقط',
  ONLY_F2: 'في الثاني فقط',
  REVIEW: 'يحتاج مراجعة'
};

const PAGE_SIZES = [25, 50, 100, 250];

/* ------------------------------------------------------------------ */
/* أدوات التطبيع النصي                                                 */
/* ------------------------------------------------------------------ */

const DIGIT_MAP: Record<string, string> = (() => {
  const map: Record<string, string> = {};
  const ar = '٠١٢٣٤٥٦٧٨٩';
  const fa = '۰۱۲۳۴۵۶۷۸۹';
  for (let i = 0; i < 10; i += 1) {
    map[ar[i]] = String(i);
    map[fa[i]] = String(i);
  }
  return map;
})();

// تحويل أي قيمة قادمة من Excel إلى نص مستقر (مهم لعمود التولد)
const toPlainString = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return '';
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof value === 'boolean') return value ? 'نعم' : 'لا';
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : '';
  return String(value);
};

// توحيد الفواصل داخل الأرقام (1,200 → 1200 | 1.200,50 → 1200.50)
const normalizeNumberSeparators = (str: string): string => {
  if (!/\d/.test(str)) return str;
  const cleaned = str.replace(/[٬\s]/g, '');
  if (/^\d{1,3}(,\d{3})+$/.test(cleaned)) return cleaned.replace(/,/g, '');
  if (/^\d+,\d{1,2}$/.test(cleaned)) return cleaned.replace(',', '.');
  if (/^\d+(\.\d+)?$/.test(cleaned)) return cleaned;
  return str;
};

// دالة لتنظيف وتطبيع النصوص العربية لتسهيل المقارنة الذكية
const normalizeText = (text: unknown): string => {
  if (text === null || text === undefined) return '';
  let str = toPlainString(text).trim();
  str = str.replace(/\u00A0/g, ' ');
  str = str.replace(/[\u064B-\u0652\u0670]/g, ''); // إزالة التشكيل
  str = str.replace(/\u0640/g, ''); // إزالة التطويل ـــ
  str = str.replace(/[أإآٱ]/g, 'ا'); // توحيد الهمزات والألف
  str = str.replace(/ة/g, 'ه'); // توحيد التاء المربوطة والهاء
  str = str.replace(/[ىی]/g, 'ي'); // توحيد الياء والألف المقصورة
  str = str.replace(/ؤ/g, 'و');
  str = str.replace(/ئ/g, 'ي');
  str = str.replace(/[٠-٩۰-۹]/g, (d) => DIGIT_MAP[d] || d);
  str = normalizeNumberSeparators(str);
  str = str.replace(/\s+/g, ' ').trim();
  return str;
};

// نسخة صلبة تُستخدم لمطابقة أسماء الأعمدة فقط (بلا مساحات + أحرف لاتينية موحّدة)
//Lowercasing here is safe: it only affects Latin letters, and this helper is never used on values
const normalizeKey = (key: string): string => normalizeText(key).replace(/\s+/g, '').toLowerCase();

// كل الأسماء البديلة مجمّعة ومطبّعة مسبقاً (تُستخدم لكشف صف الرؤوس)
const NORMALIZED_ALIASES: string[][] = STANDARD_COLUMNS.map((col) =>
  Array.from(new Set([col, ...(COLUMN_ALIASES[col] || [])].map(normalizeKey).filter(Boolean)))
);
const ALL_ALIASES: string[] = NORMALIZED_ALIASES.reduce((acc, list) => acc.concat(list), []);

/* ------------------------------------------------------------------ */
/* مسافة ليفنشتاين محسّنة (سطران فقط + قطع مبكر)                        */
/* ------------------------------------------------------------------ */

// تُعيد المسافة إذا كانت <= maxDist، وتعيد -1 في غير ذلك
const levenshteinBounded = (a: string, b: string, maxDist: number): number => {
  const la = a.length;
  const lb = b.length;
  if (Math.abs(la - lb) > maxDist) return -1;
  if (la === 0) return lb <= maxDist ? lb : -1;
  if (lb === 0) return la <= maxDist ? la : -1;

  let prev = new Array(lb + 1);
  let cur = new Array(lb + 1);
  for (let j = 0; j <= lb; j += 1) prev[j] = j;

  const outOfBand = maxDist + 1;

  for (let i = 1; i <= la; i += 1) {
    cur[0] = i;
    const code = a.charCodeAt(i - 1);
    const from = Math.max(1, i - maxDist);
    const to = Math.min(lb, i + maxDist);
    let rowMin = i;

    for (let j = 1; j < from; j += 1) cur[j] = outOfBand;
    for (let j = from; j <= to; j += 1) {
      const cost = code === b.charCodeAt(j - 1) ? 0 : 1;
      const del = prev[j] + 1;
      const ins = cur[j - 1] + 1;
      const sub = prev[j - 1] + cost;
      const v = del < ins ? (del < sub ? del : sub) : (ins < sub ? ins : sub);
      cur[j] = v;
      if (v < rowMin) rowMin = v;
    }
    for (let j = to + 1; j <= lb; j += 1) cur[j] = outOfBand;

    if (rowMin > maxDist) return -1;
    const tmp = prev;
    prev = cur;
    cur = tmp;
  }

  const d = prev[lb];
  return d <= maxDist ? d : -1;
};

// نسبة التشابه، وتعيد 0 إذا كانت أقل من الحد المطلوب
const similarityAbove = (a: string, b: string, threshold: number): number => {
  if (a === b) return 100;
  const longerLen = a.length > b.length ? a.length : b.length;
  if (longerLen === 0) return 100;
  const maxDist = Math.ceil(((100 - threshold) / 100) * longerLen);
  const dist = levenshteinBounded(a, b, maxDist);
  if (dist < 0) return 0;
  const score = ((longerLen - dist) / longerLen) * 100;
  return score >= threshold ? score : 0;
};

/* ------------------------------------------------------------------ */
/* كشف صف الرؤوس وقراءة الورقة                                           */
/* ------------------------------------------------------------------ */

// عدد الصفوف التي نفحصها للعثور على صف الرؤوس (يغطي صفوف العناوين المدمجة)
const MAX_HEADER_SCAN = 30;

// درجة تقريبية لصف يُعتبر صف رؤوس (تطابق تام = 10، بادئة = 3، احتواء = 1)
const scoreHeaderRow = (row: CellValue[]): number => {
  let score = 0;
  for (let i = 0; i < row.length; i += 1) {
    const text = normalizeKey(toPlainString(row[i]));
    if (!text) continue;
    if (ALL_ALIASES.indexOf(text) !== -1) { score += 10; continue; }
    if (ALL_ALIASES.some((a) => a.length > 2 && text.startsWith(a))) { score += 3; continue; }
    if (ALL_ALIASES.some((a) => a.length > 2 && text.includes(a))) score += 1;
  }
  return score;
};

const detectHeaderRowIndex = (grid: CellValue[][]): number => {
  const limit = Math.min(grid.length, MAX_HEADER_SCAN);
  let bestIndex = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i += 1) {
    const score = scoreHeaderRow(grid[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestScore > 0 ? bestIndex : 0;
};

/**
 * قراءة الورقة مع تجاهل أي صفوف عناوين مسبقة (مثل الصف التجميعي المدمج).
 * لا نعتمد على sheet_to_json في كشف الرؤوس لأنه يلتقط الصف الأول دائماً.
 */
const readSheetRows = (XLSX: any, ws: any): { rows: Row[]; columns: ColumnInfo[]; headerScore: number } => {
  const grid: CellValue[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, raw: true });
  if (!grid || grid.length === 0) return { rows: [], columns: [], headerScore: 0 };

  const headerIndex = detectHeaderRowIndex(grid);
  const headerRow = grid[headerIndex];
  const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { c: 0, r: 0 }, e: { c: headerRow.length - 1, r: grid.length - 1 } };
  const width = Math.max(headerRow.length, range.e.c - range.s.c + 1);
  const headerAbsoluteRow = range.s.r + headerIndex;

  // Keep the leaf header as the data key and attach merged parent headers as context.
  const keys: string[] = [];
  const columns: ColumnInfo[] = [];
  const occurrences = new Map<string, number>();
  const groupsByColumn = new Map<number, Array<{ row: number; label: string }>>();
  (ws['!merges'] || []).forEach((merge: any) => {
    if (merge.e.r >= headerAbsoluteRow) return;
    const parentCell = ws[XLSX.utils.encode_cell(merge.s)];
    const parentLabel = parentCell ? toPlainString(parentCell.v ?? parentCell.w).trim() : '';
    if (!parentLabel) return;
    for (let column = Math.max(merge.s.c, range.s.c); column <= Math.min(merge.e.c, range.e.c); column += 1) {
      const groups = groupsByColumn.get(column) || [];
      groups.push({ row: merge.s.r, label: parentLabel });
      groupsByColumn.set(column, groups);
    }
  });

  for (let idx = 0; idx < width; idx += 1) {
    const header = toPlainString(headerRow[idx]).trim();
    const physicalColumn = range.s.c + idx;
    const name = header || `عمود ${physicalColumn + 1}`;
    const seen = occurrences.get(name) || 0;
    occurrences.set(name, seen + 1);
    const key = seen === 0 ? name : `${name} (${seen + 1})`;
    const group = Array.from(new Map(
      (groupsByColumn.get(physicalColumn) || [])
        .sort((a, b) => a.row - b.row)
        .map((item) => [item.row, item.label])
    ).values()).join(' / ');
    keys.push(key);
    columns.push({ key, header: header || name, group, label: group ? `${group} › ${header || name}` : header || name });
  }

  const rows: Row[] = [];
  for (let r = headerIndex + 1; r < grid.length; r += 1) {
    const source = grid[r];
    const row: Row = {};
    let hasValue = false;
    for (let c = 0; c < keys.length; c += 1) {
      const v = source[c];
      row[keys[c]] = v === undefined || v === null ? '' : v;
      if (v !== undefined && v !== null && v !== '') hasValue = true;
    }
    if (hasValue) rows.push(row);
  }
  return { rows, columns, headerScore: scoreHeaderRow(headerRow) };
};

/* ------------------------------------------------------------------ */
/* حل أسماء الأعمدة مرة واحدة لكل ملف                                  */
/* ------------------------------------------------------------------ */

const buildColumnMap = (rows: Row[]): ColumnMap => {
  const pool = new Map<string, string>(); // normalized key -> original key
  const keys: string[] = rows.length ? Object.keys(rows[0]) : [];
  for (let i = 1; i < rows.length; i += 1) Object.keys(rows[i]).forEach((k) => keys.push(k));
  keys.forEach((k) => {
    const nk = normalizeKey(k);
    if (nk && !pool.has(nk)) pool.set(nk, k);
  });

  const entries = Array.from(pool.entries());
  const used = new Set<string>();
  const map: ColumnMap = {};

  const claim = (key: string | undefined | null): string | null => {
    if (key && !used.has(key)) {
      used.add(key);
      return key;
    }
    return null;
  };

  const tiers: Array<{ col: string; names: string[] }> = [];
  STANDARD_COLUMNS.forEach((col) => {
    const names = [col, ...(COLUMN_ALIASES[col] || [])].map(normalizeKey).filter(Boolean);
    tiers.push({ col, names: Array.from(new Set(names)) });
  });

  // 1) مطابقة تامة لاسم العمود (أو أحد أسمائه)
  tiers.forEach(({ col, names }) => {
    const hit = entries.find(([nk]) => nk === names.find((a) => a === nk));
    map[col] = hit ? claim(hit[1]) : null;
  });

  // 2) بادئة العمود
  tiers.forEach(({ col, names }) => {
    if (map[col]) return;
    const hit = entries.find(([nk, orig]) => !used.has(orig) && names.some((a) => a.length > 2 && nk.startsWith(a)));
    map[col] = hit ? claim(hit[1]) : null;
  });

  // 3) يحتوي اسم العمود
  tiers.forEach(({ col, names }) => {
    if (map[col]) return;
    const hit = entries.find(([nk, orig]) => !used.has(orig) && names.some((a) => a.length > 2 && nk.includes(a)));
    map[col] = hit ? claim(hit[1]) : null;
  });

  return map;
};

const buildComparisonColumns = (file1: LoadedFile, file2: LoadedFile): ComparisonColumn[] => {
  const columns: ComparisonColumn[] = STANDARD_COLUMNS.map((field) => ({
    field,
    key1: file1.colMap[field],
    key2: file2.colMap[field]
  }));
  const standardKeys1 = new Set(Object.values(file1.colMap).filter((key): key is string => key !== null));
  const standardKeys2 = new Set(Object.values(file2.colMap).filter((key): key is string => key !== null));
  const getExtras = (file: LoadedFile, standardKeys: Set<string>) => file.columns
    .filter(({ key }) => !standardKeys.has(key));
  const extras1 = getExtras(file1, standardKeys1);
  const extras2 = getExtras(file2, standardKeys2);
  const duplicateHeaders = new Set<string>();

  [extras1, extras2].forEach((extras) => {
    const counts = new Map<string, number>();
    extras.forEach(({ header }) => {
      const normalized = normalizeKey(header);
      if (normalized) counts.set(normalized, (counts.get(normalized) || 0) + 1);
    });
    counts.forEach((count, header) => { if (count > 1) duplicateHeaders.add(header); });
  });

  const toIdentity = ({ header, group }: ColumnInfo) => {
    const normalizedHeader = normalizeKey(header);
    return duplicateHeaders.has(normalizedHeader)
      ? normalizeKey(`${group} ${header}`)
      : normalizedHeader;
  };
  const extraKeys1 = new Map<string, ColumnInfo>();
  const extraKeys2 = new Map<string, ColumnInfo>();
  extras1.forEach((column) => {
    const identity = toIdentity(column);
    if (identity && !extraKeys1.has(identity)) extraKeys1.set(identity, column);
  });
  extras2.forEach((column) => {
    const identity = toIdentity(column);
    if (identity && !extraKeys2.has(identity)) extraKeys2.set(identity, column);
  });

  const extraFields = new Set([...extraKeys1.keys(), ...extraKeys2.keys()]);
  extraFields.forEach((identity) => {
    const column1 = extraKeys1.get(identity);
    const column2 = extraKeys2.get(identity);
    columns.push({
      field: column1?.label || column2?.label || identity,
      key1: column1?.key || null,
      key2: column2?.key || null
    });
  });

  return columns;
};

const pickValue = (row: Row | null | undefined, colMap: ColumnMap | null, col: string): CellValue => {
  if (!row || !colMap) return '';
  const key = colMap[col];
  if (!key) return '';
  const v = row[key];
  return v === undefined || v === null ? '' : v;
};

const extractIdentity = (row: Row, colMap: ColumnMap) => ({
  natId: normalizeText(pickValue(row, colMap, 'الرقم الوطني')),
  id: normalizeText(pickValue(row, colMap, 'رقم ID')),
  name: normalizeText(pickValue(row, colMap, 'اسم الموظف الثلاثي')),
  birth: normalizeText(pickValue(row, colMap, 'التولد'))
});

const hasEmployeeIdentity = (row: Row, colMap: ColumnMap): boolean => {
  const identity = extractIdentity(row, colMap);
  return Boolean(identity.natId || identity.name || identity.id);
};

/* ------------------------------------------------------------------ */
/* محرك المقارنة                                                        */
/* ------------------------------------------------------------------ */

const runComparison = (
  file1: LoadedFile,
  file2: LoadedFile,
  threshold: number,
  comparisonColumns: ComparisonColumn[]
): { results: CompareRecord[]; stats: Stats } => {
  const rows1 = file1.rows.filter((row) => hasEmployeeIdentity(row, file1.colMap));
  const rows2 = file2.rows.filter((row) => hasEmployeeIdentity(row, file2.colMap));
  const cm1 = file1.colMap;
  const cm2 = file2.colMap;

  // فهرسة بيانات الملف الأول لتسريع البحث
  const byNatId = new Map<string, number[]>();
  const byId = new Map<string, number[]>();
  const byNameBirth = new Map<string, number[]>();
  const byName = new Map<string, number[]>();
  const nameBuckets = new Map<string, number[]>(); // أول حرفين -> فهارس

  const add = (map: Map<string, number[]>, key: string, i: number) => {
    if (!key) return;
    const list = map.get(key);
    if (list) list.push(i);
    else map.set(key, [i]);
  };

  const identities1 = new Array(rows1.length);
  const flags = new Uint8Array(rows1.length).fill(1);

  for (let i = 0; i < rows1.length; i += 1) {
    const ident = extractIdentity(rows1[i], cm1);
    identities1[i] = ident;
    add(byNatId, ident.natId, i);
    add(byId, ident.id, i);
    add(byNameBirth, ident.name && ident.birth ? `${ident.name}_${ident.birth}` : '', i);
    add(byName, ident.name, i);
    if (ident.name) {
      const bucketKey = ident.name.length > 1 ? ident.name.slice(0, 2) : ident.name.slice(0, 1);
      const list = nameBuckets.get(bucketKey);
      if (list) list.push(i);
      else nameBuckets.set(bucketKey, [i]);
    }
  }

  const firstUnmatched = (list: number[] | undefined): number => {
    if (!list) return -1;
    for (let k = 0; k < list.length; k += 1) {
      if (flags[list[k]]) return list[k];
    }
    return -1;
  };

  const getCandidates = (name: string): number[] => {
    const key2 = name.length > 1 ? name.slice(0, 2) : name.slice(0, 1);
    const primary = nameBuckets.get(key2) || [];
    if (primary.length >= 50 || primary.length >= MAX_FUZZY_CANDIDATES) return primary;
    const primarySet = new Set(primary);
    const extra = nameBuckets.get(name.slice(0, 1)) || [];
    const merged = primary.slice();
    for (let k = 0; k < extra.length && merged.length < MAX_FUZZY_CANDIDATES; k += 1) {
      if (!primarySet.has(extra[k])) merged.push(extra[k]);
    }
    return merged;
  };

  const results: CompareRecord[] = [];
  const stats: Stats = {
    totalF1: rows1.length,
    totalF2: rows2.length,
    exact: 0,
    diff: 0,
    onlyF1: 0,
    onlyF2: 0,
    review: 0
  };

  for (let i2 = 0; i2 < rows2.length; i2 += 1) {
    const r2 = rows2[i2];
    const ident2 = extractIdentity(r2, cm2);

    let i1 = -1;
    let matchType: MatchType = null;
    let matchScore = 100;

    if (ident2.natId) i1 = firstUnmatched(byNatId.get(ident2.natId));
    if (i1 < 0 && ident2.id) i1 = firstUnmatched(byId.get(ident2.id));
    if (i1 < 0 && ident2.name && ident2.birth) i1 = firstUnmatched(byNameBirth.get(`${ident2.name}_${ident2.birth}`));

    if (i1 < 0 && ident2.name) {
      const exactIdx = firstUnmatched(byName.get(ident2.name));
      if (exactIdx >= 0) {
        i1 = exactIdx;
        matchType = 'NAME';
      } else {
        // المطابقة التقريبية على مرشحين محصورين (سريع)
        const candidates = getCandidates(ident2.name);
        let best = -1;
        let bestScore = 0;
        for (let k = 0; k < candidates.length; k += 1) {
          const idx = candidates[k];
          if (!flags[idx]) continue;
          const score = similarityAbove(ident2.name, identities1[idx].name, threshold);
          if (score > bestScore) {
            bestScore = score;
            best = idx;
          }
        }
        if (best >= 0) {
          i1 = best;
          matchScore = bestScore;
          matchType = bestScore >= FUZZY_HIGH_CONFIDENCE ? 'FUZZY_HIGH' : 'REVIEW';
        }
      }
    }

    if (i1 < 0) {
      stats.onlyF2 += 1;
      const identSearch = `${ident2.name} ${ident2.natId} ${ident2.id}`;
      results.push({
        status: 'ONLY_F2', r1: null, r2, i1: -1, i2,
        diffs: [], matchType: null, score: 0, search: normalizeText(identSearch)
      });
      continue;
    }

    if (!matchType) {
      if (ident2.natId) matchType = 'NAT_ID';
      else if (ident2.id) matchType = 'ID';
      else matchType = 'NAME_BIRTH';
    }

    flags[i1] = 0;
    const r1 = rows1[i1];

    // تحليل الاختلافات بين السجلين
    const diffs: FieldDiff[] = [];
    let isExact = true;
    for (let c = 0; c < comparisonColumns.length; c += 1) {
      const { field, key1, key2 } = comparisonColumns[c];
      const val1 = key1 ? r1[key1] ?? '' : '';
      const val2 = key2 ? r2[key2] ?? '' : '';
      const nVal1 = normalizeText(val1);
      const nVal2 = normalizeText(val2);

      if (nVal1 !== nVal2) {
        isExact = false;
        diffs.push({ field, val1: toPlainString(val1), val2: toPlainString(val2), type: 'DIFFERENT' });
      } else {
        // فرق شكلي فقط (تنسيق/مسافات/فواصل): يُسجَّل للعرض لكنه لا يغيّر حالة السجل
        const raw1 = toPlainString(val1).trim();
        const raw2 = toPlainString(val2).trim();
        if (raw1 !== raw2) {
          diffs.push({ field, val1: raw1, val2: raw2, type: 'FORMAL' });
        }
      }
    }

    if (matchType === 'REVIEW') stats.review += 1;
    else if (isExact) stats.exact += 1;
    else stats.diff += 1;

    const ident1 = identities1[i1];
    results.push({
      status: matchType === 'REVIEW' ? 'REVIEW' : isExact ? 'EXACT' : 'MATCH_DIFF',
      r1, r2, i1, i2, diffs, matchType, score: matchScore,
      search: normalizeText(`${ident1.name} ${ident1.natId} ${ident1.id}`)
    });
  }

  // الموظفون المتبقون في الملف الأول (غير موجودين في الثاني)
  for (let i = 0; i < rows1.length; i += 1) {
    if (!flags[i]) continue;
    stats.onlyF1 += 1;
    const ident1 = identities1[i];
    results.push({
      status: 'ONLY_F1', r1: rows1[i], r2: null, i1: i, i2: -1,
      diffs: [], matchType: null, score: 0,
      search: normalizeText(`${ident1.name} ${ident1.natId} ${ident1.id}`)
    });
  }

  return { results, stats };
};

/* ------------------------------------------------------------------ */
/* مكونات مساعدة                                                       */
/* ------------------------------------------------------------------ */

const StatusBadge = ({ status }: { status: MatchStatus }) => {
  switch (status) {
    case 'EXACT':
      return <span className="px-2 py-1 bg-green-100 text-green-700 rounded-md text-xs font-bold inline-flex items-center gap-1"><CheckCircle size={14} /> متطابق</span>;
    case 'MATCH_DIFF':
      return <span className="px-2 py-1 bg-yellow-100 text-yellow-700 rounded-md text-xs font-bold inline-flex items-center gap-1"><AlertTriangle size={14} /> يوجد اختلافات</span>;
    case 'ONLY_F1':
      return <span className="px-2 py-1 bg-red-100 text-red-700 rounded-md text-xs font-bold inline-flex items-center gap-1"><XCircle size={14} /> في الأول فقط</span>;
    case 'ONLY_F2':
      return <span className="px-2 py-1 bg-blue-100 text-blue-700 rounded-md text-xs font-bold inline-flex items-center gap-1"><Info size={14} /> في الثاني فقط</span>;
    case 'REVIEW':
      return <span className="px-2 py-1 bg-orange-100 text-orange-700 rounded-md text-xs font-bold inline-flex items-center gap-1"><AlertCircle size={14} /> يحتاج مراجعة</span>;
    default:
      return <span className="px-2 py-1 bg-slate-100 text-slate-600 rounded-md text-xs font-bold">غير معروف</span>;
  }
};

const getFieldValue = (row: Row | null | undefined, colMap: ColumnMap | null, col: string): string => {
  if (!row || !colMap) return '-';
  const key = colMap[col];
  if (!key) return '-';
  const v = row[key];
  if (v === undefined || v === null || v === '') return '-';
  return toPlainString(v);
};

const getRawFieldValue = (row: Row | null | undefined, key: string | null): string => {
  if (!row || !key) return '-';
  const value = row[key];
  return value === undefined || value === null || value === '' ? '-' : toPlainString(value);
};

interface ResultRowProps {
  r: CompareRecord;
  cm1: ColumnMap;
  cm2: ColumnMap;
  onSelect: (r: CompareRecord) => void;
}

const ResultRow = React.memo(function ResultRow({ r, cm1, cm2, onSelect }: ResultRowProps) {
  const rRef = r.r1 || r.r2;
  const refMap = r.r1 ? cm1 : cm2;
  return (
    <tr className="hover:bg-slate-50 transition-colors">
      <td className="px-6 py-3 whitespace-nowrap"><StatusBadge status={r.status} /></td>
      <td className="px-6 py-3 font-medium text-slate-700">{getFieldValue(rRef, refMap, 'الرقم الوطني')}</td>
      <td className="px-6 py-3 font-bold text-blue-900">{getFieldValue(rRef, refMap, 'اسم الموظف الثلاثي')}</td>
      <td className="px-6 py-3 text-slate-500 whitespace-nowrap">
        {getFieldValue(rRef, refMap, 'الصفة الوظيفية')} - {getFieldValue(rRef, refMap, 'موقع الوحدة الإدارية')}
      </td>
      <td className="px-6 py-3 text-center">
        <button
          onClick={() => onSelect(r)}
          className="inline-flex items-center gap-1 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-slate-600 hover:bg-blue-50 hover:text-blue-600 hover:border-blue-300 transition-colors font-medium text-xs">
          <Eye size={16} /> التفاصيل
        </button>
      </td>
    </tr>
  );
});

interface RecordModalProps {
  record: CompareRecord;
  cm1: ColumnMap;
  cm2: ColumnMap;
  columns: ComparisonColumn[];
  onClose: () => void;
}

function RecordModal({ record, cm1, cm2, columns, onClose }: RecordModalProps) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('keydown', onKey);
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = prevOverflow;
    };
  }, [onClose]);

  const diffMap = useMemo(() => {
    const map: Record<string, DiffType> = {};
    (record.diffs || []).forEach((d) => { map[d.field] = d.type; });
    return map;
  }, [record]);

  const diffCount = (record.diffs || []).filter((d) => d.type === 'DIFFERENT').length;

  const card = (row: Row | null, colMap: ColumnMap, title: string, active: string, tint: string) => {
    const isMissing = !row;
    return (
      <div className={`bg-white border rounded-xl overflow-hidden shadow-sm ${isMissing ? 'opacity-50' : active}`}>
        <div className={`px-4 py-3 border-b ${tint}`}>
          <h4 className={`font-bold flex items-center gap-2 ${isMissing ? 'text-slate-500' : 'text-slate-800'}`}>
            <FileText size={18} /> {title}
          </h4>
        </div>
        {isMissing ? (
          <div className="p-8 text-center text-slate-400 font-medium">غير موجود</div>
        ) : (
          <table className="w-full text-sm">
            <tbody className="divide-y divide-slate-100">
              {columns.map(({ field, key1, key2 }) => {
                const type = diffMap[field];
                const isDiff = type === 'DIFFERENT';
                const isFormal = type === 'FORMAL';
                const rowClass = isDiff ? 'bg-red-50' : isFormal ? 'bg-yellow-50' : '';
                const textClass = isDiff ? 'text-red-700' : isFormal ? 'text-yellow-700' : 'text-slate-800';
                return (
                  <tr key={field} className={rowClass}>
                    <td className="px-4 py-2 font-bold text-slate-600 w-1/3 bg-slate-50/50">{field}</td>
                    <td className={`px-4 py-2 font-medium ${textClass}`}>
                      {STANDARD_COLUMNS.includes(field)
                        ? getFieldValue(row, colMap, field)
                        : getRawFieldValue(row, row === record.r1 ? key1 : key2)}
                      {isFormal && <span className="mr-2 text-[10px] bg-yellow-200 text-yellow-800 px-1 rounded inline-block">اختلاف شكلي</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    );
  };

  return (
    <div
      className="fixed inset-0 bg-slate-900/60 flex items-center justify-center z-50 p-4"
      role="dialog"
      aria-modal="true"
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden">
        <div className="bg-slate-50 px-6 py-4 border-b flex items-center justify-between">
          <div>
            <h3 className="font-bold text-lg text-slate-800 flex items-center gap-2">
              تفاصيل مقارنة السجل
              <StatusBadge status={record.status} />
            </h3>
            <p className="text-sm text-slate-500 mt-1">
              الرقم الوطني: {getFieldValue(record.r1 || record.r2, record.r1 ? cm1 : cm2, 'الرقم الوطني')}
              {diffCount > 0 && <span className="text-red-600 font-bold"> — {diffCount} حقل مختلف</span>}
            </p>
          </div>
          <button onClick={onClose} aria-label="إغلاق" className="text-slate-400 hover:text-red-500 transition-colors bg-white p-2 rounded-full shadow-sm">
            <XCircle size={24} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 bg-slate-100/50">
          <div className="grid md:grid-cols-2 gap-6">
            {card(record.r1, cm1, 'بيانات الملف الأول (المرجعي)', 'border-blue-200', 'bg-blue-50 border-blue-100')}
            {card(record.r2, cm2, 'بيانات الملف الثاني (الجديد)', 'border-indigo-200', 'bg-indigo-50 border-indigo-100')}
          </div>
        </div>
      </div>
    </div>
  );
}

interface StatCardProps {
  label: string;
  value: string | number;
  onClick?: () => void;
  active?: boolean;
  tone?: string;
}

const StatCard = ({ label, value, onClick, active, tone = '' }: StatCardProps) => {
  const base = `p-4 rounded-xl border shadow-sm text-center transition-colors ${tone}`;
  if (!onClick) {
    return (
      <div className={`${base} bg-white border-slate-200`}>
        <p className="text-xs font-bold text-slate-500 mb-1">{label}</p>
        <p className="text-2xl font-black text-slate-800">{value}</p>
      </div>
    );
  }
  return (
    <button type="button" onClick={onClick} className={`${base} ${active ? 'ring-2 ring-offset-1 ring-blue-500' : ''}`}>
      <p className="text-xs font-bold mb-1">{label}</p>
      <p className="text-2xl font-black">{value}</p>
    </button>
  );
};

/* ------------------------------------------------------------------ */
/* التطبيق                                                             */
/* ------------------------------------------------------------------ */

export default function App() {
  const [isXlsxLoaded, setIsXlsxLoaded] = useState<boolean>(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // تحميل مكتبة إكسل ديناميكياً لتجنب مشاكل بناء المشروع
  useEffect(() => {
    if (window.XLSX) {
      setIsXlsxLoaded(true);
      return undefined;
    }
    let script = document.getElementById(XLSX_SCRIPT_ID) as HTMLScriptElement | null;
    if (!script) {
      script = document.createElement('script');
      script.id = XLSX_SCRIPT_ID;
      script.src = XLSX_CDN;
      script.async = true;
      document.head.appendChild(script);
    }
    const onLoad = () => setIsXlsxLoaded(true);
    const onError = () => setErrorMsg("فشل تحميل مكتبة التعامل مع Excel. تحقق من اتصالك بالإنترنت.");
    const timer = window.setTimeout(() => {
      if (!window.XLSX) setErrorMsg("انتهت مهلة تحميل مكتبة Excel. حدّث الصفحة أو تحقق من الإنترنت.");
    }, XLSX_LOAD_TIMEOUT);
    script.addEventListener('load', onLoad);
    script.addEventListener('error', onError);
    return () => {
      window.clearTimeout(timer);
      script.removeEventListener('load', onLoad);
      script.removeEventListener('error', onError);
    };
  }, []);

  const [file1, setFile1] = useState<LoadedFile | null>(null);
  const [file2, setFile2] = useState<LoadedFile | null>(null);
  const [selectedColumnFields, setSelectedColumnFields] = useState<string[]>([]);
  const availableComparisonColumns = useMemo(
    () => file1 && file2 ? buildComparisonColumns(file1, file2) : [],
    [file1, file2]
  );
  const selectedComparisonColumns = useMemo(
    () => availableComparisonColumns.filter(({ field }) => selectedColumnFields.includes(field)),
    [availableComparisonColumns, selectedColumnFields]
  );

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [results, setResults] = useState<CompareRecord[] | null>(null);
  const [stats, setStats] = useState<Stats | null>(null);

  const [filter, setFilter] = useState<FilterKey>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedRecord, setSelectedRecord] = useState<CompareRecord | null>(null);
  const [fuzzyThreshold, setFuzzyThreshold] = useState<number>(85);

  const [page, setPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(25);
  const [sortKey, setSortKey] = useState<SortKey>('default');
  const timerRef = useRef<number | null>(null);

  useEffect(() => () => { if (timerRef.current !== null) window.clearTimeout(timerRef.current); }, []);
  useEffect(() => {
    if (file1 && file2) {
      setSelectedColumnFields(availableComparisonColumns.map(({ field }) => field));
    } else {
      setSelectedColumnFields([]);
    }
  }, [file1, file2, availableComparisonColumns]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, fileNum: 1 | 2) => {
    setErrorMsg(null);
    const input = e.target;
    const file = input.files && input.files[0];
    input.value = ''; // يسمح باختيار نفس الملف مرة أخرى بعد الخطأ
    if (!file) return;

    const lowerName = file.name.toLowerCase();
    if (!/\.(xlsx|xlsm|xls|csv)$/.test(lowerName)) {
      setErrorMsg("يرجى رفع ملف بصيغة Excel (.xlsx أو .xls)");
      return;
    }
    if (!isXlsxLoaded) {
      setErrorMsg("مكتبة Excel لم تُحمّل بعد. انتظر قليلاً ثم أعد المحاولة.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        if (!XLSX) throw new Error('XLSX غير متاح');
        const buffer = evt.target ? evt.target.result : null;
        if (!buffer) throw new Error('لا توجد بيانات');
        const wb = XLSX.read(new Uint8Array(buffer as ArrayBuffer), { type: 'array', cellDates: true });
        if (!wb.SheetNames.length) throw new Error('لا توجد أوراق عمل');

        const sheetNames: string[] = wb.SheetNames;
        const candidates = sheetNames.map((sheetName) => {
          const parsed = readSheetRows(XLSX, wb.Sheets[sheetName]);
          const colMap = buildColumnMap(parsed.rows);
          const recognizedCount = Object.values(colMap).filter(Boolean).length;
          const hasIdentity = Boolean(colMap['الرقم الوطني'] || colMap['رقم ID'] || colMap['اسم الموظف الثلاثي']);
          return {
            sheetName,
            ...parsed,
            colMap,
            score: recognizedCount * 1000 + (hasIdentity ? 10000 : 0) + Math.min(parsed.rows.length, 999) + parsed.headerScore
          };
        });
        const selected = candidates.reduce((best, candidate) => candidate.score > best.score ? candidate : best);
        const { sheetName, rows, columns, colMap } = selected;

        if (!rows.length) {
          setErrorMsg(`الملف ${file.name} فارغ أو لا يحتوي على بيانات.`);
          return;
        }

        if (!colMap['الرقم الوطني'] && !colMap['رقم ID'] && !colMap['اسم الموظف الثلاثي']) {
          setErrorMsg(`تعذر العثور على ورقة بيانات تحتوي أعمدة تعريف (الرقم الوطني / رقم ID / اسم الموظف) في الملف ${file.name}.`);
          return;
        }

        const payload: LoadedFile = { name: file.name, sheetName, rows, columns, colMap };
        if (fileNum === 1) setFile1(payload);
        else setFile2(payload);
      } catch (err) {
        console.error(err);
        setErrorMsg("حدث خطأ أثناء قراءة الملف. قد يكون تالفاً أو غير مدعوم.");
      }
    };
    reader.onerror = () => setErrorMsg("تعذر قراءة الملف من الجهاز.");
    reader.readAsArrayBuffer(file);
  };

  const resetAll = () => {
    setResults(null);
    setStats(null);
    setFile1(null);
    setFile2(null);
    setFilter('ALL');
    setSearchQuery('');
    setSelectedRecord(null);
    setPage(1);
    setSortKey('default');
  };

  const editComparisonSettings = () => {
    setResults(null);
    setStats(null);
    setFilter('ALL');
    setSearchQuery('');
    setSelectedRecord(null);
    setPage(1);
    setSortKey('default');
  };

  const toggleComparisonColumn = (field: string) => {
    setSelectedColumnFields((current) => current.includes(field)
      ? current.filter((selected) => selected !== field)
      : [...current, field]);
  };

  const processComparison = () => {
    if (!file1 || !file2) {
      setErrorMsg("يجب رفع الملفين أولاً لبدء المقارنة.");
      return;
    }
    if (selectedComparisonColumns.length === 0) {
      setErrorMsg("اختر عموداً واحداً على الأقل لمقارنة بياناته.");
      return;
    }
    setErrorMsg(null);
    setSelectedRecord(null);
    setIsProcessing(true);

    // تأخير بسيط يتيح للواجهة عرض حالة التحميل قبل بدء المعالجة الثقيلة
    if (timerRef.current !== null) window.clearTimeout(timerRef.current);
    timerRef.current = window.setTimeout(() => {
      try {
        const { results: res, stats: st } = runComparison(file1, file2, fuzzyThreshold, selectedComparisonColumns);
        setResults(res);
        setStats(st);
        setFilter('ALL');
        setSearchQuery('');
        setPage(1);
        setSortKey('default');
      } catch (err) {
        console.error(err);
        setErrorMsg("حدث خطأ أثناء معالجة البيانات والمقارنة.");
      } finally {
        setIsProcessing(false);
        timerRef.current = null;
      }
    }, 60);
  };

  const exportToExcel = () => {
    const XLSX = window.XLSX;
    if (!results || !stats || !XLSX || !file1 || !file2) return;
    setErrorMsg(null);

    try {
      const wb = XLSX.utils.book_new();
      const cm1 = file1.colMap;
      const cm2 = file2.colMap;
      const comparisonColumns = selectedComparisonColumns;

      // 1. شيت الملخص
      const summaryData = [
        ['الإحصائية', 'القيمة'],
        ['إجمالي الملف الأول (المرجعي)', stats.totalF1],
        ['إجمالي الملف الثاني (الجديد)', stats.totalF2],
        ['متطابق تماماً', stats.exact],
        ['متطابق مع اختلافات', stats.diff],
        ['موجود في الأول فقط (محذوف)', stats.onlyF1],
        ['موجود في الثاني فقط (جديد)', stats.onlyF2],
        ['يحتاج مراجعة (تطابق تقريبي)', stats.review]
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), 'الملخص والإحصائيات');

      // دالة تنسيق السجل للتصدير
      const formatRecordForExport = (r: CompareRecord): Record<string, string> => {
        const base: Record<string, string> = {
          'الحالة': STATUS_LABELS[r.status] || r.status,
          'طريقة المطابقة': r.matchType || '-',
          'نسبة التشابه': r.matchType ? String(Math.round(r.score)) : '',
          'الرقم الوطني': toPlainString(pickValue(r.r1 || r.r2, r.r1 ? cm1 : cm2, 'الرقم الوطني')),
          'اسم الموظف': toPlainString(pickValue(r.r1 || r.r2, r.r1 ? cm1 : cm2, 'اسم الموظف الثلاثي')),
          'الحقول المختلفة': String((r.diffs || []).filter((d) => d.type === 'DIFFERENT').length),
          'قائمة الاختلافات': (r.diffs || []).map((d) => `${d.field}: "${d.val1}" ≠ "${d.val2}"`).join(' | ')
        };
        comparisonColumns.forEach(({ field, key1, key2 }) => {
          base[`${field} (الأول)`] = key1 && r.r1 ? toPlainString(r.r1[key1]) : '';
          base[`${field} (الثاني)`] = key2 && r.r2 ? toPlainString(r.r2[key2]) : '';
        });
        return base;
      };

      const appendJsonSheet = (name: string, data: Array<Record<string, string>>) => {
        if (!data.length) return;
        const ws = XLSX.utils.json_to_sheet(data);
        const headerCount = Object.keys(data[0]).length;
        ws['!cols'] = Object.keys(data[0]).map((k) => ({ wch: Math.min(40, Math.max(12, k.length + 4)) }));
        if (headerCount > 0) {
          ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length, c: headerCount - 1 } }) };
        }
        XLSX.utils.book_append_sheet(wb, ws, name);
      };

      // 2. شيت جميع البيانات
      appendJsonSheet('جميع النتائج التفصيلية', results.map(formatRecordForExport));
      // 3. شيت الاختلافات والمراجعة
      appendJsonSheet('السجلات المختلفة والمراجعة', results.filter((r) => r.status === 'MATCH_DIFF' || r.status === 'REVIEW').map(formatRecordForExport));
      // 4. شيت المفقودين
      appendJsonSheet('سجلات محذوفة (بالأول فقط)', results.filter((r) => r.status === 'ONLY_F1').map(formatRecordForExport));
      // 5. شيت الجدد
      appendJsonSheet('سجلات جديدة (بالثاني فقط)', results.filter((r) => r.status === 'ONLY_F2').map(formatRecordForExport));

      XLSX.writeFile(wb, 'تقرير_مقارنة_الموظفين.xlsx');
    } catch (err) {
      console.error(err);
      setErrorMsg("حدث خطأ أثناء محاولة التصدير لملف إكسل.");
    }
  };

  const filteredResults = useMemo<CompareRecord[]>(() => {
    if (!results) return [];
    let temp = results;
    if (filter !== 'ALL') temp = temp.filter((r) => r.status === filter);

    const q = normalizeText(searchQuery);
    if (q) temp = temp.filter((r) => r.search.includes(q));

    if (sortKey !== 'default' && file1 && file2) {
      const dir = sortKey === 'name_asc' ? 1 : -1;
      temp = temp.slice().sort((a, b) => {
        const an = normalizeText(pickValue(a.r1 || a.r2, a.r1 ? file1.colMap : file2.colMap, 'اسم الموظف الثلاثي'));
        const bn = normalizeText(pickValue(b.r1 || b.r2, b.r1 ? file1.colMap : file2.colMap, 'اسم الموظف الثلاثي'));
        return dir * an.localeCompare(bn, 'ar');
      });
    }
    return temp;
  }, [results, filter, searchQuery, sortKey, file1, file2]);

  useEffect(() => { setPage(1); }, [filter, searchQuery, sortKey, results]);

  const pageCount = Math.max(1, Math.ceil(filteredResults.length / pageSize));
  const safePage = Math.min(page, pageCount);
  const pageItems = useMemo(
    () => filteredResults.slice((safePage - 1) * pageSize, safePage * pageSize),
    [filteredResults, safePage, pageSize]
  );

  const handleSelect = useCallback((r: CompareRecord) => setSelectedRecord(r), []);
  const closeModal = useCallback(() => setSelectedRecord(null), []);

  const toggleFilter = (key: MatchStatus) => setFilter((prev) => (prev === key ? 'ALL' : key));

  const cards: StatCardConfig[] = stats ? [
    { key: 'EXACT', label: 'متطابق بالكامل', value: stats.exact, cls: 'bg-green-50 border-green-200 text-green-700 hover:bg-green-100' },
    { key: 'MATCH_DIFF', label: 'متطابق مع اختلاف', value: stats.diff, cls: 'bg-yellow-50 border-yellow-200 text-yellow-700 hover:bg-yellow-100' },
    { key: 'REVIEW', label: 'يحتاج مراجعة', value: stats.review, cls: 'bg-orange-50 border-orange-200 text-orange-700 hover:bg-orange-100' },
    { key: 'ONLY_F1', label: 'في الأول فقط (محذوف)', value: stats.onlyF1, cls: 'bg-red-50 border-red-200 text-red-700 hover:bg-red-100' },
    { key: 'ONLY_F2', label: 'في الثاني فقط (جديد)', value: stats.onlyF2, cls: 'bg-blue-50 border-blue-200 text-blue-700 hover:bg-blue-100' }
  ] : [];

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans" dir="rtl">

      {/* Header */}
      <header className="bg-white border-b shadow-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg">
              <Users className="text-white" size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-800">مقارن ملفات الموظفين الذكي</h1>
              <p className="text-xs text-slate-500">نظام تحليل ومطابقة السجلات بدقة عالية</p>
            </div>
          </div>
          {results && (
            <div className="flex items-center gap-2">
              <button
                onClick={editComparisonSettings}
                className="text-sm flex items-center gap-2 bg-blue-50 hover:bg-blue-100 text-blue-700 px-4 py-2 rounded-lg transition-colors font-medium">
                <Settings size={16} />
                تعديل أعمدة المقارنة
              </button>
              <button
                onClick={resetAll}
                className="text-sm flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg transition-colors font-medium">
                <RefreshCw size={16} />
                مقارنة جديدة
              </button>
            </div>
          )}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">

        {/* Error Alert */}
        {errorMsg && (
          <div className="mb-6 bg-red-50 border-r-4 border-red-500 text-red-700 p-4 rounded-lg flex items-center gap-3 shadow-sm">
            <AlertTriangle size={20} />
            <p className="font-medium">{errorMsg}</p>
            <button onClick={() => setErrorMsg(null)} aria-label="إغلاق التنبيه" className="mr-auto text-red-400 hover:text-red-700">
              <XCircle size={20} />
            </button>
          </div>
        )}

        {/* Upload Section - Only show if no results yet */}
        {!results && (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 mb-8">
            <div className="mb-6 flex items-center gap-2 text-slate-800 font-bold text-lg">
              <FileSpreadsheet size={24} className="text-blue-500" />
              <h2>استيراد الملفات</h2>
            </div>

            <div className="grid md:grid-cols-2 gap-8 mb-8">
              {([
                { num: 1 as const, file: file1, title: 'الملف الأول (المرجعي / القديم)', doneCls: 'border-blue-400 bg-blue-50', iconCls: 'text-blue-500', nameCls: 'text-blue-700 border-blue-200' },
                { num: 2 as const, file: file2, title: 'الملف الثاني (الجديد / المراد مقارنته)', doneCls: 'border-indigo-400 bg-indigo-50', iconCls: 'text-indigo-500', nameCls: 'text-indigo-700 border-indigo-200' }
              ] as Array<{ num: 1 | 2; file: LoadedFile | null; title: string; doneCls: string; iconCls: string; nameCls: string }>).map(({ num, file, title, doneCls, iconCls, nameCls }) => (
                <div key={num} className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-all ${file ? doneCls : 'border-slate-300 hover:bg-slate-50'}`}>
                  <input
                    type="file"
                    disabled={!isXlsxLoaded}
                    accept=".xlsx,.xlsm,.xls,.csv"
                    onChange={(e) => handleFileUpload(e, num)}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-not-allowed"
                  />
                  <FileText className={`mx-auto mb-4 ${file ? iconCls : 'text-slate-400'}`} size={48} />
                  <h3 className="font-bold text-lg text-slate-700 mb-2">{title}</h3>
                  {file ? (
                    <div>
                      <p className={`text-sm font-bold mb-1 bg-white inline-block px-3 py-1 rounded-full border ${nameCls}`}>{file.name}</p>
                      <p className="text-xs text-slate-500 mt-2">
                        <span className="font-bold text-slate-700">{file.rows.length}</span> سجل تم التعرف عليه
                        <span className="mx-1">·</span>
                        <span className="font-bold text-slate-700">{file.columns.length}</span> عمود في ورقة {file.sheetName}
                      </p>
                    </div>
                  ) : (
                    <p className="text-sm text-slate-500 font-medium">{!isXlsxLoaded ? 'جاري تحميل المكتبات...' : 'اضغط لاختيار ملف Excel'}</p>
                  )}
                </div>
              ))}
            </div>

            {file1 && file2 && (
              <section className="mb-6 rounded-xl border border-slate-200 bg-slate-50 p-4" aria-labelledby="comparison-columns-title">
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <h3 id="comparison-columns-title" className="font-bold text-slate-800">أعمدة مقارنة البيانات</h3>
                    <p className="mt-1 text-xs text-slate-500">
                      اختر الحقول التي تريد فحص اختلافاتها. تُستخدم أعمدة الهوية المتاحة للمطابقة بين السجلات بشكل مستقل.
                    </p>
                  </div>
                  <div className="flex items-center gap-3 text-sm">
                    <span className="font-bold text-slate-700">
                      {selectedComparisonColumns.length} من {availableComparisonColumns.length} محدد
                    </span>
                    <button
                      type="button"
                      onClick={() => setSelectedColumnFields(availableComparisonColumns.map(({ field }) => field))}
                      className="font-bold text-blue-700 hover:text-blue-900">
                      تحديد الكل
                    </button>
                    <button
                      type="button"
                      onClick={() => setSelectedColumnFields([])}
                      className="font-bold text-slate-600 hover:text-slate-900">
                      إلغاء التحديد
                    </button>
                  </div>
                </div>
                <div className="mt-4 grid max-h-64 grid-cols-1 gap-2 overflow-y-auto rounded-lg border border-slate-200 bg-white p-3 sm:grid-cols-2 lg:grid-cols-3">
                  {availableComparisonColumns.map(({ field }, index) => (
                    <label
                      key={field}
                      htmlFor={`comparison-column-${index}`}
                      className="flex cursor-pointer items-start gap-2 rounded-md p-2 text-sm text-slate-700 hover:bg-blue-50">
                      <input
                        id={`comparison-column-${index}`}
                        type="checkbox"
                        checked={selectedColumnFields.includes(field)}
                        onChange={() => toggleComparisonColumn(field)}
                        className="mt-0.5 accent-blue-600"
                      />
                      <span>{field}</span>
                    </label>
                  ))}
                </div>
                {selectedComparisonColumns.length === 0 && (
                  <p role="alert" className="mt-2 text-sm font-medium text-red-600">
                    اختر عموداً واحداً على الأقل لتفعيل المقارنة.
                  </p>
                )}
              </section>
            )}

            {/* Settings & Action */}
            <div className="flex flex-col md:flex-row items-center justify-between border-t pt-6">
              <div className="flex items-center gap-3 mb-4 md:mb-0 bg-slate-50 p-3 rounded-lg border">
                <Settings size={20} className="text-slate-500" />
                <div className="flex flex-col">
                  <label htmlFor="fuzzy-range" className="text-xs font-bold text-slate-700">عتبة المطابقة التقريبية للأسماء (%)</label>
                  <input
                    id="fuzzy-range"
                    type="range"
                    min="50"
                    max="100"
                    step="1"
                    value={fuzzyThreshold}
                    onChange={(e) => setFuzzyThreshold(Number(e.target.value))}
                    className="w-48 mt-1"
                  />
                  <span className="text-xs text-slate-500 text-center mt-1">{fuzzyThreshold}% — كلما زادت العتبة قلّت المطابقات التقريبية</span>
                </div>
              </div>

              <button
                onClick={processComparison}
                disabled={!file1 || !file2 || selectedComparisonColumns.length === 0 || isProcessing || !isXlsxLoaded}
                className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-xl shadow-md transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed text-lg">
                {isProcessing ? <RefreshCw className="animate-spin" size={24} /> : <CheckCircle size={24} />}
                {isProcessing ? 'جاري التحليل والمطابقة...' : 'بدء عملية المقارنة'}
              </button>
            </div>
          </div>
        )}

        {/* Results Section */}
        {results && stats && file1 && file2 && (
          <div className="space-y-6">
            {/* Stats Dashboard */}
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
              <StatCard
                label="إجمالي الموظفين (برقم وطني أو اسم أو ID)"
                value={`${stats.totalF1} / ${stats.totalF2}`}
              />
              {cards.map((c) => (
                <StatCard
                  key={c.key}
                  label={c.label}
                  value={c.value}
                  onClick={() => toggleFilter(c.key)}
                  active={filter === c.key}
                  tone={c.cls}
                />
              ))}
            </div>

            {/* Toolbar */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col md:flex-row justify-between items-center gap-4">
              <div className="flex items-center gap-3 w-full md:w-auto flex-wrap">
                <div className="relative w-full md:w-80">
                  <input
                    type="text"
                    placeholder="بحث بالاسم أو الرقم الوطني..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-4 pr-10 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none text-sm"
                  />
                  <Search className="absolute right-3 top-2.5 text-slate-400" size={18} />
                </div>
                <div className="flex items-center gap-2 bg-slate-100 px-3 py-2 rounded-lg">
                  <Filter size={18} className="text-slate-500" />
                  <select value={filter} onChange={(e) => setFilter(e.target.value as FilterKey)} className="bg-transparent text-sm font-medium text-slate-700 outline-none cursor-pointer">
                    <option value="ALL">عرض الكل ({results.length})</option>
                    <option value="EXACT">المتطابقين ({stats.exact})</option>
                    <option value="MATCH_DIFF">المختلفين ({stats.diff})</option>
                    <option value="REVIEW">للمراجعة ({stats.review})</option>
                    <option value="ONLY_F1">في الأول فقط ({stats.onlyF1})</option>
                    <option value="ONLY_F2">في الثاني فقط ({stats.onlyF2})</option>
                  </select>
                </div>
                <select
                  value={sortKey}
                  onChange={(e) => setSortKey(e.target.value as SortKey)}
                  className="bg-slate-100 text-sm font-medium text-slate-700 outline-none cursor-pointer px-3 py-2 rounded-lg"
                >
                  <option value="default">ترتيب: حسب الترتيب الأصلي</option>
                  <option value="name_asc">ترتيب: الاسم تصاعدي</option>
                  <option value="name_desc">ترتيب: الاسم تنازلي</option>
                </select>
              </div>

              <button onClick={exportToExcel} className="w-full md:w-auto flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-6 py-2.5 rounded-lg transition-colors">
                <Download size={18} />
                تصدير التقرير Excel
              </button>
            </div>

            {/* Data Table */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-right">
                  <thead className="bg-slate-50 text-slate-600 font-bold border-b border-slate-200">
                    <tr>
                      <th className="px-6 py-4">الحالة</th>
                      <th className="px-6 py-4">الرقم الوطني</th>
                      <th className="px-6 py-4">اسم الموظف</th>
                      <th className="px-6 py-4">معلومات أساسية (من الملف الأول)</th>
                      <th className="px-6 py-4 text-center">إجراءات</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {pageItems.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-6 py-12 text-center text-slate-500">
                          <Search size={48} className="mx-auto mb-4 opacity-20" />
                          <p className="text-lg font-medium">لا توجد نتائج مطابقة للبحث أو التصفية الحالية</p>
                        </td>
                      </tr>
                    ) : (
                      pageItems.map((r) => (
                        <ResultRow key={`${r.status}-${r.i1}-${r.i2}`} r={r} cm1={file1.colMap} cm2={file2.colMap} onSelect={handleSelect} />
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-3 border-t border-slate-200 px-6 py-3 text-sm text-slate-600">
                <div className="flex items-center gap-2">
                  <span>عرض</span>
                  <select
                    value={pageSize}
                    onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}
                    className="bg-slate-100 rounded-md px-2 py-1 outline-none cursor-pointer"
                  >
                    {PAGE_SIZES.map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                  <span>
                    {filteredResults.length === 0 ? 0 : (safePage - 1) * pageSize + 1}
                    {' - '}
                    {Math.min(safePage * pageSize, filteredResults.length)}
                    {' من '}{filteredResults.length} سجل
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setPage((p) => Math.max(1, p - 1))}
                    disabled={safePage <= 1}
                    className="px-3 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed inline-flex items-center gap-1"
                  >
                    <ChevronRight size={16} /> السابق
                  </button>
                  <span className="px-2">صفحة {safePage} من {pageCount}</span>
                  <button
                    type="button"
                    onClick={() => setPage((p) => Math.min(pageCount, p + 1))}
                    disabled={safePage >= pageCount}
                    className="px-3 py-1.5 rounded-md border border-slate-300 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed inline-flex items-center gap-1"
                  >
                    التالي <ChevronLeft size={16} />
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Details Modal */}
      {selectedRecord && file1 && file2 && (
        <RecordModal record={selectedRecord} cm1={file1.colMap} cm2={file2.colMap} columns={selectedComparisonColumns} onClose={closeModal} />
      )}
    </div>
  );
}
