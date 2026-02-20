import React, { useState, useMemo, useEffect, useRef } from "react";
import { ExtractedData } from "../types";
import * as XLSX from "xlsx";

interface Props {
  data: ExtractedData;
  categories: Record<string, string>;
  comments: Record<string, string>;
  mappings: Record<string, string>;
  types: Record<string, string>;
  onUpdateCategory: (cleanName: string, category: string) => void;
  onUpdateComment: (cleanName: string, comment: string) => void;
  onUpdateType: (cleanName: string, type: string) => void;
  onMergePayees: (
    targetName: string,
    category: string,
    selection: { index: number; desc: string }[]
  ) => void;
  onUpdateMappings: (
    selection: { index: number; desc: string }[],
    newTargetName: string | null
  ) => void;
  onBulkUpdateCategory: (cleanNames: string[], category: string) => void;
  onBulkUpdateComment: (cleanNames: string[], comment: string) => void;
  onBulkUpdateType: (cleanNames: string[], type: string) => void;
  onBulkApplyMappings: (newMap: Record<string, string>) => void;

  // Sync Toggles
  isEntitySynced: boolean;
  isCategorySynced: boolean;
  isTypeSynced: boolean;
  isCommentSynced: boolean;
  onToggleEntitySync: () => void;
  onToggleCategorySync: () => void;
  onToggleTypeSync: () => void;
  onToggleCommentSync: () => void;

  // Undo/Redo
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
}

// --- Text Cleaning & Matching Logic ---

const BANKING_TERMS = [
  "DD",
  "DDR",
  "DIRECT DEBIT",
  "DIR DEB",
  "D/D",
  "D/DR",
  "MEMO DD",
  "VAR DD",
  "AUDDIS",
  "SO",
  "STO",
  "S/O",
  "STNDG ORDER",
  "STANDING ORDER",
  "BGC",
  "BACS",
  "BACS CREDIT",
  "BACS PYMT",
  "B.G.C.",
  "BGC/FBP",
  "BGC/FPI",
  "FPI",
  "FPO",
  "FP",
  "FASTER PYMT",
  "FST PYMT",
  "FAST PAY",
  "FP/BGC",
  "CHQ",
  "CHEQUE",
  "CHQ IN",
  "CHQ PAID",
  "C/Q",
  "CQ",
  "CQ IN",
  "ATM",
  "CASH",
  "CASH WDL",
  "WDL",
  "WITHDRAWAL",
  "LINK",
  "CDM",
  "POS",
  "DEB",
  "DEBIT CARD",
  "DC",
  "VISA",
  "MC",
  "MASTERCARD",
  "CHAPS",
  "CHAPS PYMT",
  "CHAP",
  "INT",
  "INTEREST",
  "INT PAID",
  "GROSS INT",
  "NET INT",
  "DIV",
  "DIVIDEND",
  "DIV PAYMT",
  "TFR",
  "TRF",
  "TRANSFER",
  "INTERNAL TFR",
  "ITR",
  "FT",
  "GIRO",
  "GIRO CREDIT",
  "GCT",
  "GIR",
  "REF:",
  "REFERENCE",
  "REF NO",
  "REFN",
  "RN",
  "INV",
  "INVOICE",
  "INV NO",
  "INV#",
  "A/C",
  "AC",
  "ACCOUNT",
  "ACC NO",
  "ACT",
  "ACC",
  "MOTO",
  "E-COM",
  "RECURRING",
  "MANDATE",
  "VALUE DATE",
  "VAL DT",
  "BOOK DATE",
  "NON-STG",
  "NON-STERLING",
  "FX FEE",
  "X-RATE",
  "AUTH",
  "AUTHORISATION",
  "APP CODE",
  "TRANS ID",
  "ORIGINATOR",
  "ORIG",
  "USER ID",
  "MEMO",
  "REMARK",
  "NOTE",
  "CONTACTLESS",
  "CNL",
  "CTLS",
  "COMMISSION",
  "COMM",
  "CMN",
  "FEE",
  "FEES",
  "MONTHLY FEE",
  "ARRANGEMENT FEE",
  "CHARGES",
  "CHG",
  "CHGS",
  "SERVICE CHG",
  "OVERDRAFT",
  "O/D",
  "OD",
  "UNAUTH O/D",
  "PENALTY",
  "RETURNED",
  "UNPAID",
  "STOPPED",
  "ADJUSTMENT",
  "ADJ",
  "CORRECTION",
  "CORR",
  "BENEFICIARY",
  "BILL",
  "BILL PAY",
  "BILL PAYMT",
  "BOND",
  "BONUS",
  "BRANCH",
  "BRH",
  "BROKER",
  "BUSINESS",
  "BUY",
  "CALL",
  "CANCELLED",
  "CAP",
  "CAPITAL",
  "CARD PYMT",
  "CARDHOLDER",
  "CASHBACK",
  "CERTIFICATE",
  "CHARGEBACK",
  "CLEARING",
  "CLOSING",
  "COLL",
  "COLLECTION",
  "COMPOUND",
  "CONSOLIDATED",
  "CONTRA",
  "CONTRACT",
  "CONTRIBUTION",
  "CONVERSION",
  "COST",
  "COUPON",
  "CR",
  "CREDIT",
  "CSD",
  "CUST",
  "DEBIT",
  "DEBT",
  "DRAWING",
  "DR",
  "DUAL",
  "DUE",
  "DUPLICATE",
  "DUTY",
  "EARLY",
  "ELECTRONIC",
  "ESCROW",
  "ESTATE",
  "EST",
  "ESTIMATE",
  "EXCESS",
  "EXCHANGE",
  "EXCL",
  "I-BANK",
  "IBAN",
  "IDENT",
  "IMMED",
  "IMMEDIATE",
  "IMPORT",
  "IMPOST",
  "JRNL",
  "JOURNAL",
  "PAID",
  "PAY",
  "PAYABLE",
  "PAYEE",
  "PAYER",
  "PAYING",
  "PAYMENT",
  "PAYMENTS",
  "ONLINE",
  "TRANSACTION",
  "AUTOMATED",
  "PYMT",
  "PMT",
  "CORD",
  "CARD",
  "CARD TRANSACTION",
  "CORD TRANSACTION",
];

const GENERIC_NOISE = [
  "TRUCK",
  "STATION",
  "STORE",
  "SHOP",
  "ONLINE",
  "PURCHASE",
  "POS",
  "CARD",
  "TRANSACTION",
  "PAYMENT",
  "BILL",
  "VALUE",
  "DATE",
  "LOC",
  "LOCAL",
  "INT",
  "INTL",
  "COM",
  "CO",
  "UK",
  "USA",
  "EU",
  "THE",
  "AND",
  "AT",
  "OF",
  "TO",
  "FOR",
  "FROM",
  "VIA",
  "IN",
  "ON",
  "BY",
  "MR",
  "MRS",
  "MS",
  "DR",
];

BANKING_TERMS.sort((a, b) => b.length - a.length);

const NOISE_PATTERNS = BANKING_TERMS.map((term) => {
  const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const startBoundary = /^\w/.test(term) ? "\\b" : "";
  const endBoundary = /\w$/.test(term) ? "\\b" : "";
  return new RegExp(`${startBoundary}${escaped}${endBoundary}`, "gi");
});

const TYPE_PATTERNS = [
  {
    key: "DD",
    patterns: [
      /DIRECT DEBIT/i,
      /DIR DEB/i,
      /MEMO DD/i,
      /VAR DD/i,
      /\bDD\b/i,
      /\bDDR\b/i,
      /AUDDIS/i,
    ],
  },
  {
    key: "SO",
    patterns: [
      /STANDING ORDER/i,
      /\bSO\b/i,
      /\bSTO\b/i,
      /\bS\/O\b/i,
      /STNDG ORDER/i,
    ],
  },
  {
    key: "FP",
    patterns: [
      /FASTER PAYMENT/i,
      /FASTER PYMT/i,
      /FST PYMT/i,
      /FP\/BGC/i,
      /\bFP\b/i,
      /\bFPS\b/i,
      /\bFPO\b/i,
      /\bFPI\b/i,
      /FAST PAY/i,
    ],
  },
  {
    key: "CARD",
    patterns: [
      /CARD TRANSACTION/i,
      /VISA/i,
      /MASTERCARD/i,
      /DEBIT CARD/i,
      /CONTACTLESS/i,
      /^CD \d/i,
      /\bDC\b/i,
      /\bPOS\b/i,
      /\bMC\b/i,
      /CARD\b/i,
    ],
  },
  {
    key: "TFR",
    patterns: [
      /ONLINE TRANSFER/i,
      /INTERNAL TFR/i,
      /TRANSFER/i,
      /\bTFR\b/i,
      /\bTRF\b/i,
      /ITR/i,
      /\bFT\b/i,
    ],
  },
  { key: "BACS", patterns: [/\bBACS\b/i] },
  { key: "BGC", patterns: [/BANK GIRO/i, /B\.G\.C\./i, /\bBGC\b/i] },
  { key: "CHQ", patterns: [/CHEQUE/i, /\bCHQ\b/i, /C\/Q/i, /CQ\b/i] },
  {
    key: "CASH",
    patterns: [
      /\bATM\b/i,
      /CASH/i,
      /WITHDRAWAL/i,
      /\bWDL\b/i,
      /\bLINK\b/i,
      /\bCDM\b/i,
    ],
  },
  {
    key: "FEE",
    patterns: [
      /FEE\b/i,
      /CHARGE/i,
      /\bCHG\b/i,
      /COMMISSION/i,
      /\bCOMM?\b/i,
      /SERVICE CHG/i,
      /MONTHLY FEE/i,
    ],
  },
  { key: "INT", patterns: [/INTEREST/i, /\bINT\b/i, /GROSS INT/i, /NET INT/i] },
  { key: "DIV", patterns: [/DIVIDEND/i, /\bDIV\b/i] },
  { key: "BILL", patterns: [/BILL PAY/i, /\bBP\b/i, /BILL\b/i] },
  { key: "SAL", patterns: [/SALARY/i, /PAYROLL/i, /WAGES/i] },
  { key: "TAX", patterns: [/HMRC/i, /VAT/i, /TAX\b/i, /COUNCIL TAX/i] },
  { key: "DEP", patterns: [/DEPOSIT/i, /\bDEP\b/i, /CREDIT/i, /\bCR\b/i] },
  { key: "CHAPS", patterns: [/\bCHAPS\b/i] },
  {
    key: "REV",
    patterns: [/REVERSAL/i, /\bREV\b/i, /RETURNED/i, /UNPAID/i, /CANCELLED/i],
  },
  { key: "REF", patterns: [/REFUND/i, /REPAYMENT/i, /\bREFD\b/i] },
  {
    key: "ADJ",
    patterns: [/ADJUSTMENT/i, /\bADJ\b/i, /CORRECTION/i, /\bCORR\b/i],
  },
  { key: "INS", patterns: [/INSURANCE/i, /\bPREM\b/i, /PREMIUM/i, /\bINS\b/i] },
  { key: "LOAN", patterns: [/LOAN/i, /MORTGAGE/i, /\bMTG\b/i, /FINANCE/i] },
  { key: "PENS", patterns: [/PENSION/i, /\bPEN\b/i] },
  { key: "RENT", patterns: [/RENT\b/i] },
  {
    key: "UTIL",
    patterns: [
      /UTILITY/i,
      /\bUTIL\b/i,
      /ELEC/i,
      /GAS\b/i,
      /WATER\b/i,
      /ENERGY/i,
    ],
  },
  {
    key: "SUB",
    patterns: [/SUBSCRIPTION/i, /\bSUB\b/i, /MEMBERSHIP/i, /CLUB\b/i],
  },
  {
    key: "ONL",
    patterns: [
      /ONLINE/i,
      /\bONL\b/i,
      /E-COM/i,
      /INTERNET/i,
      /WEB\b/i,
      /WWW\./i,
    ],
  },
  { key: "PHON", patterns: [/TELEPHONE/i, /PHONE/i, /MOBILE/i, /\bTEL\b/i] },
  { key: "GIFT", patterns: [/GIFT/i, /DONATION/i, /CHARITY/i] },
  { key: "OTHR", patterns: [/MISC/i, /OTHER/i] },
];

const cleanPayeeHelper = (rawName: string) => {
  if (!rawName) return "UNKNOWN";
  let cleaned = rawName.toUpperCase();
  NOISE_PATTERNS.forEach((pattern) => {
    cleaned = cleaned.replace(pattern, " ");
  });
  cleaned = cleaned.replace(/\s+/g, " ").trim();
  return cleaned.length < 2 ? rawName.trim() || "UNKNOWN" : cleaned;
};

const getSmartMatchKey = (rawName: string) => {
  let text = rawName.toUpperCase();
  NOISE_PATTERNS.forEach((pattern) => {
    text = text.replace(pattern, " ");
  });
  GENERIC_NOISE.forEach((word) => {
    text = text.replace(new RegExp(`\\b${word}\\b`, "g"), " ");
  });
  text = text.replace(/\b\d+\b/g, (match) => {
    if (match.length === 8 || match.length === 6) return match;
    return " ";
  });
  text = text.replace(/[^A-Z0-9\s&]/g, " ");
  return text.replace(/\s+/g, " ").trim();
};

const levenshtein = (a: string, b: string): number => {
  const an = a ? a.length : 0;
  const bn = b ? b.length : 0;
  if (an === 0) return bn;
  if (bn === 0) return an;
  const matrix = Array(bn + 1)
    .fill(null)
    .map(() => Array(an + 1).fill(null));
  for (let i = 0; i <= bn; i++) matrix[i][0] = i;
  for (let j = 0; j <= an; j++) matrix[0][j] = j;
  for (let i = 1; i <= bn; i++) {
    for (let j = 1; j <= an; j++) {
      const cost = b.charAt(i - 1) === a.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }
  return matrix[bn][an];
};

const parseBankDate = (dateStr: string): Date | null => {
  if (!dateStr) return null;
  let d = new Date(dateStr);
  if (!isNaN(d.getTime())) return d;
  return null;
};

const getLongestCommonPrefix = (strs: string[]) => {
  if (!strs.length) return "";
  let prefix = strs[0];
  for (let i = 1; i < strs.length; i++) {
    while (strs[i].indexOf(prefix) !== 0) {
      prefix = prefix.substring(0, prefix.length - 1);
      if (prefix === "") return "";
    }
  }
  return prefix;
};

const parseTrimRules = (rulesText: string): RegExp[] => {
  const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const rawRules = rulesText
    .split(/\r?\n|\||\s+OR\s+/i)
    .map((r) => r.trim())
    .filter((r) => r.length > 0);
  const regexes: RegExp[] = [];

  const DATE_PATTERN =
    "(?:\\b\\d{1,2}[\\/\\-\\.]\\d{1,2}[\\/\\-\\.]\\d{2,4}\\b|\\b\\d{1,2}\\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*(?:\\s*\\d{2,4})?\\b|\\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{2,4}\\b|\\b\\d{6}\\b)";
  const NUMBER_PATTERN = "\\b\\d+\\b";
  const LETTER_PATTERN = "\\b[A-Za-z0-9\\-]{1,5}\\b";
  const MIX_PATTERN = "\\b[A-Za-z0-9]+\\b";
  const FUZZY_DATE_PATTERN =
    "\\b\\d*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\d*\\b";

  rawRules.forEach((rule) => {
    let regexStr = "";
    let lastIndex = 0;
    const tokenRegex = /\[(.*?)\]/g;
    let match;

    while ((match = tokenRegex.exec(rule)) !== null) {
      const before = rule.substring(lastIndex, match.index);
      if (before) {
        regexStr += escapeRegExp(before).replace(/\s+/g, "\\s*");
      }
      const content = match[1].toLowerCase();
      if (content === "date") regexStr += DATE_PATTERN;
      else if (content === "numbers" || content === "number")
        regexStr += NUMBER_PATTERN;
      else if (content === "letter") regexStr += LETTER_PATTERN;
      else if (content === "mix") regexStr += MIX_PATTERN;
      else if (content === "fuzzy date") regexStr += FUZZY_DATE_PATTERN;
      else if (content.startsWith('"') && content.endsWith('"')) {
        regexStr += escapeRegExp(content.slice(1, -1));
      } else {
        regexStr += escapeRegExp("[" + match[1] + "]");
      }
      lastIndex = tokenRegex.lastIndex;
    }
    const remainder = rule.substring(lastIndex);
    if (remainder) regexStr += escapeRegExp(remainder).replace(/\s+/g, "\\s*");

    try {
      if (regexStr) regexes.push(new RegExp(regexStr, "i"));
    } catch (e) {}
  });
  return regexes;
};

interface DrilldownProps {
  group: {
    rawRows: {
      originalIndex: number;
      date: string;
      desc: string;
      amount: number;
      isCredit: boolean;
      parsedDate: Date | null;
    }[];
  };
  selectedDrillRows: Set<number>;
  mappings: Record<string, string>;
  onToggle: (rowId: number) => void;
  onBulkToggle?: (rowIds: number[], select: boolean) => void;
}

const DrilldownRows: React.FC<DrilldownProps> = ({
  group,
  selectedDrillRows,
  mappings,
  onToggle,
  onBulkToggle,
}) => {
  const [sortConfig, setSortConfig] = useState<{
    key: string;
    dir: "asc" | "desc";
  } | null>(null);
  const [lastClickedIndex, setLastClickedIndex] = useState<number | null>(null);

  const displayRows = useMemo(() => {
    let rows = [...group.rawRows];
    if (sortConfig) {
      rows.sort((a, b) => {
        let valA: any = a[sortConfig.key as keyof typeof a];
        let valB: any = b[sortConfig.key as keyof typeof b];

        if (sortConfig.key === "mapped") {
          valA = mappings[a.desc] || "-";
          valB = mappings[b.desc] || "-";
        } else if (sortConfig.key === "in") {
          valA = a.isCredit ? a.amount : 0;
          valB = b.isCredit ? b.amount : 0;
        } else if (sortConfig.key === "out") {
          valA = !a.isCredit ? a.amount : 0;
          valB = !b.isCredit ? b.amount : 0;
        }

        if (valA === valB) return 0;
        if (valA == null) return 1;
        if (valB == null) return -1;

        const res =
          typeof valA === "string" ? valA.localeCompare(valB) : valA - valB;

        return sortConfig.dir === "asc" ? res : -res;
      });
    }
    return rows;
  }, [group.rawRows, sortConfig, mappings]);

  const handleSort = (key: string) => {
    setSortConfig((prev) => {
      if (prev?.key === key) {
        return { key, dir: prev.dir === "asc" ? "desc" : "asc" };
      }
      return { key, dir: "asc" };
    });
  };

  const handleCheckboxClick = (
    e: React.MouseEvent,
    rowIndex: number,
    originalIndex: number
  ) => {
    const isChecked = !selectedDrillRows.has(originalIndex);
    const shiftKey = e.shiftKey;

    if (shiftKey && lastClickedIndex !== null && onBulkToggle) {
      const start = Math.min(lastClickedIndex, rowIndex);
      const end = Math.max(lastClickedIndex, rowIndex);
      const idsToToggle = displayRows
        .slice(start, end + 1)
        .map((r) => r.originalIndex);
      onBulkToggle(idsToToggle, isChecked);
    } else {
      onToggle(originalIndex);
    }
    setLastClickedIndex(rowIndex);
  };

  return (
    <tr className="bg-slate-50 shadow-inner">
      <td colSpan={10} className="p-0">
        <div className="max-h-60 overflow-y-auto border-y border-slate-200">
          <table className="w-full text-xs">
            <thead className="bg-slate-100 text-slate-500 font-bold sticky top-0 z-10">
              <tr className="select-none">
                <th className="p-2 w-10 text-center">Select</th>
                <th
                  className="p-2 text-left cursor-pointer hover:bg-slate-200"
                  onClick={() => handleSort("date")}
                >
                  Date{" "}
                  {sortConfig?.key === "date" &&
                    (sortConfig.dir === "asc" ? "▲" : "▼")}
                </th>
                <th
                  className="p-2 text-left cursor-pointer hover:bg-slate-200"
                  onClick={() => handleSort("desc")}
                >
                  Original Description{" "}
                  {sortConfig?.key === "desc" &&
                    (sortConfig.dir === "asc" ? "▲" : "▼")}
                </th>
                <th
                  className="p-2 text-left cursor-pointer hover:bg-slate-200"
                  onClick={() => handleSort("mapped")}
                >
                  Mapped Entity{" "}
                  {sortConfig?.key === "mapped" &&
                    (sortConfig.dir === "asc" ? "▲" : "▼")}
                </th>
                <th
                  className="p-2 text-right cursor-pointer hover:bg-slate-200"
                  onClick={() => handleSort("in")}
                >
                  In{" "}
                  {sortConfig?.key === "in" &&
                    (sortConfig.dir === "asc" ? "▲" : "▼")}
                </th>
                <th
                  className="p-2 text-right cursor-pointer hover:bg-slate-200"
                  onClick={() => handleSort("out")}
                >
                  Out{" "}
                  {sortConfig?.key === "out" &&
                    (sortConfig.dir === "asc" ? "▲" : "▼")}
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {displayRows.map((row, i) => (
                <tr key={i} className="hover:bg-blue-50/50 transition-colors">
                  <td className="p-2 text-center">
                    <input
                      type="checkbox"
                      checked={selectedDrillRows.has(row.originalIndex)}
                      onChange={() => {}}
                      onClick={(e) =>
                        handleCheckboxClick(e, i, row.originalIndex)
                      }
                      className="rounded border-slate-300 text-blue-600 focus:ring-0 cursor-pointer"
                    />
                  </td>
                  <td className="p-2 text-slate-500 font-mono whitespace-nowrap">
                    {row.date}
                  </td>
                  <td className="p-2 text-slate-700 font-medium">{row.desc}</td>
                  <td className="p-2 text-slate-400 italic">
                    {mappings[row.desc] || "-"}
                  </td>
                  <td className="p-2 text-right font-mono font-bold text-emerald-600">
                    {row.isCredit ? row.amount.toFixed(2) : ""}
                  </td>
                  <td className="p-2 text-right font-mono font-bold text-slate-600">
                    {!row.isCredit ? row.amount.toFixed(2) : ""}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </td>
    </tr>
  );
};

export const PivotView: React.FC<Props> = ({
  data,
  categories,
  comments,
  mappings,
  types,
  onUpdateCategory,
  onUpdateComment,
  onUpdateType,
  onMergePayees,
  onUpdateMappings,
  onBulkUpdateCategory,
  onBulkUpdateComment,
  onBulkUpdateType,
  onBulkApplyMappings,
  isEntitySynced,
  isCategorySynced,
  isTypeSynced,
  isCommentSynced,
  onToggleEntitySync,
  onToggleCategorySync,
  onToggleTypeSync,
  onToggleCommentSync,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
}) => {
  const [pivotMode, setPivotMode] = useState<"PAYEE" | "CATEGORY">("PAYEE");
  const [selectedSourceFilter, setSelectedSourceFilter] =
    useState<string>("ALL");
  const [hideCategorized, setHideCategorized] = useState(false);
  const [lastUsedCategory, setLastUsedCategory] = useState<string | null>(null);

  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set());
  const [expandedSubRows, setExpandedSubRows] = useState<Set<string>>(
    new Set()
  );

  const [sortConfig, setSortConfig] = useState<{
    key: string;
    direction: "asc" | "desc";
  }>({ key: "totalOut", direction: "desc" });
  const [payeeFilter, setPayeeFilter] = useState("");

  const [startDate, setStartDate] = useState<string>("");
  const [endDate, setEndDate] = useState<string>("");

  const [selectedGroups, setSelectedGroups] = useState<Set<string>>(new Set());
  const [lastSelectedGroupIdx, setLastSelectedGroupIdx] = useState<number>(-1);
  const [selectedDrillRows, setSelectedDrillRows] = useState<Set<number>>(
    new Set()
  );

  const [isMergeModalOpen, setIsMergeModalOpen] = useState(false);
  const [mergeName, setMergeName] = useState("");
  const [mergeCategory, setMergeCategory] = useState("");
  const [mergeMode, setMergeMode] = useState<"main_merge" | "drill_move">(
    "main_merge"
  );
  const [mergeCandidates, setMergeCandidates] = useState<string[]>([]);

  const [isManualCleanModalOpen, setIsManualCleanModalOpen] = useState(false);
  const [manualCleanText, setManualCleanText] = useState("");

  const [isAdvancedCleanModalOpen, setIsAdvancedCleanModalOpen] =
    useState(false);
  const [advCleanMode, setAdvCleanMode] = useState<"FIXED" | "DELIMITER">(
    "FIXED"
  );
  const [advCleanValue, setAdvCleanValue] = useState("");

  const [isAiTrimModalOpen, setIsAiTrimModalOpen] = useState(false);
  const [aiTrimRules, setAiTrimRules] = useState(
    `[date][number][letter]\n[date] OR [numbers]\n["Automated"]\n[date]`
  );
  const [trimMode, setTrimMode] = useState<
    "MATCH_ONLY" | "START_TO_MATCH" | "MATCH_TO_END"
  >("START_TO_MATCH");

  const [uncheckedPreviewItems, setUncheckedPreviewItems] = useState<
    Set<string>
  >(new Set());
  const [lastUncheckedInteraction, setLastUncheckedInteraction] = useState<{
    index: number;
    isChecked: boolean;
  } | null>(null);

  const [previewSort, setPreviewSort] = useState<{
    key: "original" | "result";
    dir: "asc" | "desc";
  }>({ key: "original", dir: "asc" });

  const [isChangeCaseModalOpen, setIsChangeCaseModalOpen] = useState(false);
  const [isManualTypeModalOpen, setIsManualTypeModalOpen] = useState(false);
  const [manualTypeValue, setManualTypeValue] = useState("");

  const [isSemiAutoModalOpen, setIsSemiAutoModalOpen] = useState(false);
  const [semiAutoThreshold, setSemiAutoThreshold] = useState<number>(3);
  const [semiAutoGroups, setSemiAutoGroups] = useState<
    {
      id: string;
      name: string;
      filterText: string;
      items: { id: string; desc: string; selected: boolean }[];
    }[]
  >([]);
  const [selectedSemiAutoGroupIds, setSelectedSemiAutoGroupIds] = useState<
    Set<string>
  >(new Set());

  const [suggestedMerges, setSuggestedMerges] = useState<
    { target: any; candidates: any[] }[]
  >([]);
  const [isSuggestionModalOpen, setIsSuggestionModalOpen] = useState(false);
  const [selectedSuggestionIds, setSelectedSuggestionIds] = useState<
    Set<string>
  >(new Set());

  const [cleanConfirmation, setCleanConfirmation] = useState<{
    newMap: Record<string, string>;
    catUpdates: Record<string, string>;
    commentUpdates: Record<string, string>;
    shouldClose?: boolean;
  } | null>(null);

  const [isCategoriseModalOpen, setIsCategoriseModalOpen] = useState(false);
  const [categoriseCategory, setCategoriseCategory] = useState("");
  const [isCommentModalOpen, setIsCommentModalOpen] = useState(false);
  const [bulkCommentText, setBulkCommentText] = useState("");

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    entity: 250,
    category: 150,
    type: 100,
    comment: 200,
    items: 80,
    in: 100,
    out: 100,
    net: 100,
  });

  const resizingRef = useRef<{
    key: string;
    startX: number;
    startWidth: number;
  } | null>(null);

  const handleStartResize = (key: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    resizingRef.current = {
      key,
      startX: e.pageX,
      startWidth: colWidths[key] || 100,
    };
    document.addEventListener("mousemove", handleGlobalMouseMove);
    document.addEventListener("mouseup", handleGlobalMouseUp);
    document.body.style.cursor = "col-resize";
  };

  const handleGlobalMouseMove = (e: MouseEvent) => {
    if (!resizingRef.current) return;
    const { key, startX, startWidth } = resizingRef.current;
    const diff = e.pageX - startX;
    setColWidths((prev) => ({
      ...prev,
      [key]: Math.max(50, startWidth + diff),
    }));
  };

  const handleGlobalMouseUp = () => {
    resizingRef.current = null;
    document.removeEventListener("mousemove", handleGlobalMouseMove);
    document.removeEventListener("mouseup", handleGlobalMouseUp);
    document.body.style.cursor = "";
  };

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "z") {
        e.preventDefault();
        if (canUndo) onUndo();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === "y") {
        e.preventDefault();
        if (canRedo) onRedo();
      }

      if (
        e.altKey &&
        e.key.toLowerCase() === "c" &&
        selectedGroups.size > 0 &&
        lastUsedCategory
      ) {
        e.preventDefault();
        const names = Array.from(selectedGroups);
        onBulkUpdateCategory(names, lastUsedCategory);
        setSelectedGroups(new Set());
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    selectedGroups,
    lastUsedCategory,
    onBulkUpdateCategory,
  ]);

  const headers = data.headers.map((h) => h.toLowerCase());
  let descIdx = headers.findIndex((h) =>
    /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
      h
    )
  );
  if (descIdx === -1) descIdx = 0;
  let dateIdx = headers.findIndex((h) => /date/i.test(h));
  if (dateIdx === -1) dateIdx = 0;

  const inIdx = headers.findIndex((h) =>
    /money\s?in|credit|deposit|receipt|paid\s?in|\bin\b/i.test(h)
  );
  const outIdx = headers.findIndex((h) =>
    /money\s?out|debit|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
  );
  const sourceIdx = headers.findIndex((h) => /source/i.test(h));

  const uniqueSources = useMemo(() => {
    if (sourceIdx === -1) return [];
    const sources = new Set<string>();
    data.rows.forEach((row) => {
      if (row[sourceIdx]) sources.add(row[sourceIdx]);
    });
    return Array.from(sources).sort();
  }, [data, sourceIdx]);

  const payeeGroups = useMemo(() => {
    const groupMap: Record<
      string,
      {
        id: string;
        cleanName: string;
        category: string;
        type: string;
        comment: string;
        count: number;
        totalIn: number;
        totalOut: number;
        rawRows: {
          originalIndex: number;
          date: string;
          desc: string;
          amount: number;
          isCredit: boolean;
          parsedDate: Date | null;
        }[];
      }
    > = {};

    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;

    data.rows.forEach((row, rowIndex) => {
      if (selectedSourceFilter !== "ALL" && sourceIdx !== -1) {
        if (row[sourceIdx] !== selectedSourceFilter) return;
      }

      const rawDesc = row[descIdx];
      const rawDate = row[dateIdx];
      if (!rawDesc) return;

      const parsedDate = parseBankDate(rawDate);
      if (start && parsedDate && parsedDate < start) return;
      if (end && parsedDate && parsedDate > end) return;

      const mappedName = mappings[rawDesc];
      const cleanName = mappedName || cleanPayeeHelper(rawDesc);
      const cat = categories[cleanName] || "UNCATEGORIZED";
      const typ = types[cleanName] || "";
      const cmt = comments[cleanName] || "";

      const inVal =
        parseFloat(row[inIdx]?.replace(/[^0-9.-]+/g, "") || "0") || 0;
      const outVal =
        parseFloat(row[outIdx]?.replace(/[^0-9.-]+/g, "") || "0") || 0;

      if (inVal === 0 && outVal === 0) return;

      if (!groupMap[cleanName]) {
        groupMap[cleanName] = {
          id: cleanName,
          cleanName,
          category: cat,
          type: typ,
          comment: cmt,
          count: 0,
          totalIn: 0,
          totalOut: 0,
          rawRows: [],
        };
      }

      groupMap[cleanName].count++;
      groupMap[cleanName].totalIn += inVal;
      groupMap[cleanName].totalOut += outVal;
      groupMap[cleanName].rawRows.push({
        originalIndex: rowIndex,
        date: rawDate || "",
        desc: rawDesc,
        amount: inVal > 0 ? inVal : outVal,
        isCredit: inVal > 0,
        parsedDate,
      });
    });

    return Object.values(groupMap);
  }, [
    data,
    descIdx,
    inIdx,
    outIdx,
    sourceIdx,
    categories,
    types,
    comments,
    mappings,
    dateIdx,
    startDate,
    endDate,
    selectedSourceFilter,
  ]);

  const allExistingEntities = useMemo(() => {
    return Array.from(new Set(payeeGroups.map((g) => g.cleanName))).sort();
  }, [payeeGroups]);

  const categoryGroups = useMemo(() => {
    const catMap: Record<
      string,
      {
        id: string;
        name: string;
        totalIn: number;
        totalOut: number;
        payees: typeof payeeGroups;
      }
    > = {};
    payeeGroups.forEach((pg) => {
      const cat = pg.category || "UNCATEGORIZED";
      if (!catMap[cat])
        catMap[cat] = {
          id: cat,
          name: cat,
          totalIn: 0,
          totalOut: 0,
          payees: [],
        };
      catMap[cat].totalIn += pg.totalIn;
      catMap[cat].totalOut += pg.totalOut;
      catMap[cat].payees.push(pg);
    });
    return Object.values(catMap);
  }, [payeeGroups]);

  const filteredPayeeGroups = useMemo(() => {
    if (!payeeFilter) return payeeGroups;
    const tokens = payeeFilter
      .split(/,|\s+OR\s+/i)
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);
    if (tokens.length === 0) return payeeGroups;
    return payeeGroups.filter((g) =>
      tokens.some((token) => g.cleanName.toLowerCase().includes(token))
    );
  }, [payeeGroups, payeeFilter]);

  const sortedDisplayList = useMemo(() => {
    if (pivotMode === "PAYEE") {
      const list = hideCategorized
        ? filteredPayeeGroups.filter(
            (g) => !g.category || g.category === "UNCATEGORIZED"
          )
        : filteredPayeeGroups;
      return [...list].sort((a, b) => {
        let aVal: any = a[sortConfig.key as keyof typeof a];
        let bVal: any = b[sortConfig.key as keyof typeof b];
        if (sortConfig.key === "net") {
          aVal = a.totalOut - a.totalIn;
          bVal = b.totalOut - b.totalIn;
        }
        if (typeof aVal === "string")
          return sortConfig.direction === "asc"
            ? aVal.localeCompare(bVal)
            : bVal.localeCompare(aVal);
        return sortConfig.direction === "asc" ? aVal - bVal : bVal - aVal;
      });
    } else {
      let list = categoryGroups
        .map((cg) => ({
          ...cg,
          payees: cg.payees.filter((pg) =>
            filteredPayeeGroups.some((f) => f.id === pg.id)
          ),
        }))
        .filter((cg) => cg.payees.length > 0);

      if (hideCategorized) {
        list = list.filter((cg) => cg.id === "UNCATEGORIZED");
      }

      return list.sort((a, b) => {
        let aVal: any =
          a[
            sortConfig.key === "cleanName"
              ? "name"
              : (sortConfig.key as keyof typeof a)
          ];
        let bVal: any =
          b[
            sortConfig.key === "cleanName"
              ? "name"
              : (sortConfig.key as keyof typeof b)
          ];
        if (sortConfig.key === "net") {
          aVal = a.totalOut - a.totalIn;
          bVal = b.totalOut - b.totalIn;
        }
        if (typeof aVal === "string")
          return sortConfig.direction === "asc"
            ? aVal.localeCompare(bVal)
            : bVal.localeCompare(aVal);
        return sortConfig.direction === "asc" ? aVal - bVal : bVal - aVal;
      });
    }
  }, [
    pivotMode,
    filteredPayeeGroups,
    categoryGroups,
    sortConfig,
    hideCategorized,
  ]);

  const visibleIds = useMemo(() => {
    if (pivotMode === "PAYEE") {
      return (sortedDisplayList as any[]).map((g) => g.id);
    } else {
      return (sortedDisplayList as any[]).flatMap((cg) =>
        cg.payees.map((p: any) => p.id)
      );
    }
  }, [sortedDisplayList, pivotMode]);

  const grandTotalIn = payeeGroups.reduce((acc, g) => acc + g.totalIn, 0);
  const grandTotalOut = payeeGroups.reduce((acc, g) => acc + g.totalOut, 0);

  const totalGroupsCount = payeeGroups.length;
  const categorizedGroupsCount = payeeGroups.filter(
    (g) => g.category && g.category !== "UNCATEGORIZED"
  ).length;
  const progressPercent =
    totalGroupsCount === 0
      ? 0
      : Math.round((categorizedGroupsCount / totalGroupsCount) * 100);

  const toggleRow = (id: string) => {
    setExpandedRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSubRow = (id: string) => {
    setExpandedSubRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const topLevelIds = useMemo(
    () => (sortedDisplayList as any[]).map((g) => g.id),
    [sortedDisplayList]
  );
  const isAllExpanded =
    topLevelIds.length > 0 && topLevelIds.every((id) => expandedRows.has(id));

  const handleExpandAll = () => {
    if (isAllExpanded) {
      setExpandedRows(new Set());
    } else {
      setExpandedRows(new Set(topLevelIds));
    }
  };

  const trimPreview = useMemo(() => {
    if (!isAiTrimModalOpen) return [];

    const regexes = parseTrimRules(aiTrimRules);
    if (regexes.length === 0) return [];

    const samples: { original: string; result: string; match: string }[] = [];
    const seen = new Set<string>();

    const sourceList =
      selectedGroups.size > 0
        ? payeeGroups.filter((g) => selectedGroups.has(g.id))
        : selectedDrillRows.size > 0
        ? payeeGroups.flatMap((g) =>
            g.rawRows
              .filter((r) => selectedDrillRows.has(r.originalIndex))
              .map((r) => ({ cleanName: r.desc }))
          )
        : payeeGroups;

    for (const item of sourceList) {
      const desc = (item as any).cleanName || (item as any).desc || "";
      if (seen.has(desc)) continue;
      seen.add(desc);

      let result = desc;
      let matchedStr = "";

      for (const regex of regexes) {
        const match = regex.exec(desc);
        if (match) {
          matchedStr = match[0];
          if (trimMode === "START_TO_MATCH") {
            result = desc.substring(match.index + match[0].length);
          } else if (trimMode === "MATCH_TO_END") {
            result = desc.substring(0, match.index);
          } else {
            result =
              desc.substring(0, match.index) +
              " " +
              desc.substring(match.index + match[0].length);
          }
          result = result.replace(/\s+/g, " ").trim();
          break;
        }
      }

      samples.push({ original: desc, result, match: matchedStr });
      if (samples.length >= 5000) break;
    }
    return samples;
  }, [
    isAiTrimModalOpen,
    aiTrimRules,
    trimMode,
    selectedGroups,
    selectedDrillRows,
    payeeGroups,
  ]);

  const sortedPreview = useMemo(() => {
    return [...trimPreview].sort((a, b) => {
      const valA = a[previewSort.key] || "";
      const valB = b[previewSort.key] || "";
      return previewSort.dir === "asc"
        ? valA.localeCompare(valB)
        : valB.localeCompare(a.original);
    });
  }, [trimPreview, previewSort]);

  const handleAiTrimSubmit = (shouldClose: boolean = true) => {
    const regexes = parseTrimRules(aiTrimRules);
    if (regexes.length === 0) return;

    const itemsToUpdate: { desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }));
      });
    } else if (selectedDrillRows.size > 0) {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex))
            itemsToUpdate.push({ desc: r.desc });
        });
      });
    } else {
      payeeGroups.forEach((g) =>
        g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }))
      );
    }

    if (itemsToUpdate.length === 0) return;

    const newMap: Record<string, string> = {};
    const nextSelectedGroups = new Set<string>();

    itemsToUpdate.forEach((item) => {
      if (uncheckedPreviewItems.has(item.desc)) return;

      let current = mappings[item.desc] || cleanPayeeHelper(item.desc);

      for (const regex of regexes) {
        const match = regex.exec(current);
        if (match) {
          if (trimMode === "START_TO_MATCH") {
            current = current.substring(match.index + match[0].length);
          } else if (trimMode === "MATCH_TO_END") {
            current = current.substring(0, match.index);
          } else {
            current =
              current.substring(0, match.index) +
              " " +
              current.substring(match.index + match[0].length);
          }
          current = current.replace(/\s+/g, " ").trim();
          break;
        }
      }

      if (!current) current = "UNKNOWN";
      newMap[item.desc] = current;

      if (!shouldClose && selectedGroups.size > 0) {
        nextSelectedGroups.add(current);
      }
    });

    const { catUpdates, commentUpdates } = analyzeCleanOperations(newMap);

    const applyChanges = () => {
      onBulkApplyMappings(newMap);
      if (!shouldClose) {
        if (selectedGroups.size > 0) setSelectedGroups(nextSelectedGroups);
      } else {
        setIsAiTrimModalOpen(false);
        setSelectedGroups(new Set());
        setSelectedDrillRows(new Set());
        setUncheckedPreviewItems(new Set());
      }
    };

    if (
      Object.keys(catUpdates).length > 0 ||
      Object.keys(commentUpdates).length > 0
    ) {
      setCleanConfirmation({ newMap, catUpdates, commentUpdates, shouldClose });
    } else {
      applyChanges();
    }
  };

  const handleAiTrimCheckbox = (
    e: React.ChangeEvent<HTMLInputElement>,
    index: number,
    itemOriginal: string
  ) => {
    const isChecked = e.target.checked;
    const shiftKey = (e.nativeEvent as any).shiftKey;

    if (shiftKey && lastUncheckedInteraction !== null) {
      const start = Math.min(lastUncheckedInteraction.index, index);
      const end = Math.max(lastUncheckedInteraction.index, index);
      const itemsToToggle = sortedPreview.slice(start, end + 1);

      setUncheckedPreviewItems((prev) => {
        const next = new Set(prev);
        itemsToToggle.forEach((item) => {
          if (isChecked) next.delete(item.original);
          else next.add(item.original);
        });
        return next;
      });
    } else {
      setUncheckedPreviewItems((prev) => {
        const next = new Set(prev);
        if (isChecked) next.delete(itemOriginal);
        else next.add(itemOriginal);
        return next;
      });
    }
    setLastUncheckedInteraction({ index, isChecked });
  };

  const handleSemiAutoScan = () => {
    const pool = payeeGroups
      .filter((g) => !g.category || g.category === "UNCATEGORIZED")
      .map((g) => ({
        id: g.id,
        cleanName: g.cleanName,
        smartKey: getSmartMatchKey(g.cleanName),
      }));
    const validPool = pool.filter((p) => p.smartKey.length > 2);
    validPool.sort((a, b) => b.cleanName.length - a.cleanName.length);

    const groups: {
      id: string;
      name: string;
      filterText: string;
      items: { id: string; desc: string; selected: boolean }[];
    }[] = [];
    const processedIds = new Set<string>();

    for (let i = 0; i < validPool.length; i++) {
      const leader = validPool[i];
      if (processedIds.has(leader.id)) continue;
      const currentGroupItems = [];
      for (let j = 0; j < validPool.length; j++) {
        if (i === j) continue;
        if (processedIds.has(validPool[j].id)) continue;
        const candidate = validPool[j];
        let isMatch = false;
        if (leader.smartKey === candidate.smartKey) {
          isMatch = true;
        } else if (
          leader.smartKey.includes(candidate.smartKey) &&
          candidate.smartKey.length >= semiAutoThreshold
        ) {
          isMatch = true;
        } else if (
          candidate.smartKey.includes(leader.smartKey) &&
          leader.smartKey.length >= semiAutoThreshold
        ) {
          isMatch = true;
        }
        if (isMatch) {
          currentGroupItems.push({
            id: candidate.id,
            desc: candidate.cleanName,
            selected: true,
          });
        }
      }
      if (currentGroupItems.length > 0) {
        currentGroupItems.unshift({
          id: leader.id,
          desc: leader.cleanName,
          selected: true,
        });
        currentGroupItems.forEach((item) => processedIds.add(item.id));
        let mergeName = leader.smartKey;
        groups.push({
          id: `group-${i}`,
          name: mergeName,
          filterText: "",
          items: currentGroupItems,
        });
      }
    }
    if (groups.length === 0) {
      alert(`No groups found with significant name overlap.`);
      return;
    }
    groups.sort((a, b) => a.name.localeCompare(b.name));
    setSemiAutoGroups(groups);
    setSelectedSemiAutoGroupIds(new Set());
  };

  const handleSemiAutoMergeAll = () => {
    const newMap: Record<string, string> = {};
    semiAutoGroups.forEach((group) => {
      group.items
        .filter((i) => i.selected)
        .forEach((item) => {
          const originalGroup = payeeGroups.find((pg) => pg.id === item.id);
          if (originalGroup) {
            originalGroup.rawRows.forEach((row) => {
              newMap[row.desc] = group.name;
            });
          }
        });
    });
    onBulkApplyMappings(newMap);
    setSemiAutoGroups([]);
    setIsSemiAutoModalOpen(false);
  };

  const handleAutoPopulateTypes = () => {
    const groupsToUpdate: { name: string; detected: string }[] = [];

    payeeGroups.forEach((group) => {
      if (!group.type || group.type.trim() === "") {
        const candidateCounts = new Map<string, number>();

        group.rawRows.forEach((row) => {
          const desc = (row.desc || "").toUpperCase();
          let bestMatchInRow = "";
          for (const term of BANKING_TERMS) {
            const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const pattern = new RegExp(`\\b${escaped}\\b`, "i");
            if (pattern.test(desc)) {
              bestMatchInRow = term;
              break;
            }
          }
          if (bestMatchInRow) {
            candidateCounts.set(
              bestMatchInRow,
              (candidateCounts.get(bestMatchInRow) || 0) + 1
            );
          }
        });

        if (candidateCounts.size > 0) {
          let winner = "";
          let maxCount = -1;
          candidateCounts.forEach((count, type) => {
            if (count > maxCount) {
              maxCount = count;
              winner = type;
            } else if (count === maxCount) {
              if (type.length > winner.length) winner = type;
            }
          });
          if (winner)
            groupsToUpdate.push({ name: group.cleanName, detected: winner });
        }
      }
    });

    if (groupsToUpdate.length === 0) {
      alert("No new types detected automatically.");
      return;
    }

    const updatesByType: Record<string, string[]> = {};
    groupsToUpdate.forEach((item) => {
      if (!updatesByType[item.detected]) updatesByType[item.detected] = [];
      updatesByType[item.detected].push(item.name);
    });

    Object.entries(updatesByType).forEach(([type, names]) => {
      onBulkUpdateType(names, type);
    });
  };

  const handleScanDuplicates = () => {
    try {
      const groups =
        pivotMode === "PAYEE"
          ? (sortedDisplayList as any[])
          : (sortedDisplayList as any[]).flatMap((cg) => cg.payees);
      const sortedGroups = [...groups].sort((a, b) => b.count - a.count);
      const suggestions: { target: any; candidates: any[] }[] = [];
      const processedIds = new Set<string>();
      for (let i = 0; i < sortedGroups.length; i++) {
        const target = sortedGroups[i];
        if (processedIds.has(target.id)) continue;
        const candidates = [];
        for (let j = i + 1; j < sortedGroups.length; j++) {
          const candidate = sortedGroups[j];
          if (processedIds.has(candidate.id)) continue;
          const nameA = (target.cleanName || "").toLowerCase();
          const nameB = (candidate.cleanName || "").toLowerCase();
          const maxLen = Math.max(nameA.length, nameB.length);
          if (maxLen === 0) continue;
          const dist = levenshtein(nameA, nameB);
          const similarity = 1 - dist / maxLen;
          const isContained = nameA.includes(nameB) || nameB.includes(nameA);
          if (similarity > 0.8 || (similarity > 0.6 && isContained)) {
            candidates.push(candidate);
          }
        }
        if (candidates.length > 0) {
          suggestions.push({ target, candidates });
          processedIds.add(target.id);
          candidates.forEach((c) => processedIds.add(c.id));
        }
      }
      if (suggestions.length === 0) {
        alert("No duplicates found.");
        return;
      }
      setSuggestedMerges(suggestions);
      setSelectedSuggestionIds(
        new Set(suggestions.flatMap((s) => s.candidates.map((c) => c.id)))
      );
      setIsSuggestionModalOpen(true);
    } catch (e: any) {
      alert("Scan failed: " + e.message);
    }
  };

  const analyzeCleanOperations = (newMap: Record<string, string>) => {
    const catUpdates: Record<string, string> = {};
    const commentUpdates: Record<string, string> = {};
    const catVotes: Record<string, Record<string, number>> = {};
    Object.entries(newMap).forEach(([desc, newClean]) => {
      const currentClean = mappings[desc] || cleanPayeeHelper(desc);
      if (currentClean === newClean) return;
      if (categories[currentClean] && !categories[newClean]) {
        if (!catVotes[newClean]) catVotes[newClean] = {};
        const c = categories[currentClean];
        catVotes[newClean][c] = (catVotes[newClean][c] || 0) + 1;
      }
      if (comments[currentClean] && !comments[newClean]) {
        commentUpdates[newClean] = comments[currentClean];
      }
    });
    Object.entries(catVotes).forEach(([target, votes]) => {
      const bestCat = Object.keys(votes).reduce((a, b) =>
        votes[a] > votes[b] ? a : b
      );
      catUpdates[target] = bestCat;
    });
    return { catUpdates, commentUpdates };
  };

  const handleAutoClean = () => {
    const itemsToUpdate: { index: number; desc: string }[] = [];
    const names = Array.from(selectedGroups);
    if (names.length > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) =>
            itemsToUpdate.push({ index: r.originalIndex, desc: r.desc })
          );
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex)) {
            itemsToUpdate.push({ index: r.originalIndex, desc: r.desc });
          }
        });
      });
    }
    if (itemsToUpdate.length === 0) return;
    const newMap: Record<string, string> = {};
    itemsToUpdate.forEach((item) => {
      let cleaned = getSmartMatchKey(item.desc);
      if (cleaned.length < 2) cleaned = cleanPayeeHelper(item.desc);
      newMap[item.desc] = cleaned;
    });
    const { catUpdates, commentUpdates } = analyzeCleanOperations(newMap);
    if (
      Object.keys(catUpdates).length > 0 ||
      Object.keys(commentUpdates).length > 0
    ) {
      setCleanConfirmation({ newMap, catUpdates, commentUpdates });
    } else {
      onBulkApplyMappings(newMap);
      setSelectedGroups(new Set());
      setSelectedDrillRows(new Set());
    }
  };

  const confirmCleanWithUpdates = () => {
    if (cleanConfirmation) {
      onBulkApplyMappings(cleanConfirmation.newMap);
      Object.entries(cleanConfirmation.catUpdates).forEach(([key, val]) =>
        onUpdateCategory(key, val)
      );
      Object.entries(cleanConfirmation.commentUpdates).forEach(([key, val]) =>
        onUpdateComment(key, val)
      );
      if (cleanConfirmation.shouldClose !== false) setIsAiTrimModalOpen(false);
    }
    setCleanConfirmation(null);
    if (cleanConfirmation?.shouldClose !== false) {
      setSelectedGroups(new Set());
      setSelectedDrillRows(new Set());
    }
    setUncheckedPreviewItems(new Set());
  };

  const confirmCleanWithoutUpdates = () => {
    if (cleanConfirmation) {
      onBulkApplyMappings(cleanConfirmation.newMap);
      if (cleanConfirmation.shouldClose !== false) setIsAiTrimModalOpen(false);
    }
    setCleanConfirmation(null);
    if (cleanConfirmation?.shouldClose !== false) {
      setSelectedGroups(new Set());
      setSelectedDrillRows(new Set());
    }
    setUncheckedPreviewItems(new Set());
  };

  const confirmMergeSuggestions = () => {
    const updates: Record<string, string> = {};
    suggestedMerges.forEach((s) => {
      const targetName = s.target.cleanName;
      s.candidates.forEach((c) => {
        if (selectedSuggestionIds.has(c.id)) {
          c.rawRows.forEach((r: any) => {
            updates[r.desc] = targetName;
          });
        }
      });
    });
    onBulkApplyMappings(updates);
    setIsSuggestionModalOpen(false);
  };

  const handleApplyPreset = (preset: "this-month" | "last-month") => {
    const now = new Date();
    let start: Date;
    let end: Date;
    if (preset === "this-month") {
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    } else {
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end = new Date(now.getFullYear(), now.getMonth(), 0);
    }
    const toDateInputString = (d: Date) => {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    };
    setStartDate(toDateInputString(start));
    setEndDate(toDateInputString(end));
  };

  const handleSort = (key: string) => {
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "desc" ? "asc" : "desc",
    }));
  };

  const handleDownloadExcel = () => {
    let currentGroups: any[] = [];
    if (pivotMode === "PAYEE") {
      currentGroups = sortedDisplayList as any[];
    } else {
      currentGroups = (sortedDisplayList as any[]).flatMap((cg) => cg.payees);
    }
    const summaryData = currentGroups.map((g) => ({
      Payee: g.cleanName,
      Category: g.category,
      Type: g.type,
      Comment: g.comment,
      Count: g.count,
      In: g.totalIn,
      Out: g.totalOut,
      Net: g.totalIn - g.totalOut,
    }));
    const breakdownData = currentGroups.flatMap((g) =>
      g.rawRows.map((r: any) => ({
        Date: r.date,
        Description: r.desc,
        Payee: g.cleanName,
        Category: g.category,
        Type: g.type,
        Comment: g.comment,
        In: r.isCredit ? r.amount : 0,
        Out: r.isCredit ? 0 : r.amount,
      }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(summaryData),
      "Summary"
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(breakdownData),
      "Detail"
    );
    XLSX.writeFile(wb, `Pivot_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const toggleSelection = (
    groupId: string,
    index: number,
    isSelected: boolean,
    shiftKey: boolean
  ) => {
    if (shiftKey && lastSelectedGroupIdx !== -1) {
      const start = Math.min(lastSelectedGroupIdx, index);
      const end = Math.max(lastSelectedGroupIdx, index);
      const ids = visibleIds.slice(start, end + 1);
      setSelectedGroups((prev) => {
        const next = new Set(prev);
        ids.forEach((id) => {
          if (isSelected) next.add(id);
          else next.delete(id);
        });
        return next;
      });
    } else {
      setSelectedGroups((prev) => {
        const next = new Set(prev);
        if (isSelected) next.add(groupId);
        else next.delete(groupId);
        return next;
      });
      setLastSelectedGroupIdx(index);
    }
  };

  const handleGroupCheckbox = (
    e: React.ChangeEvent<HTMLInputElement>,
    groupIndex: number,
    groupId: string
  ) => {
    toggleSelection(
      groupId,
      groupIndex,
      e.target.checked,
      (e.nativeEvent as any).shiftKey
    );
  };

  const handleSelectAllVisible = () => {
    const allVisibleSelected =
      visibleIds.length > 0 && visibleIds.every((id) => selectedGroups.has(id));
    setSelectedGroups((prev) => {
      const next = new Set(prev);
      if (allVisibleSelected) {
        visibleIds.forEach((id) => next.delete(id));
      } else {
        visibleIds.forEach((id) => next.add(id));
      }
      return next;
    });
  };

  const selectedStats = useMemo(() => {
    let count = 0,
      sumIn = 0,
      sumOut = 0,
      net = 0;
    const items: { index: number; desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id)) {
          count++;
          sumIn += g.totalIn;
          sumOut += g.totalOut;
          g.rawRows.forEach((r) =>
            items.push({ index: r.originalIndex, desc: r.desc })
          );
        }
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex)) {
            count++;
            if (r.isCredit) sumIn += r.amount;
            else sumOut += r.amount;
            items.push({ index: r.originalIndex, desc: r.desc });
          }
        });
      });
    }
    net = sumIn - sumOut;
    return { count, sumIn, sumOut, net, selectionItems: items };
  }, [payeeGroups, selectedGroups, selectedDrillRows]);

  const openMergeModal = () => {
    try {
      let candidates: string[] = [];

      if (selectedGroups.size > 0) {
        setMergeMode("main_merge");
        candidates = payeeGroups
          .filter((g) => selectedGroups.has(g.id))
          .map((g) => g.cleanName || "")
          .filter(Boolean);
      } else if (selectedDrillRows.size > 0) {
        setMergeMode("drill_move");
        const descList: string[] = [];
        payeeGroups.forEach((g) => {
          if (g.rawRows) {
            g.rawRows.forEach((r) => {
              if (selectedDrillRows.has(r.originalIndex))
                descList.push(r.desc || "");
            });
          }
        });
        candidates = descList.filter(Boolean);
      }

      if (candidates.length > 0) {
        const uniqueCandidates = Array.from(new Set(candidates)).sort(
          (a, b) => a.length - b.length
        );
        setMergeCandidates(uniqueCandidates);
        let bestName = uniqueCandidates[0];
        const manualTargets = new Set<string>();
        if (mappings) {
          Object.values(mappings).forEach((v) => {
            if (typeof v === "string" && v) manualTargets.add(v);
          });
        }
        const candidatesIsManual = uniqueCandidates.filter((c) =>
          manualTargets.has(c)
        );
        if (candidatesIsManual.length > 0) {
          bestName = candidatesIsManual.sort((a, b) => a.length - b.length)[0];
        }
        setMergeName(bestName || "");
      } else {
        setMergeName("");
        setMergeCandidates([]);
      }
    } catch (error) {
      setMergeName("");
      setMergeCandidates([]);
    }
    setMergeCategory("");
    setIsMergeModalOpen(true);
  };

  const handleBulkToggleDrillRows = (rowIds: number[], select: boolean) => {
    setSelectedDrillRows((prev) => {
      const next = new Set(prev);
      rowIds.forEach((id) => {
        if (select) next.add(id);
        else next.delete(id);
      });
      return next;
    });
  };

  const handleManualCleanSubmit = () => {
    const itemsToUpdate: { desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }));
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex))
            itemsToUpdate.push({ desc: r.desc });
        });
      });
    }
    if (itemsToUpdate.length === 0) return;
    const newMap: Record<string, string> = {};
    const cleanText = manualCleanText.trim();
    if (cleanText) {
      const escapedText = cleanText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const regex = new RegExp(escapedText, "gi");
      itemsToUpdate.forEach((item) => {
        const currentEntity =
          mappings[item.desc] || cleanPayeeHelper(item.desc);
        let newEntity = currentEntity
          .replace(regex, "")
          .replace(/\s+/g, " ")
          .trim();
        if (!newEntity) newEntity = currentEntity;
        newMap[item.desc] = newEntity;
      });
      const { catUpdates, commentUpdates } = analyzeCleanOperations(newMap);
      if (
        Object.keys(catUpdates).length > 0 ||
        Object.keys(commentUpdates).length > 0
      ) {
        setCleanConfirmation({ newMap, catUpdates, commentUpdates });
        setIsManualCleanModalOpen(false);
      } else {
        onBulkApplyMappings(newMap);
        setIsManualCleanModalOpen(false);
        setSelectedGroups(new Set());
        setSelectedDrillRows(new Set());
      }
    } else {
      setIsManualCleanModalOpen(false);
    }
  };

  const openManualCleanModal = () => {
    const selectedPayees = payeeGroups.filter((g) => selectedGroups.has(g.id));
    if (selectedPayees.length > 0) {
      const texts = selectedPayees.map((g) => g.cleanName);
      setManualCleanText(getLongestCommonPrefix(texts).trim());
    } else {
      setManualCleanText("");
    }
    setIsManualCleanModalOpen(true);
  };
  const openAdvancedCleanModal = () => {
    setAdvCleanMode("FIXED");
    setAdvCleanValue("");
    setIsAdvancedCleanModalOpen(true);
  };
  const handleAdvancedCleanSubmit = () => {
    const itemsToUpdate: { desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }));
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex))
            itemsToUpdate.push({ desc: r.desc });
        });
      });
    }
    if (itemsToUpdate.length === 0) return;
    const newMap: Record<string, string> = {};
    itemsToUpdate.forEach((item) => {
      let result = item.desc;
      if (advCleanMode === "FIXED") {
        const len = parseInt(advCleanValue) || 0;
        if (len > 0 && len < item.desc.length) {
          result = item.desc.substring(len);
        }
      } else {
        if (advCleanValue) {
          const idx = item.desc.indexOf(advCleanValue);
          if (idx !== -1) {
            result = item.desc.substring(idx + advCleanValue.length);
          }
        }
      }
      result = result.trim();
      if (!result) result = mappings[item.desc] || cleanPayeeHelper(item.desc);
      newMap[item.desc] = result;
    });
    const { catUpdates, commentUpdates } = analyzeCleanOperations(newMap);
    if (
      Object.keys(catUpdates).length > 0 ||
      Object.keys(commentUpdates).length > 0
    ) {
      setCleanConfirmation({ newMap, catUpdates, commentUpdates });
    } else {
      onBulkApplyMappings(newMap);
      setSelectedGroups(new Set());
      setSelectedDrillRows(new Set());
    }
    setIsAdvancedCleanModalOpen(false);
  };

  const applyCaseChange = (
    mode: "SENTENCE" | "LOWER" | "UPPER" | "TITLE" | "TOGGLE"
  ) => {
    const itemsToUpdate: { desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }));
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex))
            itemsToUpdate.push({ desc: r.desc });
        });
      });
    }
    if (itemsToUpdate.length === 0) return;
    const newMap: Record<string, string> = {};
    itemsToUpdate.forEach((item) => {
      const current = mappings[item.desc] || cleanPayeeHelper(item.desc);
      let converted = current;
      switch (mode) {
        case "SENTENCE":
          converted =
            current.charAt(0).toUpperCase() + current.slice(1).toLowerCase();
          break;
        case "LOWER":
          converted = current.toLowerCase();
          break;
        case "UPPER":
          converted = current.toUpperCase();
          break;
        case "TITLE":
          converted = current
            .toLowerCase()
            .split(" ")
            .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
            .join(" ");
          break;
        case "TOGGLE":
          converted = current
            .split("")
            .map((c) =>
              c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()
            )
            .join("");
          break;
      }
      newMap[item.desc] = converted;
    });
    const { catUpdates, commentUpdates } = analyzeCleanOperations(newMap);
    if (
      Object.keys(catUpdates).length > 0 ||
      Object.keys(commentUpdates).length > 0
    ) {
      setCleanConfirmation({ newMap, catUpdates, commentUpdates });
    } else {
      onBulkApplyMappings(newMap);
      setSelectedGroups(new Set());
      setSelectedDrillRows(new Set());
    }
    setIsChangeCaseModalOpen(false);
  };

  const openManualTypeModal = () => {
    setManualTypeValue("");
    setIsManualTypeModalOpen(true);
  };
  const handleManualTypeSubmit = () => {
    const names = Array.from(selectedGroups);
    onBulkUpdateType(names, manualTypeValue);
    setIsManualTypeModalOpen(false);
    setSelectedGroups(new Set());
  };

  const openCategoriseModal = () => {
    setCategoriseCategory("");
    setIsCategoriseModalOpen(true);
  };
  const handleCategoriseSubmit = () => {
    const names = Array.from(selectedGroups);
    onBulkUpdateCategory(names, categoriseCategory);
    setLastUsedCategory(categoriseCategory);
    setIsCategoriseModalOpen(false);
    setSelectedGroups(new Set());
  };

  const openCommentModal = () => {
    setBulkCommentText("");
    setIsCommentModalOpen(true);
  };
  const handleCommentSubmit = () => {
    const names = Array.from(selectedGroups);
    onBulkUpdateComment(names, bulkCommentText);
    setIsCommentModalOpen(false);
    setSelectedGroups(new Set());
  };

  const handleRevertNames = () => {
    const itemsToUpdate: { desc: string }[] = [];
    if (selectedGroups.size > 0) {
      payeeGroups.forEach((g) => {
        if (selectedGroups.has(g.id))
          g.rawRows.forEach((r) => itemsToUpdate.push({ desc: r.desc }));
      });
    } else {
      payeeGroups.forEach((g) => {
        g.rawRows.forEach((r) => {
          if (selectedDrillRows.has(r.originalIndex))
            itemsToUpdate.push({ desc: r.desc });
        });
      });
    }
    if (itemsToUpdate.length === 0) return;
    const newMap: Record<string, string> = {};
    itemsToUpdate.forEach((item) => {
      newMap[item.desc] = item.desc;
    });
    onBulkApplyMappings(newMap);
    setSelectedGroups(new Set());
    setSelectedDrillRows(new Set());
  };

  const isAllVisibleSelected =
    visibleIds.length > 0 && visibleIds.every((id) => selectedGroups.has(id));
  const isIndeterminate = selectedGroups.size > 0 && !isAllVisibleSelected;

  return (
    <div className="flex flex-col h-full bg-slate-50 relative overflow-hidden">
      {/* Progress Bar Header */}
      <div className="w-full h-1 bg-slate-200 shrink-0">
        <div
          className="h-full bg-emerald-500 transition-all duration-500 ease-out"
          style={{ width: `${progressPercent}%` }}
        />
      </div>

      <div className="bg-white border-b border-slate-200 px-8 py-4 flex flex-col gap-4 shadow-sm shrink-0">
        <div className="flex justify-between items-center">
          <div className="flex items-center gap-6">
            <div>
              <h2 className="text-xl font-black text-slate-800 uppercase tracking-tight">
                Pivot & Insights
              </h2>
            </div>
            <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200">
              <button
                onClick={() => setPivotMode("PAYEE")}
                className={`px-4 py-1.5 text-xs font-bold rounded-md transition-all ${
                  pivotMode === "PAYEE"
                    ? "bg-white text-blue-600 shadow-sm"
                    : "text-slate-500 hover:text-slate-700"
                }`}
              >
                BY PAYEE
              </button>
              <button
                onClick={() => setPivotMode("CATEGORY")}
                className={`px-4 py-1.5 text-xs font-bold rounded-md transition-all ${
                  pivotMode === "CATEGORY"
                    ? "bg-white text-blue-600 shadow-sm"
                    : "text-slate-500 hover:text-slate-700"
                }`}
              >
                BY CATEGORY
              </button>
            </div>
            <div className="w-px h-8 bg-slate-200 mx-1"></div>

            <button
              onClick={() => setHideCategorized(!hideCategorized)}
              title="Hide all rows that already have a category assigned"
              className={`px-4 py-1.5 text-xs font-bold rounded-lg transition-colors border ${
                hideCategorized
                  ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                  : "bg-white text-slate-500 border-slate-200 hover:bg-slate-50"
              }`}
            >
              {hideCategorized ? "👀 SHOW ALL" : "🎯 FOCUS MODE"}
            </button>

            <div className="w-px h-8 bg-slate-200 mx-1"></div>
            <div className="flex bg-slate-100 rounded-lg p-1">
              <button
                disabled={!canUndo}
                onClick={onUndo}
                className="p-2 text-slate-600 hover:bg-white hover:text-blue-600 hover:shadow-sm rounded-md disabled:opacity-30 transition-all"
                title="Undo (Ctrl+Z)"
              >
                <svg
                  width="16"
                  height="16"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M3 7v6h6" />
                  <path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13" />
                </svg>
              </button>
              <div className="w-px bg-slate-300 my-1 mx-1"></div>
              <button
                disabled={!canRedo}
                onClick={onRedo}
                className="p-2 text-slate-600 hover:bg-white hover:text-blue-600 hover:shadow-sm rounded-md disabled:opacity-30 transition-all"
                title="Redo (Ctrl+Y)"
              >
                <svg
                  width="16"
                  height="16"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M21 7v6h-6" />
                  <path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3L21 13" />
                </svg>
              </button>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <div className="flex gap-6 text-right">
              <div>
                <div className="text-[9px] font-bold text-slate-400 uppercase">
                  Out
                </div>
                <div className="text-base font-black text-slate-800">
                  £{grandTotalOut.toLocaleString()}
                </div>
              </div>
              <div>
                <div className="text-[9px] font-bold text-slate-400 uppercase">
                  In
                </div>
                <div className="text-base font-black text-emerald-600">
                  £{grandTotalIn.toLocaleString()}
                </div>
              </div>
            </div>
            <button
              onClick={handleScanDuplicates}
              className="px-4 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-bold rounded-lg shadow-sm border border-indigo-200 flex items-center gap-2"
            >
              <span>🔍</span> Find Duplicates
            </button>
            <button
              onClick={handleDownloadExcel}
              className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm"
            >
              Export XLSX
            </button>
          </div>
        </div>
        {uniqueSources.length > 0 && (
          <div className="flex items-center gap-2 border-t border-slate-100 pt-3">
            <span className="text-[10px] font-bold text-slate-400 uppercase mr-2">
              Filter Source:
            </span>
            <button
              onClick={() => setSelectedSourceFilter("ALL")}
              className={`px-3 py-1 text-[10px] font-bold rounded-full border transition-all uppercase ${
                selectedSourceFilter === "ALL"
                  ? "bg-slate-800 text-white border-slate-800"
                  : "bg-white text-slate-500 border-slate-200 hover:border-slate-300"
              }`}
            >
              Combined (All)
            </button>
            {uniqueSources.map((src) => (
              <button
                key={src}
                onClick={() => setSelectedSourceFilter(src)}
                className={`px-3 py-1 text-[10px] font-bold rounded-full border transition-all uppercase max-w-[200px] truncate ${
                  selectedSourceFilter === src
                    ? "bg-blue-600 text-white border-blue-600"
                    : "bg-white text-slate-500 border-slate-200 hover:border-slate-300"
                }`}
              >
                {src}
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="bg-white border-b border-slate-200 px-8 py-2 flex items-center gap-6 z-20 shrink-0">
        <div className="flex items-center gap-2">
          <span className="text-[9px] font-black text-slate-400 uppercase">
            Period:
          </span>
          <input
            type="date"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
            className="bg-slate-50 border border-slate-200 rounded px-2 py-0.5 text-[10px] font-bold"
          />
          <input
            type="date"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            className="bg-slate-50 border border-slate-200 rounded px-2 py-0.5 text-[10px] font-bold"
          />
          <button
            onClick={() => handleApplyPreset("this-month")}
            className="text-[9px] font-bold text-blue-600 uppercase hover:underline"
          >
            This Month
          </button>
          <button
            onClick={() => handleApplyPreset("last-month")}
            className="text-[9px] font-bold text-blue-600 uppercase hover:underline"
          >
            Last Month
          </button>
        </div>
        <div className="relative flex-1 max-w-xs">
          <input
            type="text"
            placeholder="Filter Payees..."
            value={payeeFilter}
            onChange={(e) => setPayeeFilter(e.target.value)}
            className="w-full pl-8 pr-2 py-1 bg-slate-50 border border-slate-200 rounded text-[10px] font-bold focus:border-blue-400 focus:ring-1 focus:ring-blue-400 outline-none transition-all"
          />
          <div className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400">
            <svg
              width="10"
              height="10"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="3"
            >
              <circle cx="11" cy="11" r="8" />
              <line x1="21" x2="16.65" y2="16.65" y1="21" />
            </svg>
          </div>
        </div>
      </div>

      <div className="flex-1 overflow-auto p-6 pb-24">
        <div className="max-w-7xl mx-auto bg-white rounded-xl shadow-sm border border-slate-200">
          <table className="w-full text-left border-collapse relative table-fixed">
            <thead className="bg-slate-50 border-b border-slate-200 text-[10px] font-black uppercase text-slate-500 sticky top-0 z-30 shadow-sm">
              <tr>
                <th className="px-4 py-3 w-12 text-center rounded-tl-xl bg-white border-r border-slate-100">
                  <input
                    type="checkbox"
                    onChange={handleSelectAllVisible}
                    checked={isAllVisibleSelected}
                    ref={(el) => {
                      if (el) el.indeterminate = isIndeterminate;
                    }}
                    className="rounded w-4 h-4 text-blue-600 focus:ring-0 cursor-pointer"
                  />
                </th>
                <th
                  className="px-4 py-3 relative border-r border-slate-100 group"
                  style={{ width: colWidths.entity }}
                >
                  <div className="flex items-center justify-between">
                    <span
                      className="cursor-pointer hover:text-blue-600 truncate"
                      onClick={() => handleSort("cleanName")}
                    >
                      ENTITY
                    </span>
                    <div className="relative group/hint flex items-center">
                      <input
                        type="checkbox"
                        checked={isEntitySynced}
                        onChange={onToggleEntitySync}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-0 cursor-pointer"
                      />
                    </div>
                  </div>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("entity", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative border-r border-slate-100 group"
                  style={{ width: colWidths.category }}
                >
                  <div className="flex items-center justify-between">
                    <span
                      className="cursor-pointer hover:text-blue-600 truncate"
                      onClick={() => handleSort("category")}
                    >
                      CATEGORY
                    </span>
                    <div className="relative group/hint flex items-center">
                      <input
                        type="checkbox"
                        checked={isCategorySynced}
                        onChange={onToggleCategorySync}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-0 cursor-pointer"
                      />
                    </div>
                  </div>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("category", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative border-r border-slate-100 group"
                  style={{ width: colWidths.type }}
                >
                  <div className="flex items-center justify-between">
                    <span
                      className="cursor-pointer hover:text-blue-600 truncate"
                      onClick={() => handleSort("type")}
                    >
                      TYPE
                    </span>
                    <div className="relative group/hint flex items-center">
                      <input
                        type="checkbox"
                        checked={isTypeSynced}
                        onChange={onToggleTypeSync}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-0 cursor-pointer"
                      />
                    </div>
                  </div>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("type", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative border-r border-slate-100 group"
                  style={{ width: colWidths.comment }}
                >
                  <div className="flex items-center justify-between">
                    <span className="cursor-pointer hover:text-blue-600 truncate">
                      COMMENT
                    </span>
                    <div className="relative group/hint flex items-center">
                      <input
                        type="checkbox"
                        checked={isCommentSynced}
                        onChange={onToggleCommentSync}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-blue-600 focus:ring-0 cursor-pointer"
                      />
                    </div>
                  </div>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("comment", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative text-center cursor-pointer hover:text-blue-600 group border-r border-slate-100"
                  style={{ width: colWidths.items }}
                  onClick={() => handleSort("count")}
                >
                  <span className="truncate">Items</span>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("items", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative text-right text-emerald-600 cursor-pointer hover:bg-emerald-50 group border-r border-slate-100"
                  style={{ width: colWidths.in }}
                  onClick={() => handleSort("totalIn")}
                >
                  <span className="truncate">In</span>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("in", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative text-right text-red-600 cursor-pointer hover:bg-red-50 group border-r border-slate-100"
                  style={{ width: colWidths.out }}
                  onClick={() => handleSort("totalOut")}
                >
                  <span className="truncate">Out</span>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("out", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 relative text-right text-slate-800 cursor-pointer hover:bg-slate-100 group border-r border-slate-100"
                  style={{ width: colWidths.net }}
                  onClick={() => handleSort("net")}
                >
                  <span className="truncate">Net</span>
                  <div
                    className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-400 z-40 transition-colors opacity-0 group-hover:opacity-100"
                    onMouseDown={(e) => handleStartResize("net", e)}
                  ></div>
                </th>
                <th
                  className="px-4 py-3 w-12 rounded-tr-xl text-center cursor-pointer hover:bg-slate-100 hover:text-blue-600 transition-colors"
                  onClick={handleExpandAll}
                  title={isAllExpanded ? "Collapse All" : "Expand All"}
                >
                  {isAllExpanded ? "▼" : "▶"}
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-[12px] font-medium text-slate-700">
              {pivotMode === "PAYEE"
                ? (sortedDisplayList as any[]).map((group, idx) => (
                    <React.Fragment key={group.id}>
                      <tr
                        className={`hover:bg-slate-50 transition-colors ${
                          selectedGroups.has(group.id) ? "bg-blue-50" : ""
                        } ${expandedRows.has(group.id) ? "bg-blue-50/20" : ""}`}
                      >
                        <td
                          className="px-4 py-3 text-center cursor-pointer border-r border-slate-100"
                          onClick={(e) => {
                            if (
                              e.target !==
                              e.currentTarget.querySelector("input")
                            ) {
                              toggleSelection(
                                group.id,
                                idx,
                                !selectedGroups.has(group.id),
                                e.shiftKey
                              );
                            }
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={selectedGroups.has(group.id)}
                            onChange={(e) =>
                              handleGroupCheckbox(e, idx, group.id)
                            }
                            className="rounded w-4 h-4 text-blue-600 focus:ring-0 cursor-pointer pointer-events-auto"
                            onClick={(e) => e.stopPropagation()}
                          />
                        </td>
                        <td
                          className="px-4 py-3 font-bold cursor-pointer border-r border-slate-100 truncate"
                          onClick={() => toggleRow(group.id)}
                        >
                          {group.cleanName}
                        </td>
                        <td className="px-4 py-3 border-r border-slate-100">
                          <input
                            type="text"
                            value={group.category}
                            list="category-options"
                            onChange={(e) => {
                              onUpdateCategory(group.cleanName, e.target.value);
                              setLastUsedCategory(e.target.value);
                            }}
                            className="w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 outline-none text-slate-500 text-[11px] truncate"
                          />
                        </td>
                        <td className="px-4 py-3 border-r border-slate-100">
                          <input
                            type="text"
                            value={group.type}
                            onChange={(e) =>
                              onUpdateType(group.cleanName, e.target.value)
                            }
                            className="w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 outline-none text-slate-500 text-[11px] truncate"
                          />
                        </td>
                        <td className="px-4 py-3 border-r border-slate-100">
                          <input
                            type="text"
                            value={group.comment}
                            onChange={(e) =>
                              onUpdateComment(group.cleanName, e.target.value)
                            }
                            className="w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 outline-none text-slate-500 text-[11px] truncate"
                            placeholder="Add note..."
                          />
                        </td>
                        <td className="px-4 py-3 text-center text-slate-400 border-r border-slate-100">
                          {group.count}
                        </td>
                        <td className="px-4 py-3 text-right text-emerald-600 font-bold border-r border-slate-100">
                          £{group.totalIn.toLocaleString()}
                        </td>
                        <td className="px-4 py-3 text-right text-slate-800 font-bold border-r border-slate-100">
                          £{group.totalOut.toLocaleString()}
                        </td>
                        <td className="px-4 py-3 text-right text-slate-900 font-bold border-r border-slate-100">
                          £{(group.totalOut - group.totalIn).toLocaleString()}
                        </td>
                        <td
                          className="px-4 py-3 text-center cursor-pointer"
                          onClick={() => toggleRow(group.id)}
                        >
                          {expandedRows.has(group.id) ? "▼" : "▶"}
                        </td>
                      </tr>
                      {expandedRows.has(group.id) && (
                        <DrilldownRows
                          group={group}
                          selectedDrillRows={selectedDrillRows}
                          mappings={mappings}
                          onToggle={(rowId: number) => {
                            setSelectedDrillRows((prev) => {
                              const n = new Set(prev);
                              if (n.has(rowId)) n.delete(rowId);
                              else n.add(rowId);
                              return n;
                            });
                          }}
                          onBulkToggle={handleBulkToggleDrillRows}
                        />
                      )}
                    </React.Fragment>
                  ))
                : (sortedDisplayList as any[]).map((cg) => (
                    <React.Fragment key={cg.id}>
                      <tr
                        className="bg-slate-50 font-black cursor-pointer"
                        onClick={() => toggleRow(cg.id)}
                      >
                        <td className="px-4 py-2 text-center border-r border-slate-100">
                          •
                        </td>
                        <td className="px-4 py-2 uppercase tracking-widest text-[10px] text-blue-600 truncate border-r border-slate-100">
                          {cg.name}
                        </td>
                        <td
                          colSpan={3}
                          className="border-r border-slate-100"
                        ></td>
                        <td className="px-4 py-2 text-center text-slate-400 border-r border-slate-100">
                          {cg.payees.length} Payees
                        </td>
                        <td className="px-4 py-2 text-right text-emerald-700 border-r border-slate-100">
                          £{cg.totalIn.toLocaleString()}
                        </td>
                        <td className="px-4 py-2 text-right text-slate-900 border-r border-slate-100">
                          £{cg.totalOut.toLocaleString()}
                        </td>
                        <td className="px-4 py-2 text-right text-slate-900 border-r border-slate-100">
                          £{(cg.totalOut - cg.totalIn).toLocaleString()}
                        </td>
                        <td className="px-4 py-2 text-center">
                          {expandedRows.has(cg.id) ? "▼" : "▶"}
                        </td>
                      </tr>
                      {expandedRows.has(cg.id) &&
                        cg.payees.map((pg: any, pIdx: number) => (
                          <React.Fragment key={pg.id}>
                            <tr
                              className={`hover:bg-slate-50 transition-colors ${
                                selectedGroups.has(pg.id) ? "bg-blue-50" : ""
                              }`}
                            >
                              <td
                                className="px-4 py-2 text-center cursor-pointer border-r border-slate-100"
                                onClick={(e) => {
                                  if (
                                    e.target !==
                                    e.currentTarget.querySelector("input")
                                  ) {
                                    toggleSelection(
                                      pg.id,
                                      pIdx,
                                      !selectedGroups.has(pg.id),
                                      e.shiftKey
                                    );
                                  }
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={selectedGroups.has(pg.id)}
                                  onChange={(e) =>
                                    handleGroupCheckbox(e, pIdx, pg.id)
                                  }
                                  className="rounded w-4 h-4 text-blue-600 focus:ring-0 cursor-pointer pointer-events-auto"
                                  onClick={(e) => e.stopPropagation()}
                                />
                              </td>
                              <td
                                className="px-4 py-2 pl-8 font-bold cursor-pointer border-r border-slate-100 truncate"
                                onClick={() => toggleSubRow(pg.id)}
                              >
                                {pg.cleanName}
                              </td>
                              <td
                                colSpan={3}
                                className="border-r border-slate-100"
                              ></td>
                              <td className="px-4 py-2 text-center text-slate-400 border-r border-slate-100">
                                {pg.count} tx
                              </td>
                              <td className="px-4 py-2 text-right text-emerald-600 border-r border-slate-100">
                                £{pg.totalIn.toLocaleString()}
                              </td>
                              <td className="px-4 py-2 text-right text-slate-600 border-r border-slate-100">
                                £{pg.totalOut.toLocaleString()}
                              </td>
                              <td className="px-4 py-2 text-right text-slate-600 border-r border-slate-100">
                                £{(pg.totalOut - pg.totalIn).toLocaleString()}
                              </td>
                              <td
                                className="px-4 py-2 text-center cursor-pointer"
                                onClick={() => toggleSubRow(pg.id)}
                              >
                                {expandedSubRows.has(pg.id) ? "▼" : "▶"}
                              </td>
                            </tr>
                            {expandedSubRows.has(pg.id) && (
                              <DrilldownRows
                                group={pg}
                                selectedDrillRows={selectedDrillRows}
                                mappings={mappings}
                                onToggle={(rowId: number) => {
                                  setSelectedDrillRows((prev) => {
                                    const n = new Set(prev);
                                    if (n.has(rowId)) n.delete(rowId);
                                    else n.add(rowId);
                                    return n;
                                  });
                                }}
                                onBulkToggle={handleBulkToggleDrillRows}
                              />
                            )}
                          </React.Fragment>
                        ))}
                    </React.Fragment>
                  ))}
            </tbody>
          </table>
        </div>
      </div>

      {(selectedGroups.size > 0 || selectedDrillRows.size > 0) && (
        <div className="absolute bottom-6 left-1/2 -translate-x-1/2 bg-slate-900 text-white px-6 py-2.5 rounded-full shadow-2xl flex items-center gap-6 z-50 animate-bounce-in">
          <div className="flex items-center gap-4 border-r border-slate-700 pr-6">
            <span className="text-[10px] font-bold uppercase text-slate-400">
              {selectedStats.count} Items
            </span>
            <span
              className={`text-sm font-mono font-bold ${
                selectedStats.net >= 0 ? "text-emerald-400" : "text-red-400"
              }`}
            >
              £{Math.abs(selectedStats.net).toLocaleString()}
            </span>
          </div>
          {selectedGroups.size > 0 && (
            <>
              <button
                onClick={openCategoriseModal}
                className="px-4 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-[10px] font-black uppercase rounded-full transition-colors"
              >
                CATEGORISE
              </button>
              <button
                onClick={openCommentModal}
                className="px-4 py-1.5 bg-cyan-600 hover:bg-cyan-500 text-white text-[10px] font-black uppercase rounded-full transition-colors"
              >
                COMMENT
              </button>
              <button
                onClick={openManualTypeModal}
                className="px-4 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
              >
                MANUAL TYPE
              </button>
            </>
          )}
          <button
            onClick={handleAutoPopulateTypes}
            className="px-4 py-1.5 bg-green-600 hover:bg-green-700 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            AUTO TYPE
          </button>
          <button
            onClick={() => {
              setIsAiTrimModalOpen(true);
              setUncheckedPreviewItems(new Set());
            }}
            className="px-4 py-1.5 bg-fuchsia-600 hover:bg-fuchsia-700 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            AI TRIM
          </button>
          <button
            onClick={handleAutoClean}
            className="px-4 py-1.5 bg-purple-600 hover:bg-purple-700 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            AUTO CLEAN
          </button>
          <button
            onClick={() => setIsSemiAutoModalOpen(true)}
            className="px-4 py-1.5 bg-teal-600 hover:bg-teal-700 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            SEMI-AUTO
          </button>
          <button
            onClick={openManualCleanModal}
            className="px-4 py-1.5 bg-white text-indigo-700 hover:bg-slate-100 text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            MANUAL CLEAN
          </button>
          <button
            onClick={openAdvancedCleanModal}
            className="px-4 py-1.5 bg-cyan-500 hover:bg-cyan-600 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            ADVANCED CLEAN
          </button>
          <button
            onClick={handleRevertNames}
            className="px-4 py-1.5 bg-slate-700 hover:bg-slate-600 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            REVERT TO ORIGINAL
          </button>
          <button
            onClick={() => setIsChangeCaseModalOpen(true)}
            className="px-4 py-1.5 bg-pink-500 hover:bg-pink-600 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            CHANGE CASE
          </button>
          <button
            onClick={openMergeModal}
            className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-[10px] font-black uppercase rounded-full transition-colors shadow-sm"
          >
            MERGE / MOVE
          </button>
          <button
            onClick={() => {
              setSelectedGroups(new Set());
              setSelectedDrillRows(new Set());
            }}
            className="text-slate-500 hover:text-white transition-colors ml-2"
          >
            ✕
          </button>
        </div>
      )}

      {isAiTrimModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-6xl h-[85vh] flex flex-col overflow-hidden">
            <div className="bg-white p-6 border-b border-slate-200 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-fuchsia-100 text-fuchsia-600 rounded-full flex items-center justify-center font-bold text-lg">
                  AI
                </div>
                <div>
                  <h3 className="text-lg font-bold text-slate-900">
                    AI Trim Engine
                  </h3>
                  <p className="text-xs text-slate-500 font-medium">
                    Define rules to clean entity names.
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <button
                  onClick={() => setIsAiTrimModalOpen(false)}
                  className="px-4 py-2 text-slate-400 text-sm font-bold"
                >
                  Close
                </button>
                <button
                  onClick={() => handleAiTrimSubmit(false)}
                  className="px-4 py-2 bg-purple-50 text-fuchsia-700 border border-fuchsia-200 hover:bg-purple-100 text-sm font-bold rounded-lg transition-colors"
                >
                  Apply
                </button>
                <button
                  onClick={() => handleAiTrimSubmit(true)}
                  className="px-6 py-2 bg-fuchsia-600 hover:bg-fuchsia-700 text-white text-sm font-bold rounded-lg shadow-lg"
                >
                  Apply & Close
                </button>
              </div>
            </div>
            <div className="flex-1 flex min-h-0 bg-slate-50 overflow-x-auto">
              <div className="w-1/3 min-w-[300px] p-6 border-r border-slate-200 flex flex-col bg-white overflow-y-auto shrink-0">
                <label className="text-xs font-bold text-slate-700 mb-2 uppercase block">
                  Trim Rules (Top Down)
                </label>
                <textarea
                  autoFocus
                  value={aiTrimRules}
                  onChange={(e) => setAiTrimRules(e.target.value)}
                  className="flex-1 w-full border border-slate-200 rounded-lg p-3 text-sm font-mono focus:ring-2 focus:ring-fuchsia-500 outline-none resize-none mb-6"
                />
                <label className="text-xs font-bold text-slate-700 mb-2 uppercase block">
                  Trim Mode
                </label>
                <div className="space-y-2 mb-4">
                  <label className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-slate-50 transition-colors">
                    <input
                      type="radio"
                      name="trimMode"
                      value="START_TO_MATCH"
                      checked={trimMode === "START_TO_MATCH"}
                      onChange={() => setTrimMode("START_TO_MATCH")}
                      className="text-fuchsia-600 focus:ring-0"
                    />
                    <div className="text-sm font-bold text-slate-900">
                      Delete Start → Match
                    </div>
                  </label>
                  <label className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-slate-50 transition-colors">
                    <input
                      type="radio"
                      name="trimMode"
                      value="MATCH_ONLY"
                      checked={trimMode === "MATCH_ONLY"}
                      onChange={() => setTrimMode("MATCH_ONLY")}
                      className="text-fuchsia-600 focus:ring-0"
                    />
                    <div className="text-sm font-bold text-slate-900">
                      Delete Match Only
                    </div>
                  </label>
                  <label className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-slate-50 transition-colors">
                    <input
                      type="radio"
                      name="trimMode"
                      value="MATCH_TO_END"
                      checked={trimMode === "MATCH_TO_END"}
                      onChange={() => setTrimMode("MATCH_TO_END")}
                      className="text-fuchsia-600 focus:ring-0"
                    />
                    <div className="text-sm font-bold text-slate-900">
                      Delete Match → End
                    </div>
                  </label>
                </div>
              </div>
              <div className="flex-1 min-w-[400px] p-6 overflow-y-auto bg-slate-50">
                <div className="bg-white rounded-lg shadow-sm border border-slate-200 overflow-hidden">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-100 text-slate-500 font-bold border-b border-slate-200">
                      <tr>
                        <th className="p-3 w-8 text-center">
                          <input
                            type="checkbox"
                            checked={
                              uncheckedPreviewItems.size === 0 &&
                              sortedPreview.length > 0
                            }
                            onChange={(e) => {
                              if (e.target.checked)
                                setUncheckedPreviewItems(new Set());
                              else
                                setUncheckedPreviewItems(
                                  new Set(sortedPreview.map((i) => i.original))
                                );
                            }}
                            className="rounded border-slate-300 text-fuchsia-600 focus:ring-0 cursor-pointer"
                          />
                        </th>
                        <th
                          className="p-3 w-1/2 cursor-pointer hover:bg-slate-200 select-none"
                          onClick={() =>
                            setPreviewSort((prev) => ({
                              key: "original",
                              dir:
                                prev.key === "original" && prev.dir === "asc"
                                  ? "desc"
                                  : "asc",
                            }))
                          }
                        >
                          Original{" "}
                          {previewSort.key === "original" &&
                            (previewSort.dir === "asc" ? "▲" : "▼")}
                        </th>
                        <th
                          className="p-3 w-1/2 cursor-pointer hover:bg-slate-200 select-none"
                          onClick={() =>
                            setPreviewSort((prev) => ({
                              key: "result",
                              dir:
                                prev.key === "result" && prev.dir === "asc"
                                  ? "desc"
                                  : "asc",
                            }))
                          }
                        >
                          Result{" "}
                          {previewSort.key === "result" &&
                            (previewSort.dir === "asc" ? "▲" : "▼")}
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {sortedPreview.map((item, idx) => (
                        <tr
                          key={idx}
                          className={`hover:bg-slate-50 ${
                            uncheckedPreviewItems.has(item.original)
                              ? "opacity-50 grayscale bg-slate-50"
                              : ""
                          }`}
                        >
                          <td className="p-3 text-center">
                            <input
                              type="checkbox"
                              checked={
                                !uncheckedPreviewItems.has(item.original)
                              }
                              onChange={(e) =>
                                handleAiTrimCheckbox(e, idx, item.original)
                              }
                              className="rounded border-slate-300 text-fuchsia-600 focus:ring-0 cursor-pointer"
                            />
                          </td>
                          <td className="p-3 text-slate-500 font-mono truncate max-w-[200px]">
                            {item.original}
                          </td>
                          <td className="p-3 font-bold truncate max-w-[200px]">
                            {item.result}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {isSemiAutoModalOpen && (
        <div className="fixed inset-0 z-[100] bg-white flex flex-col animate-in fade-in duration-200">
          <div className="px-8 py-4 border-b border-slate-200 flex justify-between items-center bg-white shadow-sm shrink-0">
            <div>
              <h3 className="text-2xl font-black text-slate-900">
                Semi-Automatic Grouping
              </h3>
              <p className="text-sm text-slate-500 font-medium">
                Iteratively group unmatched transactions by similarity.
              </p>
            </div>
            <div className="flex items-center gap-4">
              <button
                onClick={handleSemiAutoMergeAll}
                className="px-6 py-2 bg-slate-900 text-white text-xs font-bold rounded-lg hover:bg-black shadow-lg transition-transform active:scale-95 uppercase tracking-wide"
              >
                Merge All Confirmed
              </button>
              <button
                onClick={() => {
                  setIsSemiAutoModalOpen(false);
                  setSemiAutoGroups([]);
                }}
                className="w-10 h-10 flex items-center justify-center rounded-full bg-slate-100 hover:bg-slate-200 text-slate-500 font-bold text-xl transition-colors"
              >
                ×
              </button>
            </div>
          </div>
          {semiAutoGroups.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center space-y-6 bg-slate-50">
              <div className="w-16 h-16 bg-teal-50 text-teal-600 rounded-2xl flex items-center justify-center text-3xl shadow-sm">
                🧩
              </div>
              <div className="max-w-md">
                <label className="block text-sm font-bold text-slate-700 mb-2 uppercase">
                  Minimum Character Match
                </label>
                <div className="flex justify-center items-center gap-4">
                  <input
                    type="number"
                    min="2"
                    max="20"
                    value={semiAutoThreshold}
                    onChange={(e) =>
                      setSemiAutoThreshold(
                        Math.max(2, parseInt(e.target.value) || 2)
                      )
                    }
                    className="w-24 p-3 border border-slate-200 rounded-xl text-center font-black text-2xl focus:ring-2 focus:ring-teal-500 shadow-sm outline-none"
                  />
                  <button
                    onClick={handleSemiAutoScan}
                    className="px-8 py-3 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded-xl shadow-lg transition-transform active:scale-95"
                  >
                    Start Analysis
                  </button>
                </div>
              </div>
            </div>
          ) : (
            <div className="flex-1 overflow-y-auto bg-slate-50 p-8 space-y-8">
              <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 pb-20">
                {semiAutoGroups.map((group) => (
                  <div
                    key={group.id}
                    className={`bg-white border rounded-xl shadow-sm overflow-hidden flex flex-col h-full transition-all duration-200 border-slate-200`}
                  >
                    <div className="p-4 border-b border-slate-100 bg-slate-50 flex items-center gap-4">
                      <div className="flex-1 min-w-0">
                        <div className="text-[10px] font-bold text-slate-400 uppercase mb-1 truncate">
                          Proposed Merge Name
                        </div>
                        <input
                          type="text"
                          list="existing-entities"
                          value={group.name}
                          onChange={(e) => {
                            const newName = e.target.value;
                            setSemiAutoGroups((prev) =>
                              prev.map((g) =>
                                g.id === group.id ? { ...g, name: newName } : g
                              )
                            );
                          }}
                          className="w-full bg-transparent font-bold text-slate-800 border-b border-slate-300 focus:border-teal-500 outline-none pb-1 text-sm truncate"
                        />
                      </div>
                    </div>
                    <div className="flex-1 overflow-y-auto max-h-60 border-b border-slate-100">
                      {group.items.map((item, itemIdx) => (
                        <div
                          key={item.id}
                          className="p-3 hover:bg-slate-50 border-b border-slate-50 last:border-0"
                        >
                          <div className="text-xs font-medium truncate text-slate-700">
                            {item.desc}
                          </div>
                        </div>
                      ))}
                    </div>
                    <div className="p-2 bg-slate-50 text-[10px] text-slate-400 font-bold text-center uppercase tracking-widest">
                      {group.items.length} Items
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {isSuggestionModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[80vh]">
            <div className="p-6 border-b border-slate-200">
              <h3 className="text-xl font-black text-slate-900">
                Duplicate Suggestions
              </h3>
            </div>
            <div className="flex-1 overflow-y-auto p-6 space-y-6">
              {suggestedMerges.map((suggestion, idx) => (
                <div
                  key={idx}
                  className="border border-slate-200 rounded-xl overflow-hidden"
                >
                  <div className="bg-slate-50 p-3 border-b border-slate-200 flex justify-between items-center">
                    <div className="font-bold text-slate-700 text-sm">
                      Target:{" "}
                      <span className="text-blue-600">
                        {suggestion.target.cleanName}
                      </span>
                    </div>
                  </div>
                  {suggestion.candidates.map((cand: any) => (
                    <label
                      key={cand.id}
                      className="flex items-center gap-3 p-3 hover:bg-slate-50 cursor-pointer border-b border-slate-100 last:border-0"
                    >
                      <input
                        type="checkbox"
                        checked={selectedSuggestionIds.has(cand.id)}
                        onChange={(e) => {
                          const next = new Set(selectedSuggestionIds);
                          if (e.target.checked) next.add(cand.id);
                          else next.delete(cand.id);
                          setSelectedSuggestionIds(next);
                        }}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-0"
                      />
                      <span className="flex-1 text-sm text-slate-600">
                        {cand.cleanName}
                      </span>
                    </label>
                  ))}
                </div>
              ))}
            </div>
            <div className="p-4 border-t border-slate-200 flex justify-end gap-3 bg-slate-50 rounded-b-xl">
              <button
                onClick={() => setIsSuggestionModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-xs uppercase hover:bg-slate-200 rounded-lg"
              >
                Cancel
              </button>
              <button
                onClick={confirmMergeSuggestions}
                className="px-6 py-2 bg-indigo-600 text-white font-bold text-xs uppercase rounded-lg shadow hover:bg-indigo-700"
              >
                Merge Selected
              </button>
            </div>
          </div>
        </div>
      )}

      {isManualCleanModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-md">
            <h3 className="text-lg font-bold text-slate-900 mb-4">
              Manual Clean
            </h3>
            <input
              type="text"
              value={manualCleanText}
              onChange={(e) => setManualCleanText(e.target.value)}
              className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-indigo-500 outline-none mb-6"
              autoFocus
            />
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setIsManualCleanModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-sm hover:bg-slate-100 rounded"
              >
                Cancel
              </button>
              <button
                onClick={handleManualCleanSubmit}
                className="px-6 py-2 bg-indigo-600 text-white font-bold text-sm rounded hover:bg-indigo-700"
              >
                Clean
              </button>
            </div>
          </div>
        </div>
      )}

      {isChangeCaseModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-sm text-center">
            <h3 className="text-lg font-bold text-slate-900 mb-6 uppercase tracking-tight">
              Change Text Case
            </h3>
            <div className="grid grid-cols-1 gap-3 mb-6">
              <button
                onClick={() => applyCaseChange("TITLE")}
                className="p-3 bg-pink-50 hover:bg-pink-100 text-pink-700 font-bold rounded-lg transition-colors border border-pink-100"
              >
                Name Case (Title Case)
              </button>
              <button
                onClick={() => applyCaseChange("UPPER")}
                className="p-3 bg-pink-50 hover:bg-pink-100 text-pink-700 font-bold rounded-lg transition-colors border border-pink-100"
              >
                UPPERCASE
              </button>
              <button
                onClick={() => applyCaseChange("LOWER")}
                className="p-3 bg-pink-50 hover:bg-pink-100 text-pink-700 font-bold rounded-lg transition-colors border border-pink-100"
              >
                lowercase
              </button>
              <button
                onClick={() => applyCaseChange("SENTENCE")}
                className="p-3 bg-pink-50 hover:bg-pink-100 text-pink-700 font-bold rounded-lg transition-colors border border-pink-100"
              >
                Sentence case
              </button>
            </div>
            <button
              onClick={() => setIsChangeCaseModalOpen(false)}
              className="text-slate-400 font-bold text-xs uppercase hover:text-slate-600"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {isManualTypeModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-md">
            <h3 className="text-lg font-bold text-slate-900 mb-4">Set Type</h3>
            <input
              type="text"
              value={manualTypeValue}
              onChange={(e) => setManualTypeValue(e.target.value)}
              className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-emerald-500 outline-none mb-6"
              autoFocus
            />
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setIsManualTypeModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-sm hover:bg-slate-100 rounded"
              >
                Cancel
              </button>
              <button
                onClick={handleManualTypeSubmit}
                className="px-6 py-2 bg-emerald-600 text-white font-bold text-sm rounded hover:bg-emerald-700"
              >
                Set Type
              </button>
            </div>
          </div>
        </div>
      )}

      {isCategoriseModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-md">
            <h3 className="text-lg font-bold text-slate-900 mb-4">
              Set Category
            </h3>
            <input
              type="text"
              value={categoriseCategory}
              onChange={(e) => setCategoriseCategory(e.target.value)}
              className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-amber-500 outline-none mb-6"
              autoFocus
            />
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setIsCategoriseModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-sm hover:bg-slate-100 rounded"
              >
                Cancel
              </button>
              <button
                onClick={handleCategoriseSubmit}
                className="px-6 py-2 bg-amber-500 text-white font-bold text-sm rounded hover:bg-amber-600"
              >
                Set Category
              </button>
            </div>
          </div>
        </div>
      )}

      {isCommentModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-md">
            <h3 className="text-lg font-bold text-slate-900 mb-4">
              Add Comment
            </h3>
            <input
              type="text"
              value={bulkCommentText}
              onChange={(e) => setBulkCommentText(e.target.value)}
              className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-cyan-500 outline-none mb-6"
              autoFocus
            />
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setIsCommentModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-sm hover:bg-slate-100 rounded"
              >
                Cancel
              </button>
              <button
                onClick={handleCommentSubmit}
                className="px-6 py-2 bg-cyan-600 text-white font-bold text-sm rounded hover:bg-cyan-700"
              >
                Add Comment
              </button>
            </div>
          </div>
        </div>
      )}

      {isMergeModalOpen && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-2xl animate-bounce-in flex flex-col max-h-[85vh]">
            <h3 className="text-xl font-black text-slate-900 mb-4">
              {mergeMode === "main_merge"
                ? "Merge Entities"
                : "Move Transactions"}
            </h3>
            <div className="flex gap-6 overflow-hidden">
              <div className="flex-1 space-y-4 overflow-y-auto">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
                    Target Entity Name
                  </label>
                  <input
                    type="text"
                    value={mergeName}
                    list="merge-options"
                    onChange={(e) => setMergeName(e.target.value)}
                    className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-indigo-500 outline-none font-bold text-slate-800"
                    autoFocus
                  />
                  <datalist id="merge-options">
                    {allExistingEntities.map((name, i) => (
                      <option key={i} value={name} />
                    ))}
                  </datalist>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
                    Set Category (Optional)
                  </label>
                  <input
                    type="text"
                    value={mergeCategory}
                    onChange={(e) => setMergeCategory(e.target.value)}
                    className="w-full p-2 border border-slate-300 rounded focus:ring-2 focus:ring-indigo-500 outline-none text-slate-800"
                    placeholder="e.g. Utilities"
                  />
                </div>
                <div className="bg-blue-50 p-3 rounded text-[11px] text-blue-700 border border-blue-100 leading-relaxed font-medium">
                  {mergeMode === "main_merge"
                    ? `Merging ${
                        selectedGroups.size
                      } group(s) will consolidate all linked descriptions under '${
                        mergeName || "the target"
                      }'.`
                    : `Moving ${
                        selectedDrillRows.size
                      } row(s) will re-map those specific descriptions to '${
                        mergeName || "the target"
                      }'.`}
                </div>
              </div>

              {mergeCandidates.length > 0 && (
                <div className="w-64 border-l border-slate-100 pl-4 flex flex-col">
                  <span className="text-[10px] font-bold text-slate-400 uppercase mb-2">
                    Suggestions
                  </span>
                  <div className="flex-1 overflow-y-auto space-y-2 pr-1">
                    {mergeCandidates.map((cand, idx) => (
                      <button
                        key={idx}
                        onClick={() => setMergeName(cand)}
                        className="w-full text-left p-2 bg-slate-50 hover:bg-indigo-50 text-[11px] text-slate-600 hover:text-indigo-700 rounded border border-slate-100 transition-colors break-words font-medium"
                      >
                        {cand}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <div className="mt-6 flex justify-end gap-3 pt-4 border-t border-slate-100">
              <button
                onClick={() => setIsMergeModalOpen(false)}
                className="px-4 py-2 text-slate-500 font-bold text-sm hover:bg-slate-100 rounded"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  if (!mergeName) return;
                  const selection: { index: number; desc: string }[] = [];
                  if (mergeMode === "main_merge") {
                    payeeGroups.forEach((g) => {
                      if (selectedGroups.has(g.id))
                        g.rawRows.forEach((r) =>
                          selection.push({
                            index: r.originalIndex,
                            desc: r.desc,
                          })
                        );
                    });
                  } else {
                    payeeGroups.forEach((g) => {
                      g.rawRows.forEach((r) => {
                        if (selectedDrillRows.has(r.originalIndex))
                          selection.push({
                            index: r.originalIndex,
                            desc: r.desc,
                          });
                      });
                    });
                  }
                  onMergePayees(mergeName, mergeCategory, selection);
                  setSelectedGroups(new Set());
                  setSelectedDrillRows(new Set());
                  setIsMergeModalOpen(false);
                }}
                disabled={!mergeName.trim()}
                className="px-6 py-2 bg-indigo-600 text-white font-bold text-sm rounded hover:bg-indigo-700 disabled:opacity-50 shadow-lg"
              >
                Confirm Action
              </button>
            </div>
          </div>
        </div>
      )}

      {cleanConfirmation && (
        <div className="fixed inset-0 bg-black/50 z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 w-full max-w-md animate-bounce-in">
            <h3 className="text-lg font-bold text-slate-900 mb-2 tracking-tight">
              Apply Existing Metadata?
            </h3>
            <p className="text-sm text-slate-500 mb-4 leading-relaxed font-medium">
              Updated entities were found that match existing records. Would you
              like to auto-populate categories and comments based on those
              records?
            </p>
            <div className="flex flex-col gap-2 mt-6">
              <button
                onClick={confirmCleanWithUpdates}
                className="px-4 py-2 bg-blue-600 text-white font-bold text-sm rounded hover:bg-blue-700 shadow-md"
              >
                Yes, Apply Metadata
              </button>
              <button
                onClick={confirmCleanWithoutUpdates}
                className="px-4 py-2 bg-white text-slate-600 border border-slate-200 font-bold text-slate-600 rounded hover:bg-slate-50"
              >
                No, Only Update Names
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
