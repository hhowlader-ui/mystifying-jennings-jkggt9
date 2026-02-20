import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { AppState, ExtractedData } from "./types";
import { FileUploader } from "./components/FileUploader";
import { ResultsTable } from "./components/ResultsTable";
import { PivotView } from "./components/PivotView";
import { ColumnMapper, MappingConfig } from "./components/ColumnMapper";
import { parseBankStatementAuto } from "./services/geminiService";
import { getPdfPageCount, splitPdf } from "./services/pdfUtils";

// --- Shared Logic for Payee Cleaning ---
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
];

BANKING_TERMS.sort((a, b) => b.length - a.length);

const NOISE_PATTERNS = BANKING_TERMS.map((term) => {
  const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const startBoundary = /^\w/.test(term) ? "\\b" : "";
  const endBoundary = /\w$/.test(term) ? "\\b" : "";
  return new RegExp(`${startBoundary}${escaped}${endBoundary}`, "gi");
});

const cleanPayee = (rawName: string) => {
  if (!rawName) return "UNKNOWN";
  let cleaned = rawName.toUpperCase();
  NOISE_PATTERNS.forEach((pattern) => {
    cleaned = cleaned.replace(pattern, " ");
  });
  cleaned = cleaned.replace(/\s+/g, " ").trim();
  return cleaned.length < 2 ? rawName.trim() || "UNKNOWN" : cleaned;
};

// --- Shared Date Helper for Sorting ---
const parseBankDate = (dateStr: string): Date | null => {
  if (!dateStr) return null;
  let d = new Date(dateStr);
  if (!isNaN(d.getTime())) return d;
  const parts = dateStr.trim().split(/[\s\/\-\.]+/);
  if (parts.length >= 2) {
    const day = parseInt(parts[0]);
    const monthStr = parts[1].toUpperCase();
    const months: Record<string, number> = {
      JAN: 0,
      FEB: 1,
      MAR: 2,
      APR: 3,
      MAY: 4,
      JUN: 5,
      JUL: 6,
      AUG: 7,
      SEP: 8,
      OCT: 9,
      NOV: 10,
      DEC: 11,
    };
    const month = months[monthStr.substring(0, 3)];
    if (month !== undefined && !isNaN(day)) {
      let year =
        parts.length >= 3 ? parseInt(parts[2]) : new Date().getFullYear();
      if (year < 100) year += 2000;
      return new Date(year, month, day);
    }
  }
  return null;
};

interface HistorySnapshot {
  data: ExtractedData;
  categories: Record<string, string>;
  comments: Record<string, string>;
  mappings: Record<string, string>;
  types: Record<string, string>;
}

export default function App() {
  const [state, setState] = useState<AppState>({
    step: "UPLOAD",
    data: null,
    error: null,
  });

  const [history, setHistory] = useState<HistorySnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([]);
  const [accumulatedUsage, setAccumulatedUsage] = useState<{
    prompt: number;
    response: number;
  }>({ prompt: 0, response: 0 });
  const [viewMode, setViewMode] = useState<"TABLE" | "PIVOT">("TABLE");
  const [categories, setCategories] = useState<Record<string, string>>({});
  const [comments, setComments] = useState<Record<string, string>>({});
  const [mappings, setMappings] = useState<Record<string, string>>({});
  const [types, setTypes] = useState<Record<string, string>>({});
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [showChunker, setShowChunker] = useState(false);
  const [showSourceSelector, setShowSourceSelector] = useState(false);
  const [totalPages, setTotalPages] = useState(0);
  const [chunkSize, setChunkSize] = useState(5);
  const [processingStatus, setProcessingStatus] = useState<string | null>(null);
  const [targetSource, setTargetSource] = useState<string>("Statement 1");
  const [newSourceName, setNewSourceName] = useState("");
  const [isAugmenting, setIsAugmenting] = useState(false);
  const [augmentIndices, setAugmentIndices] = useState<number[]>([]);
  const [pendingSpreadsheet, setPendingSpreadsheet] = useState<{
    headers: string[];
    rows: any[][];
    file: File;
  } | null>(null);
  const [showMapper, setShowMapper] = useState(false);
  const [pendingSourceLabel, setPendingSourceLabel] = useState<string>("");

  const PRICE_PER_1M_INPUT = 0.075;
  const PRICE_PER_1M_OUTPUT = 0.3;
  const calculateCost = (prompt: number, response: number) => {
    const inputCost = (prompt / 1_000_000) * PRICE_PER_1M_INPUT;
    const outputCost = (response / 1_000_000) * PRICE_PER_1M_OUTPUT;
    return inputCost + outputCost;
  };

  const totalCost = calculateCost(
    accumulatedUsage.prompt,
    accumulatedUsage.response
  );

  const normalizeData = (
    data: ExtractedData,
    sourceLabel: string = "Statement 1"
  ): ExtractedData => {
    const MANUAL_HEADER = "Manual Calc";
    const SOURCE_HEADER = "Source";
    let headers = [...data.headers];
    let rows = data.rows.map((row) => [...row]);

    if (!headers.includes(SOURCE_HEADER)) {
      headers.unshift(SOURCE_HEADER);
      rows = rows.map((row) => [sourceLabel, ...row]);
    }

    const manualIndex = headers.indexOf(MANUAL_HEADER);
    if (manualIndex === -1) {
      headers.push(MANUAL_HEADER);
      rows = rows.map((row) => [...row, ""]);
    } else {
      rows = rows.map((row) => {
        if (row.length < headers.length) {
          return [...row, ...Array(headers.length - row.length).fill("")];
        }
        return row;
      });
    }
    return { headers, rows };
  };

  const truncateData = (data: ExtractedData): ExtractedData => {
    const descIdx = data.headers.findIndex((h) =>
      /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
        h
      )
    );
    if (descIdx === -1) return data;
    const newRows = data.rows.map((row) => {
      const newRow = [...row];
      if (newRow[descIdx] && newRow[descIdx].length > 50) {
        newRow[descIdx] = newRow[descIdx].substring(0, 50);
      }
      return newRow;
    });
    return { ...data, rows: newRows };
  };

  const saveToHistory = () => {
    if (!state.data) return;
    const snapshot: HistorySnapshot = {
      data: state.data,
      categories: { ...categories },
      comments: { ...comments },
      mappings: { ...mappings },
      types: { ...types },
    };
    setHistory((prev) => {
      const next = [...prev, snapshot];
      return next.length > 50 ? next.slice(next.length - 50) : next;
    });
    setRedoStack([]);
  };

  const handleUndo = () => {
    if (history.length === 0 || !state.data) return;
    const previous = history[history.length - 1];
    const newHistory = history.slice(0, -1);
    const currentSnapshot: HistorySnapshot = {
      data: state.data,
      categories: { ...categories },
      comments: { ...comments },
      mappings: { ...mappings },
      types: { ...types },
    };
    setRedoStack((prev) => [currentSnapshot, ...prev]);
    setHistory(newHistory);
    setState((prev) => ({ ...prev, data: previous.data }));
    setCategories(previous.categories);
    setComments(previous.comments);
    setMappings(previous.mappings);
    setTypes(previous.types || {});
  };

  const handleRedo = () => {
    if (redoStack.length === 0 || !state.data) return;
    const next = redoStack[0];
    const newRedo = redoStack.slice(1);
    const currentSnapshot: HistorySnapshot = {
      data: state.data,
      categories: { ...categories },
      comments: { ...comments },
      mappings: { ...mappings },
      types: { ...types },
    };
    setHistory((prev) => [...prev, currentSnapshot]);
    setRedoStack(newRedo);
    setState((prev) => ({ ...prev, data: next.data }));
    setCategories(next.categories);
    setComments(next.comments);
    setMappings(next.mappings);
    setTypes(next.types || {});
  };

  const modifyData = (
    transformer: (
      currentRows: string[][],
      currentHeaders: string[]
    ) => string[][],
    skipHistory: boolean = false
  ) => {
    if (!skipHistory) saveToHistory();
    setState((prev) => {
      if (!prev.data) return prev;
      const newRows = transformer(prev.data.rows, prev.data.headers);
      return {
        ...prev,
        data: { ...prev.data, rows: newRows },
      };
    });
  };

  const handleToggleColumn = (
    columnName: string,
    deriveValue: (row: string[], headers: string[]) => string
  ) => {
    if (!state.data) return;
    saveToHistory();
    const currentData = state.data;
    const headerIdx = currentData.headers.indexOf(columnName);
    let newHeaders = [...currentData.headers];
    let newRows = currentData.rows.map((r) => [...r]);

    if (headerIdx !== -1) {
      newHeaders.splice(headerIdx, 1);
      newRows = newRows.map((row) => {
        const r = [...row];
        r.splice(headerIdx, 1);
        return r;
      });
    } else {
      let insertIndex = newHeaders.length;
      const descRegex =
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i;
      const descIdx = newHeaders.findIndex((h) => descRegex.test(h));
      if (columnName === "Entity") {
        if (descIdx !== -1) insertIndex = descIdx + 1;
      } else if (columnName === "Category") {
        const entityIdx = newHeaders.indexOf("Entity");
        if (entityIdx !== -1) insertIndex = entityIdx + 1;
        else if (descIdx !== -1) insertIndex = descIdx + 1;
      } else if (columnName === "Type") {
        const catIdx = newHeaders.indexOf("Category");
        const entityIdx = newHeaders.indexOf("Entity");
        if (catIdx !== -1) insertIndex = catIdx + 1;
        else if (entityIdx !== -1) insertIndex = entityIdx + 1;
        else if (descIdx !== -1) insertIndex = descIdx + 1;
      } else if (columnName === "Comment") {
        const typeIdx = newHeaders.indexOf("Type");
        const catIdx = newHeaders.indexOf("Category");
        if (typeIdx !== -1) insertIndex = typeIdx + 1;
        else if (catIdx !== -1) insertIndex = catIdx + 1;
        else if (descIdx !== -1) insertIndex = descIdx + 1;
      } else {
        const manualIdx = newHeaders.indexOf("Manual Calc");
        if (manualIdx !== -1) insertIndex = manualIdx;
      }
      newHeaders.splice(insertIndex, 0, columnName);
      newRows = newRows.map((row) => {
        const newRow = [...row];
        const val = deriveValue(row, currentData.headers);
        newRow.splice(insertIndex, 0, val);
        return newRow;
      });
    }
    setState((prev) => ({
      ...prev,
      data: { ...prev.data!, headers: newHeaders, rows: newRows },
    }));
  };

  const handleToggleEntityColumn = () => {
    handleToggleColumn("Entity", (row, headers) => {
      const descIdx = headers.findIndex((h) =>
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
          h
        )
      );
      const desc = descIdx !== -1 ? row[descIdx] : "";
      return mappings[desc] || cleanPayee(desc);
    });
  };

  const handleToggleCategoryColumn = () => {
    handleToggleColumn("Category", (row, headers) => {
      const descIdx = headers.findIndex((h) =>
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
          h
        )
      );
      const desc = descIdx !== -1 ? row[descIdx] : "";
      const entity = mappings[desc] || cleanPayee(desc);
      return categories[entity] || "";
    });
  };

  const handleToggleTypeColumn = () => {
    handleToggleColumn("Type", (row, headers) => {
      const descIdx = headers.findIndex((h) =>
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
          h
        )
      );
      const desc = descIdx !== -1 ? row[descIdx] : "";
      const entity = mappings[desc] || cleanPayee(desc);
      return types[entity] || "";
    });
  };

  const handleToggleCommentColumn = () => {
    handleToggleColumn("Comment", (row, headers) => {
      const descIdx = headers.findIndex((h) =>
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
          h
        )
      );
      const desc = descIdx !== -1 ? row[descIdx] : "";
      const entity = mappings[desc] || cleanPayee(desc);
      return comments[entity] || "";
    });
  };

  const handleSaveProject = () => {
    if (!state.data) return;
    const projectData = {
      ...state.data,
      categories,
      comments,
      mappings,
      types,
      accumulatedUsage,
    };
    const jsonString = JSON.stringify(projectData, null, 2);
    const blob = new Blob([jsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `fastscan_project_${new Date()
      .toISOString()
      .slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const cleanNumber = (str: string) => {
    if (!str) return 0;
    const clean = str.replace(/[^0-9.-]+/g, "");
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  };

  const calculateSimilarity = (s1: string, s2: string) => {
    if (!s1 || !s2) return 0;
    const a = s1.toLowerCase().trim();
    const b = s2.toLowerCase().trim();
    if (a === b) return 100;
    if (a.includes(b) || b.includes(a)) return 80;
    const wordsA = a.split(/\W+/).filter((w) => w.length > 2);
    const wordsB = b.split(/\W+/).filter((w) => w.length > 2);
    if (wordsA.length === 0 || wordsB.length === 0) return 0;
    const intersection = wordsA.filter((w) => wordsB.includes(w));
    return (intersection.length / Math.max(wordsA.length, wordsB.length)) * 100;
  };

  const performAugmentation = (
    currentData: ExtractedData,
    newScanData: ExtractedData,
    indicesToFix: number[]
  ) => {
    const headers = currentData.headers.map((h) => h.toLowerCase());
    const newHeaders = newScanData.headers.map((h) => h.toLowerCase());
    const getColIdx = (hList: string[], patterns: RegExp[]) =>
      hList.findIndex((h) => patterns.some((p) => p.test(h)));
    const datePatterns = [/date/i];
    const descPatterns = [/desc|details|payee|narrative|memo|transaction/i];
    const inPatterns = [/credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i];
    const outPatterns = [
      /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i,
    ];
    const typePatterns = [/type|code/i];

    const cDateIdx = getColIdx(headers, datePatterns);
    const cDescIdx = getColIdx(headers, descPatterns);
    const cInIdx = getColIdx(headers, inPatterns);
    const cOutIdx = getColIdx(headers, outPatterns);
    const cTypeIdx = getColIdx(headers, typePatterns);

    const nDateIdx = getColIdx(newHeaders, datePatterns);
    const nDescIdx = getColIdx(newHeaders, descPatterns);
    const nInIdx = getColIdx(newHeaders, inPatterns);
    const nOutIdx = getColIdx(newHeaders, outPatterns);
    const nTypeIdx = getColIdx(newHeaders, typePatterns);

    if (cDescIdx === -1 || nDescIdx === -1) return currentData;

    const sortedIndices = [...indicesToFix].sort((a, b) => a - b);
    const newRows = currentData.rows.map((row) => [...row]);
    const newScanRows = newScanData.rows;
    let lastMatchIndex = -1;

    sortedIndices.forEach((idx) => {
      const targetRow = newRows[idx];
      const targetDate = cDateIdx !== -1 ? targetRow[cDateIdx] : "";
      const targetDesc = cDescIdx !== -1 ? targetRow[cDescIdx] : "";
      const targetIn = cInIdx !== -1 ? cleanNumber(targetRow[cInIdx]) : 0;
      const targetOut = cOutIdx !== -1 ? cleanNumber(targetRow[cOutIdx]) : 0;
      const targetHasAmount = targetIn !== 0 || targetOut !== 0;
      let bestMatchIndex = -1;
      let bestScore = -Infinity;
      const startIndex = lastMatchIndex + 1;

      for (let sIdx = startIndex; sIdx < newScanRows.length; sIdx++) {
        const scanRow = newScanRows[sIdx];
        const scanDate = nDateIdx !== -1 ? scanRow[nDateIdx] : "";
        const scanDesc = nDescIdx !== -1 ? scanRow[nDescIdx] : "";
        const scanIn = nInIdx !== -1 ? scanRow[nInIdx] : 0;
        const scanOut = nOutIdx !== -1 ? scanRow[nOutIdx] : 0;
        const scanHasAmount = scanIn !== 0 || scanOut !== 0;
        let score = 0;

        if (targetDate && scanDate) {
          if (targetDate === scanDate) score += 30;
          else score -= 50;
        } else if (!targetDate && scanDate) {
          score += 5;
        }

        if (targetHasAmount && scanHasAmount) {
          const matchIn =
            Math.abs(targetIn - cleanNumber(String(scanIn))) < 0.05;
          const matchOut =
            Math.abs(targetOut - cleanNumber(String(scanOut))) < 0.05;
          if (matchIn && matchOut) score += 40;
          else score -= 40;
        } else if (!targetHasAmount && scanHasAmount) {
          score += 20;
        } else if (targetHasAmount && !scanHasAmount) {
          score -= 10;
        }

        const sim = calculateSimilarity(targetDesc, scanDesc);
        score += sim * 0.3;
        const distance = sIdx - startIndex;
        if (distance < 5) score += 5 - distance;

        if (score > bestScore) {
          bestScore = score;
          bestMatchIndex = sIdx;
        }
      }

      if (bestMatchIndex !== -1 && bestScore > 20) {
        const bestMatchRow = newScanRows[bestMatchIndex];
        const newDesc = String(bestMatchRow[nDescIdx] || "");
        const oldDesc = String(targetRow[cDescIdx] || "");
        if (newDesc.length >= oldDesc.length) {
          newRows[idx][cDescIdx] = newDesc.substring(0, 50);
        }
        const newIn = nInIdx !== -1 ? String(bestMatchRow[nInIdx]) : "";
        const newOut = nOutIdx !== -1 ? String(bestMatchRow[nOutIdx]) : "";
        if (targetIn === 0 && targetOut === 0) {
          if (cInIdx !== -1) newRows[idx][cInIdx] = newIn;
          if (cOutIdx !== -1) newRows[idx][cOutIdx] = newOut;
        }
        if (cTypeIdx !== -1 && nTypeIdx !== -1) {
          const newType = String(bestMatchRow[nTypeIdx]);
          if (!targetRow[cTypeIdx] && newType) {
            newRows[idx][cTypeIdx] = newType;
          }
        }
        lastMatchIndex = bestMatchIndex;
      }
    });
    return { ...currentData, rows: newRows };
  };

  const executeProcessing = async (
    file: File,
    isAppending: boolean,
    sourceLabel: string
  ) => {
    const contextHeaders = state.data
      ? state.data.headers.filter((h) => h !== "Manual Calc" && h !== "Source")
      : undefined;
    const rawData = await parseBankStatementAuto(file, contextHeaders);

    if (rawData.usage) {
      setAccumulatedUsage((prev) => ({
        prompt: prev.prompt + (rawData.usage?.promptTokens || 0),
        response: prev.response + (rawData.usage?.responseTokens || 0),
      }));
    }

    const truncatedRawData = truncateData(rawData);
    let mergedData: ExtractedData;

    if (isAugmenting && state.data) {
      saveToHistory();
      mergedData = performAugmentation(
        state.data,
        truncatedRawData,
        augmentIndices
      );
    } else if (state.data) {
      saveToHistory();
      const targetHeaders = state.data.headers;
      const sourceIdx = targetHeaders.indexOf("Source");
      const alignedRows = truncatedRawData.rows.map((row) => {
        return targetHeaders.map((targetH, idx) => {
          if (idx === sourceIdx) return sourceLabel;
          if (targetH === "Manual Calc") return "";
          const matchIdx = truncatedRawData.headers.findIndex(
            (h) => h.toLowerCase() === targetH.toLowerCase()
          );
          if (matchIdx !== -1 && matchIdx < row.length) {
            return row[matchIdx];
          }
          return "";
        });
      });
      mergedData = {
        headers: targetHeaders,
        rows: [...state.data.rows, ...alignedRows],
      };
    } else {
      mergedData = normalizeData(truncatedRawData, sourceLabel);
    }
    return mergedData;
  };

  const handleFileSelect = async (file: File) => {
    if (file.type === "application/json" || file.name.endsWith(".json")) {
      try {
        const text = await file.text();
        const json = JSON.parse(text);
        if (json.headers && Array.isArray(json.rows)) {
          const normalized = normalizeData(json, "Imported");
          setState({ step: "RESULTS", data: normalized, error: null });
          if (json.categories) setCategories(json.categories);
          if (json.comments) setComments(json.comments);
          if (json.mappings) setMappings(json.mappings);
          if (json.types) setTypes(json.types);
          if (json.accumulatedUsage) setAccumulatedUsage(json.accumulatedUsage);
          setHistory([]);
          setRedoStack([]);
          setIsAugmenting(false);
          setAugmentIndices([]);
          return;
        } else {
          throw new Error("Invalid project file format");
        }
      } catch (e: any) {
        setState((prev) => ({
          ...prev,
          error: "Failed to load project: " + e.message,
        }));
        return;
      }
    }

    if (/\.(csv|xls|xlsx)$/i.test(file.name)) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const readFn = XLSX.read || (XLSX as any).default?.read;
          const utilsFn = XLSX.utils || (XLSX as any).default?.utils;
          if (!readFn || !utilsFn)
            throw new Error(
              "XLSX library not loaded correctly. Please refresh."
            );

          const workbook = readFn(data, { type: "array", cellDates: true });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = utilsFn.sheet_to_json(firstSheet, {
            header: 1,
            defval: "",
          }) as any[][];

          if (json.length > 0) {
            const headers = json[0].map((h: any) => String(h || ""));
            const rows = json.slice(1).map((r) => {
              if (!Array.isArray(r)) return [];
              return r.map((c) => {
                if (c instanceof Date) {
                  return c.toLocaleDateString("en-GB", {
                    day: "2-digit",
                    month: "short",
                    year: "2-digit",
                  });
                }
                return c !== null && c !== undefined ? String(c) : "";
              });
            });

            setPendingSpreadsheet({ headers, rows, file });
            if (state.data && !isAugmenting) {
              const sourceIdx = state.data.headers.indexOf("Source");
              const existingSources =
                sourceIdx !== -1
                  ? Array.from(
                      new Set(state.data.rows.map((r) => r[sourceIdx]))
                    ).filter(Boolean)
                  : ["Statement 1"];
              if (existingSources.length > 0)
                setTargetSource(
                  String(existingSources[existingSources.length - 1])
                );
              setShowSourceSelector(true);
            } else {
              setPendingSourceLabel("Statement 1");
              setShowMapper(true);
            }
          }
        } catch (error: any) {
          setState((prev) => ({
            ...prev,
            error:
              "Spreadsheet read failed: " + (error.message || "Unknown error"),
          }));
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    if (state.data && !isAugmenting) {
      setPendingFile(file);
      const sourceIdx = state.data.headers.indexOf("Source");
      const existingSources =
        sourceIdx !== -1
          ? Array.from(
              new Set(state.data.rows.map((r) => r[sourceIdx]))
            ).filter(Boolean)
          : ["Statement 1"];
      if (existingSources.length > 0)
        setTargetSource(String(existingSources[existingSources.length - 1]));
      setShowSourceSelector(true);
      return;
    }

    if (file.type === "application/pdf" && !isAugmenting) {
      try {
        const count = await getPdfPageCount(file);
        setTotalPages(count);
        setPendingFile(file);
        setShowChunker(true);
        return;
      } catch (e) {}
    }

    setState((prev) => ({ ...prev, step: "PROCESSING", error: null }));
    try {
      const newData = await executeProcessing(
        file,
        !!state.data,
        "Statement 1"
      );
      setState({ step: "RESULTS", data: newData, error: null });
      if (isAugmenting) {
        setIsAugmenting(false);
        setAugmentIndices([]);
      }
    } catch (e: any) {
      setState((prev) => ({
        ...prev,
        step: prev.data ? "RESULTS" : "UPLOAD",
        error: e.message || "Failed to read file.",
      }));
    }
  };

  const handleAugmentTrigger = (indices: number[]) => {
    setAugmentIndices(indices);
    setIsAugmenting(true);
    setState((prev) => ({ ...prev, step: "UPLOAD" }));
  };

  const confirmChunkProcessing = async () => {
    if (!pendingFile) return;
    setShowChunker(false);
    const finalSourceLabel = showSourceSelector
      ? targetSource === "NEW_SOURCE"
        ? newSourceName || "New Statement"
        : targetSource
      : "Statement 1";
    setShowSourceSelector(false);

    try {
      setProcessingStatus(`Initializing scan for ${finalSourceLabel}...`);
      setState((prev) => ({ ...prev, step: "PROCESSING", error: null }));
      const chunkGenerator = splitPdf(pendingFile, chunkSize);
      let currentHeaders = state.data ? state.data.headers : undefined;
      let currentRows = state.data ? state.data.rows : [];
      let batchCount = 0;
      const totalBatches = Math.ceil(totalPages / chunkSize);

      for await (const chunk of chunkGenerator) {
        batchCount++;
        setProcessingStatus(
          `Scanning batch ${batchCount} of ${totalBatches}...`
        );
        const contextHeaders = currentHeaders
          ? currentHeaders.filter((h) => h !== "Manual Calc" && h !== "Source")
          : undefined;
        const rawData = await parseBankStatementAuto(chunk, contextHeaders);

        if (rawData.usage) {
          setAccumulatedUsage((prev) => ({
            prompt: prev.prompt + (rawData.usage?.promptTokens || 0),
            response: prev.response + (rawData.usage?.responseTokens || 0),
          }));
        }

        const truncatedRawData = truncateData(rawData);
        if (!currentHeaders) {
          const normalized = normalizeData(truncatedRawData, finalSourceLabel);
          currentHeaders = normalized.headers;
          currentRows = [...currentRows, ...normalized.rows];
        } else {
          const sourceIdx = currentHeaders.indexOf("Source");
          const alignedRows = truncatedRawData.rows.map((row) => {
            return currentHeaders!.map((targetH, idx) => {
              if (idx === sourceIdx) return finalSourceLabel;
              if (targetH === "Manual Calc") return "";
              const sourceIdxRaw = truncatedRawData.headers.findIndex(
                (h) => h.toLowerCase() === targetH.toLowerCase()
              );
              if (sourceIdxRaw !== -1 && sourceIdxRaw < row.length)
                return row[sourceIdxRaw];
              return "";
            });
          });
          currentRows = [...currentRows, ...alignedRows];
        }
        setState({
          step: "RESULTS",
          data: { headers: currentHeaders!, rows: currentRows },
          error: null,
        });
        await new Promise((r) => setTimeout(r, 100));
      }
    } catch (e: any) {
      setState((prev) => ({
        ...prev,
        error: "Batch processing failed: " + e.message,
      }));
    } finally {
      setProcessingStatus(null);
      setPendingFile(null);
    }
  };

  const confirmSourceSelection = async () => {
    const finalSourceLabel =
      targetSource === "NEW_SOURCE"
        ? newSourceName || "New Statement"
        : targetSource;
    setShowSourceSelector(false);
    if (pendingSpreadsheet) {
      setPendingSourceLabel(finalSourceLabel);
      setShowMapper(true);
      return;
    }
    if (!pendingFile) return;
    if (pendingFile.type === "application/pdf") {
      try {
        const count = await getPdfPageCount(pendingFile);
        if (count > 5) {
          setTotalPages(count);
          setShowChunker(true);
          return;
        }
      } catch (e) {}
    }
    setState((prev) => ({ ...prev, step: "PROCESSING", error: null }));
    try {
      const newData = await executeProcessing(
        pendingFile,
        true,
        finalSourceLabel
      );
      setState({ step: "RESULTS", data: newData, error: null });
    } catch (e: any) {
      setState((prev) => ({ ...prev, step: "RESULTS", error: e.message }));
    } finally {
      setPendingFile(null);
    }
  };

  const handleMapperConfirm = (mapping: MappingConfig) => {
    if (!pendingSpreadsheet) return;
    setShowMapper(false);
    const stdHeaders = [
      "Date",
      "Description",
      "Money Out",
      "Money In",
      "Balance",
    ];
    const transformedRows = pendingSpreadsheet.rows.map((row) => {
      let date =
        mapping.dateIdx !== -1 ? String(row[mapping.dateIdx] || "") : "";
      const desc =
        mapping.descIdx !== -1 ? String(row[mapping.descIdx] || "") : "";
      const balance =
        mapping.balanceIdx !== -1 ? String(row[mapping.balanceIdx] || "") : "";
      let moneyIn = "";
      let moneyOut = "";

      if (/^\d{5}(\.\d+)?$/.test(date)) {
        const serial = parseFloat(date);
        if (serial > 30000 && serial < 60000) {
          const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
          const day = String(d.getUTCDate()).padStart(2, "0");
          const monthNames = [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
          ];
          const month = monthNames[d.getUTCMonth()];
          const year = d.getUTCFullYear();
          date = `${day} ${month} ${year}`;
        }
      }

      if (mapping.isSingleAmountColumn) {
        const valStr = row[mapping.amountIdx];
        if (valStr) {
          const val = parseFloat(String(valStr).replace(/[^0-9.-]/g, ""));
          if (!isNaN(val)) {
            if (val > 0) moneyIn = val.toFixed(2);
            else if (val < 0) moneyOut = Math.abs(val).toFixed(2);
          }
        }
      } else {
        if (mapping.inIdx !== -1) moneyIn = String(row[mapping.inIdx] || "");
        if (mapping.outIdx !== -1) moneyOut = String(row[mapping.outIdx] || "");
      }
      return [date, desc, moneyOut, moneyIn, balance];
    });

    const newData: ExtractedData = {
      headers: stdHeaders,
      rows: transformedRows,
    };
    let finalData: ExtractedData;

    if (state.data) {
      saveToHistory();
      const targetHeaders = state.data.headers;
      const sourceIdx = targetHeaders.indexOf("Source");
      const alignedRows = newData.rows.map((row) => {
        return targetHeaders.map((targetH, i) => {
          if (i === sourceIdx)
            return pendingSourceLabel || pendingSpreadsheet.file.name;
          if (targetH === "Manual Calc") return "";
          const matchIdx = newData.headers.indexOf(targetH);
          if (matchIdx !== -1) return row[matchIdx];
          return "";
        });
      });
      finalData = {
        headers: targetHeaders,
        rows: [...state.data.rows, ...alignedRows],
      };
    } else {
      finalData = normalizeData(newData, pendingSourceLabel || "Imported");
    }
    setState({ step: "RESULTS", data: finalData, error: null });
    setPendingSpreadsheet(null);
  };

  const handleUpdateCategory = (cleanName: string, category: string) => {
    saveToHistory();
    setCategories((prev) => ({ ...prev, [cleanName]: category }));
  };

  const handleBulkUpdateCategory = (cleanNames: string[], category: string) => {
    saveToHistory();
    setCategories((prev) => {
      const next = { ...prev };
      cleanNames.forEach((name) => {
        next[name] = category;
      });
      return next;
    });
  };

  const handleUpdateComment = (cleanName: string, comment: string) => {
    saveToHistory();
    setComments((prev) => ({ ...prev, [cleanName]: comment }));
  };

  const handleBulkUpdateComment = (cleanNames: string[], comment: string) => {
    saveToHistory();
    setComments((prev) => {
      const next = { ...prev };
      cleanNames.forEach((name) => {
        next[name] = comment;
      });
      return next;
    });
  };

  const handleUpdateType = (cleanName: string, type: string) => {
    saveToHistory();
    setTypes((prev) => ({ ...prev, [cleanName]: type }));
  };

  const handleBulkUpdateType = (cleanNames: string[], type: string) => {
    saveToHistory();
    setTypes((prev) => {
      const next = { ...prev };
      cleanNames.forEach((name) => {
        next[name] = type;
      });
      return next;
    });
  };

  const handleMergePayees = (
    targetName: string,
    category: string,
    selection: { index: number; desc: string }[]
  ) => {
    saveToHistory();
    if (category)
      setCategories((prev) => ({ ...prev, [targetName]: category }));
    setMappings((prev) => {
      const next = { ...prev };
      selection.forEach((item) => {
        next[item.desc] = targetName;
      });
      return next;
    });
  };

  const handleUpdateMappings = (
    selection: { index: number; desc: string }[],
    newTargetName: string | null
  ) => {
    saveToHistory();
    setMappings((prev) => {
      const next = { ...prev };
      selection.forEach((item) => {
        if (newTargetName) next[item.desc] = newTargetName;
        else delete next[item.desc];
      });
      return next;
    });
  };

  const handleBulkApplyMappings = (newMap: Record<string, string>) => {
    saveToHistory();
    setMappings((prev) => ({ ...prev, ...newMap }));
    if (state.data && state.data.headers.includes("Entity")) {
      const entityIdx = state.data.headers.indexOf("Entity");
      const descIdx = state.data.headers.findIndex((h) =>
        /desc|details|payee|narrative|memo|transaction|particulars|row\s?label|account|reference/i.test(
          h
        )
      );
      if (entityIdx !== -1 && descIdx !== -1) {
        modifyData((rows) => {
          return rows.map((row) => {
            const newRow = [...row];
            const desc = row[descIdx];
            const entity = newMap[desc] || mappings[desc] || cleanPayee(desc);
            newRow[entityIdx] = entity;
            return newRow;
          });
        }, true);
      }
    }
  };

  const handleUpdateCell = (
    rowIndex: number,
    colIndex: number,
    value: string,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => {
    modifyData((rows, headers) => {
      const newRows = [...rows];
      newRows[rowIndex] = [...newRows[rowIndex]];
      newRows[rowIndex][colIndex] = value;

      const safeManualIdx = headers.indexOf("Manual Calc");
      const safeCreditIdx = headers.findIndex((h) =>
        /credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i.test(h)
      );
      const safeDebitIdx = headers.findIndex((h) =>
        /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
      );

      if (safeManualIdx !== -1) {
        let startCalcRow = colIndex === safeManualIdx ? rowIndex + 1 : rowIndex;
        if (startCalcRow === 0) startCalcRow = 1;

        for (let i = startCalcRow; i < newRows.length; i++) {
          if (colIndex === safeManualIdx && i === rowIndex) continue;
          if (newRows[i] === rows[i]) newRows[i] = [...newRows[i]];

          const prevBalStr = newRows[i - 1][safeManualIdx];
          const creditStr =
            safeCreditIdx !== -1 ? newRows[i][safeCreditIdx] : "0";
          const debitStr = safeDebitIdx !== -1 ? newRows[i][safeDebitIdx] : "0";

          const prevBal = parseFloat(
            prevBalStr?.replace(/[^0-9.-]+/g, "") || "0"
          );
          const credit = parseFloat(
            creditStr?.replace(/[^0-9.-]+/g, "") || "0"
          );
          const debit = parseFloat(debitStr?.replace(/[^0-9.-]+/g, "") || "0");

          if (!isNaN(prevBal)) {
            const newBal =
              prevBal +
              (isNaN(credit) ? 0 : credit) -
              (isNaN(debit) ? 0 : debit);
            newRows[i][safeManualIdx] = newBal.toFixed(2);
          }
        }
      }
      return newRows;
    }, true);
  };

  const handleBulkMoveValues = (
    indices: number[],
    sourceColIdx: number,
    targetColIdx: number,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      indices.forEach((idx) => {
        if (idx < newRows.length) {
          const val = newRows[idx][sourceColIdx];
          if (val && val.trim() !== "") {
            newRows[idx][targetColIdx] = val;
            newRows[idx][sourceColIdx] = "";
          }
        }
      });
      return newRows;
    });
  };

  const handleBulkFixNumbers = (
    indices: number[],
    colIdx: number,
    operation: "NEG_SUFFIX" | "CLEAN" | "FORCE_NEG" | "FORCE_POS" | "INVERT",
    calcConfig?: any
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      indices.forEach((idx) => {
        if (idx < newRows.length) {
          let val = newRows[idx][colIdx] || "";
          if (!val.trim()) return;
          if (operation === "CLEAN") val = val.replace(/[^0-9.-]/g, "");
          else if (operation === "NEG_SUFFIX") {
            const isNegative =
              /\(.*\)|[0-9].*(dr|db|out|d)\b|\b(dr|db|out|d)\b/i.test(val);
            const num = parseFloat(val.replace(/[^0-9.-]/g, ""));
            if (!isNaN(num))
              val = (isNegative ? -Math.abs(num) : num).toFixed(2);
          } else {
            const num = parseFloat(val.replace(/[^0-9.-]/g, ""));
            if (!isNaN(num)) {
              if (operation === "FORCE_NEG") val = (-Math.abs(num)).toFixed(2);
              if (operation === "FORCE_POS") val = Math.abs(num).toFixed(2);
              if (operation === "INVERT") val = (num * -1).toFixed(2);
            }
          }
          newRows[idx][colIdx] = val;
        }
      });
      return newRows;
    });
  };

  const handleBulkSetValue = (
    indices: number[],
    colIdx: number,
    value: string,
    calcConfig?: any
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      indices.forEach((idx) => {
        if (idx < newRows.length) newRows[idx][colIdx] = value;
      });
      return newRows;
    });
  };

  const handleSwapValues = (
    indices: number[],
    colIdx1: number,
    colIdx2: number,
    calcConfig?: any
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      let minRowIndex = newRows.length;
      indices.forEach((idx) => {
        if (idx < newRows.length) {
          const val1 = newRows[idx][colIdx1] || "";
          const val2 = newRows[idx][colIdx2] || "";
          newRows[idx][colIdx1] = val2;
          newRows[idx][colIdx2] = val1;
          if (idx < minRowIndex) minRowIndex = idx;
        }
      });
      if (calcConfig && minRowIndex < newRows.length) {
        const { creditColIdx, debitColIdx, manualColIdx } = calcConfig;
        let startCalcRow = minRowIndex === 0 ? 1 : minRowIndex;
        for (let i = startCalcRow; i < newRows.length; i++) {
          const prevBalStr = newRows[i - 1][manualColIdx];
          const creditStr =
            creditColIdx !== -1 ? newRows[i][creditColIdx] : "0";
          const debitStr = debitColIdx !== -1 ? newRows[i][debitColIdx] : "0";
          const prevBal = parseFloat(
            prevBalStr?.replace(/[^0-9.-]+/g, "") || "0"
          );
          const credit = parseFloat(
            creditStr?.replace(/[^0-9.-]+/g, "") || "0"
          );
          const debit = parseFloat(debitStr?.replace(/[^0-9.-]+/g, "") || "0");
          if (!isNaN(prevBal)) {
            const newBal =
              prevBal +
              (isNaN(credit) ? 0 : credit) -
              (isNaN(debit) ? 0 : debit);
            newRows[i][manualColIdx] = newBal.toFixed(2);
          }
        }
      }
      return newRows;
    });
  };

  const handleSwapRows = (indices: number[]) => {
    if (indices.length !== 2) return;
    modifyData((rows) => {
      const newRows = [...rows];
      const [idx1, idx2] = indices;
      if (idx1 < newRows.length && idx2 < newRows.length)
        [newRows[idx1], newRows[idx2]] = [newRows[idx2], newRows[idx1]];
      return newRows;
    });
  };

  const handleSwapRowContents = (
    indices: number[],
    excludeColIndices: number[]
  ) => {
    if (indices.length !== 2) return;
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      const [idx1, idx2] = indices;
      if (idx1 < newRows.length && idx2 < newRows.length) {
        const row1 = newRows[idx1],
          row2 = newRows[idx2];
        for (let i = 0; i < row1.length; i++)
          if (!excludeColIndices.includes(i))
            [row1[i], row2[i]] = [row2[i], row1[i]];
      }
      return newRows;
    });
  };

  const handleBulkAddYear = (
    indices: number[],
    colIdx: number,
    year: string
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      indices.forEach((idx) => {
        if (idx < newRows.length) {
          const currentCell = newRows[idx][colIdx];
          let val =
            currentCell === undefined || currentCell === null
              ? ""
              : String(currentCell).trim();
          if (val) {
            if (val.includes("/"))
              newRows[idx][colIdx] = `${
                val.endsWith("/") ? val.slice(0, -1) : val
              }/${year}`;
            else if (val.includes("."))
              newRows[idx][colIdx] = `${
                val.endsWith(".") ? val.slice(0, -1) : val
              }.${year}`;
            else if (val.includes("-"))
              newRows[idx][colIdx] = `${
                val.endsWith("-") ? val.slice(0, -1) : val
              }-${year}`;
            else newRows[idx][colIdx] = `${val} ${year}`;
          }
        }
      });
      return newRows;
    });
  };

  const handleShiftColumn = (
    colIndex: number,
    direction: "up" | "down",
    startRowIndex: number = 0
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      const lastIdx = newRows.length - 1;
      const start = Math.max(0, Math.min(startRowIndex, lastIdx));
      if (direction === "up") {
        for (let i = start; i < lastIdx; i++)
          newRows[i][colIndex] = newRows[i + 1][colIndex];
        newRows[lastIdx][colIndex] = "";
      } else {
        for (let i = lastIdx; i > start; i--)
          newRows[i][colIndex] = newRows[i - 1][colIndex];
        newRows[start][colIndex] = "";
      }
      return newRows;
    });
  };

  const handleSortRows = (colIdx: number, direction: "asc" | "desc") => {
    modifyData((rows, headers) => {
      const newRows = [...rows];
      const manualIdx = headers.indexOf("Manual Calc");
      const creditIdx = headers.findIndex((h) =>
        /credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i.test(h)
      );
      const debitIdx = headers.findIndex((h) =>
        /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
      );

      newRows.sort((a, b) => {
        const valA = (a[colIdx] || "").trim();
        const valB = (b[colIdx] || "").trim();
        const dateA = parseBankDate(valA);
        const dateB = parseBankDate(valB);

        let comparison = 0;
        if (dateA && dateB) {
          comparison = dateA.getTime() - dateB.getTime();
        } else {
          const numA = parseFloat(valA.replace(/[^0-9.-]+/g, ""));
          const numB = parseFloat(valB.replace(/[^0-9.-]+/g, ""));
          if (!isNaN(numA) && !isNaN(numB)) comparison = numA - numB;
          else comparison = valA.localeCompare(valB);
        }
        return direction === "asc" ? comparison : -comparison;
      });

      if (manualIdx !== -1) {
        for (let i = 0; i < newRows.length; i++) {
          const prevBal =
            i === 0 ? 0 : parseFloat(newRows[i - 1][manualIdx] || "0");
          const credit =
            creditIdx !== -1
              ? parseFloat(
                  newRows[i][creditIdx]?.replace(/[^0-9.-]+/g, "") || "0"
                )
              : 0;
          const debit =
            debitIdx !== -1
              ? parseFloat(
                  newRows[i][debitIdx]?.replace(/[^0-9.-]+/g, "") || "0"
                )
              : 0;
          newRows[i][manualIdx] = (
            prevBal +
            (isNaN(credit) ? 0 : credit) -
            (isNaN(debit) ? 0 : debit)
          ).toFixed(2);
        }
      }
      return newRows;
    });
  };

  const handleShiftRowValues = (
    indices: number[],
    direction: "left" | "right"
  ) => {
    modifyData((rows) => {
      const newRows = rows.map((r) => [...r]);
      indices.forEach((idx) => {
        if (idx < newRows.length) {
          const row = newRows[idx];
          if (direction === "right") {
            for (let i = row.length - 1; i > 0; i--) row[i] = row[i - 1];
            row[0] = "";
          } else {
            for (let i = 0; i < row.length - 1; i++) row[i] = row[i + 1];
            row[row.length - 1] = "";
          }
        }
      });
      return newRows;
    });
  };

  const handleClearRows = (indicesToClear: number[]) => {
    modifyData((rows) => {
      const indicesSet = new Set(indicesToClear);
      return rows.map((row, idx) =>
        indicesSet.has(idx) ? new Array(row.length).fill("") : row
      );
    });
  };

  const handleInsertRow = (insertIndex: number) => {
    modifyData((rows, headers) => {
      const newRows = [...rows];
      const newRow = new Array(headers.length).fill("");
      if (insertIndex >= 0 && insertIndex <= newRows.length)
        newRows.splice(insertIndex, 0, newRow);
      else newRows.push(newRow);
      return newRows;
    });
  };

  const handleMoveRows = (indices: number[], direction: "up" | "down") => {
    modifyData((rows) => {
      const newRows = [...rows];
      const sortedIndices = [...indices].sort((a, b) => a - b);
      const movedSet = new Set<number>();
      if (direction === "up") {
        for (const idx of sortedIndices) {
          if (
            idx > 0 &&
            (!indices.includes(idx - 1) || movedSet.has(idx - 1))
          ) {
            [newRows[idx], newRows[idx - 1]] = [newRows[idx - 1], newRows[idx]];
            movedSet.add(idx);
          }
        }
      } else {
        for (let i = sortedIndices.length - 1; i >= 0; i--) {
          const idx = sortedIndices[i];
          if (
            idx < newRows.length - 1 &&
            (!indices.includes(idx + 1) || movedSet.has(idx + 1))
          ) {
            [newRows[idx], newRows[idx + 1]] = [newRows[idx + 1], newRows[idx]];
            movedSet.add(idx);
          }
        }
      }
      return newRows;
    });
  };

  const handleReverseRows = (indices: number[]) => {
    modifyData((rows) => {
      const newRows = [...rows];
      const sortedIndices = [...indices].sort((a, b) => a - b);
      const rowsToReverse = sortedIndices.map((idx) => newRows[idx]).reverse();
      sortedIndices.forEach((idx, i) => {
        newRows[idx] = rowsToReverse[i];
      });
      return newRows;
    });
  };

  const handleWaterfallFill = (
    colIndex: number,
    startIndex: number = 0,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => {
    modifyData((rows, headers) => {
      const newRows = [...rows];
      const sourceValue = newRows[startIndex][colIndex];

      for (let i = startIndex + 1; i < newRows.length; i++) {
        if (newRows[i] === rows[i]) newRows[i] = [...newRows[i]];
        newRows[i][colIndex] = sourceValue;
      }

      const manualIdx = headers.indexOf("Manual Calc");
      const creditIdx = headers.findIndex((h) =>
        /credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i.test(h)
      );
      const debitIdx = headers.findIndex((h) =>
        /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
      );

      if (manualIdx !== -1) {
        let startCalcRow = startIndex > 0 ? startIndex : 1;
        for (let i = startCalcRow; i < newRows.length; i++) {
          if (newRows[i] === rows[i]) newRows[i] = [...newRows[i]];
          const prevBalStr = newRows[i - 1][manualIdx];
          const prevBal = parseFloat(
            prevBalStr?.replace(/[^0-9.-]+/g, "") || "0"
          );
          const credit =
            creditIdx !== -1
              ? parseFloat(
                  newRows[i][creditIdx]?.replace(/[^0-9.-]+/g, "") || "0"
                )
              : 0;
          const debit =
            debitIdx !== -1
              ? parseFloat(
                  newRows[i][debitIdx]?.replace(/[^0-9.-]+/g, "") || "0"
                )
              : 0;
          if (!isNaN(prevBal)) {
            newRows[i][manualIdx] = (
              prevBal +
              (isNaN(credit) ? 0 : credit) -
              (isNaN(debit) ? 0 : debit)
            ).toFixed(2);
          }
        }
      }
      return newRows;
    });
  };

  const handleAddMore = () => setState((prev) => ({ ...prev, step: "UPLOAD" }));

  const reset = () => {
    setState({ step: "UPLOAD", data: null, error: null });
    setHistory([]);
    setRedoStack([]);
    setCategories({});
    setComments({});
    setMappings({});
    setTypes({});
    setIsAugmenting(false);
    setAugmentIndices([]);
    setAccumulatedUsage({ prompt: 0, response: 0 });
  };

  const toggleFullScreen = () => {
    if (!document.fullscreenElement)
      document.documentElement.requestFullscreen();
    else if (document.exitFullscreen) document.exitFullscreen();
  };

  return (
    <div className="h-screen flex flex-col bg-slate-50 font-sans relative">
      <header className="bg-white border-b border-slate-200 px-6 py-3 shadow-sm z-10 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-blue-600 rounded-md flex items-center justify-center font-bold text-white shadow-inner">
            ⚡
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-lg font-bold text-slate-900 leading-tight">
                FastScan
              </h1>
              <span className="px-1.5 py-0.5 bg-slate-100 text-slate-500 text-[9px] font-black rounded uppercase border border-slate-200">
                v10-ariana
              </span>
            </div>
            <p className="text-[10px] text-blue-500 font-bold uppercase tracking-widest">
              Sequential OCR Mode
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {processingStatus && (
            <div className="mr-4 px-3 py-1 bg-yellow-100 text-yellow-800 text-xs font-bold rounded-full animate-pulse flex items-center gap-2">
              <span className="w-2 h-2 bg-yellow-600 rounded-full"></span>
              {processingStatus}
            </div>
          )}
          {state.data && state.data.rows.length > 0 && (
            <>
              <div className="mr-2 px-2 py-1 bg-blue-50 text-blue-700 text-[10px] font-black rounded border border-blue-100 uppercase">
                {state.data.rows.length} Rows Cached
              </div>
              <div className="mr-4 px-2 py-1 bg-emerald-50 text-emerald-700 text-[10px] font-black rounded border border-emerald-100 uppercase">
                Est. Cost: ${totalCost.toFixed(4)}
              </div>
            </>
          )}
          <span className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
          <span className="text-[10px] font-mono text-slate-400">
            GEMINI-FLASH-LITE
          </span>
          <div className="w-px h-6 bg-slate-200 mx-2"></div>
          <button
            onClick={toggleFullScreen}
            className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-all"
            title="Toggle Full Screen"
          >
            <svg
              width="18"
              height="18"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3" />
            </svg>
          </button>
        </div>
      </header>

      <main
        className={`flex-1 relative ${
          state.step === "RESULTS"
            ? "overflow-hidden flex flex-col"
            : "overflow-auto"
        }`}
      >
        {state.error && (
          <div className="fixed top-16 left-1/2 -translate-x-1/2 bg-red-600 text-white px-4 py-2 rounded-full shadow-2xl z-50 flex items-center gap-2 text-sm animate-bounce">
            <span className="font-bold">{state.error}</span>
            <button
              onClick={() => setState((s) => ({ ...s, error: null }))}
              className="hover:opacity-70"
            >
              ×
            </button>
          </div>
        )}

        {state.step === "UPLOAD" && (
          <div className="h-full flex flex-col items-center justify-center px-4">
            <div className="w-full max-w-lg text-center mb-8">
              {isAugmenting ? (
                <>
                  <div className="inline-block px-3 py-1 bg-amber-100 text-amber-800 rounded-full text-xs font-bold mb-4 uppercase">
                    Augmenting {augmentIndices.length} Rows
                  </div>
                  <h2 className="text-3xl font-black text-slate-900 mb-2">
                    Re-Scan Source File
                  </h2>
                  <p className="text-slate-500 font-medium">
                    Upload original document for selected rows.
                  </p>
                </>
              ) : (
                <>
                  <h2 className="text-3xl font-black text-slate-900 mb-2">
                    {state.data ? "Import Data." : "Instant CSV Export."}
                  </h2>
                  <p className="text-slate-500 font-medium">
                    {state.data
                      ? "Drop next page or new statement."
                      : "Drop statement or project."}
                  </p>
                </>
              )}
            </div>
            <FileUploader
              onFileSelect={handleFileSelect}
              onBack={
                state.data
                  ? () => {
                      setState((prev) => ({ ...prev, step: "RESULTS" }));
                      setIsAugmenting(false);
                      setAugmentIndices([]);
                    }
                  : () => {}
              }
              isLoading={false}
            />
          </div>
        )}

        {state.step === "PROCESSING" && (
          <div className="h-full flex flex-col items-center justify-center">
            <div className="relative">
              <div className="w-24 h-24 border-[6px] border-slate-100 border-t-blue-600 rounded-full animate-spin"></div>
              <div className="absolute inset-0 flex items-center justify-center text-blue-600 font-black">
                ⚡
              </div>
            </div>
            <h3 className="mt-8 text-xl font-black text-slate-900 uppercase tracking-tighter">
              {processingStatus || "Processing..."}
            </h3>
          </div>
        )}

        {state.step === "RESULTS" && state.data && (
          <>
            <div className="bg-white border-b border-slate-200 px-4 py-2 flex justify-center shrink-0 z-30 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
              <div className="flex bg-slate-100 p-1 rounded-lg">
                <button
                  onClick={() => setViewMode("TABLE")}
                  className={`px-4 py-1.5 text-xs font-bold rounded-md transition-all ${
                    viewMode === "TABLE"
                      ? "bg-white text-slate-800 shadow-sm"
                      : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  Grid Editor
                </button>
                <button
                  onClick={() => setViewMode("PIVOT")}
                  className={`px-4 py-1.5 text-xs font-bold rounded-md transition-all ${
                    viewMode === "PIVOT"
                      ? "bg-white text-slate-800 shadow-sm"
                      : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  Pivot Analysis
                </button>
              </div>
            </div>
            <div className="flex-1 relative overflow-hidden">
              {viewMode === "TABLE" ? (
                <ResultsTable
                  data={state.data}
                  onReset={reset}
                  onAddMore={handleAddMore}
                  onAugment={handleAugmentTrigger}
                  onUpdateCell={handleUpdateCell}
                  onClearRows={handleClearRows}
                  onInsertRow={handleInsertRow}
                  onMoveRows={handleMoveRows}
                  onReverseRows={handleReverseRows}
                  onBulkMoveValues={handleBulkMoveValues}
                  onBulkAddYear={handleBulkAddYear}
                  onBulkFixNumbers={handleBulkFixNumbers}
                  onBulkSetValue={handleBulkSetValue}
                  onShiftColumn={handleShiftColumn}
                  onSortRows={handleSortRows}
                  onShiftRowValues={handleShiftRowValues}
                  onSwapValues={handleSwapValues}
                  onSwapRows={handleSwapRows}
                  onSwapRowContents={handleSwapRowContents}
                  onWaterfallFill={handleWaterfallFill}
                  onUndo={handleUndo}
                  onRedo={handleRedo}
                  onSave={handleSaveProject}
                  canUndo={history.length > 0}
                  canRedo={redoStack.length > 0}
                />
              ) : (
                <PivotView
                  data={state.data}
                  categories={categories}
                  comments={comments}
                  mappings={mappings}
                  types={types}
                  onUpdateCategory={handleUpdateCategory}
                  onUpdateComment={handleUpdateComment}
                  onUpdateType={handleUpdateType}
                  onMergePayees={handleMergePayees}
                  onUpdateMappings={handleUpdateMappings}
                  onBulkUpdateCategory={handleBulkUpdateCategory}
                  onBulkUpdateComment={handleBulkUpdateComment}
                  onBulkUpdateType={handleBulkUpdateType}
                  onBulkApplyMappings={handleBulkApplyMappings}
                  isEntitySynced={state.data.headers.includes("Entity")}
                  isCategorySynced={state.data.headers.includes("Category")}
                  isTypeSynced={state.data.headers.includes("Type")}
                  isCommentSynced={state.data.headers.includes("Comment")}
                  onToggleEntitySync={handleToggleEntityColumn}
                  onToggleCategorySync={handleToggleCategoryColumn}
                  onToggleTypeSync={handleToggleTypeColumn}
                  onToggleCommentSync={handleToggleCommentColumn}
                  onUndo={handleUndo}
                  onRedo={handleRedo}
                  canUndo={history.length > 0}
                  canRedo={redoStack.length > 0}
                />
              )}
            </div>
          </>
        )}
      </main>

      {showChunker && pendingFile && !showSourceSelector && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6">
            <h3 className="text-xl font-black text-slate-900 mb-2">
              Large PDF Detected
            </h3>
            <p className="text-sm text-slate-500 mb-4">
              {pendingFile.name} ({totalPages} Pages)
            </p>
            <div className="mb-6">
              <label className="block text-sm font-bold text-slate-700 mb-2">
                Pages per batch?
              </label>
              <input
                type="number"
                min="1"
                max="50"
                value={chunkSize}
                onChange={(e) =>
                  setChunkSize(Math.max(1, parseInt(e.target.value) || 1))
                }
                className="w-24 p-2 border border-slate-300 rounded font-bold text-center outline-none"
              />
            </div>
            <div className="flex justify-end gap-3">
              <button
                onClick={() => {
                  setShowChunker(false);
                  setPendingFile(null);
                }}
                className="px-4 py-2 text-slate-500 font-bold text-sm"
              >
                Cancel
              </button>
              <button
                onClick={confirmChunkProcessing}
                className="px-6 py-2 bg-blue-600 text-white font-bold text-sm rounded"
              >
                Start Batch Scan
              </button>
            </div>
          </div>
        </div>
      )}

      {showSourceSelector && (pendingFile || pendingSpreadsheet) && (
        <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6">
            <h3 className="text-xl font-black text-slate-900 mb-4">
              Import Data
            </h3>
            <div className="mb-6 space-y-4">
              {state.data && state.data.headers.includes("Source") && (
                <label className="flex items-center gap-2 cursor-pointer p-3 border rounded-lg hover:bg-slate-50">
                  <input
                    type="radio"
                    name="source"
                    value={String(
                      state.data.rows[state.data.rows.length - 1][
                        state.data.headers.indexOf("Source")
                      ] || ""
                    )}
                    checked={targetSource !== "NEW_SOURCE"}
                    onChange={(e) => setTargetSource(e.target.value)}
                  />
                  <div className="text-sm font-bold">
                    Append to Existing ({targetSource})
                  </div>
                </label>
              )}
              <label className="flex items-center gap-2 cursor-pointer p-3 border rounded-lg hover:bg-slate-50">
                <input
                  type="radio"
                  name="source"
                  value="NEW_SOURCE"
                  checked={targetSource === "NEW_SOURCE"}
                  onChange={() => setTargetSource("NEW_SOURCE")}
                />
                <div className="text-sm font-bold">Create New Statement</div>
              </label>
              {targetSource === "NEW_SOURCE" && (
                <input
                  type="text"
                  autoFocus
                  placeholder="Enter Name"
                  value={newSourceName}
                  onChange={(e) => setNewSourceName(e.target.value)}
                  className="w-full p-2 border rounded font-bold outline-none"
                />
              )}
            </div>
            <div className="flex justify-end gap-3">
              <button
                onClick={() => {
                  setShowSourceSelector(false);
                  setPendingFile(null);
                  setPendingSpreadsheet(null);
                }}
                className="px-4 py-2 text-slate-500 font-bold text-sm"
              >
                Cancel
              </button>
              <button
                onClick={confirmSourceSelection}
                className="px-6 py-2 bg-blue-600 text-white font-bold text-sm rounded"
              >
                Confirm Import
              </button>
            </div>
          </div>
        </div>
      )}

      {showMapper && pendingSpreadsheet && (
        <ColumnMapper
          headers={pendingSpreadsheet.headers}
          filename={pendingSpreadsheet.file.name}
          onConfirm={handleMapperConfirm}
          onCancel={() => {
            setShowMapper(false);
            setPendingSpreadsheet(null);
          }}
        />
      )}
    </div>
  );
}
