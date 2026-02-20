import React, {
  useState,
  useEffect,
  useRef,
  useMemo,
  useCallback,
} from "react";
import { ExtractedData } from "../types";
import { TableVirtuoso } from "react-virtuoso";

interface Props {
  data: ExtractedData;
  onReset: () => void;
  onAddMore: () => void;
  onAugment: (indices: number[]) => void;
  onUpdateCell: (
    rowIdx: number,
    colIdx: number,
    value: string,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onClearRows: (indices: number[]) => void;
  onInsertRow: (index: number) => void;
  onMoveRows: (indices: number[], direction: "up" | "down") => void;
  onReverseRows: (indices: number[]) => void;
  onBulkMoveValues: (
    indices: number[],
    sourceColIdx: number,
    targetColIdx: number,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onBulkAddYear: (indices: number[], colIdx: number, year: string) => void;
  onBulkFixNumbers: (
    indices: number[],
    colIdx: number,
    operation: "NEG_SUFFIX" | "CLEAN" | "FORCE_NEG" | "FORCE_POS" | "INVERT",
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onBulkSetValue: (
    indices: number[],
    colIdx: number,
    value: string,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onShiftColumn: (
    colIdx: number,
    direction: "up" | "down",
    startRowIndex: number
  ) => void;
  onSortRows: (colIdx: number, direction: "asc" | "desc") => void;
  onShiftRowValues: (indices: number[], direction: "left" | "right") => void;
  onSwapValues: (
    indices: number[],
    colIdx1: number,
    colIdx2: number,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onSwapRows: (indices: number[]) => void;
  onSwapRowContents: (indices: number[], excludeColIndices: number[]) => void;
  onWaterfallFill: (
    colIdx: number,
    startRowIdx: number,
    calcConfig?: {
      creditColIdx: number;
      debitColIdx: number;
      manualColIdx: number;
    }
  ) => void;
  onUndo: () => void;
  onRedo: () => void;
  onSave: () => void;
  canUndo: boolean;
  canRedo: boolean;
}

type ResizeState = {
  index: number | "diff";
  startX: number;
  startWidth: number;
};

// --- ROW CELLS (Content Only) ---
const RowCells = ({
  row,
  rowIdx,
  isSelected,
  manualColIndex,
  creditColIndex,
  debitColIndex,
  balanceColIndex,
  onToggleSelection,
  onCellChange,
  onFocusCell,
  dragState,
}: any) => {
  const manualVal = parseFloat(
    String(row[manualColIndex] || "").replace(/[^0-9.-]+/g, "")
  );
  const balanceCellStr =
    balanceColIndex !== -1 ? row[balanceColIndex] || "" : "";
  const isBalancePresent = balanceCellStr && balanceCellStr.trim() !== "";

  let diffFormatted = "";
  let isDiffZero = false;

  if (balanceColIndex !== -1 && isBalancePresent) {
    const balanceVal = parseFloat(balanceCellStr.replace(/[^0-9.-]+/g, ""));
    const diff = isNaN(manualVal) ? 0 - balanceVal : manualVal - balanceVal;
    diffFormatted = diff.toFixed(2);
    isDiffZero = Math.abs(diff) < 0.01;
  }

  const handleCellDragStart = (e: React.MouseEvent, c: number) => {
    e.preventDefault();
    e.stopPropagation();
    dragState.onDragStart(rowIdx, c);
  };

  const isCellInDrag = (c: number) => {
    if (!dragState.isDragging || !dragState.start || !dragState.end)
      return false;
    if (c !== dragState.start.c) return false;
    const minR = Math.min(dragState.start.r, dragState.end.r);
    const maxR = Math.max(dragState.start.r, dragState.end.r);
    return rowIdx >= minR && rowIdx <= maxR && rowIdx !== dragState.start.r;
  };

  return (
    <>
      <td
        className={`p-3 text-center border-r border-slate-100 cursor-pointer select-none ${
          isSelected
            ? "text-blue-600"
            : "text-slate-300 group-hover:text-slate-500"
        }`}
        onClick={(e) => onToggleSelection(rowIdx, e)}
      >
        {isSelected ? (
          <div className="w-4 h-4 bg-blue-600 rounded mx-auto"></div>
        ) : (
          <div className="w-4 h-4 border border-slate-300 rounded mx-auto bg-white"></div>
        )}
      </td>
      {row.map((cell: string, colIdx: number) => (
        <td
          key={colIdx}
          className={`p-0 border-r border-slate-100 relative 
                        ${colIdx === manualColIndex ? "bg-indigo-50/30" : ""}
                        ${isCellInDrag(colIdx) ? "bg-blue-50" : ""}
                    `}
          onMouseEnter={() => dragState.onMouseEnter(rowIdx, colIdx)}
        >
          <input
            type="text"
            value={cell || ""}
            onFocus={() => onFocusCell(rowIdx, colIdx)}
            onChange={(e) => onCellChange(rowIdx, colIdx, e.target.value)}
            className={`w-full h-full p-3 bg-transparent outline-none text-sm text-slate-700 font-medium focus:bg-white focus:ring-2 focus:ring-blue-500 focus:z-10 absolute inset-0 
                            ${
                              colIdx === manualColIndex
                                ? "text-indigo-700 font-bold font-mono text-right"
                                : ""
                            }
                            ${
                              colIdx === creditColIndex
                                ? "text-emerald-700"
                                : ""
                            }
                            ${colIdx === debitColIndex ? "text-red-700" : ""}
                        `}
          />

          {dragState.focusedCell?.r === rowIdx &&
            dragState.focusedCell?.c === colIdx && (
              <div
                className="absolute bottom-0 right-0 w-2.5 h-2.5 bg-blue-600 border border-white cursor-crosshair z-20"
                onMouseDown={(e) => handleCellDragStart(e, colIdx)}
                onDoubleClick={(e) =>
                  dragState.onHandleDoubleClick(e, rowIdx, colIdx)
                }
                title="Drag to fill down"
              />
            )}

          {isCellInDrag(colIdx) && (
            <div className="absolute inset-0 border-2 border-blue-500 pointer-events-none z-20"></div>
          )}
        </td>
      ))}
      <td className="p-3 text-right font-mono font-bold text-sm bg-slate-50/50">
        {balanceColIndex !== -1 && isBalancePresent && (
          <span
            className={`px-2 py-1 rounded text-xs ${
              isDiffZero
                ? "bg-green-100 text-green-700"
                : "bg-red-100 text-red-700"
            }`}
          >
            {diffFormatted}
          </span>
        )}
      </td>
    </>
  );
};

export const ResultsTable: React.FC<Props> = ({
  data,
  onReset,
  onAddMore,
  onAugment,
  onUpdateCell,
  onClearRows,
  onInsertRow,
  onMoveRows,
  onReverseRows,
  onBulkMoveValues,
  onBulkAddYear,
  onBulkFixNumbers,
  onBulkSetValue,
  onShiftColumn,
  onSortRows,
  onShiftRowValues,
  onSwapValues,
  onSwapRows,
  onSwapRowContents,
  onWaterfallFill,
  onUndo,
  onRedo,
  onSave,
  canUndo,
  canRedo,
}) => {
  const manualColIndex = data.headers.indexOf("Manual Calc");
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const selectAllRef = useRef<HTMLInputElement>(null);

  const [dateColIndex, setDateColIndex] = useState<number>(() =>
    data.headers.findIndex((h) => /date/i.test(h))
  );
  const [creditColIndex, setCreditColIndex] = useState<number>(() =>
    data.headers.findIndex((h) =>
      /credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i.test(h)
    )
  );
  const [debitColIndex, setDebitColIndex] = useState<number>(() =>
    data.headers.findIndex((h) =>
      /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
    )
  );
  const [balanceColIndex, setBalanceColIndex] = useState<number>(() =>
    data.headers.findIndex((h) => /balance/i.test(h) && h !== "Manual Calc")
  );

  const [sourceFilter, setSourceFilter] = useState<string>("ALL");
  const sourceColIndex = data.headers.indexOf("Source");

  const [sortState, setSortState] = useState<{
    colIdx: number;
    direction: "asc" | "desc";
  } | null>(null);

  const uniqueSources = useMemo(() => {
    if (sourceColIndex === -1) return [];
    const sources = new Set<string>();
    data.rows.forEach((r) => {
      if (r[sourceColIndex]) sources.add(r[sourceColIndex]);
    });
    return Array.from(sources).sort();
  }, [data.rows, sourceColIndex]);

  const isRowVisible = useCallback(
    (rowIndex: number) => {
      if (sourceFilter === "ALL") return true;
      if (sourceColIndex === -1) return true;
      return data.rows[rowIndex][sourceColIndex] === sourceFilter;
    },
    [data.rows, sourceFilter, sourceColIndex]
  );

  const visibleRowIndices = useMemo(() => {
    const indices: number[] = [];
    data.rows.forEach((_, i) => {
      if (isRowVisible(i)) indices.push(i);
    });
    return indices;
  }, [data.rows, isRowVisible]);

  useEffect(() => {
    setDateColIndex(data.headers.findIndex((h) => /date/i.test(h)));
    setCreditColIndex(
      data.headers.findIndex((h) =>
        /credit|money\s?in|deposit|receipt|paid\s?in|\bin\b/i.test(h)
      )
    );
    setDebitColIndex(
      data.headers.findIndex((h) =>
        /debit|money\s?out|withdrawal|payment|paid\s?out|\bout\b/i.test(h)
      )
    );
    setBalanceColIndex(
      data.headers.findIndex((h) => /balance/i.test(h) && h !== "Manual Calc")
    );
  }, [data.headers]);

  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [lastSelectedIndex, setLastSelectedIndex] = useState<number>(-1);
  const [confirmReset, setConfirmReset] = useState(false);

  const stateRef = useRef({
    selectedRows,
    lastSelectedIndex,
    visibleRowIndices,
  });
  useEffect(() => {
    stateRef.current = { selectedRows, lastSelectedIndex, visibleRowIndices };
  }, [selectedRows, lastSelectedIndex, visibleRowIndices]);

  const [moveSourceCol, setMoveSourceCol] = useState(0);
  const [moveTargetCol, setMoveTargetCol] = useState(0);
  const [bulkYear, setBulkYear] = useState(new Date().getFullYear().toString());
  const [fixColIndex, setFixColIndex] = useState(0);
  const [fixOperation, setFixOperation] = useState<
    "NEG_SUFFIX" | "CLEAN" | "FORCE_NEG" | "FORCE_POS" | "INVERT"
  >("NEG_SUFFIX");
  const [bulkSetCol, setBulkSetCol] = useState(0);
  const [bulkSetValue, setBulkSetValue] = useState("");

  const [focusedCell, setFocusedCell] = useState<{
    r: number;
    c: number;
  } | null>(null);
  const [dragStart, setDragStart] = useState<{ r: number; c: number } | null>(
    null
  );
  const [dragEnd, setDragEnd] = useState<{ r: number; c: number } | null>(null);
  const [isDraggingHandle, setIsDraggingHandle] = useState(false);
  const [autofillMenu, setAutofillMenu] = useState<{
    r: number;
    c: number;
    x: number;
    y: number;
  } | null>(null);

  const isAllVisibleSelected =
    visibleRowIndices.length > 0 &&
    visibleRowIndices.every((i) => selectedRows.has(i));
  const isIndeterminate = selectedRows.size > 0 && !isAllVisibleSelected;

  const [colWidths, setColWidths] = useState<number[]>([]);
  const [diffColWidth, setDiffColWidth] = useState(130);
  const resizingRef = useRef<ResizeState | null>(null);

  const CHAR_PX = 8.5;
  const PADDING_PX = 30;

  useEffect(() => {
    if (data.headers.length > 0) {
      setColWidths((prev) => {
        if (prev.length === data.headers.length) return prev;
        return data.headers.map((h, i) => {
          const lower = h.toLowerCase();
          let chars = 15;
          if (/date/.test(lower)) chars = 15;
          else if (
            /desc|details|payee|narrative|memo|transaction|particulars/.test(
              lower
            )
          )
            chars = 50;
          else if (/type/.test(lower)) chars = 15;
          else if (/out|debit|payment|withdrawal/.test(lower)) chars = 15;
          else if (/in|credit|receipt|deposit/.test(lower)) chars = 15;
          else if (/manual/.test(lower)) chars = 20;
          else if (/balance/.test(lower)) chars = 20;
          else if (/entity|category|comment/.test(lower)) chars = 30;
          return Math.floor(chars * CHAR_PX + PADDING_PX);
        });
      });
      if (colWidths.length === 0)
        setDiffColWidth(Math.floor(15 * CHAR_PX + PADDING_PX));
    }
  }, [data.headers]);

  const startResize = (index: number | "diff", e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    let startW = 100;
    if (index === "diff") startW = diffColWidth;
    else if (typeof index === "number") startW = colWidths[index] || 100;
    resizingRef.current = { index, startX: e.pageX, startWidth: startW };
    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
    document.body.style.cursor = "col-resize";
  };

  const handleMouseMove = useCallback((e: globalThis.MouseEvent) => {
    if (!resizingRef.current) return;
    const { index, startX, startWidth } = resizingRef.current;
    const newWidth = Math.max(30, startWidth + (e.pageX - startX));
    if (index === "diff") setDiffColWidth(newWidth);
    else if (typeof index === "number") {
      setColWidths((prev) => {
        const next = [...prev];
        next[index] = newWidth;
        return next;
      });
    }
  }, []);

  const handleMouseUp = useCallback(() => {
    resizingRef.current = null;
    document.removeEventListener("mousemove", handleMouseMove);
    document.removeEventListener("mouseup", handleMouseUp);
    document.body.style.cursor = "";
  }, [handleMouseMove]);

  useEffect(() => {
    if (confirmReset) {
      const timer = setTimeout(() => setConfirmReset(false), 3000);
      return () => clearTimeout(timer);
    }
  }, [confirmReset]);

  useEffect(() => {
    if (selectAllRef.current)
      selectAllRef.current.indeterminate = isIndeterminate;
  }, [isIndeterminate]);

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
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [canUndo, canRedo, onUndo, onRedo]);

  const getCalcConfig = useCallback(
    () =>
      manualColIndex !== -1
        ? {
            creditColIdx: creditColIndex,
            debitColIdx: debitColIndex,
            manualColIdx: manualColIndex,
          }
        : undefined,
    [manualColIndex, creditColIndex, debitColIndex]
  );

  useEffect(() => {
    const handleGlobalMouseUp = () => {
      if (isDraggingHandle && dragStart && dragEnd) {
        if (dragStart.c === dragEnd.c && dragStart.r !== dragEnd.r) {
          const col = dragStart.c;
          const startRow = Math.min(dragStart.r, dragEnd.r);
          const endRow = Math.max(dragStart.r, dragEnd.r);
          const sourceValue = data.rows[dragStart.r][col];
          const indicesToUpdate: number[] = [];
          for (let i = startRow; i <= endRow; i++) {
            if (i !== dragStart.r && isRowVisible(i)) indicesToUpdate.push(i);
          }
          if (indicesToUpdate.length > 0)
            onBulkSetValue(indicesToUpdate, col, sourceValue, getCalcConfig());
        }
        setIsDraggingHandle(false);
        setDragStart(null);
        setDragEnd(null);
      } else if (isDraggingHandle) {
        setIsDraggingHandle(false);
        setDragStart(null);
        setDragEnd(null);
      }
    };
    window.addEventListener("mouseup", handleGlobalMouseUp);
    return () => window.removeEventListener("mouseup", handleGlobalMouseUp);
  }, [
    isDraggingHandle,
    dragStart,
    dragEnd,
    data.rows,
    onBulkSetValue,
    getCalcConfig,
    isRowVisible,
  ]);

  const handleHandleDoubleClick = (
    e: React.MouseEvent,
    r: number,
    c: number
  ) => {
    e.preventDefault();
    e.stopPropagation();
    const cellValue = data.rows[r][c];
    if (!cellValue || cellValue.trim() === "") return;
    setAutofillMenu({ r, c, x: e.clientX, y: e.clientY });
  };

  const handleAutofillAction = (type: "NEXT_VALUE" | "COMPLETE") => {
    if (!autofillMenu) return;
    const { r, c } = autofillMenu;
    if (type === "NEXT_VALUE") {
      const sourceValue = data.rows[r][c];
      const indicesToUpdate: number[] = [];
      for (let i = r + 1; i < data.rows.length; i++) {
        if (!isRowVisible(i)) continue;
        const val = data.rows[i][c];
        if (val && val.trim() !== "") break;
        indicesToUpdate.push(i);
      }
      if (indicesToUpdate.length > 0)
        onBulkSetValue(indicesToUpdate, c, sourceValue, getCalcConfig());
    } else {
      onWaterfallFill(c, r, getCalcConfig());
    }
    setAutofillMenu(null);
  };

  const handleCellChange = useCallback(
    (rowIndex: number, colIndex: number, value: string) => {
      onUpdateCell(rowIndex, colIndex, value, getCalcConfig());
    },
    [onUpdateCell, getCalcConfig]
  );

  const toggleRowSelection = useCallback(
    (index: number, e: React.MouseEvent) => {
      const { selectedRows, lastSelectedIndex, visibleRowIndices } =
        stateRef.current;
      const multiSelect = e.ctrlKey || e.metaKey;
      const shiftSelect = e.shiftKey;

      if (shiftSelect && lastSelectedIndex !== -1) {
        const lastIdxPos = visibleRowIndices.indexOf(lastSelectedIndex);
        const currIdxPos = visibleRowIndices.indexOf(index);
        if (lastIdxPos !== -1 && currIdxPos !== -1) {
          const start = Math.min(lastIdxPos, currIdxPos);
          const end = Math.max(lastIdxPos, currIdxPos);
          const newSelected = new Set(selectedRows);
          for (let i = start; i <= end; i++) {
            newSelected.add(visibleRowIndices[i]);
          }
          setSelectedRows(newSelected);
        }
      } else {
        if (
          !multiSelect &&
          selectedRows.has(index) &&
          selectedRows.size === 1
        ) {
          setSelectedRows(new Set());
          setLastSelectedIndex(-1);
          return;
        }
        const newSelected = new Set(multiSelect ? selectedRows : []);
        if (newSelected.has(index)) newSelected.delete(index);
        else newSelected.add(index);
        setSelectedRows(newSelected);
        setLastSelectedIndex(index);
      }
    },
    []
  );

  const handleSelectAll = () => {
    if (isAllVisibleSelected) setSelectedRows(new Set());
    else setSelectedRows(new Set(visibleRowIndices));
  };

  const handleManualCheck = () => {
    const cleanNum = (str: string) => {
      if (!str) return 0;
      const cleaned = str.replace(/[^0-9.-]+/g, "");
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    };

    let totalIn = 0;
    let totalOut = 0;
    let startBal: number | null = null;
    let endBal: number | null = null;
    let startBalIndex = -1;

    data.rows.forEach((row, idx) => {
      const inVal = creditColIndex !== -1 ? cleanNum(row[creditColIndex]) : 0;
      const outVal = debitColIndex !== -1 ? cleanNum(row[debitColIndex]) : 0;

      totalIn += inVal;
      totalOut += outVal;

      if (balanceColIndex !== -1) {
        const balStr = row[balanceColIndex];
        if (balStr && balStr.trim() !== "") {
          const val = cleanNum(balStr);
          if (startBal === null) {
            startBal = val;
            startBalIndex = idx;
          }
          endBal = val;
        }
      }
    });

    if (startBal === null || endBal === null) {
      alert(
        "⚠️ Reconciliation Error: Could not isolate a stated Opening and Closing balance in the extracted text."
      );
      return;
    }

    const startRowIn =
      creditColIndex !== -1
        ? cleanNum(data.rows[startBalIndex][creditColIndex])
        : 0;
    const startRowOut =
      debitColIndex !== -1
        ? cleanNum(data.rows[startBalIndex][debitColIndex])
        : 0;

    const adjTotalIn = totalIn - startRowIn;
    const adjTotalOut = totalOut - startRowOut;

    const calculatedEnd = startBal + adjTotalIn - adjTotalOut;
    const difference = calculatedEnd - endBal;

    if (Math.abs(difference) < 0.02) {
      alert(
        `✅ MATCH CONFIRMED\n\nStated Balance: ${endBal.toLocaleString(
          undefined,
          { minimumFractionDigits: 2 }
        )}\nCalculated: ${calculatedEnd.toLocaleString(undefined, {
          minimumFractionDigits: 2,
        })}\n\nLogic: Local extraction & sum successful.`
      );
    } else {
      alert(
        `❌ DISCREPANCY DETECTED\n\nStated Balance: ${endBal.toLocaleString(
          undefined,
          { minimumFractionDigits: 2 }
        )}\nCalculated: ${calculatedEnd.toLocaleString(undefined, {
          minimumFractionDigits: 2,
        })}\nDiff: ${difference.toLocaleString(undefined, {
          minimumFractionDigits: 2,
        })}\n\nNote: Checked ${data.rows.length} extracted lines.`
      );
    }
  };

  const handleDownloadCSV = () => {
    const cleanNumber = (str: string) => {
      const c = str.replace(/[^0-9.-]+/g, "");
      const n = parseFloat(c);
      return isNaN(n) ? 0 : n;
    };
    const headers = [...data.headers, "Difference"];
    const csvRows = data.rows.map((row: string[]) => {
      const manualVal = cleanNumber(row[manualColIndex] || "");
      const balanceCellStr =
        balanceColIndex !== -1 ? row[balanceColIndex] || "" : "";
      const isBalancePresent = balanceCellStr && balanceCellStr.trim() !== "";
      let diff = "";
      if (balanceColIndex !== -1 && isBalancePresent) {
        const balanceVal = cleanNumber(balanceCellStr);
        diff = (manualVal - balanceVal).toFixed(2);
      }
      const escapedRow = row.map(
        (cell) => `"${(cell || "").replace(/"/g, '""')}"`
      );
      escapedRow.push(`"${diff}"`);
      return escapedRow.join(",");
    });
    const csvContent = [headers.join(","), ...csvRows].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute(
      "download",
      `reconciliation_${new Date().toISOString().slice(0, 10)}.csv`
    );
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleSmartInsertRow = () => {
    let insertIndex = data.rows.length;
    if (selectedRows.size > 0)
      insertIndex = Math.max(...(Array.from(selectedRows) as number[])) + 1;
    onInsertRow(insertIndex);
  };

  const handleLocalMoveRows = (direction: "up" | "down") => {
    const selectedIndices = (Array.from(selectedRows) as number[]).sort(
      (a, b) => a - b
    );
    if (selectedIndices.length === 0) return;

    onMoveRows(selectedIndices, direction);

    const movedMap = new Map<number, number>();

    if (direction === "up") {
      for (const idx of selectedIndices) {
        if (idx > 0 && !selectedRows.has(idx - 1)) {
          movedMap.set(idx, idx - 1);
        } else {
          let canMove = true;
          if (idx === 0) canMove = false;
          if (canMove) {
            const isBlockedAtTop = selectedIndices[0] === 0;
            if (!isBlockedAtTop) {
              movedMap.set(idx, idx - 1);
            }
          }
        }
      }
    } else {
      for (let i = selectedIndices.length - 1; i >= 0; i--) {
        const idx = selectedIndices[i];
        if (idx < data.rows.length - 1) {
          const isBlockedAtBottom =
            selectedIndices[selectedIndices.length - 1] ===
            data.rows.length - 1;
          if (!isBlockedAtBottom) {
            movedMap.set(idx, idx + 1);
          }
        }
      }
    }

    if (movedMap.size > 0) {
      const nextSelection = new Set<number>();
      selectedRows.forEach((idx) => {
        if (movedMap.has(idx)) nextSelection.add(movedMap.get(idx)!);
        else nextSelection.add(idx);
      });
      setSelectedRows(nextSelection);
    }
  };

  const handleHeaderSort = (colIdx: number) => {
    const nextDir =
      sortState?.colIdx === colIdx && sortState.direction === "asc"
        ? "desc"
        : "asc";
    onSortRows(colIdx, nextDir);
    setSortState({ colIdx, direction: nextDir });
  };

  const selectedIndices = (Array.from(selectedRows) as number[]).sort(
    (a, b) => a - b
  );
  const totalTableWidth =
    colWidths.reduce((sum, w) => sum + w, 0) + diffColWidth + 48;
  const totalVisibleCount = visibleRowIndices.length;
  const isFiltered = sourceFilter !== "ALL";

  const dragState = {
    isDragging: isDraggingHandle,
    start: dragStart,
    end: dragEnd,
    focusedCell: focusedCell,
    onDragStart: (r: number, c: number) => {
      setDragStart({ r, c });
      setDragEnd({ r, c });
      setIsDraggingHandle(true);
    },
    onMouseEnter: (r: number, c: number) => {
      if (isDraggingHandle && dragStart && c === dragStart.c)
        setDragEnd({ r, c });
    },
    onHandleDoubleClick: handleHandleDoubleClick,
  };

  const virtuosoComponents = useMemo(
    () => ({
      Scroller: React.forwardRef<HTMLDivElement, any>((props, ref) => (
        <div {...props} ref={ref} className="overflow-auto custom-scrollbar" />
      )),
      Table: ({ context, ...props }: any) => (
        <table
          {...props}
          className="border-collapse text-sm table-fixed bg-white shadow-sm rounded-lg"
          style={{
            ...props.style,
            width: Math.max(100, context?.totalTableWidth || 100),
          }}
        />
      ),
      TableRow: ({ item: rowIdx, context, ...props }: any) => {
        const isSelected = context?.selectedRows?.has(rowIdx);
        return (
          <tr
            {...props}
            className={`group hover:bg-slate-50 transition-colors ${
              isSelected ? "bg-blue-50" : ""
            }`}
          />
        );
      },
    }),
    []
  );

  return (
    <div className="flex flex-col h-full bg-slate-50 font-sans relative">
      {/* --- Toolbar Header --- */}
      <div className="bg-white px-6 py-4 border-b border-slate-200 flex items-center justify-between shadow-sm z-20">
        <div>
          <h1 className="text-xl font-black text-slate-800 tracking-tight uppercase">
            Reconciliation Mode
          </h1>
          <div className="flex items-center gap-2 mt-1">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-widest">
              {isFiltered
                ? `${totalVisibleCount} / ${data.rows.length}`
                : data.rows.length}{" "}
              TRANSACTIONS
            </span>
            <span className="text-slate-300">•</span>
            <span className="text-xs font-bold text-blue-600 uppercase tracking-widest">
              Auto-Recalc Active
            </span>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex bg-slate-100 rounded-lg p-1 mr-4">
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

          <button
            onClick={onAddMore}
            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded uppercase tracking-wide transition-colors"
          >
            Add Next Scan
          </button>
          <button
            onClick={handleSmartInsertRow}
            className="px-4 py-2 bg-white border-2 border-slate-200 hover:border-slate-300 text-slate-700 text-xs font-bold rounded uppercase tracking-wide transition-colors"
          >
            {selectedRows.size > 0 ? "+ Row (Below)" : "+ Row"}
          </button>
          <button
            onClick={handleManualCheck}
            className="px-4 py-2 bg-violet-600 hover:bg-violet-700 text-white text-xs font-bold rounded uppercase tracking-wide shadow-sm flex items-center gap-2 transition-colors"
          >
            <span>🧮</span> Manual Check
          </button>
          <button
            onClick={onSave}
            className="px-4 py-2 bg-emerald-500 hover:bg-emerald-600 text-white text-xs font-bold rounded uppercase tracking-wide shadow-sm flex items-center gap-2 transition-colors"
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="3"
            >
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
              <polyline points="17 21 17 13 7 13 7 21" />
              <polyline points="7 3 7 8 15 8" />
            </svg>
            Save Work
          </button>
          <button
            onClick={handleDownloadCSV}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded uppercase tracking-wide shadow-sm flex items-center gap-2 transition-colors"
          >
            Download CSV
          </button>

          <div className="w-px h-8 bg-slate-200 mx-2"></div>

          {!confirmReset ? (
            <button
              onClick={() => setConfirmReset(true)}
              className="px-4 py-2 text-slate-400 hover:text-red-500 text-xs font-bold uppercase tracking-wide border border-transparent hover:border-red-100 rounded transition-colors"
            >
              Reset
            </button>
          ) : (
            <button
              onClick={onReset}
              className="px-4 py-2 bg-red-600 text-white text-xs font-bold rounded animate-pulse uppercase tracking-wide"
            >
              Confirm?
            </button>
          )}
        </div>
      </div>

      {/* --- Column Control Bar --- */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 flex items-center gap-6 overflow-x-auto shadow-sm z-10">
        <div className="flex gap-4 border-r border-slate-200 pr-6">
          <div className="flex flex-col gap-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Date Column
            </label>
            <div className="relative">
              <select
                className="appearance-none pl-3 pr-8 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none min-w-[120px]"
                value={dateColIndex}
                onChange={(e) => setDateColIndex(Number(e.target.value))}
              >
                <option value={-1}>- Select -</option>
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Receipts (+)
            </label>
            <div className="relative">
              <select
                className="appearance-none pl-3 pr-8 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none min-w-[120px]"
                value={creditColIndex}
                onChange={(e) => setCreditColIndex(Number(e.target.value))}
              >
                <option value={-1}>- Select -</option>
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Payments (-)
            </label>
            <div className="relative">
              <select
                className="appearance-none pl-3 pr-8 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none min-w-[120px]"
                value={debitColIndex}
                onChange={(e) => setDebitColIndex(Number(e.target.value))}
              >
                <option value={-1}>- Select -</option>
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
            </div>
          </div>
          {uniqueSources.length > 0 && (
            <div className="flex flex-col gap-1 border-l border-slate-200 pl-4 ml-4">
              <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                Filter Source
              </label>
              <div className="relative">
                <select
                  className="appearance-none pl-3 pr-8 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none min-w-[140px]"
                  value={sourceFilter}
                  onChange={(e) => {
                    setSourceFilter(e.target.value);
                    setSelectedRows(new Set());
                    setLastSelectedIndex(-1);
                  }}
                >
                  <option value="ALL">Show All ({data.rows.length})</option>
                  {uniqueSources.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mr-1">
            Edit Rows:
          </span>
          <div className="flex bg-slate-100 rounded p-1">
            <button
              disabled={selectedRows.size === 0}
              onClick={() => onClearRows(selectedIndices)}
              className="px-2 py-1 text-xs font-bold text-red-500 hover:bg-white rounded disabled:opacity-30 transition-colors"
            >
              DEL
            </button>
            <div className="w-px bg-slate-300 mx-1"></div>
            <button
              disabled={selectedRows.size === 0}
              onClick={() => onReverseRows(selectedIndices)}
              className="px-2 py-1 text-xs font-bold text-slate-600 hover:bg-white rounded disabled:opacity-30 hover:text-blue-600"
              title="Reverse Selection Order"
            >
              REV
            </button>
            <div className="w-px bg-slate-300 mx-1"></div>
            <button
              disabled={selectedRows.size === 0}
              onClick={() => handleLocalMoveRows("up")}
              className="px-2 py-1 text-xs font-bold text-slate-600 hover:bg-white rounded disabled:opacity-30 hover:text-blue-600"
            >
              ▲
            </button>
            <button
              disabled={selectedRows.size === 0}
              onClick={() => handleLocalMoveRows("down")}
              className="px-2 py-1 text-xs font-bold text-slate-600 hover:bg-white rounded disabled:opacity-30 hover:text-blue-600"
            >
              ▼
            </button>
            <div className="w-px bg-slate-300 mx-1"></div>
            <button
              disabled={selectedRows.size === 0}
              onClick={() => onShiftRowValues(selectedIndices, "left")}
              className="px-2 py-1 text-xs font-bold text-slate-600 hover:bg-white rounded disabled:opacity-30 hover:text-blue-600"
            >
              {"<<"}
            </button>
            <button
              disabled={selectedRows.size === 0}
              onClick={() => onShiftRowValues(selectedIndices, "right")}
              className="px-2 py-1 text-xs font-bold text-slate-600 hover:bg-white rounded disabled:opacity-30 hover:text-blue-600"
            >
              {">>"}
            </button>
          </div>

          <button
            disabled={
              selectedRows.size === 0 ||
              creditColIndex === -1 ||
              debitColIndex === -1
            }
            onClick={() =>
              onSwapValues(
                selectedIndices,
                creditColIndex,
                debitColIndex,
                getCalcConfig()
              )
            }
            className="ml-2 px-3 py-1.5 bg-purple-50 text-purple-700 border border-purple-200 rounded text-xs font-bold hover:bg-purple-100 disabled:opacity-50 flex items-center gap-1 transition-colors"
          >
            <span className="text-[10px]">⇄</span> Swap I/O
          </button>

          {selectedRows.size === 2 && (
            <>
              <button
                onClick={() => onSwapRows(selectedIndices)}
                className="ml-2 px-3 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded text-xs font-bold hover:bg-blue-100 flex items-center gap-1 transition-colors"
              >
                <span>⇅</span> Swap Rows
              </button>
              <button
                onClick={() =>
                  onSwapRowContents(selectedIndices, [balanceColIndex])
                }
                className="ml-2 px-3 py-1.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded text-xs font-bold hover:bg-indigo-100 flex items-center gap-1 transition-colors"
              >
                <span>⇋</span> Swap Content
              </button>
            </>
          )}

          <button
            disabled={selectedRows.size === 0}
            onClick={() => onAugment(selectedIndices)}
            className="ml-2 px-3 py-1.5 bg-amber-100 text-amber-700 border border-amber-200 rounded text-xs font-bold hover:bg-amber-200 disabled:opacity-50 flex items-center gap-1 transition-colors shadow-sm"
          >
            <span className="text-[10px]">✨</span> Augment Scan
          </button>
        </div>

        <div className="ml-auto flex items-center gap-4 pl-4 border-l border-slate-200">
          <div className="flex flex-col gap-1">
            <div className="flex items-center justify-between">
              <span className="text-[10px] font-bold text-slate-400 uppercase">
                Bulk Fix Numbers
              </span>
            </div>
            <div className="flex items-center gap-1">
              <select
                className="text-xs border border-slate-200 rounded p-1 max-w-[80px] bg-slate-50 focus:border-blue-500 outline-none"
                value={fixColIndex}
                onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                  setFixColIndex(Number(e.target.value))
                }
              >
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
              <select
                className="text-xs border border-slate-200 rounded p-1 max-w-[110px] bg-slate-50 focus:border-blue-500 outline-none"
                value={fixOperation}
                onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                  setFixOperation(e.target.value as any)
                }
              >
                <option value="NEG_SUFFIX">D/DR/DB/OUT → (-)</option>
                <option value="CLEAN">Clean Text</option>
                <option value="FORCE_NEG">Force (-)</option>
                <option value="FORCE_POS">Force (+)</option>
                <option value="INVERT">Invert Sign</option>
              </select>
              <button
                disabled={selectedRows.size === 0}
                onClick={() =>
                  onBulkFixNumbers(
                    selectedIndices,
                    fixColIndex,
                    fixOperation,
                    getCalcConfig()
                  )
                }
                className="px-2 py-1 bg-indigo-500 hover:bg-indigo-600 text-white rounded text-xs font-bold disabled:opacity-50 transition-colors uppercase"
              >
                FIX
              </button>
            </div>
          </div>

          <div className="w-px bg-slate-200 h-8"></div>

          <div className="flex flex-col gap-1">
            <span className="text-[10px] font-bold text-slate-400 uppercase">
              Move Column
            </span>
            <div className="flex items-center gap-1">
              <select
                className="text-xs border border-slate-200 rounded p-1 max-w-[80px] bg-slate-50 focus:border-blue-500 outline-none"
                value={moveSourceCol}
                onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                  setMoveSourceCol(Number(e.target.value))
                }
              >
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
              <span className="text-[10px] text-slate-400">→</span>
              <select
                className="text-xs border border-slate-200 rounded p-1 max-w-[80px] bg-slate-50 focus:border-blue-500 outline-none"
                value={moveTargetCol}
                onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                  setMoveTargetCol(Number(e.target.value))
                }
              >
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
              <button
                disabled={
                  selectedRows.size === 0 || moveSourceCol === moveTargetCol
                }
                onClick={() =>
                  onBulkMoveValues(
                    selectedIndices,
                    moveSourceCol,
                    moveTargetCol,
                    getCalcConfig()
                  )
                }
                className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-xs font-bold disabled:opacity-50 transition-colors uppercase"
              >
                MOVE
              </button>
            </div>
          </div>

          <div className="w-px bg-slate-200 h-8"></div>

          <div className="flex flex-col gap-1">
            <span className="text-[10px] font-bold text-slate-400 uppercase">
              Set Value
            </span>
            <div className="flex items-center gap-1">
              <select
                className="text-xs border border-slate-200 rounded p-1 max-w-[80px] bg-slate-50 focus:border-blue-500 outline-none"
                value={bulkSetCol}
                onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                  setBulkSetCol(Number(e.target.value))
                }
              >
                {data.headers.map((h, i) => (
                  <option key={i} value={i}>
                    {h}
                  </option>
                ))}
              </select>
              <input
                type="text"
                value={bulkSetValue}
                onChange={(e) => setBulkSetValue(e.target.value)}
                className="w-20 p-1 border border-slate-200 bg-slate-50 focus:border-blue-500 outline-none rounded text-xs text-center"
                placeholder="Value"
              />
              <button
                disabled={selectedRows.size === 0}
                onClick={() =>
                  onBulkSetValue(
                    selectedIndices,
                    bulkSetCol,
                    bulkSetValue,
                    getCalcConfig()
                  )
                }
                className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-xs font-bold disabled:opacity-50 transition-colors uppercase"
              >
                SET
              </button>
            </div>
          </div>

          <div className="w-px bg-slate-200 h-8"></div>

          <div className="flex flex-col gap-1">
            <span className="text-[10px] font-bold text-slate-400 uppercase">
              Append Year
            </span>
            <div className="flex items-center gap-1">
              <input
                type="text"
                value={bulkYear}
                onChange={(e) => setBulkYear(e.target.value)}
                className="w-12 p-1 border border-slate-200 bg-slate-50 focus:border-blue-500 outline-none rounded text-xs text-center"
              />
              <button
                disabled={selectedRows.size === 0 || dateColIndex === -1}
                onClick={() =>
                  onBulkAddYear(selectedIndices, dateColIndex, bulkYear)
                }
                className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-xs font-bold disabled:opacity-50 transition-colors uppercase"
              >
                ADD
              </button>
            </div>
          </div>
        </div>
      </div>

      <div ref={tableContainerRef} className="flex-1 bg-slate-200 p-6 relative">
        <TableVirtuoso
          style={{ height: "100%", width: "100%" }}
          data={visibleRowIndices}
          context={{ selectedRows, totalTableWidth }}
          components={virtuosoComponents}
          fixedHeaderContent={() => (
            <tr className="bg-slate-900 text-white uppercase text-xs font-bold tracking-wider">
              <th className="sticky top-0 z-30 w-12 p-3 text-center border-r border-slate-700 bg-slate-800 first:rounded-tl-lg shadow-sm">
                <input
                  type="checkbox"
                  ref={selectAllRef}
                  checked={isAllVisibleSelected}
                  onChange={handleSelectAll}
                  className="rounded bg-slate-700 border-slate-600"
                />
              </th>
              {data.headers.map((h, i) => {
                let headerClass =
                  "sticky top-0 z-20 p-3 text-left border-r border-slate-700 relative group overflow-hidden bg-slate-900 select-none shadow-sm";
                if (i === dateColIndex) headerClass += " text-yellow-400";
                if (i === creditColIndex) headerClass += " text-emerald-400";
                if (i === debitColIndex) headerClass += " text-red-400";
                if (i === manualColIndex)
                  headerClass += " bg-indigo-900 text-indigo-200";
                const w = colWidths[i] || 120;
                return (
                  <th key={i} className={headerClass} style={{ width: w }}>
                    <div className="flex items-center justify-between">
                      <span
                        className="truncate cursor-pointer hover:text-blue-300 transition-colors"
                        onClick={() => handleHeaderSort(i)}
                      >
                        {h}
                        {sortState?.colIdx === i && (
                          <span className="ml-1 opacity-100">
                            {sortState.direction === "asc" ? "▲" : "▼"}
                          </span>
                        )}
                      </span>
                      <div className="flex flex-col gap-0.5 opacity-50 group-hover:opacity-100 transition-opacity mr-2">
                        <button
                          onClick={() => onShiftColumn(i, "up", 0)}
                          className="hover:text-white text-slate-400"
                        >
                          <svg
                            width="8"
                            height="8"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="4"
                          >
                            <polyline points="18 15 12 9 6 15" />
                          </svg>
                        </button>
                        <button
                          onClick={() => onShiftColumn(i, "down", 0)}
                          className="hover:text-white text-slate-400"
                        >
                          <svg
                            width="8"
                            height="8"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="4"
                          >
                            <polyline points="6 9 12 15 18 9" />
                          </svg>
                        </button>
                      </div>
                    </div>
                    <div className="text-[9px] mt-1 font-normal opacity-70 truncate">
                      {i === dateColIndex && "DATE (DATE)"}
                      {i === creditColIndex && "MONEY IN (+)"}
                      {i === debitColIndex && "MONEY OUT (-)"}
                      {i === balanceColIndex && "BALANCE"}
                      {i === manualColIndex && "MANUAL CALC"}
                    </div>
                    <div
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-blue-500 z-50"
                      onMouseDown={(e) => startResize(i, e)}
                    ></div>
                  </th>
                );
              })}
              <th
                className="sticky top-0 z-20 p-3 text-right border-l border-slate-700 bg-slate-900 last:rounded-tr-lg relative shadow-sm"
                style={{ width: diffColWidth }}
              >
                <span className="truncate">DIFFERENCE</span>
                <div
                  className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-blue-500 z-50"
                  onMouseDown={(e) => startResize("diff", e)}
                ></div>
              </th>
            </tr>
          )}
          itemContent={(index, rowIdx) => (
            <RowCells
              key={rowIdx}
              row={data.rows[rowIdx]}
              rowIdx={rowIdx}
              isSelected={selectedRows.has(rowIdx)}
              colWidths={colWidths}
              manualColIndex={manualColIndex}
              creditColIndex={creditColIndex}
              debitColIndex={debitColIndex}
              balanceColIndex={balanceColIndex}
              onToggleSelection={toggleRowSelection}
              onCellChange={handleCellChange}
              onFocusCell={(r: number, c: number) => setFocusedCell({ r, c })}
              dragState={dragState}
            />
          )}
        />
      </div>

      {autofillMenu && (
        <>
          <div
            className="fixed inset-0 z-40"
            onClick={() => setAutofillMenu(null)}
          ></div>
          <div
            className="fixed bg-white rounded-lg shadow-xl border border-slate-200 z-50 flex flex-col p-1 animate-in fade-in zoom-in-95 duration-100 min-w-[180px]"
            style={{
              top: Math.min(window.innerHeight - 100, autofillMenu.y + 5),
              left: Math.min(window.innerWidth - 200, autofillMenu.x + 5),
            }}
          >
            <div className="px-3 py-1 text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">
              Autofill Options
            </div>
            <button
              onClick={() => handleAutofillAction("NEXT_VALUE")}
              className="flex items-center gap-2 px-3 py-2 text-left text-xs font-semibold text-slate-700 hover:bg-blue-50 hover:text-blue-700 rounded-md transition-colors"
            >
              <span className="text-slate-400">⬇</span> Fill Down (To Next
              Value)
            </button>
            <button
              onClick={() => handleAutofillAction("COMPLETE")}
              className="flex items-center gap-2 px-3 py-2 text-left text-xs font-semibold text-slate-700 hover:bg-purple-50 hover:text-purple-700 rounded-md transition-colors"
            >
              <span className="text-slate-400">⚡</span> Fill Down Complete
              (Waterfall)
            </button>
          </div>
        </>
      )}
    </div>
  );
};
