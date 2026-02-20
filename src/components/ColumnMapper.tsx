import React, { useState, useEffect } from "react";

interface Props {
  headers: string[];
  filename: string;
  onConfirm: (mapping: MappingConfig) => void;
  onCancel: () => void;
}

export interface MappingConfig {
  dateIdx: number;
  descIdx: number;
  inIdx: number; // Used if isSingleAmountColumn is false
  outIdx: number; // Used if isSingleAmountColumn is false
  amountIdx: number; // Used if isSingleAmountColumn is true
  balanceIdx: number;
  isSingleAmountColumn: boolean;
}

export const ColumnMapper: React.FC<Props> = ({
  headers,
  filename,
  onConfirm,
  onCancel,
}) => {
  const [isSingleAmount, setIsSingleAmount] = useState(false);

  // Initialize with "best guess" logic
  const [mapping, setMapping] = useState<MappingConfig>({
    dateIdx: -1,
    descIdx: -1,
    inIdx: -1,
    outIdx: -1,
    amountIdx: -1,
    balanceIdx: -1,
    isSingleAmountColumn: false,
  });

  useEffect(() => {
    // Auto-detect columns based on common names
    const newMapping = { ...mapping };
    headers.forEach((h, idx) => {
      const lower = h.toLowerCase();
      if (
        (lower.includes("date") || lower.includes("dt")) &&
        newMapping.dateIdx === -1
      )
        newMapping.dateIdx = idx;
      if (
        (lower.includes("desc") ||
          lower.includes("narrative") ||
          lower.includes("details") ||
          lower.includes("transaction")) &&
        newMapping.descIdx === -1
      )
        newMapping.descIdx = idx;
      if (
        (lower.includes("balance") || lower.includes("bal")) &&
        newMapping.balanceIdx === -1
      )
        newMapping.balanceIdx = idx;

      if (lower.includes("amount") || lower.includes("value")) {
        if (newMapping.amountIdx === -1) newMapping.amountIdx = idx;
      }
      if (
        (lower.includes("in") ||
          lower.includes("credit") ||
          lower.includes("receipt")) &&
        newMapping.inIdx === -1
      )
        newMapping.inIdx = idx;
      if (
        (lower.includes("out") ||
          lower.includes("debit") ||
          lower.includes("payment")) &&
        newMapping.outIdx === -1
      )
        newMapping.outIdx = idx;
    });

    // Heuristic: If we found "Amount" but not In/Out, default to Single Column mode
    if (
      newMapping.amountIdx !== -1 &&
      newMapping.inIdx === -1 &&
      newMapping.outIdx === -1
    ) {
      setIsSingleAmount(true);
    }

    setMapping((prev) => ({ ...prev, ...newMapping }));
  }, [headers]);

  const handleConfirm = () => {
    onConfirm({ ...mapping, isSingleAmountColumn: isSingleAmount });
  };

  const getSelectClass = (isSelected: boolean) =>
    `w-full p-2 border rounded-lg text-sm font-bold outline-none focus:ring-2 focus:ring-blue-500 ${
      isSelected
        ? "border-blue-300 bg-white text-slate-800"
        : "border-slate-200 bg-slate-50 text-slate-400"
    }`;

  return (
    <div className="fixed inset-0 bg-black/50 z-[100] flex items-center justify-center p-4 backdrop-blur-sm">
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[90vh] animate-bounce-in">
        <div className="bg-slate-50 p-6 border-b border-slate-200">
          <div className="flex items-center gap-3 mb-1">
            <div className="w-10 h-10 bg-green-100 text-green-600 rounded-lg flex items-center justify-center font-black shadow-inner">
              xls
            </div>
            <div>
              <h3 className="text-xl font-black text-slate-900">Map Columns</h3>
              <p className="text-xs text-slate-500 font-bold truncate max-w-[300px]">
                {filename}
              </p>
            </div>
          </div>
        </div>

        <div className="p-6 overflow-y-auto flex-1 space-y-6">
          {/* Mode Toggle */}
          <div className="bg-blue-50/50 p-4 rounded-lg border border-blue-100">
            <label className="flex items-center gap-3 cursor-pointer">
              <div
                className={`w-12 h-6 rounded-full p-1 transition-colors ${
                  isSingleAmount ? "bg-blue-600" : "bg-slate-300"
                }`}
              >
                <div
                  className={`w-4 h-4 bg-white rounded-full shadow-sm transition-transform ${
                    isSingleAmount ? "translate-x-6" : "translate-x-0"
                  }`}
                ></div>
              </div>
              <input
                type="checkbox"
                className="hidden"
                checked={isSingleAmount}
                onChange={(e) => setIsSingleAmount(e.target.checked)}
              />
              <div>
                <div className="text-sm font-bold text-slate-800">
                  Single 'Amount' Column?
                </div>
                <div className="text-[10px] text-slate-500">
                  Enable if values use +/- signs in one column
                </div>
              </div>
            </label>
          </div>

          <div className="space-y-4">
            {/* Date */}
            <div className="grid grid-cols-3 items-center gap-4">
              <label className="text-right text-xs font-black text-slate-500 uppercase">
                Date
              </label>
              <div className="col-span-2">
                <select
                  className={getSelectClass(mapping.dateIdx !== -1)}
                  value={mapping.dateIdx}
                  onChange={(e) =>
                    setMapping({ ...mapping, dateIdx: Number(e.target.value) })
                  }
                >
                  <option value={-1}>- Select Column -</option>
                  {headers.map((h, i) => (
                    <option key={i} value={i}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Description */}
            <div className="grid grid-cols-3 items-center gap-4">
              <label className="text-right text-xs font-black text-slate-500 uppercase">
                Description
              </label>
              <div className="col-span-2">
                <select
                  className={getSelectClass(mapping.descIdx !== -1)}
                  value={mapping.descIdx}
                  onChange={(e) =>
                    setMapping({ ...mapping, descIdx: Number(e.target.value) })
                  }
                >
                  <option value={-1}>- Select Column -</option>
                  {headers.map((h, i) => (
                    <option key={i} value={i}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Amount Logic */}
            {isSingleAmount ? (
              <div className="grid grid-cols-3 items-center gap-4">
                <label className="text-right text-xs font-black text-slate-500 uppercase">
                  Amount (+/-)
                </label>
                <div className="col-span-2">
                  <select
                    className={getSelectClass(mapping.amountIdx !== -1)}
                    value={mapping.amountIdx}
                    onChange={(e) =>
                      setMapping({
                        ...mapping,
                        amountIdx: Number(e.target.value),
                      })
                    }
                  >
                    <option value={-1}>- Select Column -</option>
                    {headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            ) : (
              <>
                <div className="grid grid-cols-3 items-center gap-4">
                  <label className="text-right text-xs font-black text-slate-500 uppercase text-emerald-600">
                    Money In
                  </label>
                  <div className="col-span-2">
                    <select
                      className={getSelectClass(mapping.inIdx !== -1)}
                      value={mapping.inIdx}
                      onChange={(e) =>
                        setMapping({
                          ...mapping,
                          inIdx: Number(e.target.value),
                        })
                      }
                    >
                      <option value={-1}>- Select Column -</option>
                      {headers.map((h, i) => (
                        <option key={i} value={i}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="grid grid-cols-3 items-center gap-4">
                  <label className="text-right text-xs font-black text-slate-500 uppercase text-red-600">
                    Money Out
                  </label>
                  <div className="col-span-2">
                    <select
                      className={getSelectClass(mapping.outIdx !== -1)}
                      value={mapping.outIdx}
                      onChange={(e) =>
                        setMapping({
                          ...mapping,
                          outIdx: Number(e.target.value),
                        })
                      }
                    >
                      <option value={-1}>- Select Column -</option>
                      {headers.map((h, i) => (
                        <option key={i} value={i}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </>
            )}

            {/* Balance */}
            <div className="grid grid-cols-3 items-center gap-4">
              <label className="text-right text-xs font-black text-slate-500 uppercase">
                Balance (Opt)
              </label>
              <div className="col-span-2">
                <select
                  className={getSelectClass(mapping.balanceIdx !== -1)}
                  value={mapping.balanceIdx}
                  onChange={(e) =>
                    setMapping({
                      ...mapping,
                      balanceIdx: Number(e.target.value),
                    })
                  }
                >
                  <option value={-1}>- Select Column -</option>
                  {headers.map((h, i) => (
                    <option key={i} value={i}>
                      {h}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        </div>

        <div className="p-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-3">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-slate-500 font-bold text-xs hover:bg-slate-200 rounded-lg uppercase tracking-wide transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={handleConfirm}
            disabled={
              mapping.dateIdx === -1 ||
              mapping.descIdx === -1 ||
              (isSingleAmount
                ? mapping.amountIdx === -1
                : mapping.inIdx === -1 && mapping.outIdx === -1)
            }
            className="px-6 py-2 bg-blue-600 text-white font-bold text-xs rounded-lg shadow-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed uppercase tracking-wide transition-colors"
          >
            Import Data
          </button>
        </div>
      </div>
    </div>
  );
};
