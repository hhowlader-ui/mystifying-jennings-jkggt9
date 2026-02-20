import React, { useCallback } from "react";

interface Props {
  onFileSelect: (file: File) => void;
  onBack: () => void;
  isLoading: boolean;
}

export const FileUploader: React.FC<Props> = ({
  onFileSelect,
  onBack,
  isLoading,
}) => {
  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        onFileSelect(e.dataTransfer.files[0]);
      }
    },
    [onFileSelect]
  );

  return (
    <div
      onDrop={handleDrop}
      onDragOver={(e) => e.preventDefault()}
      className="w-full max-w-md p-12 border-2 border-dashed border-slate-300 rounded-xl text-center hover:border-blue-500 hover:bg-blue-50 transition-colors cursor-pointer bg-white shadow-sm"
    >
      <input
        type="file"
        id="file-upload"
        className="hidden"
        onChange={(e) => e.target.files && onFileSelect(e.target.files[0])}
        accept=".csv,.xls,.xlsx,.pdf,.json"
      />
      <label
        htmlFor="file-upload"
        className="cursor-pointer flex flex-col items-center"
      >
        <div className="text-4xl mb-4">📄</div>
        <span className="text-sm font-bold text-slate-700">
          Click to upload or drag and drop
        </span>
        <span className="text-xs text-slate-500 mt-2">
          CSV, Excel, PDF, or JSON Project
        </span>
      </label>

      {isLoading && (
        <div className="mt-6 text-sm font-bold text-blue-600 animate-pulse">
          Processing file...
        </div>
      )}
    </div>
  );
};
