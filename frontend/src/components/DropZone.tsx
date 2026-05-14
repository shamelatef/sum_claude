import React, { useCallback, useState, useRef } from 'react';

interface DropZoneProps {
  onFile: (file: File) => void;
  loading: boolean;
}

export const DropZone: React.FC<DropZoneProps> = ({ onFile, loading }) => {
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  const validate = (file: File): boolean => {
    const ok = /\.(xlsx|xls)$/i.test(file.name);
    if (!ok) setError('Please upload an .xlsx or .xls file.');
    else setError('');
    return ok;
  };

  const handleFiles = useCallback(
    (files: FileList | null) => {
      if (!files || files.length === 0) return;
      const file = files[0];
      if (validate(file)) onFile(file);
    },
    [onFile]
  );

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    handleFiles(e.dataTransfer.files);
  };

  return (
    <div className="w-full">
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        onClick={() => !loading && inputRef.current?.click()}
        className={`
          relative flex flex-col items-center justify-center gap-4
          border-2 border-dashed rounded-2xl p-12 cursor-pointer
          transition-all duration-200 select-none
          ${dragging
            ? 'border-vois-600 bg-vois-50 scale-[1.01]'
            : 'border-vois-300 bg-white hover:border-vois-500 hover:bg-vois-50'
          }
          ${loading ? 'opacity-60 pointer-events-none' : ''}
        `}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => handleFiles(e.target.files)}
        />

        {loading ? (
          <div className="flex flex-col items-center gap-3">
            <svg className="animate-spin h-10 w-10 text-vois-600" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
            </svg>
            <span className="text-vois-700 font-medium">Parsing file…</span>
          </div>
        ) : (
          <>
            <div className="w-16 h-16 rounded-2xl bg-vois-100 flex items-center justify-center">
              <svg className="w-8 h-8 text-vois-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            </div>
            <div className="text-center">
              <p className="text-lg font-semibold text-vois-800">
                {dragging ? 'Drop to upload' : 'Drag & drop your Excel file'}
              </p>
              <p className="text-sm text-gray-500 mt-1">or click to browse — .xlsx, .xls supported</p>
            </div>
            <div className="flex gap-2">
              <span className="px-3 py-1 bg-vois-100 text-vois-700 rounded-full text-xs font-medium">.xlsx</span>
              <span className="px-3 py-1 bg-vois-100 text-vois-700 rounded-full text-xs font-medium">.xls</span>
              <span className="px-3 py-1 bg-gray-100 text-gray-600 rounded-full text-xs font-medium">Max 20 MB</span>
            </div>
          </>
        )}
      </div>
      {error && (
        <p className="mt-2 text-sm text-red-600 flex items-center gap-1">
          <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
            <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
          </svg>
          {error}
        </p>
      )}
    </div>
  );
};
