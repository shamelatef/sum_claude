import React, { useState } from 'react';
import { RawRow } from '../types';

interface ExcelPreviewProps {
  headers: string[];
  rows: RawRow[];
  totalRows: number;
}

const PREVIEW_ROWS = 8;

export const ExcelPreview: React.FC<ExcelPreviewProps> = ({ headers, rows, totalRows }) => {
  const [showAll, setShowAll] = useState(false);
  const displayed = showAll ? rows : rows.slice(0, PREVIEW_ROWS);

  return (
    <div className="w-full">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2">
          <div className="w-2 h-2 rounded-full bg-green-500" />
          <span className="text-sm font-medium text-gray-700">
            Excel Preview
          </span>
          <span className="px-2 py-0.5 bg-vois-100 text-vois-700 rounded-full text-xs font-semibold">
            {totalRows} rows · {headers.length} columns
          </span>
        </div>
        {rows.length > PREVIEW_ROWS && (
          <button
            onClick={() => setShowAll((s) => !s)}
            className="text-xs text-vois-600 hover:text-vois-800 font-medium transition-colors"
          >
            {showAll ? 'Show less' : `Show all ${totalRows} rows`}
          </button>
        )}
      </div>

      <div className="overflow-auto rounded-xl border border-gray-200 shadow-sm max-h-64">
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-vois-800 text-white sticky top-0">
              <th className="px-3 py-2 text-left font-semibold text-vois-200 w-8">#</th>
              {headers.map((h) => (
                <th key={h} className="px-3 py-2 text-left font-semibold whitespace-nowrap">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayed.map((row, i) => (
              <tr
                key={i}
                className={i % 2 === 0 ? 'bg-white' : 'bg-vois-50'}
              >
                <td className="px-3 py-1.5 text-gray-400 font-mono">{i + 1}</td>
                {headers.map((h) => (
                  <td key={h} className="px-3 py-1.5 text-gray-700 whitespace-nowrap max-w-[180px] truncate">
                    {row[h] !== undefined && row[h] !== null ? String(row[h]) : ''}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {!showAll && rows.length > PREVIEW_ROWS && (
        <p className="text-center text-xs text-gray-400 mt-2">
          Showing {PREVIEW_ROWS} of {totalRows} rows
        </p>
      )}
    </div>
  );
};
