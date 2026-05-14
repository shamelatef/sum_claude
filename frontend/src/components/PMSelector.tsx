import React, { useState, useRef, useEffect } from 'react';

interface PMSelectorProps {
  pms: string[];
  selected: string[];
  onChange: (selected: string[]) => void;
}

export const PMSelector: React.FC<PMSelectorProps> = ({ pms, selected, onChange }) => {
  const [search, setSearch] = useState('');
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const filtered = pms.filter((pm) =>
    pm.toLowerCase().includes(search.toLowerCase())
  );

  const allSelected = selected.length === pms.length && pms.length > 0;
  const noneSelected = selected.length === 0;

  const toggle = (pm: string) => {
    if (selected.includes(pm)) {
      onChange(selected.filter((s) => s !== pm));
    } else {
      onChange([...selected, pm]);
    }
  };

  const selectAll = () => onChange([...pms]);
  const clearAll = () => onChange([]);

  return (
    <div className="w-full" ref={containerRef}>
      <p className="text-sm text-gray-500 mb-3">
        Choose which Project Managers to include in the generated presentation.
      </p>

      {/* Summary chips */}
      <div className="flex flex-wrap gap-2 mb-3 min-h-[36px]">
        {selected.length === 0 && (
          <span className="text-sm text-gray-400 italic py-1">No PMs selected</span>
        )}
        {selected.slice(0, 5).map((pm) => (
          <span
            key={pm}
            className="inline-flex items-center gap-1 px-3 py-1 bg-vois-100 text-vois-800 rounded-full text-xs font-medium"
          >
            {pm}
            <button
              onClick={() => toggle(pm)}
              className="hover:text-vois-600 ml-0.5 text-vois-500"
            >
              ×
            </button>
          </span>
        ))}
        {selected.length > 5 && (
          <span className="px-3 py-1 bg-vois-200 text-vois-800 rounded-full text-xs font-medium">
            +{selected.length - 5} more
          </span>
        )}
      </div>

      {/* Dropdown trigger */}
      <div
        className="relative"
      >
        <button
          type="button"
          onClick={() => setOpen((o) => !o)}
          className="w-full flex items-center justify-between px-4 py-3 rounded-xl border border-gray-300 bg-white hover:border-vois-400 focus:outline-none focus:ring-2 focus:ring-vois-500 transition-colors text-sm"
        >
          <span className={selected.length === 0 ? 'text-gray-400' : 'text-gray-700 font-medium'}>
            {selected.length === 0
              ? 'Select Project Managers…'
              : `${selected.length} of ${pms.length} PM${pms.length !== 1 ? 's' : ''} selected`}
          </span>
          <svg
            className={`w-4 h-4 text-gray-400 transition-transform duration-200 ${open ? 'rotate-180' : ''}`}
            fill="none" stroke="currentColor" viewBox="0 0 24 24"
          >
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
          </svg>
        </button>

        {open && (
          <div className="absolute z-30 mt-2 w-full bg-white border border-gray-200 rounded-2xl shadow-2xl overflow-hidden">
            {/* Search */}
            <div className="p-3 border-b border-gray-100">
              <div className="relative">
                <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                </svg>
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search PMs…"
                  className="w-full pl-9 pr-3 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-vois-500"
                  autoFocus
                />
              </div>
            </div>

            {/* Actions */}
            <div className="flex items-center gap-2 px-3 py-2 border-b border-gray-100 bg-gray-50">
              <button
                onClick={selectAll}
                disabled={allSelected}
                className="text-xs font-semibold text-vois-600 hover:text-vois-800 disabled:text-gray-300 disabled:cursor-not-allowed transition-colors"
              >
                Select All ({pms.length})
              </button>
              <span className="text-gray-300">|</span>
              <button
                onClick={clearAll}
                disabled={noneSelected}
                className="text-xs font-semibold text-gray-500 hover:text-red-600 disabled:text-gray-300 disabled:cursor-not-allowed transition-colors"
              >
                Clear All
              </button>
              <span className="ml-auto text-xs text-gray-400">
                {selected.length} selected
              </span>
            </div>

            {/* PM list */}
            <ul className="max-h-56 overflow-y-auto">
              {filtered.length === 0 && (
                <li className="px-4 py-3 text-sm text-gray-400 text-center">No PMs match your search.</li>
              )}
              {filtered.map((pm) => {
                const isSelected = selected.includes(pm);
                return (
                  <li
                    key={pm}
                    onClick={() => toggle(pm)}
                    className={`
                      flex items-center gap-3 px-4 py-2.5 cursor-pointer transition-colors text-sm
                      ${isSelected
                        ? 'bg-vois-50 text-vois-800'
                        : 'text-gray-700 hover:bg-gray-50'
                      }
                    `}
                  >
                    <div
                      className={`
                        w-4 h-4 rounded border-2 flex items-center justify-center flex-shrink-0 transition-colors
                        ${isSelected
                          ? 'bg-vois-600 border-vois-600'
                          : 'border-gray-300 bg-white'
                        }
                      `}
                    >
                      {isSelected && (
                        <svg className="w-2.5 h-2.5 text-white" fill="currentColor" viewBox="0 0 12 12">
                          <path d="M10 3L5 8.5 2 5.5" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" fill="none" />
                        </svg>
                      )}
                    </div>
                    <span className={isSelected ? 'font-medium' : ''}>{pm}</span>
                  </li>
                );
              })}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
};
