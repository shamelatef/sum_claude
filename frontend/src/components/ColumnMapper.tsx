import React from 'react';
import { ColumnMapping } from '../types';

interface Field {
  key: keyof ColumnMapping;
  label: string;
  description: string;
  icon: string;
}

const FIELDS: Field[] = [
  { key: 'pmName',       label: 'PM Name',       description: 'Project Manager full name', icon: '👤' },
  { key: 'projectName',  label: 'Project Name',   description: 'Name or title of the project', icon: '📋' },
  { key: 'projectId',    label: 'Project ID',     description: 'Unique project code or reference', icon: '🔖' },
  { key: 'gateApproved', label: 'Gate Approved',  description: 'Project gate / approval stage', icon: '✅' },
];

interface ColumnMapperProps {
  headers: string[];
  mapping: ColumnMapping;
  onFieldChange: (field: keyof ColumnMapping, value: string) => void;
}

export const ColumnMapper: React.FC<ColumnMapperProps> = ({ headers, mapping, onFieldChange }) => {
  return (
    <div className="w-full">
      <p className="text-sm text-gray-500 mb-4">
        Map your Excel columns to the required fields. Auto-detected suggestions are pre-filled.
      </p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        {FIELDS.map((field) => {
          const value = mapping[field.key];
          const isSet = value !== '';
          return (
            <div
              key={field.key}
              className={`
                rounded-xl border p-4 transition-all duration-150
                ${isSet
                  ? 'border-vois-400 bg-vois-50 shadow-sm'
                  : 'border-gray-200 bg-white'
                }
              `}
            >
              <div className="flex items-center gap-2 mb-2">
                <span className="text-lg">{field.icon}</span>
                <div>
                  <p className="text-sm font-semibold text-gray-800">{field.label}</p>
                  <p className="text-xs text-gray-400">{field.description}</p>
                </div>
                {isSet && (
                  <span className="ml-auto">
                    <svg className="w-4 h-4 text-green-500" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                  </span>
                )}
              </div>
              <select
                value={value}
                onChange={(e) => onFieldChange(field.key, e.target.value)}
                className={`
                  w-full text-sm rounded-lg border px-3 py-2 appearance-none
                  focus:outline-none focus:ring-2 focus:ring-vois-500 focus:border-transparent
                  transition-colors cursor-pointer
                  ${isSet
                    ? 'border-vois-300 bg-white text-vois-800 font-medium'
                    : 'border-gray-300 bg-white text-gray-500'
                  }
                `}
              >
                <option value="">— Select column —</option>
                {headers.map((h) => (
                  <option key={h} value={h}>{h}</option>
                ))}
              </select>
            </div>
          );
        })}
      </div>
    </div>
  );
};
