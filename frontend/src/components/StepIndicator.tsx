import React from 'react';
import { Step } from '../types';

interface StepDef {
  key: Step;
  label: string;
  icon: React.ReactNode;
}

const STEPS: StepDef[] = [
  {
    key: 'upload',
    label: 'Upload',
    icon: (
      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
      </svg>
    ),
  },
  {
    key: 'map',
    label: 'Map Columns',
    icon: (
      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
      </svg>
    ),
  },
  {
    key: 'select',
    label: 'Select PMs',
    icon: (
      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z" />
      </svg>
    ),
  },
  {
    key: 'generate',
    label: 'Generate',
    icon: (
      <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
          d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
      </svg>
    ),
  },
];

const ORDER: Step[] = ['upload', 'map', 'select', 'generate'];

interface StepIndicatorProps {
  current: Step;
}

export const StepIndicator: React.FC<StepIndicatorProps> = ({ current }) => {
  const currentIdx = ORDER.indexOf(current);

  return (
    <div className="flex items-center w-full max-w-2xl mx-auto">
      {STEPS.map((step, i) => {
        const stepIdx = ORDER.indexOf(step.key);
        const done = stepIdx < currentIdx;
        const active = stepIdx === currentIdx;

        return (
          <React.Fragment key={step.key}>
            <div className="flex flex-col items-center gap-1.5 flex-shrink-0">
              <div
                className={`
                  w-9 h-9 rounded-full flex items-center justify-center
                  border-2 transition-all duration-300
                  ${done
                    ? 'bg-vois-600 border-vois-600 text-white'
                    : active
                    ? 'bg-white border-vois-600 text-vois-600 shadow-md shadow-vois-200'
                    : 'bg-white border-gray-200 text-gray-300'
                  }
                `}
              >
                {done ? (
                  <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
                    <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                  </svg>
                ) : (
                  step.icon
                )}
              </div>
              <span
                className={`text-xs font-medium whitespace-nowrap ${
                  active ? 'text-vois-700' : done ? 'text-vois-500' : 'text-gray-400'
                }`}
              >
                {step.label}
              </span>
            </div>

            {i < STEPS.length - 1 && (
              <div
                className={`
                  flex-1 h-0.5 mx-2 mb-5 rounded transition-all duration-300
                  ${stepIdx < currentIdx ? 'bg-vois-500' : 'bg-gray-200'}
                `}
              />
            )}
          </React.Fragment>
        );
      })}
    </div>
  );
};
