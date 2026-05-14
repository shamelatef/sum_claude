import { useState, useEffect } from 'react';
import { ColumnMapping } from '../types';

function autoDetect(headers: string[], keywords: string[]): string {
  const lower = headers.map((h) => h.toLowerCase());
  for (const kw of keywords) {
    const idx = lower.findIndex((h) => h.includes(kw));
    if (idx !== -1) return headers[idx];
  }
  return '';
}

export function useColumnMapping(headers: string[]) {
  const [mapping, setMapping] = useState<ColumnMapping>({
    pmName: '',
    projectName: '',
    projectId: '',
    gateApproved: '',
  });

  useEffect(() => {
    if (headers.length === 0) return;
    setMapping({
      pmName:       autoDetect(headers, ['oit project manager', 'pm', 'manager', 'owner', 'responsible']),
      projectName:  autoDetect(headers, ['title', 'project name', 'name', 'project title']),
      projectId:    autoDetect(headers, ['projectno', 'id', 'code', 'number', 'ref']),
      gateApproved: autoDetect(headers, ['gate approved', 'gate', 'approved', 'stage', 'phase', 'status']),
    });
  }, [headers]);

  const setField = (field: keyof ColumnMapping, value: string) => {
    setMapping((prev) => ({ ...prev, [field]: value }));
  };

  const isComplete =
    mapping.pmName !== '' &&
    mapping.projectName !== '' &&
    mapping.projectId !== '' &&
    mapping.gateApproved !== '';

  return { mapping, setField, isComplete };
}
