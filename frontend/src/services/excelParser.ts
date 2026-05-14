import * as XLSX from 'xlsx';
import { RawRow, ColumnMapping, PMGroup, GateGroup, Project } from '../types';
import { cleanPMName, normalizeGate } from '../utils/nameUtils';

export interface ParsedExcelData {
  headers: string[];
  rows: RawRow[];
  totalRows: number;
}

export async function parseExcelFile(file: File): Promise<ParsedExcelData> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const jsonData = XLSX.utils.sheet_to_json<RawRow>(sheet, {
    defval: '',
    raw: false,
  });

  const headers =
    jsonData.length > 0
      ? Object.keys(jsonData[0]).filter((h) => h.trim() !== '')
      : [];

  const rows = jsonData.filter((row) =>
    Object.values(row).some((v) => v !== '' && v !== null && v !== undefined)
  );

  return { headers, rows, totalRows: rows.length };
}

const EXCLUDED_PROGRESS = new Set(['cancelled', 'closed']);

function findProgressColumn(rows: RawRow[]): string | null {
  if (rows.length === 0) return null;
  const key = Object.keys(rows[0]).find((k) => k.toLowerCase().trim() === 'progress');
  return key ?? null;
}

export function groupData(
  rows: RawRow[],
  mapping: ColumnMapping,
  selectedPMs: string[]
): PMGroup[] {
  const pmMap = new Map<string, Map<string, Project[]>>();
  const progressCol = findProgressColumn(rows);

  for (const row of rows) {
    // Skip cancelled / closed projects
    if (progressCol) {
      const progress = String(row[progressCol] ?? '').trim().toLowerCase();
      if (EXCLUDED_PROGRESS.has(progress)) continue;
    }

    const rawPM = row[mapping.pmName];
    const pmName = cleanPMName(rawPM !== undefined ? String(rawPM) : '');
    if (!pmName) continue;
    if (selectedPMs.length > 0 && !selectedPMs.includes(pmName)) continue;

    const projectName = String(row[mapping.projectName] ?? '').trim() || 'N/A';
    const projectId   = String(row[mapping.projectId]   ?? '').trim() || 'N/A';
    const gate = normalizeGate(
      row[mapping.gateApproved] !== undefined ? String(row[mapping.gateApproved]) : ''
    );

    if (!pmMap.has(pmName)) pmMap.set(pmName, new Map());
    const gateMap = pmMap.get(pmName)!;
    if (!gateMap.has(gate)) gateMap.set(gate, []);

    const project: Project = { projectName, projectId, gateApproved: gate };
    const existing = gateMap.get(gate)!;
    const isDuplicate = existing.some(
      (p) => p.projectId === projectId && p.projectName === projectName
    );
    if (!isDuplicate) existing.push(project);
  }

  const pmGroups: PMGroup[] = [];

  for (const [pmName, gateMap] of pmMap) {
    const gates: GateGroup[] = [];
    let totalProjects = 0;

    const sortedGates = Array.from(gateMap.keys()).sort((a, b) => {
      if (a === 'Unspecified') return 1;
      if (b === 'Unspecified') return -1;
      return a.localeCompare(b, undefined, { numeric: true });
    });

    for (const gate of sortedGates) {
      const projects = gateMap.get(gate)!;
      gates.push({ gate, projects });
      totalProjects += projects.length;
    }

    pmGroups.push({ pmName, totalProjects, gates });
  }

  pmGroups.sort((a, b) => a.pmName.localeCompare(b.pmName));
  return pmGroups;
}
