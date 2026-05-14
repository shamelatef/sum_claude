export interface RawRow {
  [key: string]: string | number | boolean | null | undefined;
}

export interface ColumnMapping {
  pmName: string;
  projectName: string;
  projectId: string;
  gateApproved: string;
}

export interface Project {
  projectName: string;
  projectId: string;
  gateApproved: string;
}

export interface GateGroup {
  gate: string;
  projects: Project[];
}

export interface PMGroup {
  pmName: string;
  totalProjects: number;
  gates: GateGroup[];
}

export type Step = 'upload' | 'map' | 'select' | 'generate';
