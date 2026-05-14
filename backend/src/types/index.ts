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

export interface ParsedExcelData {
  headers: string[];
  rows: RawRow[];
}

export interface GenerateRequest {
  rows: RawRow[];
  columnMapping: ColumnMapping;
  selectedPMs: string[];
}

export interface SlideConfig {
  width: number;
  height: number;
  margin: {
    top: number;
    bottom: number;
    left: number;
    right: number;
  };
  header: {
    height: number;
  };
  footer: {
    height: number;
  };
}
