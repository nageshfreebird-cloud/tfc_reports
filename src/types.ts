export type TestPhase = 'Baseline' | 'Midline' | 'Endline';

export interface PhaseStats {
  KNOW: number;
  READ: number;
  SPELL: number;
  CWR: number;
  CWS: number;
  Total: number;
}

export interface ClassStats {
  assessedCount: number;
  baseline: PhaseStats;
  midline: PhaseStats;
  endline: PhaseStats;
}

export interface SchoolData {
  District: string;
  Mandal: string;
  School_Name: string;
  classes: {
    3?: ClassStats;
    4?: ClassStats;
    5?: ClassStats;
  };
}

export interface AppState {
  schools: SchoolData[];
}

export interface PhaseConfig {
  aboutText: string;
  nextStepsTitle: string;
  nextStepsBullets: string[]; // Or paragraph text
}

export interface ReportConfig {
  parameters: {
    KNOW: { label: string; max: number };
    READ: { label: string; max: number };
    SPELL: { label: string; max: number };
    CWR: { label: string; max: number };
    CWS: { label: string; max: number };
  };
  phases: {
    Baseline: PhaseConfig;
    Midline: PhaseConfig;
    Endline: PhaseConfig;
  };
}
