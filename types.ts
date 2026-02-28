
export enum DefectStatus {
  PENDING = 'PENDING',
  PROCESSING = 'PROCESSING',
  COMPLETED = 'COMPLETED',
  URGENT = 'URGENT'
}

export interface Activity {
  id: string;
  title: string;
  category: string;
  timeLabel: string;
  status: DefectStatus;
  type: 'problem' | 'repair' | 'safety';
}

export interface ChartData {
  name: string;
  detected: number;
  processed: number;
  operators: number;
}
