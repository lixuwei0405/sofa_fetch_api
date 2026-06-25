export interface TaskRow {
  id: string;
  originalIndex: number;
  status: 'pending' | 'running' | 'success' | 'failed';
  report?: string;
}
