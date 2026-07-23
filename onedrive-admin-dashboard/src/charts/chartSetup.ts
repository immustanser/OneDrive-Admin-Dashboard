import {
  Chart,
  ArcElement,
  BarElement,
  CategoryScale,
  LinearScale,
  LineElement,
  PointElement,
  Tooltip,
  Legend,
  Filler
} from 'chart.js';

let registered = false;

export function ensureChartsRegistered(): void {
  if (registered) {
    return;
  }
  Chart.register(ArcElement, BarElement, CategoryScale, LinearScale, LineElement, PointElement, Tooltip, Legend, Filler);
  registered = true;
}

export const CHART_PALETTE: string[] = ['#0A83AE', '#0078D4', '#2B88D8', '#83D0F5', '#107C10', '#D29200', '#D13438', '#5C2D91'];
