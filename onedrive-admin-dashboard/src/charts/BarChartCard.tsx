import * as React from 'react';
import { Bar } from 'react-chartjs-2';
import { ensureChartsRegistered } from './chartSetup';
import styles from './ChartCard.module.scss';

ensureChartsRegistered();

export interface IBarChartCardProps {
  title: string;
  labels: string[];
  data: number[];
  label?: string;
  horizontal?: boolean;
}

export const BarChartCard: React.FC<IBarChartCardProps> = ({ title, labels, data, label = 'Storage Used (GB)', horizontal = true }) => (
  <div className={styles.chartCard}>
    <h4 className={styles.chartTitle}>{title}</h4>
    <div className={styles.chartBody}>
      <Bar
        data={{
          labels,
          datasets: [{ label, data, backgroundColor: '#0A83AE', borderRadius: 4, maxBarThickness: 22 }]
        }}
        options={{
          indexAxis: horizontal ? 'y' : 'x',
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { grid: { display: !horizontal } },
            y: { grid: { display: horizontal } }
          }
        }}
      />
    </div>
  </div>
);
