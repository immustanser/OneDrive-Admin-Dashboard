import * as React from 'react';
import { Line } from 'react-chartjs-2';
import { ensureChartsRegistered } from './chartSetup';
import styles from './ChartCard.module.scss';

ensureChartsRegistered();

export interface ILineChartCardProps {
  title: string;
  labels: string[];
  data: number[];
  label?: string;
}

export const LineChartCard: React.FC<ILineChartCardProps> = ({ title, labels, data, label = 'Storage Used (GB)' }) => (
  <div className={styles.chartCard}>
    <h4 className={styles.chartTitle}>{title}</h4>
    <div className={styles.chartBody}>
      <Line
        data={{
          labels,
          datasets: [{
            label,
            data,
            borderColor: '#0078D4',
            backgroundColor: 'rgba(0, 120, 212, 0.12)',
            fill: true,
            tension: 0.35,
            pointRadius: 3
          }]
        }}
        options={{
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: false } }
        }}
      />
    </div>
  </div>
);
