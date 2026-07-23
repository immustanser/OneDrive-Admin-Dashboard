import * as React from 'react';
import { Pie } from 'react-chartjs-2';
import { ensureChartsRegistered, CHART_PALETTE } from './chartSetup';
import styles from './ChartCard.module.scss';

ensureChartsRegistered();

export interface IPieChartCardProps {
  title: string;
  labels: string[];
  data: number[];
}

export const PieChartCard: React.FC<IPieChartCardProps> = ({ title, labels, data }) => (
  <div className={styles.chartCard}>
    <h4 className={styles.chartTitle}>{title}</h4>
    <div className={styles.chartBody}>
      <Pie
        data={{
          labels,
          datasets: [{ data, backgroundColor: CHART_PALETTE, borderWidth: 0 }]
        }}
        options={{
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } } }
        }}
      />
    </div>
  </div>
);
