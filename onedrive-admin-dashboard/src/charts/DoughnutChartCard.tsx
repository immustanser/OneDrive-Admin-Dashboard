import * as React from 'react';
import { Doughnut } from 'react-chartjs-2';
import { ensureChartsRegistered } from './chartSetup';
import styles from './ChartCard.module.scss';

ensureChartsRegistered();

export interface IDoughnutChartCardProps {
  title: string;
  labels: string[];
  data: number[];
  colors?: string[];
}

export const DoughnutChartCard: React.FC<IDoughnutChartCardProps> = ({ title, labels, data, colors }) => (
  <div className={styles.chartCard}>
    <h4 className={styles.chartTitle}>{title}</h4>
    <div className={styles.chartBody}>
      <Doughnut
        data={{
          labels,
          datasets: [{ data, backgroundColor: colors ?? ['#107C10', '#D13438'], borderWidth: 0 }]
        }}
        options={{
          responsive: true,
          maintainAspectRatio: false,
          cutout: '65%',
          plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } } }
        }}
      />
    </div>
  </div>
);
