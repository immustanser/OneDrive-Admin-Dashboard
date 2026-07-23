import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IKpiCardData } from '../../models';
import styles from './KpiCard.module.scss';

const STATUS_CLASS: Record<IKpiCardData['status'], string> = {
  success: styles.statusSuccess,
  warning: styles.statusWarning,
  danger: styles.statusDanger,
  info: styles.statusInfo,
  neutral: styles.statusNeutral
};

export const KpiCard: React.FC<{ data: IKpiCardData }> = ({ data }) => (
  <div className={`${styles.kpiCard} ${STATUS_CLASS[data.status]}`}>
    <div className={styles.iconWrap}>
      <Icon iconName={data.icon} />
    </div>
    <div className={styles.body}>
      <span className={styles.title}>{data.title}</span>
      <span className={styles.value}>{data.value}</span>
      {data.subtitle && <span className={styles.subtitle}>{data.subtitle}</span>}
    </div>
  </div>
);
