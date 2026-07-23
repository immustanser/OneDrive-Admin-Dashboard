import * as React from 'react';
import { RiskLevel } from '../../models';
import styles from './Common.module.scss';

export interface IRiskBadgeProps {
  level: RiskLevel;
}

const RISK_CLASS: Record<RiskLevel, string> = {
  Critical: styles.badgeCritical,
  High: styles.badgeDanger,
  Medium: styles.badgeWarning,
  Low: styles.badgeInfo
};

export const RiskBadge: React.FC<IRiskBadgeProps> = ({ level }) => (
  <span className={`${styles.badge} ${RISK_CLASS[level]}`}>
    <span className={styles.badgeDot} />
    {level}
  </span>
);
