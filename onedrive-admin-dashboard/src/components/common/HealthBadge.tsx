import * as React from 'react';
import { HealthLevel } from '../../utils/health';
import styles from './Common.module.scss';

export interface IHealthBadgeProps {
  level: HealthLevel;
}

const HEALTH_CLASS: Record<HealthLevel, string> = {
  Healthy: styles.badgeSuccess,
  Warning: styles.badgeWarning,
  'High Risk': styles.badgeOrange,
  Critical: styles.badgeCritical
};

export const HealthBadge: React.FC<IHealthBadgeProps> = ({ level }) => (
  <span className={`${styles.badge} ${HEALTH_CLASS[level]}`}>
    <span className={styles.badgeDot} />
    {level}
  </span>
);
