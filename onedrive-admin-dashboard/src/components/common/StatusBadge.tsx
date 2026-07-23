import * as React from 'react';
import { OneDriveStatus } from '../../models';
import styles from './Common.module.scss';

export interface IStatusBadgeProps {
  status: OneDriveStatus;
}

const STATUS_CLASS: Record<OneDriveStatus, string> = {
  Active: styles.badgeSuccess,
  Warning: styles.badgeWarning,
  Inactive: styles.badgeDanger
};

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => (
  <span className={`${styles.badge} ${STATUS_CLASS[status]}`}>
    <span className={styles.badgeDot} />
    {status}
  </span>
);
