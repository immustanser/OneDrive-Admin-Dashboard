import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Common.module.scss';

export interface IEmptyStateProps {
  icon?: string;
  title: string;
  description?: string;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({ icon = 'Inbox', title, description }) => (
  <div className={styles.emptyState}>
    <Icon iconName={icon} className={styles.emptyStateIcon} />
    <div style={{ fontWeight: 600 }}>{title}</div>
    {description && <div style={{ fontSize: 13, marginTop: 4 }}>{description}</div>}
  </div>
);
