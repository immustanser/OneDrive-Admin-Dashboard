import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Common.module.scss';

export interface ISectionCardProps {
  title: string;
  icon?: string;
  actions?: React.ReactNode;
  children: React.ReactNode;
  className?: string;
}

export const SectionCard: React.FC<ISectionCardProps> = ({ title, icon, actions, children, className }) => (
  <div className={`${styles.card} ${className ?? ''}`}>
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
      <h3 className={styles.sectionTitle}>
        {icon && <Icon iconName={icon} />}
        {title}
      </h3>
      {actions}
    </div>
    {children}
  </div>
);
