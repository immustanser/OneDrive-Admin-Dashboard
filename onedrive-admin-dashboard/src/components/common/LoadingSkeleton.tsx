import * as React from 'react';
import styles from './Common.module.scss';

export interface ILoadingSkeletonProps {
  height?: number | string;
  width?: number | string;
  borderRadius?: number | string;
  className?: string;
}

export const LoadingSkeleton: React.FC<ILoadingSkeletonProps> = ({ height = 16, width = '100%', borderRadius, className }) => (
  <div
    className={`${styles.skeleton} ${className ?? ''}`}
    style={{ height, width, borderRadius }}
  />
);

export const KpiCardSkeleton: React.FC = () => (
  <div className={styles.card} style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
    <LoadingSkeleton height={14} width="50%" />
    <LoadingSkeleton height={28} width="70%" />
    <LoadingSkeleton height={10} width="40%" />
  </div>
);

export const ChartCardSkeleton: React.FC = () => (
  <div className={styles.card} style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
    <LoadingSkeleton height={16} width="40%" />
    <LoadingSkeleton height={220} borderRadius={8} />
  </div>
);
