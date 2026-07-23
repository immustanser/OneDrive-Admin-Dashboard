import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from './Common.module.scss';

export interface IErrorStateProps {
  message: string;
  onRetry?: () => void;
}

export const ErrorState: React.FC<IErrorStateProps> = ({ message, onRetry }) => (
  <div className={styles.errorState}>
    <Icon iconName="ErrorBadge" style={{ fontSize: 32 }} />
    <div style={{ fontWeight: 600 }}>Something went wrong</div>
    <div style={{ fontSize: 13 }}>{message}</div>
    {onRetry && <DefaultButton text="Retry" onClick={onRetry} />}
  </div>
);
