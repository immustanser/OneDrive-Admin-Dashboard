import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IThemeContextValue {
  theme: IReadonlyTheme | undefined;
  isDarkTheme: boolean;
}

const defaultValue: IThemeContextValue = { theme: undefined, isDarkTheme: false };

export const ThemeContext = React.createContext<IThemeContextValue>(defaultValue);

export const ThemeProvider: React.FC<{ theme: IReadonlyTheme | undefined; isDarkTheme: boolean; children: React.ReactNode }> = ({
  theme,
  isDarkTheme,
  children
}) => {
  const value = React.useMemo(() => ({ theme, isDarkTheme }), [theme, isDarkTheme]);
  return <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>;
};

export function useThemeContext(): IThemeContextValue {
  return React.useContext(ThemeContext);
}
