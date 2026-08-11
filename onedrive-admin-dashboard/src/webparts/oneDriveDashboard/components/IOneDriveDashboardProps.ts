import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IOneDriveDashboardProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  theme: IReadonlyTheme | undefined;
  useMockData: boolean;
  apiBaseUrl: string;
  apiResourceUri: string;
}
