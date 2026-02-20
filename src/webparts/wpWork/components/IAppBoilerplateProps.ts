/**
 * IAppBoilerplateProps.ts
 * Props interface for AppBoilerplate component
 */

import { IWebPartContext } from '@microsoft/sp-webpart-base';

export interface IAppBoilerplateProps {
  /**
   * SPFx web part context for accessing SharePoint APIs
   */
  context: IWebPartContext;

  /**
   * Dark theme flag from SPFx
   */
  isDarkTheme: boolean;

  /**
   * Current user's display name
   */
  userDisplayName: string;
}
