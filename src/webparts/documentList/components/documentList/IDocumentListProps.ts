import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ServiceScope, DisplayMode } from '@microsoft/sp-core-library';

export interface IDocumentListProps {
  /**
   * Web part display mode. Used for inline editing of the web part title
   */
  displayMode: DisplayMode;
  /**
   * The absolute URL of the current web
   */
  siteUrl: string;
  /**
   * Current context Service Scope
   */
  serviceScope: ServiceScope;
  /**
   * The title of the web part
   */
  title: string;
  /**
   * Event handler after updating the web part title
   */
  updateProperty: (value: string) => void;
  /**
   * Document library URL
   */
  doclibUrl: string;
  /**
   * Layout type
   */
  layoutType: string;
  /**
   * Date format
   */
  dateFormat: string;  
  /**
   * Show folder or only files
   */
  showFolder: boolean;
  /**
   * Current context for Configure button
   */
  currentContext: IWebPartContext;
}
