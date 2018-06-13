import { IWebPartContext } from '@microsoft/sp-webpart-base';
export interface IConfigContainerProps {
  currentContext: IWebPartContext;
  iconText?: string;
  description?: string;
  buttonText?: string;
  displayButton?: boolean;
}
