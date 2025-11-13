import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INewsCarouselWebPartProps } from '../NewsCarouselWebPart';

export interface INewsCarouselProps extends INewsCarouselWebPartProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}