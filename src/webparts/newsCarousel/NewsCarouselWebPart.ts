import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'NewsCarouselWebPartStrings';
import NewsCarousel from './components/NewsCarousel';
import { INewsCarouselProps } from './components/INewsCarouselProps';

export interface INewsCarouselWebPartProps {
  title?: string;
  itemsToShow?: number;
  showArrows?: boolean;
  autoPlay?: boolean;
  autoPlayInterval?: number;
}

export default class NewsCarouselWebPart extends BaseClientSideWebPart<INewsCarouselWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // Ensure the web part container allows 100% width
    this.domElement.style.width = '100%!important';
    this.domElement.style.maxWidth = 'none';
    
    // Also ensure parent containers don't constrain width
    let parent = this.domElement.parentElement;
    while (parent) {
      // Check if it's the ms-SPLegacyFabricBlock or other SharePoint containers
      if (parent.classList.contains('ms-SPLegacyFabricBlock') || 
          parent.classList.contains('CanvasSection') ||
          parent.classList.contains('ControlZone') ||
          parent.hasAttribute('data-sp-webpart')) {
        parent.style.width = '100%!important';
        parent.style.maxWidth = 'none';
      }
      parent = parent.parentElement;
    }
    
    const element: React.ReactElement<INewsCarouselProps> = React.createElement(
      NewsCarousel,
      {
        title: this.properties.title || '',
        itemsToShow: this.properties.itemsToShow || 3,
        showArrows: this.properties.showArrows !== undefined ? this.properties.showArrows : true,
        autoPlay: this.properties.autoPlay !== undefined ? this.properties.autoPlay : false,
        autoPlayInterval: this.properties.autoPlayInterval || 5000,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    // Keep this stable - only increment if you need to migrate old property structures
    // Since all properties are optional with defaults, no version bump needed for normal changes
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title'
                }),
                PropertyPaneSlider('itemsToShow', {
                  label: 'Number of items to show',
                  min: 3,
                  max: 5,
                  value: this.properties.itemsToShow || 3
                }),
                PropertyPaneToggle('showArrows', {
                  label: 'Show navigation arrows'
                }),
                PropertyPaneToggle('autoPlay', {
                  label: 'Auto-play carousel'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}