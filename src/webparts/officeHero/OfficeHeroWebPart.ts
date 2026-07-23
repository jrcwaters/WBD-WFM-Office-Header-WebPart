import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OfficeHeroWebPartStrings';
import OfficeHero from './components/OfficeHero';
import type {
  IOfficeHeroProps,
  IQuickLink,
  QuickLinksResult
} from './components/IOfficeHeroProps';

/**
 * Full property set for the web part. The property pane (wired in a later step)
 * writes to these; for now the render path uses hardcoded demo data so the
 * layout can be reviewed before the plumbing goes in.
 */
export interface IOfficeHeroWebPartProps {
  officeName: string;
  addressLine: string;
  openingHours: string;
  postRoomHours: string;
  backgroundImage: string;
  imageAltText: string;

  showNotice: boolean;
  noticeLabel: string;
  noticeTitle: string;
  noticeDetail: string;
  noticeCtaText: string;
  noticeCtaUrl: string;

  contact1RoleLabel: string;
  contact1JobTitle: string;
  contact2RoleLabel: string;
  contact2JobTitle: string;
  contact3RoleLabel: string;
  contact3JobTitle: string;
  contact4RoleLabel: string;
  contact4JobTitle: string;
}

// -----------------------------------------------------------------------------
// STATIC PHASE ONLY — hardcoded demo data. Replaced by property-pane values and
// the SharePoint list query in later steps.
// -----------------------------------------------------------------------------
const DEMO_QUICK_LINKS: IQuickLink[] = [
  { title: 'Report a fault', url: '#', iconName: 'Warning' },
  { title: 'Fire & evacuation', url: '#', iconName: 'Running' },
  { title: 'First aid & wellbeing', url: '#', iconName: 'Health' },
  { title: 'Contact facilities', url: '#', iconName: 'Mail' }
];

export default class OfficeHeroWebPart extends BaseClientSideWebPart<IOfficeHeroWebPartProps> {
  public render(): void {
    const demo: IOfficeHeroProps = {
      officeName: 'London — Fleet Place',
      addressLine: '10 Fleet Place, London EC4M 7RB',
      openingHours: 'Open 08:00–18:00',
      postRoomHours: 'Post room closes 17:30',
      backgroundImageUrl: undefined,
      imageAltText: undefined,

      facts: [
        { iconName: 'Train', text: 'Blackfriars · 5 min' },
        { iconName: 'PartlyCloudyDay', text: '14° · Light cloud' },
        { iconName: 'Wheelchair', text: 'Step-free access' },
        { iconName: 'Coffee', text: 'Café until 16:00' }
      ],

      showNotice: true,
      notice: {
        label: 'Facilities notice',
        title: 'Air-conditioning offline on floors 3–5',
        detail: 'Engineers are on site; expected resolved by 15:00.',
        ctaText: 'Details',
        ctaUrl: '#'
      },

      contacts: [
        { roleLabel: 'Head of office', displayName: 'Priya Nair', jobTitle: 'Office Director', email: 'priya.nair@example.com' },
        { roleLabel: 'Facilities manager', displayName: 'Tom Blake', jobTitle: 'Facilities & Workplace Manager', email: 'tom.blake@example.com' },
        { roleLabel: 'Reception supervisor', displayName: 'Sofia Almeida', jobTitle: 'Front of House Lead', email: 'sofia.almeida@example.com' },
        { roleLabel: 'Post room supervisor', displayName: undefined, jobTitle: undefined, email: undefined }
      ],

      isEditMode: false,
      getQuickLinks: (): Promise<QuickLinksResult> => Promise.resolve(DEMO_QUICK_LINKS)
    };

    const element: React.ReactElement<IOfficeHeroProps> = React.createElement(OfficeHero, demo);
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.OfficeGroupName,
              groupFields: [
                PropertyPaneLabel('setupNotice', {
                  text: strings.StaticPhaseLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
