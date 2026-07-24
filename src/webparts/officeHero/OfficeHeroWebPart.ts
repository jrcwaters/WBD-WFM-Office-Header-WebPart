import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import {
  PropertyFieldPeoplePicker,
  PrincipalType,
  IPropertyFieldGroupOrPerson
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

import * as strings from 'OfficeHeroWebPartStrings';
import OfficeHero from './components/OfficeHero';
import type {
  IOfficeHeroProps,
  IOfficeFact,
  IContact,
  IQuickLink,
  QuickLinksResult
} from './components/IOfficeHeroProps';

const QUICK_LINKS_LIST_TITLE: string = 'Office Quick Links';

export interface IOfficeHeroWebPartProps {
  officeName: string;
  addressLine: string;
  openingHours: string;
  postRoomHours: string;
  backgroundImage: string;
  imageAltText: string;

  facts: IOfficeFact[];

  showNotice: boolean;
  noticeLabel: string;
  noticeTitle: string;
  noticeDetail: string;
  noticeCtaText: string;
  noticeCtaUrl: string;

  contact1RoleLabel: string;
  contact1Person: IPropertyFieldGroupOrPerson[];
  contact1JobTitle: string;
  contact2RoleLabel: string;
  contact2Person: IPropertyFieldGroupOrPerson[];
  contact2JobTitle: string;
  contact3RoleLabel: string;
  contact3Person: IPropertyFieldGroupOrPerson[];
  contact3JobTitle: string;
  contact4RoleLabel: string;
  contact4Person: IPropertyFieldGroupOrPerson[];
  contact4JobTitle: string;
}

/** Raw shape of an "Office Quick Links" list item (only the selected fields). */
interface IQuickLinkListItem {
  Title: string;
  LinkUrl?: { Url?: string; Description?: string };
  IconName?: string;
  SortOrder?: number;
}

export default class OfficeHeroWebPart extends BaseClientSideWebPart<IOfficeHeroWebPartProps> {
  private _sp: SPFI;
  private _backgroundImageResult: IFilePickerResult | undefined;

  protected onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    if (this.properties.backgroundImage) {
      this._backgroundImageResult = {
        fileAbsoluteUrl: this.properties.backgroundImage,
        fileName: '',
        fileNameWithoutExtension: ''
      } as IFilePickerResult;
    }
    return super.onInit();
  }

  public render(): void {
    const showNotice: boolean = !!this.properties.showNotice;

    const viewProps: IOfficeHeroProps = {
      officeName: this.properties.officeName || 'Our office',
      addressLine: this.properties.addressLine,
      openingHours: this.properties.openingHours,
      postRoomHours: this.properties.postRoomHours,
      backgroundImageUrl: this.properties.backgroundImage || undefined,
      imageAltText: this.properties.imageAltText || undefined,

      facts: (this.properties.facts || []).filter((f: IOfficeFact) => f && f.text && f.text.trim().length > 0),

      showNotice,
      notice: showNotice
        ? {
            label: this.properties.noticeLabel || strings.NoticeLabelLabel,
            title: this.properties.noticeTitle,
            detail: this.properties.noticeDetail,
            ctaText: this.properties.noticeCtaText,
            ctaUrl: this.properties.noticeCtaUrl
          }
        : undefined,

      contacts: [
        this._buildContact(this.properties.contact1RoleLabel, this.properties.contact1Person, this.properties.contact1JobTitle),
        this._buildContact(this.properties.contact2RoleLabel, this.properties.contact2Person, this.properties.contact2JobTitle),
        this._buildContact(this.properties.contact3RoleLabel, this.properties.contact3Person, this.properties.contact3JobTitle),
        this._buildContact(this.properties.contact4RoleLabel, this.properties.contact4Person, this.properties.contact4JobTitle)
      ],

      isEditMode: this.displayMode === DisplayMode.Edit,
      getQuickLinks: this._getQuickLinks.bind(this)
    };

    const element: React.ReactElement<IOfficeHeroProps> = React.createElement(OfficeHero, viewProps);
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * The property-control packages bundle slightly different @microsoft/sp-*
   * typings, so WebPartContext is not structurally identical to the
   * BaseComponentContext they expect. Bridge the type skew in one place.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private get _controlContext(): any {
    return this.context;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // --- Data mapping ----------------------------------------------------------

  private _buildContact(
    roleLabel: string,
    person: IPropertyFieldGroupOrPerson[] | undefined,
    jobTitle: string
  ): IContact {
    const first: IPropertyFieldGroupOrPerson | undefined =
      person && person.length > 0 ? person[0] : undefined;
    return {
      roleLabel: roleLabel || '',
      displayName: first ? first.fullName : undefined,
      // Job title is a deliberate manual field (not the Entra job title).
      jobTitle: jobTitle || undefined,
      email: first ? this._accountName(first) : undefined
    };
  }

  /** Best account name for userphoto.aspx: the email, else the login's UPN part. */
  private _accountName(person: IPropertyFieldGroupOrPerson): string | undefined {
    if (person.email && person.email.trim().length > 0) {
      return person.email;
    }
    if (person.login && person.login.indexOf('|') !== -1) {
      return person.login.substring(person.login.lastIndexOf('|') + 1);
    }
    return person.login;
  }

  // --- Quick links query (PnPjs) ---------------------------------------------

  private async _getQuickLinks(): Promise<QuickLinksResult> {
    try {
      const items: IQuickLinkListItem[] = await this._sp.web.lists
        .getByTitle(QUICK_LINKS_LIST_TITLE)
        .items.select('Title', 'LinkUrl', 'IconName', 'SortOrder')
        .orderBy('SortOrder', true)
        .top(4)();

      return items
        .map((item: IQuickLinkListItem): IQuickLink => ({
          title: item.Title,
          url: (item.LinkUrl && item.LinkUrl.Url) || '',
          iconName: item.IconName
        }))
        .filter((link: IQuickLink) => link.title && link.url);
    } catch (error) {
      if (this._isListMissing(error)) {
        // Missing list: surfaced to page editors only (handled by the component).
        return undefined;
      }
      // Any other failure: log and omit the section entirely (readers see nothing).
      // eslint-disable-next-line no-console
      console.error('[Office Hero] Quick links query failed.', error);
      return [];
    }
  }

  private _isListMissing(error: unknown): boolean {
    const status: number | undefined = (error as { status?: number }).status;
    const message: string = (error as { message?: string }).message || '';
    return status === 404 || /does not exist|not found/i.test(message);
  }

  // --- Property pane ---------------------------------------------------------

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // Toggling the notice enables/disables the notice fields, so refresh the pane.
    if (propertyPath === 'showNotice') {
      this.context.propertyPane.refresh();
    }
    this.render();
  }

  private _validateUrl(value: string): string {
    if (!value || value.trim().length === 0) {
      return '';
    }
    try {
      const parsed: URL = new URL(value);
      return parsed.protocol === 'http:' || parsed.protocol === 'https:' ? '' : strings.FieldInvalidUrl;
    } catch {
      return strings.FieldInvalidUrl;
    }
  }

  private _contactFields(
    index: number,
    header: string,
    person: IPropertyFieldGroupOrPerson[]
  ): IPropertyPaneField<unknown>[] {
    return [
      PropertyPaneLabel(`contact${index}Header`, { text: header }),
      PropertyPaneTextField(`contact${index}RoleLabel`, { label: strings.ContactRoleLabelLabel }),
      PropertyFieldPeoplePicker(`contact${index}Person`, {
        label: strings.ContactPersonLabel,
        initialData: person || [],
        allowDuplicate: false,
        multiSelect: false,
        principalType: [PrincipalType.Users],
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        context: this._controlContext,
        properties: this.properties,
        key: `contact${index}PersonId`,
        searchTextLimit: 4,
        deferredValidationTime: 0
      }),
      PropertyPaneTextField(`contact${index}JobTitle`, { label: strings.ContactJobTitleLabel })
    ] as IPropertyPaneField<unknown>[];
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const noticeOff: boolean = !this.properties.showNotice;

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.OfficeGroupName,
              groupFields: [
                PropertyPaneTextField('officeName', {
                  label: strings.OfficeNameLabel,
                  onGetErrorMessage: (value: string): string =>
                    value && value.trim().length > 0 ? '' : strings.FieldRequired
                }),
                PropertyPaneTextField('addressLine', { label: strings.AddressLineLabel }),
                PropertyPaneTextField('openingHours', { label: strings.OpeningHoursLabel }),
                PropertyPaneTextField('postRoomHours', { label: strings.PostRoomHoursLabel }),
                PropertyFieldFilePicker('backgroundImage', {
                  context: this._controlContext,
                  properties: this.properties,
                  key: 'backgroundImageId',
                  label: strings.BackgroundImageLabel,
                  buttonLabel: strings.ChooseImageButton,
                  filePickerResult: this._backgroundImageResult as IFilePickerResult,
                  accepts: ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'],
                  onSave: (result: IFilePickerResult): void => {
                    this.properties.backgroundImage = result.fileAbsoluteUrl || '';
                    this._backgroundImageResult = result;
                    this.render();
                  },
                  onChanged: (result: IFilePickerResult): void => {
                    this._backgroundImageResult = result;
                  },
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this)
                }),
                PropertyPaneTextField('imageAltText', {
                  label: strings.ImageAltTextLabel,
                  description: strings.ImageAltTextDescription
                })
              ]
            },
            {
              groupName: strings.OfficeFactsGroupName,
              groupFields: [
                PropertyFieldCollectionData('facts', {
                  key: 'facts',
                  label: strings.FactsLabel,
                  panelHeader: strings.FactsPanelHeader,
                  panelDescription: strings.FactsPanelDescription,
                  manageBtnLabel: strings.FactsManageButton,
                  value: this.properties.facts,
                  enableSorting: true,
                  fields: [
                    {
                      id: 'iconName',
                      title: strings.FactsIconColumn,
                      type: CustomCollectionFieldType.fabricIcon
                    },
                    {
                      id: 'text',
                      title: strings.FactsTextColumn,
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ]
                })
              ]
            },
            {
              groupName: strings.FacilitiesNoticeGroupName,
              groupFields: [
                PropertyPaneToggle('showNotice', { label: strings.ShowNoticeLabel }),
                PropertyPaneTextField('noticeLabel', { label: strings.NoticeLabelLabel, disabled: noticeOff }),
                PropertyPaneTextField('noticeTitle', { label: strings.NoticeTitleLabel, disabled: noticeOff }),
                PropertyPaneTextField('noticeDetail', {
                  label: strings.NoticeDetailLabel,
                  multiline: true,
                  disabled: noticeOff
                }),
                PropertyPaneTextField('noticeCtaText', { label: strings.NoticeCtaTextLabel, disabled: noticeOff }),
                PropertyPaneTextField('noticeCtaUrl', {
                  label: strings.NoticeCtaUrlLabel,
                  disabled: noticeOff,
                  onGetErrorMessage: (value: string): string => this._validateUrl(value)
                })
              ]
            },
            {
              groupName: strings.KeyContactsGroupName,
              groupFields: [
                ...this._contactFields(1, strings.Contact1Header, this.properties.contact1Person),
                PropertyPaneHorizontalRule(),
                ...this._contactFields(2, strings.Contact2Header, this.properties.contact2Person),
                PropertyPaneHorizontalRule(),
                ...this._contactFields(3, strings.Contact3Header, this.properties.contact3Person),
                PropertyPaneHorizontalRule(),
                ...this._contactFields(4, strings.Contact4Header, this.properties.contact4Person)
              ]
            }
          ]
        }
      ]
    };
  }
}
