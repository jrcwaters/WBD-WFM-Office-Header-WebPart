import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import * as strings from 'OfficeHeroWebPartStrings';
import OfficeHero from './components/OfficeHero';
import type {
  IOfficeHeroProps,
  IOfficeData,
  IOfficeFact,
  IContact,
  IQuickLink,
  IFacilitiesNotice,
  OfficeLoadResult,
  QuickLinksResult
} from './components/IOfficeHeroProps';

const OFFICE_INFO_LIST: string = 'Office Information';
const OFFICE_FACTS_LIST: string = 'Office Facts';
const OFFICE_PEOPLE_LIST: string = 'Office People';
const QUICK_LINKS_LIST: string = 'Office Quick Links';

/** The four fixed contact roles (also the choices on the Office People list's Role column). */
const CONTACT_ROLES: string[] = [
  'Head of office',
  'Facilities manager',
  'Reception supervisor',
  'Post room supervisor'
];

export interface IOfficeHeroWebPartProps {
  /** Title of the selected row in the Office Information list. */
  officeKey: string;
}

// --- Raw list item shapes (only the selected fields) -------------------------
interface IOfficeInfoItem {
  Id: number;
  Title: string;
  AddressLine?: string;
  OpeningHours?: string;
  PostRoomHours?: string;
  BackgroundImage?: { Url?: string };
  ImageAltText?: string;
  ShowNotice?: boolean;
  NoticeLabel?: string;
  NoticeTitle?: string;
  NoticeDetail?: string;
  NoticeCtaText?: string;
  NoticeCtaUrl?: { Url?: string };
}
interface IFactItem {
  Title: string;
  IconName?: string;
  SortOrder?: number;
}
interface IPersonValue {
  Title?: string;
  EMail?: string;
  Name?: string;
}
interface IPeopleItem {
  Role?: string;
  JobTitle?: string;
  Person?: IPersonValue;
}
interface IQuickLinkItem {
  Title: string;
  LinkUrl?: { Url?: string };
  IconName?: string;
  SortOrder?: number;
}

interface IResolvedOfficeInfo {
  id: number;
  officeName: string;
  addressLine?: string;
  openingHours?: string;
  postRoomHours?: string;
  backgroundImageUrl?: string;
  imageAltText?: string;
  showNotice: boolean;
  notice?: IFacilitiesNotice;
}

export default class OfficeHeroWebPart extends BaseClientSideWebPart<IOfficeHeroWebPartProps> {
  private _sp: SPFI;
  private _officeOptions: IPropertyPaneDropdownOption[] = [];
  private _officeOptionsLoaded: boolean = false;

  protected onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IOfficeHeroProps> = React.createElement(OfficeHero, {
      officeKey: this.properties.officeKey || undefined,
      isEditMode: this.displayMode === DisplayMode.Edit,
      loadData: this._loadData.bind(this)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // --- Data loading ----------------------------------------------------------

  /** Loads everything the hero needs for one office. Resolves to data or 'notFound'. */
  private async _loadData(officeKey: string): Promise<OfficeLoadResult> {
    const info: IResolvedOfficeInfo | undefined = await this._getOfficeInfo(officeKey);
    if (!info) {
      return 'notFound';
    }

    const [facts, contacts, quickLinks]: [IOfficeFact[], IContact[], QuickLinksResult] =
      await Promise.all([
        this._getFacts(info.id),
        this._getContacts(info.id),
        this._getQuickLinks(info.id)
      ]);

    const data: IOfficeData = {
      officeName: info.officeName,
      addressLine: info.addressLine,
      openingHours: info.openingHours,
      postRoomHours: info.postRoomHours,
      backgroundImageUrl: info.backgroundImageUrl,
      imageAltText: info.imageAltText,
      facts,
      showNotice: info.showNotice,
      notice: info.notice,
      contacts,
      quickLinks
    };
    return data;
  }

  private async _getOfficeInfo(officeKey: string): Promise<IResolvedOfficeInfo | undefined> {
    try {
      const items: IOfficeInfoItem[] = await this._sp.web.lists
        .getByTitle(OFFICE_INFO_LIST)
        .items.select(
          'Id',
          'Title',
          'AddressLine',
          'OpeningHours',
          'PostRoomHours',
          'BackgroundImage',
          'ImageAltText',
          'ShowNotice',
          'NoticeLabel',
          'NoticeTitle',
          'NoticeDetail',
          'NoticeCtaText',
          'NoticeCtaUrl'
        )
        .filter(`Title eq '${officeKey.replace(/'/g, "''")}'`)
        .top(1)();

      if (!items || items.length === 0) {
        return undefined;
      }
      const it: IOfficeInfoItem = items[0];
      const showNotice: boolean = !!it.ShowNotice;
      return {
        id: it.Id,
        officeName: it.Title,
        addressLine: it.AddressLine || undefined,
        openingHours: it.OpeningHours || undefined,
        postRoomHours: it.PostRoomHours || undefined,
        backgroundImageUrl: (it.BackgroundImage && it.BackgroundImage.Url) || undefined,
        imageAltText: it.ImageAltText || undefined,
        showNotice,
        notice: showNotice
          ? {
              label: it.NoticeLabel || 'Facilities notice',
              title: it.NoticeTitle || '',
              detail: it.NoticeDetail || undefined,
              ctaText: it.NoticeCtaText || undefined,
              ctaUrl: (it.NoticeCtaUrl && it.NoticeCtaUrl.Url) || undefined
            }
          : undefined
      };
    } catch (error) {
      // Missing list or query failure: treat as "office not found" (editor sees a hint).
      if (!this._isListMissing(error)) {
        // eslint-disable-next-line no-console
        console.error('[Office Hero] Office Information query failed.', error);
      }
      return undefined;
    }
  }

  private async _getFacts(officeId: number): Promise<IOfficeFact[]> {
    try {
      const items: IFactItem[] = await this._sp.web.lists
        .getByTitle(OFFICE_FACTS_LIST)
        .items.select('Title', 'IconName', 'SortOrder')
        .filter(`OfficeId eq ${officeId}`)
        .orderBy('SortOrder', true)();

      return items
        .map((i: IFactItem): IOfficeFact => ({ iconName: i.IconName || undefined, text: i.Title }))
        .filter((f: IOfficeFact) => !!f.text && f.text.trim().length > 0);
    } catch (error) {
      if (!this._isListMissing(error)) {
        // eslint-disable-next-line no-console
        console.error('[Office Hero] Office Facts query failed.', error);
      }
      return [];
    }
  }

  private async _getContacts(officeId: number): Promise<IContact[]> {
    const byRole: { [role: string]: IPeopleItem } = {};
    try {
      const items: IPeopleItem[] = await this._sp.web.lists
        .getByTitle(OFFICE_PEOPLE_LIST)
        .items.select('Role', 'JobTitle', 'Person/Title', 'Person/EMail', 'Person/Name')
        .expand('Person')
        .filter(`OfficeId eq ${officeId}`)();

      for (const it of items) {
        // First person wins per role slot.
        if (it.Role && !byRole[it.Role]) {
          byRole[it.Role] = it;
        }
      }
    } catch (error) {
      if (!this._isListMissing(error)) {
        // eslint-disable-next-line no-console
        console.error('[Office Hero] Office People query failed.', error);
      }
      // Fall through: every slot renders as vacant.
    }

    return CONTACT_ROLES.map((role: string): IContact => {
      const row: IPeopleItem | undefined = byRole[role];
      const person: IPersonValue | undefined = row ? row.Person : undefined;
      return {
        roleLabel: role,
        displayName: person ? person.Title : undefined,
        jobTitle: (row && row.JobTitle) || undefined,
        email: person ? this._personAccount(person) : undefined
      };
    });
  }

  private async _getQuickLinks(officeId: number): Promise<QuickLinksResult> {
    try {
      const items: IQuickLinkItem[] = await this._sp.web.lists
        .getByTitle(QUICK_LINKS_LIST)
        .items.select('Title', 'LinkUrl', 'IconName', 'SortOrder')
        .filter(`OfficeId eq ${officeId}`)
        .orderBy('SortOrder', true)
        .top(4)();

      return items
        .map((item: IQuickLinkItem): IQuickLink => ({
          title: item.Title,
          url: (item.LinkUrl && item.LinkUrl.Url) || '',
          iconName: item.IconName
        }))
        .filter((link: IQuickLink) => !!link.title && !!link.url);
    } catch (error) {
      if (this._isListMissing(error)) {
        // Missing list: surfaced to page editors only (handled by the component).
        return undefined;
      }
      // eslint-disable-next-line no-console
      console.error('[Office Hero] Quick links query failed.', error);
      return [];
    }
  }

  /** Best account name for userphoto.aspx: the email, else the login's UPN part. */
  private _personAccount(person: IPersonValue): string | undefined {
    if (person.EMail && person.EMail.trim().length > 0) {
      return person.EMail;
    }
    if (person.Name && person.Name.indexOf('|') !== -1) {
      return person.Name.substring(person.Name.lastIndexOf('|') + 1);
    }
    return person.Name || undefined;
  }

  private _isListMissing(error: unknown): boolean {
    const status: number | undefined = (error as { status?: number }).status;
    const message: string = (error as { message?: string }).message || '';
    return status === 404 || /does not exist|not found/i.test(message);
  }

  // --- Property pane ---------------------------------------------------------

  /** Populate the office dropdown from the Office Information list when the pane opens. */
  protected onPropertyPaneConfigurationStart(): void {
    if (this._officeOptionsLoaded) {
      return;
    }
    this._getOfficeOptions()
      .then((options: IPropertyPaneDropdownOption[]): void => {
        this._officeOptions = options;
        this._officeOptionsLoaded = true;
        this.context.propertyPane.refresh();
      })
      .catch((): void => {
        this._officeOptionsLoaded = true;
      });
  }

  private async _getOfficeOptions(): Promise<IPropertyPaneDropdownOption[]> {
    try {
      const items: Array<{ Title: string }> = await this._sp.web.lists
        .getByTitle(OFFICE_INFO_LIST)
        .items.select('Title')
        .orderBy('Title', true)
        .top(500)();
      return items
        .filter((i) => !!i.Title)
        .map((i): IPropertyPaneDropdownOption => ({ key: i.Title, text: i.Title }));
    } catch (error) {
      if (!this._isListMissing(error)) {
        // eslint-disable-next-line no-console
        console.error('[Office Hero] Could not load the office list for the picker.', error);
      }
      return [];
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // Re-render so the component reloads for the newly selected office.
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const noOffices: boolean = this._officeOptions.length === 0;
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneDropdown('officeKey', {
                  label: strings.OfficeSelectLabel,
                  options: this._officeOptions,
                  disabled: noOffices
                }),
                PropertyPaneLabel('officeHint', {
                  text: noOffices ? strings.NoOfficesHint : strings.OfficeSelectDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
