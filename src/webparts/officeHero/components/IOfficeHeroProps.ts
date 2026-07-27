/**
 * View-model for the Office Hero component.
 *
 * Data is now list-driven: the web part passes the selected office key and a
 * loader; the component fetches on mount and renders the resolved `IOfficeData`.
 * The individual shapes (`IContact`, `IOfficeFact`, `IFacilitiesNotice`,
 * `IQuickLink`) are unchanged — only how they are populated moved from the
 * property pane to SharePoint lists.
 */

export interface IQuickLink {
  title: string;
  url: string;
  /** A Fluent UI icon name (e.g. "Warning"). Unknown/blank => no icon rendered. */
  iconName?: string;
}

export interface IContact {
  roleLabel: string;
  /** Empty/undefined => the slot renders as "Currently vacant". */
  displayName?: string;
  jobTitle?: string;
  /** Used to build the userphoto.aspx URL. */
  email?: string;
}

export interface IFacilitiesNotice {
  label: string;
  title: string;
  detail?: string;
  ctaText?: string;
  /** Blank => the notice renders without a button rather than a dead link. */
  ctaUrl?: string;
}

/**
 * An "office fact" chip in the "Today at this office" area (nearest station,
 * step-free access, café hours, weather, ...). Icon is optional and validated;
 * an unknown/blank name renders no icon.
 */
export interface IOfficeFact {
  iconName?: string;
  text: string;
}

/**
 * Resolved quick-links result:
 *   IQuickLink[] – zero or more links (empty array = list exists but is empty)
 *   undefined    – the list is missing or the query failed
 */
export type QuickLinksResult = IQuickLink[] | undefined;

/** Everything the hero renders for one office, resolved from the lists. */
export interface IOfficeData {
  officeName: string;
  addressLine?: string;
  openingHours?: string;
  postRoomHours?: string;
  backgroundImageUrl?: string;
  imageAltText?: string;

  facts: IOfficeFact[];

  showNotice: boolean;
  notice?: IFacilitiesNotice;

  /** Always four slots; a slot with no person renders as "Currently vacant". */
  contacts: IContact[];

  quickLinks: QuickLinksResult;
}

/** Result of loading an office: its data, or the sentinel that it wasn't found. */
export type OfficeLoadResult = IOfficeData | 'notFound';

export interface IOfficeHeroProps {
  /** The selected office key (Office Information Title). Undefined => not configured. */
  officeKey?: string;
  /** True in page edit mode (page editors see the "configure / not found" hints). */
  isEditMode: boolean;
  /** Loads everything for the office, once, on mount. Resolves to data or 'notFound'. */
  loadData: (officeKey: string) => Promise<OfficeLoadResult>;
}
