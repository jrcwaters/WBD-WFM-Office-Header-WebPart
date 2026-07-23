/**
 * View-model for the Office Hero component. The web part maps property-pane
 * values and the quick-links list query onto these shapes; the component itself
 * stays presentational (plus the one on-mount quick-links fetch it owns).
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
 * Result of the quick-links query:
 *   IQuickLink[] – zero or more links (empty array = list exists but is empty)
 *   undefined    – the list is missing or the query failed
 */
export type QuickLinksResult = IQuickLink[] | undefined;

export interface IOfficeHeroProps {
  // --- Office -------------------------------------------------------------
  officeName: string;
  addressLine?: string;
  openingHours?: string;
  postRoomHours?: string;
  backgroundImageUrl?: string;
  imageAltText?: string;

  // --- Facilities notice --------------------------------------------------
  showNotice: boolean;
  notice?: IFacilitiesNotice;

  // --- Key contacts (always four slots) -----------------------------------
  contacts: IContact[];

  // --- Quick links --------------------------------------------------------
  /** True when the page is in edit mode (page editors see the "list missing" hint). */
  isEditMode: boolean;
  /** Fetched once on mount and cached in component state. */
  getQuickLinks: () => Promise<QuickLinksResult>;
}
