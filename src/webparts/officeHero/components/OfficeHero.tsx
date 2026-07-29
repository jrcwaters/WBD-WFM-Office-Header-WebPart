import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { getIcon } from '@fluentui/react/lib/Styling';

import styles from './OfficeHero.module.scss';
import type {
  IOfficeHeroProps,
  IOfficeData,
  IContact,
  IQuickLink,
  IOfficeFact,
  OfficeLoadResult
} from './IOfficeHeroProps';

const META_SEPARATOR: string = ' · '; // &nbsp;·&nbsp;

interface ILoadState {
  status: 'loading' | 'ready' | 'notFound' | 'misconfigured';
  data?: IOfficeData;
}

/** Joins only the non-empty meta segments so a stranded separator never renders. */
function buildMetaLine(segments: Array<string | undefined>): string {
  return segments
    .map((s) => (s ? s.trim() : ''))
    .filter((s) => s.length > 0)
    .join(META_SEPARATOR);
}

/**
 * Renders a Fluent icon only when the name is a registered icon; otherwise
 * nothing (so unknown/blank names never show a broken glyph). Icons are muted
 * or theme-neutral — never gold.
 */
function renderIcon(iconName: string | undefined, className: string): JSX.Element {
  if (!iconName || getIcon(iconName) === undefined) {
    return <React.Fragment />;
  }
  return <Icon iconName={iconName} className={className} aria-hidden={true} />;
}

/** A neutral person glyph used when a contact has no usable photo. */
function PersonGlyph(): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <path
        fill="currentColor"
        d="M12 12a5 5 0 1 0-5-5 5 5 0 0 0 5 5Zm0 2c-4.4 0-8 2.2-8 4.9V20h16v-1.1c0-2.7-3.6-4.9-8-4.9Z"
      />
    </svg>
  );
}

/** Office facts — a vertical list in the "Today" column, or a condensed strip beneath the notice. */
function FactList(props: { facts: IOfficeFact[]; condensed?: boolean }): JSX.Element {
  const { facts, condensed } = props;
  if (facts.length === 0) {
    return <React.Fragment />;
  }
  return (
    <div className={condensed ? styles.factsCondensed : styles.factsList}>
      {facts.map((fact: IOfficeFact, i: number) => (
        <span key={i} className={styles.fact}>
          {renderIcon(fact.iconName, styles.factIcon)}
          <span>{fact.text}</span>
        </span>
      ))}
    </div>
  );
}

/** Quick-link tiles — two-per-row in the right column (default), or 4-across when they stand alone. */
function QuickLinkTiles(props: { links: IQuickLink[]; paired: boolean }): JSX.Element {
  const { links, paired } = props;
  return (
    <div className={paired ? styles.quickPair : styles.quickGrid}>
      {links.map((link: IQuickLink, i: number) => (
        <a key={i} className={styles.tile} href={link.url} target="_blank" rel="noreferrer">
          {renderIcon(link.iconName, styles.tileIcon)}
          <span className={styles.tileLabel}>{link.title}</span>
        </a>
      ))}
    </div>
  );
}

/** One contacts column. Owns its own photo-error state so a 404 falls back cleanly. */
function ContactColumn(props: { contact: IContact }): JSX.Element {
  const { roleLabel, displayName, jobTitle, email } = props.contact;
  const [photoFailed, setPhotoFailed] = React.useState<boolean>(false);

  const isVacant: boolean = !displayName || displayName.trim().length === 0;
  const photoUrl: string | undefined =
    email && email.trim().length > 0
      ? `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`
      : undefined;
  const showPhoto: boolean = !!photoUrl && !photoFailed;

  return (
    <div className={styles.contactCol}>
      <h3 className={styles.roleHeading}>{roleLabel}</h3>

      {isVacant ? (
        <p className={styles.contactVacant}>Currently vacant</p>
      ) : (
        <div className={styles.contactPerson}>
          {showPhoto ? (
            <img
              className={styles.contactPhoto}
              src={photoUrl}
              alt=""
              onError={() => setPhotoFailed(true)}
            />
          ) : (
            <span className={styles.contactPhotoPlaceholder} aria-hidden="true">
              <PersonGlyph />
            </span>
          )}
          <div className={styles.contactText}>
            <p className={styles.contactName}>{displayName}</p>
            {jobTitle ? <p className={styles.contactJob}>{jobTitle}</p> : null}
          </div>
        </div>
      )}
    </div>
  );
}

/** Static charcoal placeholder while the office data loads (no motion, no layout jump). */
function HeroSkeleton(): JSX.Element {
  return (
    <section className={styles.officeHero} aria-hidden="true">
      <div className={styles.inner}>
        <div className={styles.card}>
          <div className={styles.hero}>
            <div className={styles.scrim} />
          </div>
          <div className={styles.contacts}>
            <div className={styles.contactsGrid} />
          </div>
        </div>
      </div>
    </section>
  );
}

/** Editor-only message when no office is selected, the row can't be found, or the list is mis-built. */
function ConfigMessage(props: { officeKey?: string; misconfigured?: boolean }): JSX.Element {
  let message: string;
  if (props.misconfigured) {
    message =
      'The Office Information list is missing the columns this web part expects — it looks ' +
      'like it was built by importing a spreadsheet. Recreate the columns with the exact ' +
      'internal names (see provisioning/Manual-Column-Setup.md in the solution).';
  } else if (props.officeKey) {
    message = `The office “${props.officeKey}” was not found in the Office Information list on this site.`;
  } else {
    message = 'Select an office in the property pane to configure this hero. Only page editors see this message.';
  }
  return (
    <section className={styles.officeHero}>
      <div className={styles.inner}>
        <div className={styles.card}>
          <div className={styles.hero}>
            <div className={styles.scrim} />
            <div className={styles.heroInner}>
              <p className={styles.eyebrow}>Our offices</p>
              <p className={styles.editorHint}>{message}</p>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

/** The hero itself, rendered from resolved office data. */
function HeroView(props: { data: IOfficeData; isEditMode: boolean }): JSX.Element {
  const { data, isEditMode } = props;
  const {
    officeName,
    addressLine,
    openingHours,
    postRoomHours,
    backgroundImageUrl,
    imageAltText,
    facts,
    showNotice,
    notice,
    contacts,
    quickLinks
  } = data;

  const metaLine: string = buildMetaLine([addressLine, openingHours, postRoomHours]);
  const heroStyle: React.CSSProperties = backgroundImageUrl
    ? { backgroundImage: `url('${backgroundImageUrl}')` }
    : {};

  const hasNotice: boolean = showNotice && !!notice;
  const hasFacts: boolean = facts.length > 0;
  const linkList: IQuickLink[] | undefined = quickLinks;
  const hasLinks: boolean = Array.isArray(linkList) && linkList.length > 0;
  const listMissing: boolean = linkList === undefined;
  // Missing list is only surfaced to page editors; readers see nothing.
  const showQuickColumn: boolean = hasLinks || (listMissing && isEditMode);
  const leftHasContent: boolean = hasNotice || hasFacts;

  const noticeCtaHref: string | undefined =
    notice && notice.ctaUrl && notice.ctaUrl.trim().length > 0 ? notice.ctaUrl : undefined;

  const quickLinksBlock = (paired: boolean): JSX.Element => (
    <React.Fragment>
      <h2 className={styles.sectionLabel}>Quick links</h2>
      {hasLinks && linkList ? (
        <QuickLinkTiles links={linkList} paired={paired} />
      ) : (
        <p className={styles.editorHint}>
          The &ldquo;Office Quick Links&rdquo; list was not found on this site. Provision it
          (see the solution README) to show quick links here. Only page editors see this message.
        </p>
      )}
    </React.Fragment>
  );

  return (
    <section className={styles.officeHero}>
      <div className={styles.inner}>
        <div className={styles.card}>
          {/* ---- Hero ---- */}
          <div className={styles.hero} style={heroStyle}>
            <div className={styles.scrim} />
            <div className={styles.heroInner}>
              {imageAltText ? (
                <span className={styles.visuallyHidden}>{imageAltText}</span>
              ) : null}

              <p className={styles.eyebrow}>Our offices</p>
              <h1 className={styles.officeName}>{officeName}</h1>
              {metaLine ? <p className={styles.meta}>{metaLine}</p> : null}

              {leftHasContent ? (
                <div className={styles.columns}>
                  <div className={`${styles.colLeft} ${showQuickColumn ? '' : styles.fullWidth}`}>
                    <h2 className={styles.sectionLabel}>Today at this office</h2>
                    {hasNotice && notice ? (
                      <div className={styles.notice}>
                        <div className={styles.noticeBody}>
                          <p className={styles.noticeLabel}>{notice.label}</p>
                          <p className={styles.noticeTitle}>{notice.title}</p>
                          {notice.detail ? <p className={styles.noticeDetail}>{notice.detail}</p> : null}
                        </div>
                        {noticeCtaHref ? (
                          <a
                            className={styles.noticeCta}
                            href={noticeCtaHref}
                            target="_blank"
                            rel="noreferrer"
                          >
                            {notice.ctaText || 'Details'}
                          </a>
                        ) : null}
                      </div>
                    ) : null}
                    {hasFacts ? <FactList facts={facts} condensed={hasNotice} /> : null}
                  </div>

                  {showQuickColumn ? (
                    <div className={styles.colRight}>{quickLinksBlock(true)}</div>
                  ) : null}
                </div>
              ) : showQuickColumn ? (
                <div className={styles.soloQuick}>{quickLinksBlock(false)}</div>
              ) : null}
            </div>
          </div>

          {/* ---- Contacts banner ---- */}
          <div className={styles.contacts}>
            <h2 className={styles.visuallyHidden}>Key contacts</h2>
            <div className={styles.contactsGrid}>
              {contacts.map((contact: IContact, i: number) => (
                <ContactColumn key={i} contact={contact} />
              ))}
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

export default function OfficeHero(props: IOfficeHeroProps): JSX.Element {
  const { officeKey, isEditMode, loadData } = props;
  const [state, setState] = React.useState<ILoadState>({ status: 'loading' });

  // Load the office (once) when the selected office changes.
  React.useEffect((): (() => void) => {
    let active: boolean = true;

    if (!officeKey) {
      setState({ status: 'notFound' });
      return (): void => {
        active = false;
      };
    }

    setState({ status: 'loading' });
    loadData(officeKey)
      .then((result: OfficeLoadResult): void => {
        if (!active) {
          return;
        }
        if (result === 'notFound' || result === 'misconfigured') {
          setState({ status: result });
        } else {
          setState({ status: 'ready', data: result });
        }
      })
      .catch((): void => {
        if (active) {
          setState({ status: 'notFound' });
        }
      });

    return (): void => {
      active = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [officeKey]);

  if (state.status === 'loading') {
    return <HeroSkeleton />;
  }

  if (state.status === 'ready' && state.data) {
    return <HeroView data={state.data} isEditMode={isEditMode} />;
  }

  // 'notFound' or 'misconfigured' — page editors get a helpful message; readers see nothing.
  return isEditMode ? (
    <ConfigMessage officeKey={officeKey} misconfigured={state.status === 'misconfigured'} />
  ) : (
    <React.Fragment />
  );
}
