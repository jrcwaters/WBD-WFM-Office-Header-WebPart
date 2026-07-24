import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { getIcon } from '@fluentui/react/lib/Styling';

import styles from './OfficeHero.module.scss';
import type {
  IOfficeHeroProps,
  IContact,
  IQuickLink,
  IOfficeFact,
  QuickLinksResult
} from './IOfficeHeroProps';

const META_SEPARATOR: string = ' · '; // &nbsp;·&nbsp;

interface IQuickLinksState {
  loaded: boolean;
  value: QuickLinksResult;
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

export default function OfficeHero(props: IOfficeHeroProps): JSX.Element {
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
    isEditMode,
    getQuickLinks
  } = props;

  const [links, setLinks] = React.useState<IQuickLinksState>({ loaded: false, value: undefined });

  // Fetch the quick links once, on mount, and cache the result in state.
  React.useEffect((): (() => void) => {
    let active: boolean = true;
    getQuickLinks()
      .then((result: QuickLinksResult): void => {
        if (active) {
          setLinks({ loaded: true, value: result });
        }
      })
      .catch((): void => {
        if (active) {
          setLinks({ loaded: true, value: undefined });
        }
      });
    return (): void => {
      active = false;
    };
    // getQuickLinks is a stable per-render closure from the web part; run once.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const metaLine: string = buildMetaLine([addressLine, openingHours, postRoomHours]);
  const heroStyle: React.CSSProperties = backgroundImageUrl
    ? { backgroundImage: `url('${backgroundImageUrl}')` }
    : {};

  // --- Layout decisions ------------------------------------------------------
  const hasNotice: boolean = showNotice && !!notice;
  const hasFacts: boolean = facts.length > 0;
  const linkList: IQuickLink[] | undefined = links.value;
  const hasLinks: boolean = Array.isArray(linkList) && linkList.length > 0;
  const listMissing: boolean = links.loaded && linkList === undefined;
  // Missing list is only surfaced to page editors; readers see nothing.
  const showQuickColumn: boolean = hasLinks || (listMissing && isEditMode);
  // The "Today at this office" (left) column has content whenever there is a
  // notice or any facts. Quick links are always pinned to the right column.
  const leftHasContent: boolean = hasNotice || hasFacts;

  const noticeCtaHref: string | undefined =
    notice && notice.ctaUrl && notice.ctaUrl.trim().length > 0 ? notice.ctaUrl : undefined;

  // Quick-links block, shared by both layouts (two-per-row in the right column;
  // 4-across only when it stands alone with no "Today" content).
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
                // Two columns: "Today at this office" on the left, quick links on the right.
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
                // No notice and no facts: quick links stand alone across the width.
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
