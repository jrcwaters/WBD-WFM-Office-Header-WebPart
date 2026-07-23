import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { getIcon } from '@fluentui/react/lib/Styling';

import styles from './OfficeHero.module.scss';
import type {
  IOfficeHeroProps,
  IContact,
  IQuickLink,
  QuickLinksResult
} from './IOfficeHeroProps';

const META_SEPARATOR: string = ' · '; // &nbsp;·&nbsp;

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

/** Renders a Fluent icon only when the name is a registered icon; otherwise nothing. */
function QuickLinkIcon(props: { iconName?: string }): JSX.Element {
  const { iconName } = props;
  // Validate against the registered Fluent icon set so an unknown or blank name
  // renders no icon (label sits flush left) rather than a broken glyph.
  if (!iconName || getIcon(iconName) === undefined) {
    return <React.Fragment />;
  }
  return <Icon iconName={iconName} className={styles.tileIcon} aria-hidden={true} />;
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
  const linkList: IQuickLink[] | undefined = links.value;
  const hasLinks: boolean = Array.isArray(linkList) && linkList.length > 0;
  const listMissing: boolean = links.loaded && linkList === undefined;
  // Missing list is only surfaced to page editors; readers see nothing.
  const showQuickColumn: boolean = hasLinks || (listMissing && isEditMode);
  // When one column is absent, the other spans the full content width.
  const leftFull: boolean = hasNotice && !showQuickColumn;
  const rightFull: boolean = showQuickColumn && !hasNotice;
  const showColumns: boolean = hasNotice || showQuickColumn;

  const noticeCtaHref: string | undefined =
    notice && notice.ctaUrl && notice.ctaUrl.trim().length > 0 ? notice.ctaUrl : undefined;

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

              {showColumns ? (
                <div className={styles.columns}>
                  {/* Left: facilities notice */}
                  {hasNotice && notice ? (
                    <div className={`${styles.colLeft} ${leftFull ? styles.fullWidth : ''}`}>
                      <h2 className={styles.sectionLabel}>Today at this office</h2>
                      <div className={styles.notice}>
                        <div className={styles.noticeBody}>
                          <p className={styles.noticeLabel}>{notice.label}</p>
                          <p className={styles.noticeTitle}>{notice.title}</p>
                          {notice.detail ? (
                            <p className={styles.noticeDetail}>{notice.detail}</p>
                          ) : null}
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
                    </div>
                  ) : null}

                  {/* Right: quick links */}
                  {showQuickColumn ? (
                    <div className={`${styles.colRight} ${rightFull ? styles.fullWidth : ''}`}>
                      <h2 className={styles.sectionLabel}>Quick links</h2>
                      {hasLinks && linkList ? (
                        <div className={`${styles.quickGrid} ${rightFull ? styles.fourAcross : ''}`}>
                          {linkList.map((link: IQuickLink, i: number) => (
                            <a
                              key={i}
                              className={styles.tile}
                              href={link.url}
                              target="_blank"
                              rel="noreferrer"
                            >
                              <QuickLinkIcon iconName={link.iconName} />
                              <span className={styles.tileLabel}>{link.title}</span>
                            </a>
                          ))}
                        </div>
                      ) : (
                        <p className={styles.editorHint}>
                          The &ldquo;Office Quick Links&rdquo; list was not found on this site.
                          Provision it (see the solution README) to show quick links here. Only
                          page editors see this message.
                        </p>
                      )}
                    </div>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>

          {/* ---- Contacts banner ---- */}
          <div className={styles.contacts}>
            <h2 className={styles.visuallyHidden}>Key contacts</h2>
            {contacts.map((contact: IContact, i: number) => (
              <ContactColumn key={i} contact={contact} />
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
