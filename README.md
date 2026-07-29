# Office Hero — SharePoint Framework web part

A full-bleed hero banner for individual **office information pages** on a
corporate intranet. It renders an office photograph behind a charcoal scrim, the
office name and details, an optional facilities notice, everyday "office facts",
a quick-links grid, and a key-contacts banner.

**Data is list-driven.** Each page just picks an office; all of its content —
details, facts, notice, contacts, and quick links — comes from SharePoint lists
on the site. Edit a list row once and every page showing that office updates.

![Office Hero — notice on and notice off states](./docs/office-hero.png)

> The screenshot is illustrative. The web part adapts to its container width (see
> [Responsive behaviour](#responsive-behaviour)).

---

## Contents

- [At a glance](#at-a-glance)
- [Prerequisites](#prerequisites)
- [Getting started (local)](#getting-started-local)
- [Build & deploy to the app catalog](#build--deploy-to-the-app-catalog)
- [Central lists](#central-lists)
- [Configuring a page](#configuring-a-page)
- [Facts, notice & weather](#facts-notice--weather)
- [Key contacts & profile photos](#key-contacts--profile-photos)
- [Responsive behaviour](#responsive-behaviour)
- [Accessibility & contrast](#accessibility--contrast)
- [Colour rule (gold)](#colour-rule-gold)
- [Out of scope / follow-ups](#out-of-scope--follow-ups)

---

## At a glance

| | |
|---|---|
| SPFx version | **1.23.2** |
| Framework | React 17.0.1 + TypeScript 5.3 |
| UI | SCSS modules (no CSS-in-JS, no Tailwind); `@fluentui/react` v8 icons |
| Data | PnPjs v3 (`@pnp/sp` 3.26.0) wired to the SPFx context; four lists on the current site |
| Config | One property: an **Office** picker (dropdown of offices) |
| Full bleed | `supportsFullBleed: true` |

---

## Prerequisites

- **Node.js `>=22.14.0 <23.0.0`** (Node 22 LTS). SPFx 1.23.2 pins this range in
  `package.json` `engines`. Verified on Node 22.22.2.
- **npm** 10+ (ships with Node 22).
- **gulp-cli** installed globally: `npm install --global gulp-cli`.
- Access to a SharePoint Online tenant **app catalog** to deploy the package.
- **PnP.PowerShell** to provision the lists: `Install-Module PnP.PowerShell -Scope CurrentUser`.

Install project dependencies:

```bash
npm install
```

---

## Getting started (local)

SPFx no longer ships the offline local workbench, so previewing uses the
**hosted workbench** on your tenant.

1. In `config/serve.json`, replace `{tenantDomain}` with your tenant, e.g.
   `contoso.sharepoint.com`.
2. Trust the local dev certificate once: `gulp trust-dev-cert`.
3. Serve: `gulp serve --nobrowser`.
4. Open `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`, add the
   **Office Hero** web part, and pick an office. (The lists must exist on the site
   you open the workbench on — see [Central lists](#central-lists).)

### Full-bleed note

`supportsFullBleed: true` only takes effect when the web part is placed in a
**full-width section**. Full-width sections are only available on pages in sites
with a **publishing-capable page layout** (communication sites, or team sites with
the publishing feature). In a standard section the web part still renders,
constrained to the section width rather than edge-to-edge.

---

## Build & deploy to the app catalog

```bash
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/office-hero.sppkg`.

1. Upload `office-hero.sppkg` to your **tenant app catalog** (or a site-collection
   app catalog) and **Deploy**. `skipFeatureDeployment: true` lets an admin make
   it available tenant-wide.
2. On each site that hosts office pages, **provision the lists** (below) once.
3. Add the web part to a **full-width section** on each office page and pick the
   office in the property pane.

---

## Central lists

The web part reads four lists on the **current site**. Provision them (and sample
data) in one step:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser   # once
./provisioning/Provision-OfficeHero.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/offices"
# optionally seed one sample person (yourself):
#   -SampleUserEmail "you@contoso.com"
```

Facts, People, and Quick Links each link back to Office Information via an
**`Office` lookup**, so one office edit updates every page.

Prefer to click through the UI? [**Manual column setup**](./provisioning/Manual-Column-Setup.md)
lists the exact column names and types to create by hand. **Do not build these lists
by importing a spreadsheet/CSV** — that gives the columns the wrong internal names
and types and skips the `Office` lookup (see
[troubleshooting](#behaviour-when-listsrows-are-missing)).

### Office Information — one row per office

| Column | Internal name | Type | Notes |
|---|---|---|---|
| Title | `Title` | Text | Office name; this is the value the page's Office picker selects |
| Address line | `AddressLine` | Text | |
| Opening hours | `OpeningHours` | Text | |
| Post room closing time | `PostRoomHours` | Text | Optional; omitted from the meta line when blank |
| Background image | `BackgroundImage` | Hyperlink | Absolute image URL |
| Image description | `ImageAltText` | Note | Optional; rendered as a visually-hidden description |
| Show facilities notice | `ShowNotice` | Yes/No | |
| Notice label / title / detail | `NoticeLabel`, `NoticeTitle`, `NoticeDetail` | Text / Text / Note | |
| Notice button text / link | `NoticeCtaText`, `NoticeCtaUrl` | Text / Hyperlink | Blank link ⇒ notice renders without a button |

### Office Facts — the "Today at this office" chips

| Column | Internal name | Type | Notes |
|---|---|---|---|
| Fact | `Title` | Text | The fact text (e.g. "Blackfriars · 5 min") |
| Icon | `IconName` | Text | A Fluent UI icon name; blank/unknown ⇒ no icon |
| Order | `SortOrder` | Number | Ascending |
| Office | `Office` | Lookup → Office Information | |

### Office People — who fills each role

| Column | Internal name | Type | Notes |
|---|---|---|---|
| Person | `Person` | Person | Users only |
| Role | `Role` | Choice | Head of office / Facilities manager / Reception supervisor / Post room supervisor |
| Job title | `JobTitle` | Text | The line beneath the name (a manual field — no Graph) |
| Office | `Office` | Lookup → Office Information | |

The banner always shows the four fixed roles. For each role it uses the first
matching person for that office; a role with no person renders **"Currently
vacant"**, so the grid never collapses to three.

### Office Quick Links — per-office links

| Column | Internal name | Type | Notes |
|---|---|---|---|
| Title | `Title` | Text | The tile label |
| Url | `LinkUrl` | Hyperlink | Destination |
| Icon | `IconName` | Text | A Fluent UI icon name |
| Order | `SortOrder` | Number | Ascending; the first four are shown |
| Office | `Office` | Lookup → Office Information | |

### Behaviour when lists/rows are missing

- **No office selected, or the office row isn't found** → an editor-only message
  (Edit mode); readers see nothing.
- **The list exists but the office reads as "not found"** → its columns don't match
  the expected internal names — almost always because the list was built by
  importing a spreadsheet/CSV. Edit mode shows a specific "missing expected columns"
  hint; rebuild the columns per
  [Manual column setup](./provisioning/Manual-Column-Setup.md).
- **Missing quick-links list** → editor-only hint; readers see the section omitted.
- **Empty facts / people / quick links** → those sections are simply omitted (or
  vacant, for a role).
- **Query failure** → logged to the console; the section is omitted, never thrown
  into the page.
- **Icons** — an unknown or blank icon name renders no icon rather than a broken
  glyph.

---

## Configuring a page

1. Provision the lists on the site (once) and add office rows, facts, people, and
   quick links.
2. Add **Office Hero** to a full-width section on the office page.
3. In the property pane, choose the **Office** from the dropdown. That is the only
   setting — everything else comes from the lists.

To spin up a new office: add a row to Office Information and its related rows;
create a page and point its Office Hero at that office. No per-page content entry.

---

## Facts, notice & weather

Facts are a generic **icon + text** list, so an office can add whatever is useful
(nearest station, step-free access, café hours, cycle store…). They fill "Today at
this office" at full size when there is no notice, and condense to a strip beneath
the notice when one is present.

**Weather** can be entered as a fact today, but a **live/auto weather value is
deliberately not built** — it needs an approved external weather API and a proxy
(a separate decision). When that is approved it would populate a weather fact
automatically.

---

## Key contacts & profile photos

- Profile photos come from `/_layouts/15/userphoto.aspx?size=M&accountname={email}`
  — **no Microsoft Graph permission required.** If the image 404s or the person
  has no photo, the column falls back to a neutral person glyph; a broken-image
  icon never renders.
- **Job title is a manual field on purpose** (the `JobTitle` column). The
  facilities-facing role name often differs from the Entra job title, and keeping
  it manual avoids a Graph permission grant. The web part does **not** call Graph
  for job title or presence.

---

## Responsive behaviour

| Breakpoint | Behaviour |
|---|---|
| **≥1280px** | Two columns — "Today at this office" (notice + facts) left, quick links (two per row) right. |
| **1024–1279px** | "Today" and quick links stack vertically; quick links go 4-across. |
| **768–1023px** | Contacts banner becomes 2×2; the office name drops to 32px. |
| **<768px** | Everything single column; quick links one-per-row; the contacts banner stacks with a top divider replacing the left one; the scrim flips to a vertical gradient. |

**Container queries vs. media queries.** Container queries would be the ideal fit
(the web part can sit in a narrower section), but the SPFx 1.23 CSS-module build
mangles `@container` rules (it drops the inner selectors), so the shipped version
uses **viewport media queries**, which the pipeline compiles correctly.

---

## Accessibility & contrast

- The office name is the only `<h1>`. Section labels ("Today at this office",
  "Quick links", and a visually-hidden "Key contacts") are `<h2>`; contact role
  headings are `<h3>`.
- The background photo is a CSS `background-image` (never an `<img>`) and is not
  exposed to assistive technology. The optional image description is rendered
  visually-hidden when populated.
- Contact photos are decorative (`alt=""`); icons are `aria-hidden`.
- Quick-link tiles are `<a>` elements (the whole tile is the hit target).
- Visible keyboard focus on every interactive element: `outline: 2px solid #FFD024; outline-offset: 2px`.
- **No motion in v1**, so `prefers-reduced-motion` has nothing to switch off (a
  no-op block documents that it was considered). The loading placeholder is static.

### Contrast ratios (WCAG 2.1)

Measured against the **worst case** for the scrim — the flat 0–40% plateau over a
pure-white photograph. Over any real/darker photo the ratios are higher.

| Foreground | Background | Ratio | Level |
|---|---|---|---|
| Office name `#F1EFE8` | scrim plateau (worst) | **9.79 : 1** | AAA |
| Meta line `#D3D1C7` | scrim plateau (worst) | **7.36 : 1** | AAA |
| Eyebrow / labels `#B4B2A9` | scrim plateau (worst) | **5.30 : 1** | AA |
| Gold `#FFD024` | scrim plateau (worst) | **7.67 : 1** | AAA |
| Gold `#FFD024` | facilities-notice card | **5.47 : 1** | AA |
| Button ink `#2C2C2A` | gold button `#FFD024` | **9.53 : 1** | AAA |

The flat 0–40% plateau in the scrim gradient guarantees these ratios independently
of the photograph — do not simplify it to a two-stop gradient.

---

## Colour rule (gold)

`#FFD024` is the single "needs your attention" accent. It appears in **exactly
three places and nowhere else**: the facilities-notice label, its CTA button, and
the keyboard focus outline. It is **not** used for icons, borders, quick-link
tiles, hover states, or anything decorative. When the notice is hidden, no gold
appears anywhere. The rule is documented at the top of
[`_tokens.scss`](./src/webparts/officeHero/components/_tokens.scss).

---

## Out of scope / follow-ups

Deliberately **not** built:

- **Live weather.** Needs an approved external API and a proxy. (Weather can be a
  static fact in the meantime.)
- **Desk & room occupancy chips.** These will come from a workplace booking
  platform via a cached server-side aggregation; the two-column "Today at this
  office" structure is kept so they can slot in beneath the notice without a
  relayout.
- **Graph lookups** of job title or presence.
- **Multi-notice support.** One notice per office row.
- **Cross-site / hub-wide lists.** The web part reads lists on the current site;
  spanning multiple site collections from one central site is a small change
  (a source-site property + `spfi(url)`).
