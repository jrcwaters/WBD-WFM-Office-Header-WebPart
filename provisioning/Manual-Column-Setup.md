# Manual column setup for the Office Hero lists

Use this if you can't (or don't want to) run the PnP PowerShell script and are
building the four lists **by hand** in the SharePoint UI.

> **Why this matters:** the web part queries columns by their **internal name** and
> joins the related lists through an **`Office` lookup**. If you create the lists by
> **importing a spreadsheet/CSV**, SharePoint names the columns after the CSV
> *headers* (so `Address line` becomes the internal name `Address_x0020_line`, not
> `AddressLine`) and makes everything plain text — no Hyperlink, Yes/No, Number,
> Person, Choice, or lookup. The web part then can't find its columns. Creating the
> columns yourself, with the names and types below, is what makes it work.

## The one trick you need

When you add a column, **the name you type becomes the internal name** (spaces get
encoded as `_x0020_`). So:

1. **Type the internal name with _no spaces_** — e.g. `AddressLine`.
2. *(Optional, cosmetic)* edit the column afterwards and change its **display name**
   to something friendlier like "Address line". Renaming the display name **never**
   changes the internal name, so the web part keeps working.

If you skip step 2 the lists just show `AddressLine`-style headers — the web part
works exactly the same either way. **Only the internal names in the tables below
matter.**

## Order of work

1. **Delete the four spreadsheet-imported lists** (they only held sample data).
2. Create **Office Information first** — the other three point a lookup at it.
3. Create Office Facts, Office People, Office Quick Links.
4. Add rows (the values in [`csv/`](./csv/) are ready to copy from).

Create each list via **Site contents → + New → List → Blank list**, and name it
**exactly** (the web part looks these up by title):

`Office Information` · `Office Facts` · `Office People` · `Office Quick Links`

---

## 1. Office Information

One row per office. The built-in **`Title`** column is the office name — leave it as
`Title`.

| Type this name (internal) | Column type | Suggested display label |
|---|---|---|
| `AddressLine` | Single line of text | Address line |
| `OpeningHours` | Single line of text | Opening hours |
| `PostRoomHours` | Single line of text | Post room closing time |
| `BackgroundImage` | **Hyperlink** | Background image |
| `ImageAltText` | Multiple lines of text (plain text) | Image description |
| `ShowNotice` | **Yes/No** | Show facilities notice |
| `NoticeLabel` | Single line of text | Notice label |
| `NoticeTitle` | Single line of text | Notice title |
| `NoticeDetail` | Multiple lines of text (plain text) | Notice detail |
| `NoticeCtaText` | Single line of text | Notice button text |
| `NoticeCtaUrl` | **Hyperlink** | Notice button link |

## 2. Office Facts

The "Today at this office" chips. The built-in **`Title`** holds the fact text —
you can rename its display label to "Fact".

| Type this name (internal) | Column type | Suggested display label |
|---|---|---|
| `IconName` | Single line of text | Icon |
| `SortOrder` | **Number** (0 decimal places) | Order |
| `Office` | **Lookup** → *Office Information*, column *Title* | Office |

## 3. Office People

Who fills each role. The built-in **`Title`** holds the person's name (it's required,
so put the name there).

| Type this name (internal) | Column type | Suggested display label |
|---|---|---|
| `Person` | **Person** (selection: *People only*) | Person |
| `Role` | **Choice** — see the four exact values below | Role |
| `JobTitle` | Single line of text | Job title |
| `Office` | **Lookup** → *Office Information*, column *Title* | Office |

The `Role` choices must be **exactly** these (the web part matches them):

```
Head of office
Facilities manager
Reception supervisor
Post room supervisor
```

## 4. Office Quick Links

Per-office links. The built-in **`Title`** is the tile label.

| Type this name (internal) | Column type | Suggested display label |
|---|---|---|
| `LinkUrl` | **Hyperlink** | Url |
| `IconName` | Single line of text | Icon |
| `SortOrder` | **Number** (0 decimal places) | Order |
| `Office` | **Lookup** → *Office Information*, column *Title* | Office |

---

## About the `Office` lookup (important)

Name the lookup column **exactly `Office`** (no spaces). SharePoint automatically
exposes a lookup's ID as `<internalName>Id` — so `Office` gives you `OfficeId`,
which is exactly what the web part filters on. Point it at the **Office
Information** list and show its **Title** column. When you add a Facts/People/Links
row, pick the office from this lookup.

## Adding rows

Open each list → **Edit in grid view** and copy the sample values from the CSVs in
[`csv/`](./csv/) as a starting point (they're just reference data — edit the URLs,
people, and text for your tenant). Add **Office Information** rows first, then the
others so the `Office` lookup has offices to point at.

## Prefer to automate it?

Everything above is exactly what [`Provision-OfficeHero.ps1`](./Provision-OfficeHero.ps1)
does in about a minute, if you can run PnP PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser   # once
./provisioning/Provision-OfficeHero.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<the-pages-site>"
```
