# Sample-data CSVs for the Office Hero lists

> ## ⚠️ Do NOT create the lists by importing these CSVs
>
> These files are **row data only.** Do **not** feed them to SharePoint's *"Import
> spreadsheet"* / *"Create list → From Excel"*. Doing so builds columns named after
> the CSV *headers* (internal name `Address_x0020_line` instead of `AddressLine`),
> makes every column plain **text**, and creates **no `Office` lookup / Person /
> Choice** columns — so the web part can't find its data and reports the office as
> *"not found."*
>
> **Create the columns first**, with the correct internal names and types:
> [`../Manual-Column-Setup.md`](../Manual-Column-Setup.md) (by hand) or
> [`../Provision-OfficeHero.ps1`](../Provision-OfficeHero.ps1) (PnP PowerShell).
> **Then** use these CSVs to paste *rows* into the existing lists.

These CSVs let you **bulk-populate rows** in the four Office Hero lists instead of
typing them one at a time. They hold sample data for two offices (London and
Bristol) — edit the URLs, people, and text to match your tenant.

## Important: create the columns first

A CSV fills in **row data only**. It cannot create the **Lookup** (`Office`),
**Person** (`Person`), or **Choice** (`Role`) columns with the internal names the
web part reads. So do this in order:

1. **Create the lists and columns** — run the provisioning script once:

   ```powershell
   ./provisioning/Provision-OfficeHero.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<yoursite>"
   ```

   (or create the columns by hand using the internal names in the main
   [README](../../README.md#central-lists)). That script also *seeds this same
   sample data*, so if you run it you can skip the CSV upload entirely — the CSVs
   are here for when you'd rather paste your own rows in bulk.

2. **Upload `Office-Information.csv` first.** The other three lists' `Office`
   column is a lookup that must match an existing Office Information **Title**, so
   the offices have to exist before you add facts / people / quick links.

3. **Then upload** `Office-Facts.csv`, `Office-People.csv`, and
   `Office-Quick-Links.csv`.

## How to upload

Open the list → **Edit in grid view** → paste the rows (the columns line up
left-to-right with the CSV), or open the CSV in Excel and copy the cells across.

> Do **not** use the "Import spreadsheet" web part — it creates a *brand-new* list
> and will not reproduce the lookup / person / choice columns the web part needs.

## Column notes

| Column | Notes |
|---|---|
| **Show facilities notice** | Use `Yes` / `No`. |
| **Background image**, **Notice button link**, **Url** | Hyperlink columns — paste the plain URL. Add friendly display text in the list afterwards if you want. |
| **Office** (Facts / People / Quick Links) | Type the office's exact **Title** (e.g. `London - Fleet Place`). SharePoint resolves it to the lookup. |
| **Person** (People) | Must be a real, resolvable user in your tenant. Replace the sample `@contoso.com` addresses. |
| **Role** (People) | Must be exactly one of: `Head of office`, `Facilities manager`, `Reception supervisor`, `Post room supervisor`. |
| **Title** (People) | The list's required default field — put the person's name here. |
| **Order** (Facts / Quick Links) | Ascending number; quick links show the first four. |
| **Icon** | A Fluent UI icon name (e.g. `Train`, `Coffee`, `Warning`, `Mail`). Blank/unknown = no icon. |

The header row in each CSV uses the **display** names you see in the list; they map
to the internal names the web part queries (see the main README's
[Central lists](../../README.md#central-lists) tables).
