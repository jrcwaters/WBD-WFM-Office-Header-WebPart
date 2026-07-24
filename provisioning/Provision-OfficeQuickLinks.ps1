<#
.SYNOPSIS
    Creates the "Office Quick Links" list that the Office Hero web part reads,
    with the exact internal column names it queries and four sample items.

.DESCRIPTION
    PnP PowerShell equivalent of office-quick-links.xml. Run it against each
    office site that hosts the Office Hero web part.

.EXAMPLE
    .\Provision-OfficeQuickLinks.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/london-fleet-place"

.NOTES
    Requires the PnP.PowerShell module:  Install-Module PnP.PowerShell -Scope CurrentUser
    Edit the sample link URLs afterwards to point at your own destinations.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $false)]
    [string] $ListTitle = "Office Quick Links"
)

$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $SiteUrl -Interactive

# --- List -------------------------------------------------------------------
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "Creating list '$ListTitle'..."
    $list = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch:$false
}
else {
    Write-Host "List '$ListTitle' already exists; ensuring columns and items."
}

# --- Columns (internal names must match what the web part selects) ----------
function Ensure-Field($internalName, $displayName, $xml) {
    $existing = Get-PnPField -List $ListTitle -Identity $internalName -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Write-Host "  Adding column '$displayName' ($internalName)"
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml | Out-Null
    }
}

Ensure-Field "LinkUrl" "Url" '<Field Type="URL" Name="LinkUrl" DisplayName="Url" Format="Hyperlink" Required="TRUE" Group="Office Hero" />'
Ensure-Field "IconName" "Icon" '<Field Type="Text" Name="IconName" DisplayName="Icon" MaxLength="255" Group="Office Hero" />'
Ensure-Field "SortOrder" "Order" '<Field Type="Number" Name="SortOrder" DisplayName="Order" Decimals="0" Group="Office Hero" />'

# --- Sample items -----------------------------------------------------------
# Edit the Url values to point at your real destinations.
$samples = @(
    @{ Title = "Report a fault";         Url = "https://contoso.sharepoint.com/sites/facilities/SitePages/report-a-fault.aspx"; Icon = "Warning"; Order = 1 },
    @{ Title = "Fire & evacuation";      Url = "https://contoso.sharepoint.com/sites/facilities/SitePages/fire-evacuation.aspx"; Icon = "Running"; Order = 2 },
    @{ Title = "First aid & wellbeing";  Url = "https://contoso.sharepoint.com/sites/facilities/SitePages/first-aid.aspx";       Icon = "Health";  Order = 3 },
    @{ Title = "Contact facilities";     Url = "mailto:facilities@contoso.com";                                                  Icon = "Mail";    Order = 4 }
)

$existingCount = (Get-PnPListItem -List $ListTitle -PageSize 10 | Measure-Object).Count
if ($existingCount -eq 0) {
    foreach ($s in $samples) {
        Write-Host "  Adding item '$($s.Title)'"
        Add-PnPListItem -List $ListTitle -Values @{
            "Title"     = $s.Title
            "LinkUrl"   = "$($s.Url), $($s.Title)"
            "IconName"  = $s.Icon
            "SortOrder" = $s.Order
        } | Out-Null
    }
}
else {
    Write-Host "List already has items; skipping sample data."
}

Write-Host "Done. '$ListTitle' is ready for the Office Hero web part." -ForegroundColor Green
