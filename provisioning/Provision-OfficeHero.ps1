<#
.SYNOPSIS
    Provisions the four SharePoint lists that drive the Office Hero web part, on
    the site where the office pages live, with sample data.

.DESCRIPTION
    Creates (idempotently):
      * Office Information  - one row per office (details + facilities notice)
      * Office Facts        - "Today at this office" facts (Office lookup)
      * Office People       - who fills each role at each office (Office lookup)
      * Office Quick Links  - per-office quick links (Office lookup)

    Facts / People / Quick Links link back to Office Information via an "Office"
    lookup column, so a single edit to an office row updates every page that
    shows that office.

.EXAMPLE
    .\Provision-OfficeHero.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/offices"

.EXAMPLE
    # Also seed one sample person (yourself) as London's Head of office:
    .\Provision-OfficeHero.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/offices" -SampleUserEmail "you@contoso.com"

.NOTES
    Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser
    Sample link/image URLs are placeholders - edit them after provisioning.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]  [string] $SiteUrl,
    [Parameter(Mandatory = $false)] [string] $SampleUserEmail
)

$ErrorActionPreference = "Stop"
$GROUP = "Office Hero"

Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-List([string]$title) {
    $list = Get-PnPList -Identity $title -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        Write-Host "Creating list '$title'"
        $list = New-PnPList -Title $title -Template GenericList -OnQuickLaunch:$false
    }
    return $list
}

function Ensure-Field([string]$list, [string]$internal, [string]$xml) {
    if ($null -eq (Get-PnPField -List $list -Identity $internal -ErrorAction SilentlyContinue)) {
        Add-PnPFieldFromXml -List $list -FieldXml $xml | Out-Null
        Write-Host "  + $list / $internal"
    }
}

function Ensure-OfficeLookup([string]$list, [string]$officeInfoId) {
    if ($null -eq (Get-PnPField -List $list -Identity "Office" -ErrorAction SilentlyContinue)) {
        $xml = "<Field Type='Lookup' Name='Office' DisplayName='Office' List='{$officeInfoId}' ShowField='Title' Required='TRUE' Group='$GROUP' />"
        Add-PnPFieldFromXml -List $list -FieldXml $xml | Out-Null
        Write-Host "  + $list / Office (lookup)"
    }
}

# --- 1. Office Information ---------------------------------------------------
Ensure-List "Office Information" | Out-Null
Ensure-Field "Office Information" "AddressLine"    "<Field Type='Text' Name='AddressLine' DisplayName='Address line' Group='$GROUP' />"
Ensure-Field "Office Information" "OpeningHours"   "<Field Type='Text' Name='OpeningHours' DisplayName='Opening hours' Group='$GROUP' />"
Ensure-Field "Office Information" "PostRoomHours"  "<Field Type='Text' Name='PostRoomHours' DisplayName='Post room closing time' Group='$GROUP' />"
Ensure-Field "Office Information" "BackgroundImage" "<Field Type='URL' Name='BackgroundImage' DisplayName='Background image' Format='Hyperlink' Group='$GROUP' />"
Ensure-Field "Office Information" "ImageAltText"   "<Field Type='Note' Name='ImageAltText' DisplayName='Image description' NumLines='2' RichText='FALSE' Group='$GROUP' />"
Ensure-Field "Office Information" "ShowNotice"     "<Field Type='Boolean' Name='ShowNotice' DisplayName='Show facilities notice' Group='$GROUP'><Default>0</Default></Field>"
Ensure-Field "Office Information" "NoticeLabel"    "<Field Type='Text' Name='NoticeLabel' DisplayName='Notice label' Group='$GROUP' />"
Ensure-Field "Office Information" "NoticeTitle"    "<Field Type='Text' Name='NoticeTitle' DisplayName='Notice title' Group='$GROUP' />"
Ensure-Field "Office Information" "NoticeDetail"   "<Field Type='Note' Name='NoticeDetail' DisplayName='Notice detail' NumLines='3' RichText='FALSE' Group='$GROUP' />"
Ensure-Field "Office Information" "NoticeCtaText"  "<Field Type='Text' Name='NoticeCtaText' DisplayName='Notice button text' Group='$GROUP' />"
Ensure-Field "Office Information" "NoticeCtaUrl"   "<Field Type='URL' Name='NoticeCtaUrl' DisplayName='Notice button link' Format='Hyperlink' Group='$GROUP' />"

$officeInfoId = (Get-PnPList -Identity "Office Information").Id

# --- 2. Office Facts --------------------------------------------------------
Ensure-List "Office Facts" | Out-Null
Set-PnPField -List "Office Facts" -Identity "Title" -Values @{ Title = "Fact" } | Out-Null
Ensure-Field "Office Facts" "IconName"  "<Field Type='Text' Name='IconName' DisplayName='Icon' Group='$GROUP' />"
Ensure-Field "Office Facts" "SortOrder" "<Field Type='Number' Name='SortOrder' DisplayName='Order' Decimals='0' Group='$GROUP' />"
Ensure-OfficeLookup "Office Facts" $officeInfoId

# --- 3. Office People -------------------------------------------------------
Ensure-List "Office People" | Out-Null
Ensure-Field "Office People" "Person"   "<Field Type='User' Name='Person' DisplayName='Person' UserSelectionMode='PeopleOnly' Group='$GROUP' />"
Ensure-Field "Office People" "Role"      "<Field Type='Choice' Name='Role' DisplayName='Role' Format='Dropdown' Group='$GROUP'><CHOICES><CHOICE>Head of office</CHOICE><CHOICE>Facilities manager</CHOICE><CHOICE>Reception supervisor</CHOICE><CHOICE>Post room supervisor</CHOICE></CHOICES></Field>"
Ensure-Field "Office People" "JobTitle"  "<Field Type='Text' Name='JobTitle' DisplayName='Job title' Group='$GROUP' />"
Ensure-OfficeLookup "Office People" $officeInfoId

# --- 4. Office Quick Links --------------------------------------------------
Ensure-List "Office Quick Links" | Out-Null
Ensure-Field "Office Quick Links" "LinkUrl"   "<Field Type='URL' Name='LinkUrl' DisplayName='Url' Format='Hyperlink' Required='TRUE' Group='$GROUP' />"
Ensure-Field "Office Quick Links" "IconName"  "<Field Type='Text' Name='IconName' DisplayName='Icon' Group='$GROUP' />"
Ensure-Field "Office Quick Links" "SortOrder" "<Field Type='Number' Name='SortOrder' DisplayName='Order' Decimals='0' Group='$GROUP' />"
Ensure-OfficeLookup "Office Quick Links" $officeInfoId

# --- Sample data (only when the lists are empty) ----------------------------
if ((Get-PnPListItem -List "Office Information" -PageSize 5 | Measure-Object).Count -eq 0) {
    Write-Host "Seeding sample offices..."

    $london = Add-PnPListItem -List "Office Information" -Values @{
        Title          = "London - Fleet Place"
        AddressLine    = "10 Fleet Place, London EC4M 7RB"
        OpeningHours   = "Open 08:00-18:00"
        PostRoomHours  = "Post room closes 17:30"
        BackgroundImage = "https://contoso.sharepoint.com/sites/offices/SiteAssets/london.jpg, London - Fleet Place"
        ImageAltText   = "The London Fleet Place building reception."
        ShowNotice     = $true
        NoticeLabel    = "Facilities notice"
        NoticeTitle    = "Air-conditioning offline on floors 3-5"
        NoticeDetail   = "Engineers are on site; expected resolved by 15:00."
        NoticeCtaText  = "Details"
        NoticeCtaUrl   = "https://contoso.sharepoint.com/sites/facilities/SitePages/aircon.aspx, Details"
    }
    $bristol = Add-PnPListItem -List "Office Information" -Values @{
        Title         = "Bristol - Temple Quay"
        AddressLine   = "3 Temple Back East, Bristol BS1 6DZ"
        OpeningHours  = "Open 08:30-17:30"
        ShowNotice    = $false
    }

    $facts = @(
        @{ O = $london.Id;  T = "Blackfriars - 5 min";      I = "Train";           S = 1 },
        @{ O = $london.Id;  T = "Step-free access";         I = "Wheelchair";      S = 2 },
        @{ O = $london.Id;  T = "Cafe open until 16:00";     I = "Coffee";          S = 3 },
        @{ O = $bristol.Id; T = "Temple Meads - 8 min walk"; I = "Train";          S = 1 },
        @{ O = $bristol.Id; T = "Step-free access";          I = "Wheelchair";     S = 2 }
    )
    foreach ($f in $facts) {
        Add-PnPListItem -List "Office Facts" -Values @{ Title = $f.T; IconName = $f.I; SortOrder = $f.S; Office = $f.O } | Out-Null
    }

    $links = @(
        @{ O = $london.Id; T = "Report a fault";        U = "https://contoso.sharepoint.com/sites/facilities/report.aspx"; I = "Warning"; S = 1 },
        @{ O = $london.Id; T = "Fire & evacuation";     U = "https://contoso.sharepoint.com/sites/facilities/fire.aspx";   I = "Running"; S = 2 },
        @{ O = $london.Id; T = "First aid & wellbeing"; U = "https://contoso.sharepoint.com/sites/facilities/firstaid.aspx"; I = "Health"; S = 3 },
        @{ O = $london.Id; T = "Contact facilities";    U = "mailto:facilities@contoso.com";                                I = "Mail";    S = 4 }
    )
    foreach ($l in $links) {
        Add-PnPListItem -List "Office Quick Links" -Values @{ Title = $l.T; LinkUrl = "$($l.U), $($l.T)"; IconName = $l.I; SortOrder = $l.S; Office = $l.O } | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($SampleUserEmail)) {
        try {
            Add-PnPListItem -List "Office People" -Values @{
                Title    = $SampleUserEmail
                Person   = $SampleUserEmail
                Role     = "Head of office"
                JobTitle = "Office Director"
                Office   = $london.Id
            } | Out-Null
            Write-Host "  Added $SampleUserEmail as London's Head of office."
        } catch {
            Write-Warning "Could not add the sample person ($SampleUserEmail): $($_.Exception.Message)"
        }
    } else {
        Write-Host "  (Skipped sample people - pass -SampleUserEmail to seed one, or add people in the Office People list.)"
    }
}

Write-Host "Done. The Office Hero lists are ready on $SiteUrl." -ForegroundColor Green
