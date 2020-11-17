[cmdletbinding(DefaultParameterSetName=$false)]
Param(
  [Parameter(ParameterSetName="Rule", Mandatory=$true)]
  [ValidateSet(
    "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AX", "AZ",
    "BA", "BB", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ", "BR", "BS",
    "BT", "BV", "BW", "BY", "BZ", "CA", "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN",
    "CO", "CR", "CU", "CV", "CW", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EC", "EE",
    "EG", "EH", "ER", "ES", "ET", "FI", "FJ", "FK", "FM", "FO", "FR", "GA", "GB", "GD", "GE", "GF",
    "GG", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS", "GT", "GU", "GW", "GY", "HK", "HM",
    "HN", "HR", "HT", "HU", "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IR", "IS", "IT", "JE", "JM",
    "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ", "LA", "LB", "LC",
    "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "ME", "MF", "MG", "MH", "MK",
    "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA",
    "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG",
    "PH", "PK", "PL", "PM", "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RS", "RU", "RW",
    "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS",
    "ST", "SV", "SX", "SY", "SZ", "TC", "TD", "TF", "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TO",
    "TR", "TT", "TV", "TW", "TZ", "UA", "UG", "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI",
    "VN", "VU", "WF", "WS", "XK", "YE", "YT", "ZA", "ZM", "ZW"
  )]
  [string[]]$Country,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleName,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleDisplayName,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [string[]]$RuleDisplayGroup,

  [Parameter(ParameterSetName="Rule", Mandatory=$false)]
  [switch]$ExcludeLocalSubnet,

  [Parameter(ParameterSetName="ListCountries", Mandatory=$true)]
  [switch]$ListCountries,

  [Parameter(Mandatory=$false)]
  [string]$MaxMindLicenseKey,

  [Parameter(Mandatory=$false)]
  [switch]$ForceDownload=$false
)

$ErrorActionPreference = "Stop"

# Constants
$APP_DIR = Join-Path $env:LOCALAPPDATA "GeoWF"
$MAXMIND_LICENSE_KEY_FILE = Join-Path $APP_DIR "maxmind_license_key.txt"
$GEOIP_URL_TEMPLATE = "https://download.maxmind.com/app/geoip_download?edition_id=GeoLite2-Country-CSV&license_key={0}&suffix=zip"
$MAXIMUM_RANGES = 9999

# Variables
$RuleUpdateDesired = $PSCmdlet.ParameterSetName -eq "Rule"

# Create AppData directory
if (!(Test-Path $APP_DIR)) {
  [void](New-Item -ItemType Directory -Force $APP_DIR)
}

# MaxMind license key
if ($MaxMindLicenseKey) {
  ($MaxMindLicenseKey -replace "\s", "") | Out-File -FilePath $MAXMIND_LICENSE_KEY_FILE -NoNewline
  Write-Information ("License key saved to `"{0}`"." -f $MAXMIND_LICENSE_KEY_FILE)
}
if (!(Test-Path $MAXMIND_LICENSE_KEY_FILE)) {
  throw "`"$MAXMIND_LICENSE_KEY_FILE`" not found. Set -MaxMindLicenseKey parameter to specify and store MaxMind license key.`r`n" +
        "Create a free account at https://www.maxmind.com/en/geolite2/signup"
}
$MaxMindLicenseKey = (Get-Content $MAXMIND_LICENSE_KEY_FILE) -replace "\s", ""
if (!$RuleUpdateDesired -and !$ForceDownload -and !$ListCountries) { exit }

# Path variables
$GeoIPURL = ($GEOIP_URL_TEMPLATE -f $MaxMindLicenseKey)
$GeoIPDir = Join-Path $APP_DIR "GeoIP"
$GeoIPZip = Join-Path $APP_DIR "GeoIP.zip"

# Delete GeoIP directory if forcing download
if ($ForceDownload -and (Test-Path $GeoIPDir)) {
  Write-Information "Deleting GeoIP data directory."
  Remove-Item $GeoIPDir -Recurse -Force -Confirm:$false
}

# Download MaxMind databases if not present
if (!(Test-Path $GeoIPDir)) {
  Write-Information "Downloading `"$GeoIPURL`"..."
  Invoke-WebRequest $GeoIPURL -OutFile $GeoIPZip
  Write-Information "Extracting `"$GeoIPZip`"..."
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::ExtractToDirectory($GeoIPZip, $APP_DIR)
  Rename-Item -Path (Get-Item -Path (Join-Path $APP_DIR "GeoLite2-Country-CSV_*") | Sort-Object "Name" -Descending)[0] -NewName "GeoIP"
  Remove-Item $GeoIPZip -Force -Confirm:$false
}

# Ensure at least one rule criteria is specified if in rule update mode
if ($RuleUpdateDesired) {
  if (!($RuleName -or $RuleDisplayName -or $RuleDisplayGroup)) {
    throw "You must specify at least one of the following parameters: -RuleName, -RuleDisplayName, -RuleDisplayGroup."
  }
} else {
  if (!$ListCountries) {
    exit
  }
}

# Import country definitions from CSV
if (!$CountryLocations)  {
  Write-Information "Importing country locations..."
  $CountryLocations  = Import-Csv (Join-Path $GeoIPDir "GeoLite2-Country-Locations-en.csv")
}

# List countries and exit
if ($ListCountries) {
  $CountryLocations | ? { $_.country_name -match "\w" } | Sort-Object "country_name" | Select-Object `
    @{ n = "CountryName"; e = { $_.country_name } },
    @{ n = "CountryCode"; e = { $_.country_iso_code } }
  exit
}

# Import GeoIP data
if (!$CountryBlocksIPv4) {
  Write-Information "Importing IPv4 country blocks..."
  $CountryBlocksIPv4 = Import-Csv (Join-Path $GeoIPDir "GeoLite2-Country-Blocks-IPv4.csv")
}

$CountriesGeonameIDs = ($CountryLocations | ? { $Country -contains $_.country_iso_code }).geoname_id

Write-Information "Getting networks..."
$Networks = ($CountryBlocksIPv4 | ? { $CountriesGeonameIDs -contains $_.geoname_id }).network
Write-Information ("{0} IP ranges retrieved." -f $Networks.Count)

if (!$ExcludeLocalSubnet) {
  $Networks += "LocalSubnet"
}

if ($Networks.Count -gt $MAXIMUM_RANGES) {
  throw "List of IP ranges exceeds the Windows Firewall maximum of $MAXIMUM_RANGES."
}

# Search Windows Firewall on selected criteria
$TargetRules = @()
if ($RuleName)         { $TargetRules += Get-NetFirewallRule -Name $RuleName                 | ? { $_.Direction -eq "Inbound" } }
if ($RuleDisplayName)  { $TargetRules += Get-NetFirewallRule -DisplayName $RuleDisplayName   | ? { $_.Direction -eq "Inbound" } }
if ($RuleDisplayGroup) { $TargetRules += Get-NetFirewallRule -DisplayGroup $RuleDisplayGroup | ? { $_.Direction -eq "Inbound" } }

# Update Windows Firewall
if (($TargetRules | measure).Count -gt 0) {
  Write-Host "Modifying these rules with $($Networks.Count) networks:"
  $TargetRules | Select-Object Direction, DisplayName, DisplayGroup, Profile
  Set-NetFirewallRule -Name $TargetRules.Name -RemoteAddress $Networks
} else {
  Write-Warning "No firewall rules matched criteria. No rules have been modified."
}
