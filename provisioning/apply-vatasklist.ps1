<#
 .SYNOPSIS
   Apply the VATaskList provisioning template using PnP.PowerShell

 USAGE
   - Interactive (recommended):
       .\apply-vatasklist.ps1 -SiteUrl https://contoso.sharepoint.com/sites/mysite

   - App-only (uses env vars CLIENT_ID and CLIENT_SECRET):
       $env:CLIENT_ID = "..."; $env:CLIENT_SECRET = "..."; .\apply-vatasklist.ps1 -SiteUrl https://contoso.sharepoint.com/sites/mysite -UseAppOnly

 REQUIREMENTS
   Install-Module PnP.PowerShell
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$SiteUrl,

  [Parameter(Mandatory=$false)]
  [string]$TemplatePath = "$(Split-Path -Parent $MyInvocation.MyCommand.Definition)\vatasklist-provisioning.xml",

  [switch]$UseAppOnly
)

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  Write-Host "PnP.PowerShell module not found. Install it using: Install-Module PnP.PowerShell" -ForegroundColor Yellow
  exit 1
}

if ($UseAppOnly) {
  if (-not $env:CLIENT_ID -or -not $env:CLIENT_SECRET) {
    Write-Host "CLIENT_ID and CLIENT_SECRET environment variables are required for app-only authentication." -ForegroundColor Red
    exit 1
  }
  Connect-PnPOnline -Url $SiteUrl -ClientId $env:CLIENT_ID -ClientSecret $env:CLIENT_SECRET -ReturnConnection
} else {
  Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection
}

Write-Host "Applying provisioning template: $TemplatePath" -ForegroundColor Cyan
Invoke-PnPProvisioningTemplate -Path $TemplatePath

Write-Host "Provisioning complete." -ForegroundColor Green
