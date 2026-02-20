# VATaskList Provisioning

This document explains how to run the TypeScript provisioning script which creates the `VATaskList` and a few supporting fields used by the boilerplate web part.

# VATaskList Provisioning

This document explains two ways to provision the `VATaskList` used by the boilerplate web part:

- A TypeScript script that uses `@pnp/sp` from Node.js (convenient for scripted CI/CD).
- A PnP Provisioning template (XML) with a PowerShell helper using `PnP.PowerShell`.

1) TypeScript provisioning (existing)

Prerequisites
- Node.js (LTS) installed
- An Azure AD app with client credentials (Client ID and Client Secret) that has permission to manage the target SharePoint site (Sites.FullControl.All or delegated permissions as appropriate).

Environment variables
- `SITE_URL` - the full SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/mysite)
- `CLIENT_ID` - Azure AD App client id
- `CLIENT_SECRET` - Azure AD App client secret

Install dependencies (project root):

```bash
npm install
```

Run the TypeScript provisioning script (from project root):

PowerShell
```powershell
$env:SITE_URL = "https://contoso.sharepoint.com/sites/mysite"
$env:CLIENT_ID = "your-client-id"
$env:CLIENT_SECRET = "your-client-secret"
npm run provision
```

What the script does
- Checks for a list named `VATaskList`. If it does not exist, creates it as a custom list and adds these fields: `TaskTitle`, `AssignedTo`, `Completed`, `DueDate`, `Details`, plus additional developer fields: `Priority` (Choice), `Status` (Choice), `Category` (Choice), `EstimatedHours` (Number).

PnP template notes
- The PnP XML also defines a `VATask` content type and associates it with the list. The `VATask` content type includes the custom fields listed above.

Default values and sample rows
- The provisioning template sets default values on several fields:
	- `Priority`: `Medium`
	- `Status`: `Open`
	- `Category`: `Task`
	- `EstimatedHours`: `1`
	- `Completed`: `FALSE`

- The template also seeds the list with three sample rows to help developers get started (Initial setup, Design schema, Sample complete).

When you apply the template the seeded items will be created in the target site and can be edited via the list UI or the web part.

2) PnP Provisioning template (XML) + PowerShell helper

A PnP provisioning template and PowerShell helper have been added at `provisioning/vatasklist-provisioning.xml` and `provisioning/apply-vatasklist.ps1`.

Install PnP.PowerShell (one-time):

PowerShell
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Apply the template (interactive - recommended):

PowerShell
```powershell
.\n+provisioning\apply-vatasklist.ps1 -SiteUrl https://contoso.sharepoint.com/sites/mysite
```

Apply using App-Only (CI/CD) with environment variables:

PowerShell
```powershell
$env:CLIENT_ID = 'your-client-id'
$env:CLIENT_SECRET = 'your-client-secret'
.
provisioning\apply-vatasklist.ps1 -SiteUrl https://contoso.sharepoint.com/sites/mysite -UseAppOnly
```

Notes
- The PowerShell helper uses `Invoke-PnPProvisioningTemplate` and supports both interactive and app-only authentication.
- For production deployments you may prefer to incorporate the provisioning template into a deployment pipeline or convert it to a PnP provisioning JSON format.
