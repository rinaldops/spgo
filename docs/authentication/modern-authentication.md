---
layout: page
title: Modern Authentication
---

Modern authentication uses an interactive OAuth device-code login (via MSAL) instead of
passing a username/password to SharePoint directly. This is the only authentication mode in
SPGo that supports MFA / Conditional Access, because it's the only one that isn't built on
`node-sp-auth`'s unattended credential strategies. This mode exists only in this fork - it is
not part of upstream SPGo.

When you run a command, VS Code will show a message with a short code and a URL
(`microsoft.com/devicelogin`). Enter the code there, sign in (and complete MFA if prompted),
and the command continues automatically. The access token is cached in memory for the VS Code
session only - it is never written to disk.

## Additional Configuration

There is no shared/well-known app to piggyback on: the community "PnP Management Shell"
multi-tenant app that used to fill this role was **deleted by the PnP team on 2024-09-09**
for security reasons. Every tenant needs its own App Registration now - same requirement
PnP PowerShell itself has had since that date.

Fastest way to create one (requires an account that can create app registrations in the
target tenant; final admin consent step needs a Global Administrator):

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPGo-ModernAuth" -Tenant <yourtenant>.onmicrosoft.com
```

This opens a browser for you to sign in and consent, then prints the new Application
(client) ID. Set `authenticationType` to `Modern` in your `SPGo.json` file and put that ID in:

```json
{
    "authenticationType": "Modern",
    "authenticationDetails": {
        "clientId": "<the clientId Register-PnPEntraIDAppForInteractiveLogin printed>"
    }
}
```

`authenticationDetails.tenantId` is optional - auto-resolved from `sharePointSiteUrl` via
SharePoint's own unauthenticated realm-discovery response. Device code login requires a
tenant-specific authority (the multi-tenant `organizations`/`common` endpoints fail with
`AADSTS50059`), which is why this can't just be skipped. Set it explicitly only if
auto-resolution fails or you need to target a different tenant than the site implies.

## How file operations work under this mode

`spsave`/`sppull` have no OAuth/MFA strategy, so none of SPGo's file-transfer commands go
through them when `authenticationType` is `Modern`. Instead they go through
[PnPjs](https://pnp.github.io/pnpjs/) (`@pnp/sp`), configured with a custom device-code auth
behavior (`src/service/pnpService.ts`) that reuses the same login already implemented in
`ModernAuthService`. This includes bulk folder operations - `SPGo: Populate local workspace`,
`SPGo: Retrieve folder`, and `SPGo: Publish local workspace` all work under Modern auth.

Non-file operations (check out, delete, get server version, discard checkout) don't need
PnPjs - they're small enough that they still go directly through a Bearer-token request
(`src/util/bearerSPRequest.ts`).

Right-clicking a folder and choosing "Publish a major/minor version" publishes everything
under it, recursively, with no glob filtering - unlike `SPGo: Publish local workspace`, which
matches against the `globPattern` you configured, this convention has no configured pattern to
go by, so it sends every file it finds.

## Testing tenant/app access before using the extension

The `probe/` folder at the root of this repo has a standalone script that performs the same
device-code login and a read-only SharePoint REST call, independent of the extension build.
Run it first to confirm your tenant currently allows this kind of login at all:

```
cd probe
node entra-access-probe.js <clientId> [siteUrl]
```

(No install step - it's zero-dependency, plain Node talking to the OAuth endpoints directly.)
