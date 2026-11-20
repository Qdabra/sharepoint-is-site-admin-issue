# SharePoint Online: IsSiteAdmin = false for a site collection administrator

A minimal repro for a change in SharePoint Online behaviour seen from
mid-November 2026. For a user who is a site collection administrator,
SharePoint reports `IsSiteAdmin = false` to an app's delegated token that has
`AllSites.Manage` and not `AllSites.FullControl`. The same user's own browser request reports `true`.
Granting the app `AllSites.FullControl` appears to make the problem go away.

The script signs in with the device-code flow, so it needs no web app and no
redirect URI. It then calls two REST endpoints several times with the same
token and prints what SharePoint returned:

| Request | What it reads |
|---|---|
| `currentuser` | `GET {site}/_api/web/currentuser?$select=Id,IsSiteAdmin`, the user's `IsSiteAdmin` property (what a browser request returns) |
| `siteuserinfolist` | `GET {site}/_api/site/rootweb/siteuserinfolist/items({id})?$select=IsSiteAdmin`, the `IsSiteAdmin` column of the user's item in the User Information List (what CSOM-based apps often read) |

Each line shows the UTC time, the status, the `IsSiteAdmin` value, and the
`SPRequestGuid` and `request-id` response headers, which Microsoft can look up.

## Requirements

- Node.js 18 or later. There are no packages to install.
- A user who is a site collection administrator of a SharePoint Online site,
  added directly as a user (not through a group).
- An Entra ID app registration you can configure (below).

## App registration

1. In the Entra admin center, go to **App registrations** > **New
   registration**. Any name; single tenant is fine; no redirect URI.
2. Under **Authentication**, set **Allow public client flows** to **Yes**.
   This is what lets the script use the device-code sign-in.
3. Under **API permissions**, add the SharePoint delegated permission
   **AllSites.Manage** (**SharePoint** > **Delegated permissions**), and grant
   admin consent.
   - Do not add `AllSites.FullControl` yet. That is the comparison run.
4. Copy the **Application (client) ID**.

## Run

```
node repro.mjs --site https://<tenant>.sharepoint.com/sites/<site> --client-id <app id> --tenant <tenant id or domain>
```

Follow the sign-in prompt as the site collection administrator. Options:

| Option | Default | Meaning |
|---|---|---|
| `--tenant` | `organizations` | Entra ID tenant ID or domain |
| `--scope` | `AllSites.Manage` | The SharePoint delegated scope to request |
| `--repeat` | `3` | How many times to make both requests with one token |
| `--delay` | `5` | Seconds between repeats |
| `--print-token` | off | Also print the access token, for a `curl` replay |

The script prints the token's `aud`, `scp`, `appid`, `tid`, `upn`, `iat` and
`exp` claims (not the token itself). It saves the results to
`results/<time>.json`, which git ignores.

In our tenant, `scp` listed every SharePoint permission the app had been
granted, not only the scope requested.

## Comparison run

1. Run the script with the app as set up above (`AllSites.Manage` only), and
   note the results.
2. Add **SharePoint** > **Delegated** > **AllSites.FullControl** to the app
   and grant admin consent.
3. Run the script again. Check that `scp` now includes `AllSites.FullControl`,
   and compare the `IsSiteAdmin` values.
4. For a browser comparison, open
   `https://<tenant>.sharepoint.com/sites/<site>/_api/web/currentuser?$select=IsSiteAdmin`
   as the same user, and note the response headers from the browser's
   developer tools.

## What to send to Microsoft Support

- The printed table from each run (UTC time, `IsSiteAdmin`, `SPRequestGuid`,
  `request-id`), and the token claims.
- Not the access token itself, unless Microsoft asks for it through their
  secure upload. If you use `--print-token`, treat the output as a password.
