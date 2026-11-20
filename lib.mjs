// Pure helpers for repro.mjs, kept separate so they can be unit tested.

export const DEFAULT_SCOPE = 'AllSites.Manage';

export const CLAIMS_SHOWN = ['aud', 'scp', 'appid', 'tid', 'upn', 'iat', 'exp'];

const USAGE = `Usage:
  node repro.mjs --site <site url> --client-id <app id> [options]

Options:
  --tenant <id or domain>  Entra ID tenant (default: organizations)
  --scope <name>           SharePoint delegated scope (default: ${DEFAULT_SCOPE})
  --repeat <n>             Times to run both requests with one token (default: 3)
  --delay <seconds>        Pause between repeats (default: 5)
  --print-token            Also print the access token (for a curl replay)`;

export function parseArgs(argv) {
    const options = {
        tenant: 'organizations',
        scope: DEFAULT_SCOPE,
        repeat: 3,
        delay: 5,
        printToken: false,
    };

    for (let i = 0; i < argv.length; i++) {
        const arg = argv[i];
        const next = () => {
            const value = argv[++i];
            if (value === undefined || value.startsWith('--')) {
                throw new Error(`${arg} needs a value.\n\n${USAGE}`);
            }
            return value;
        };

        switch (arg) {
            case '--site':
                options.site = next();
                break;
            case '--client-id':
                options.clientId = next();
                break;
            case '--tenant':
                options.tenant = next();
                break;
            case '--scope':
                options.scope = next();
                break;
            case '--repeat':
                options.repeat = positiveInteger(arg, next());
                break;
            case '--delay':
                options.delay = nonNegativeNumber(arg, next());
                break;
            case '--print-token':
                options.printToken = true;
                break;
            case '--help':
            case '-h':
                throw new Error(USAGE);
            default:
                throw new Error(`Unknown option ${arg}.\n\n${USAGE}`);
        }
    }

    if (!options.site || !options.clientId) {
        throw new Error(`--site and --client-id are required.\n\n${USAGE}`);
    }

    options.site = siteBase(options.site);
    return options;
}

function positiveInteger(name, value) {
    const number = Number(value);
    if (!Number.isInteger(number) || number < 1) {
        throw new Error(`${name} must be a whole number of at least 1.`);
    }
    return number;
}

function nonNegativeNumber(name, value) {
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0) {
        throw new Error(`${name} must be a number of seconds, 0 or more.`);
    }
    return number;
}

/** The site URL without query, fragment or trailing slash. */
export function siteBase(siteUrl) {
    const url = new URL(siteUrl);
    if (url.protocol !== 'https:') {
        throw new Error('--site must be an https:// SharePoint site URL.');
    }
    return `${url.origin}${url.pathname}`.replace(/\/+$/, '');
}

/** A delegated scope for the site's SharePoint host, e.g. https://contoso.sharepoint.com/AllSites.Manage. */
export function sharePointScope(siteUrl, scopeName) {
    return `${new URL(siteUrl).origin}/${scopeName}`;
}

/** What a browser checks: the current user's IsSiteAdmin property. */
export function currentUserUrl(site) {
    return `${site}/_api/web/currentuser?$select=Id,IsSiteAdmin`;
}

/**
 * What CSOM-based apps often check: the IsSiteAdmin column of the user's item in
 * the root web's User Information List (SiteUserInfoList).
 */
export function siteUserInfoItemUrl(site, userId) {
    return `${site}/_api/site/rootweb/siteuserinfolist/items(${userId})?$select=IsSiteAdmin`;
}

/** Decodes a JWT's payload without verifying it, and keeps only CLAIMS_SHOWN. */
export function tokenClaims(accessToken) {
    const parts = accessToken.split('.');
    if (parts.length < 2) {
        throw new Error('The access token is not a JWT.');
    }

    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    const claims = {};
    for (const name of CLAIMS_SHOWN) {
        if (payload[name] === undefined) {
            continue;
        }
        claims[name] =
            name === 'iat' || name === 'exp'
                ? new Date(payload[name] * 1000).toISOString()
                : payload[name];
    }
    return claims;
}

/** The parts of a SharePoint response that Microsoft Support can look up. */
export function describeResponse(label, status, headers, body, sentAt) {
    let isSiteAdmin = null;
    let id = null;
    try {
        const json = JSON.parse(body);
        isSiteAdmin = typeof json.IsSiteAdmin === 'boolean' ? json.IsSiteAdmin : null;
        id = Number.isInteger(json.Id) ? json.Id : null;
    } catch {
        // Not JSON (e.g. an HTML error page); leave the values null.
    }

    return {
        request: label,
        sentUtc: sentAt.toISOString(),
        status,
        isSiteAdmin,
        id,
        spRequestGuid: headers.get('sprequestguid'),
        requestId: headers.get('request-id'),
        date: headers.get('date'),
        healthScore: headers.get('x-sharepointhealthscore'),
    };
}

export function formatRow(r) {
    const value = r.isSiteAdmin === null ? '(none)' : String(r.isSiteAdmin);
    return [
        r.sentUtc,
        r.request.padEnd(16),
        String(r.status).padEnd(3),
        `IsSiteAdmin=${value}`.padEnd(18),
        `SPRequestGuid=${r.spRequestGuid ?? '(none)'}`,
        `request-id=${r.requestId ?? '(none)'}`,
    ].join('  ');
}
