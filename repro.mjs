// Minimal repro: does SharePoint Online report IsSiteAdmin=false to a delegated
// token for a user who is a site collection administrator?
//
// Signs in with the device-code flow, then calls two REST endpoints with the
// same token, several times, and prints what SharePoint returned.
// No dependencies; needs Node 18 or later.

import { mkdirSync, writeFileSync } from 'node:fs';
import { setTimeout as sleep } from 'node:timers/promises';
import {
    currentUserUrl,
    describeResponse,
    formatRow,
    parseArgs,
    sharePointScope,
    siteUserInfoItemUrl,
    tokenClaims,
} from './lib.mjs';

async function signIn({ tenant, clientId, scope }) {
    const authority = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0`;

    const start = await postForm(`${authority}/devicecode`, {
        client_id: clientId,
        scope,
    });
    if (!start.ok) {
        throw new Error(`Device code request failed: ${JSON.stringify(start.json)}`);
    }

    console.log(`\n${start.json.message}\n`);

    let interval = (start.json.interval ?? 5) * 1000;
    const deadline = Date.now() + (start.json.expires_in ?? 900) * 1000;

    while (Date.now() < deadline) {
        await sleep(interval);
        const result = await postForm(`${authority}/token`, {
            grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
            client_id: clientId,
            device_code: start.json.device_code,
        });

        if (result.ok) {
            return result.json.access_token;
        }

        const error = result.json.error;
        if (error === 'authorization_pending') {
            continue;
        }
        if (error === 'slow_down') {
            interval += 5000;
            continue;
        }
        throw new Error(`Sign-in failed: ${result.json.error_description ?? error}`);
    }

    throw new Error('Sign-in timed out.');
}

async function postForm(url, fields) {
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams(fields),
    });
    return { ok: response.ok, json: await response.json() };
}

async function callSharePoint(label, url, accessToken) {
    const sentAt = new Date();
    const response = await fetch(url, {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: 'application/json;odata=nometadata',
        },
    });
    const body = await response.text();
    return describeResponse(label, response.status, response.headers, body, sentAt);
}

async function main() {
    const options = parseArgs(process.argv.slice(2));
    const scope = sharePointScope(options.site, options.scope);

    const accessToken = await signIn({ ...options, scope });
    const claims = tokenClaims(accessToken);

    console.log('Token claims (the token itself is not shown):');
    for (const [name, value] of Object.entries(claims)) {
        console.log(`  ${name}: ${value}`);
    }
    if (options.printToken) {
        console.log(`\nAccess token (treat as a password):\n${accessToken}`);
    }
    console.log('');

    const results = [];
    for (let run = 1; run <= options.repeat; run++) {
        const user = await callSharePoint(
            'currentuser',
            currentUserUrl(options.site),
            accessToken,
        );
        results.push(user);
        console.log(formatRow(user));

        if (user.id !== null) {
            const item = await callSharePoint(
                'siteuserinfolist',
                siteUserInfoItemUrl(options.site, user.id),
                accessToken,
            );
            results.push(item);
            console.log(formatRow(item));
        }

        if (run < options.repeat && options.delay > 0) {
            await sleep(options.delay * 1000);
        }
    }

    mkdirSync('results', { recursive: true });
    const file = `results/${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    writeFileSync(
        file,
        JSON.stringify({ site: options.site, scope, claims, results }, null, 2),
    );
    console.log(`\nSaved ${file} (no token in it).`);
}

main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
});
