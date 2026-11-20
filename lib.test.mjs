import { strict as assert } from 'node:assert';
import { describe, it } from 'node:test';
import {
    currentUserUrl,
    describeResponse,
    formatRow,
    parseArgs,
    sharePointScope,
    siteBase,
    siteUserInfoItemUrl,
    tokenClaims,
} from './lib.mjs';

const SITE = 'https://contoso.sharepoint.com/sites/Team';

function jwt(payload) {
    const encode = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
    return `${encode({ alg: 'none' })}.${encode(payload)}.signature`;
}

describe('parseArgs', () => {
    it('applies defaults and normalises the site URL', () => {
        const options = parseArgs(['--site', `${SITE}/?a=1`, '--client-id', 'app']);

        assert.deepEqual(options, {
            site: SITE,
            clientId: 'app',
            tenant: 'organizations',
            scope: 'AllSites.Manage',
            repeat: 3,
            delay: 5,
            printToken: false,
        });
    });

    it('reads every option', () => {
        const options = parseArgs([
            '--site', SITE,
            '--client-id', 'app',
            '--tenant', 'contoso.onmicrosoft.com',
            '--scope', 'AllSites.FullControl',
            '--repeat', '5',
            '--delay', '0',
            '--print-token',
        ]);

        assert.equal(options.tenant, 'contoso.onmicrosoft.com');
        assert.equal(options.scope, 'AllSites.FullControl');
        assert.equal(options.repeat, 5);
        assert.equal(options.delay, 0);
        assert.equal(options.printToken, true);
    });

    it('requires the site and client id', () => {
        assert.throws(() => parseArgs(['--site', SITE]), /--client-id are required/);
    });

    it('rejects a missing value, an unknown option and a bad repeat', () => {
        assert.throws(() => parseArgs(['--site', '--client-id', 'x']), /--site needs a value/);
        assert.throws(() => parseArgs(['--bogus']), /Unknown option --bogus/);
        assert.throws(
            () => parseArgs(['--site', SITE, '--client-id', 'a', '--repeat', '0']),
            /at least 1/,
        );
    });
});

describe('URLs', () => {
    it('strips query, fragment and trailing slash, and requires https', () => {
        assert.equal(siteBase(`${SITE}/#x`), SITE);
        assert.throws(() => siteBase('http://contoso.sharepoint.com/sites/Team'), /https/);
    });

    it('builds the scope on the SharePoint host', () => {
        assert.equal(
            sharePointScope(SITE, 'AllSites.Manage'),
            'https://contoso.sharepoint.com/AllSites.Manage',
        );
    });

    it('builds both request URLs', () => {
        assert.equal(currentUserUrl(SITE), `${SITE}/_api/web/currentuser?$select=Id,IsSiteAdmin`);
        assert.equal(
            siteUserInfoItemUrl(SITE, 11),
            `${SITE}/_api/site/rootweb/siteuserinfolist/items(11)?$select=IsSiteAdmin`,
        );
    });
});

describe('tokenClaims', () => {
    it('keeps only the listed claims and shows times as ISO', () => {
        const claims = tokenClaims(
            jwt({ aud: 'sp', scp: 'AllSites.Manage', iat: 0, exp: 3600, secret: 'x' }),
        );

        assert.deepEqual(claims, {
            aud: 'sp',
            scp: 'AllSites.Manage',
            iat: '1970-01-01T00:00:00.000Z',
            exp: '1970-01-01T01:00:00.000Z',
        });
    });

    it('rejects a value that is not a JWT', () => {
        assert.throws(() => tokenClaims('opaque'), /not a JWT/);
    });
});

describe('describeResponse', () => {
    const headers = new Headers({
        SPRequestGuid: 'guid-1',
        'request-id': 'rid-1',
        Date: 'Sat, 21 Nov 2026 00:00:00 GMT',
    });
    const sentAt = new Date('2026-11-21T00:00:00Z');

    it('reads IsSiteAdmin, Id and the lookup headers', () => {
        const result = describeResponse(
            'currentuser', 200, headers, '{"Id":11,"IsSiteAdmin":false}', sentAt,
        );

        assert.deepEqual(result, {
            request: 'currentuser',
            sentUtc: '2026-11-21T00:00:00.000Z',
            status: 200,
            isSiteAdmin: false,
            id: 11,
            spRequestGuid: 'guid-1',
            requestId: 'rid-1',
            date: 'Sat, 21 Nov 2026 00:00:00 GMT',
            healthScore: null,
        });
    });

    it('leaves the values null for a body that is not JSON', () => {
        const result = describeResponse('currentuser', 403, headers, '<html>', sentAt);

        assert.equal(result.isSiteAdmin, null);
        assert.equal(result.id, null);
        assert.match(formatRow(result), /403\s+IsSiteAdmin=\(none\)/);
    });
});
