import type { WebPartContext } from './spShim';

type SpPageCtxGlobal = {
  webAbsoluteUrl?: string;
  userEmail?: string;
  userDisplayName?: string;
  userId?: number;
  webTitle?: string;
};

declare global {
  interface Window {
    _spPageContextInfo?: SpPageCtxGlobal;
  }
}

let _digest = '';
let _digestExpiry = 0;

const getDigest = async (siteUrl: string): Promise<string> => {
  if (_digest && Date.now() < _digestExpiry) return _digest;
  try {
    const res = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' },
    });
    if (res.ok) {
      const json = await res.json();
      _digest = String(json?.FormDigestValue || '');
      _digestExpiry = Date.now() + 25 * 60 * 1000;
    }
  } catch { /* ignore */ }
  return _digest;
};

export const createSpContext = async (): Promise<WebPartContext> => {
  const spCtx = window._spPageContextInfo;
  const siteUrl = spCtx?.webAbsoluteUrl || window.location.origin;
  let userEmail = spCtx?.userEmail || '';
  let userDisplayName = spCtx?.userDisplayName || '';
  let userId: number | undefined = spCtx?.userId;
  const webTitle = spCtx?.webTitle || '';

  if (!userEmail) {
    try {
      const res = await fetch(
        `${siteUrl}/_api/web/currentuser?$select=Email,Title,Id`,
        { credentials: 'include', headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } }
      );
      if (res.ok) {
        const json = await res.json();
        userEmail = String(json?.Email || '');
        userDisplayName = String(json?.Title || '');
        userId = Number(json?.Id) || undefined;
      }
    } catch { /* ignore */ }
  }

  const spGet = async (url: string, _config: null, options?: RequestInit): Promise<Response> => {
    const { headers: callerHeaders, ...restOpts } = options || {};
    return fetch(url, {
      credentials: 'include',
      ...restOpts,
      method: 'GET',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'odata-version': '',
        ...(callerHeaders as Record<string, string> || {}),
      },
    });
  };

  const spPost = async (
    url: string,
    _config: null,
    options?: { headers?: Record<string, string>; body?: BodyInit | null }
  ): Promise<Response> => {
    const digest = await getDigest(siteUrl);
    const { headers: callerHeaders, body } = options || {};
    return fetch(url, {
      credentials: 'include',
      method: 'POST',
      body: body ?? null,
      headers: {
        Accept: 'application/json;odata=nometadata',
        'odata-version': '',
        ...(digest ? { 'X-RequestDigest': digest } : {}),
        ...(callerHeaders || {}),
      },
    });
  };

  const ctx: WebPartContext = {
    pageContext: {
      web: { absoluteUrl: siteUrl, title: webTitle },
      site: { absoluteUrl: siteUrl },
      user: { email: userEmail, displayName: userDisplayName },
      legacyPageContext: userId ? { userId } : {},
    },
    spHttpClient: { get: spGet, post: spPost },
    httpClient: {
      post: async (url: string, _config: null, options?: RequestInit): Promise<Response> =>
        fetch(url, { method: 'POST', ...options }),
    },
    msGraphClientFactory: {
      getClient: async (): Promise<never> => {
        throw new Error('Microsoft Graph is not configured in standalone React mode.');
      },
    },
    sdks: { microsoftTeams: null },
    isServedFromLocalhost: false,
  };

  return ctx;
};
