// Compatibility shim for @microsoft/sp-webpart-base and @microsoft/sp-http.
// Aliased by Vite (and TypeScript paths) so all existing component imports work unchanged.

// ── Graph client interface ─────────────────────────────────────────────────

interface GraphCall {
  select(fields: string): this;
  get(): Promise<unknown>;
  post(body: unknown): Promise<unknown>;
}

interface GraphClient {
  api(path: string): GraphCall;
}

// ── WebPartContext ─────────────────────────────────────────────────────────

export interface WebPartContext {
  pageContext: {
    web: {
      absoluteUrl: string;
      title?: string;
    };
    site?: {
      absoluteUrl?: string;
    };
    user: {
      email: string;
      displayName: string;
      loginName?: string;
      userPrincipalName?: string;
    };
    legacyPageContext?: Record<string, unknown>;
  };
  spHttpClient: {
    get(url: string, config: null, options?: RequestInit): Promise<Response>;
    post(
      url: string,
      config: null,
      options?: { headers?: Record<string, string>; body?: BodyInit | null }
    ): Promise<Response>;
  };
  httpClient: {
    post(url: string, config: null, options?: RequestInit): Promise<Response>;
  };
  msGraphClientFactory: {
    getClient(version: string): Promise<GraphClient>;
  };
  sdks: {
    microsoftTeams: null;
  };
  isServedFromLocalhost?: boolean;
}

// ── SPHttpClient stub ──────────────────────────────────────────────────────

export const SPHttpClient = {
  configurations: { v1: null as null },
};

export type SPHttpClientResponse = Response;

// ── HttpClient stub ───────────────────────────────────────────────────────

export const HttpClient = {
  configurations: { v1: null as null },
};

export type HttpClientResponse = Response;

// ── IReadonlyTheme (from @microsoft/sp-component-base) ───────────────────

export interface IReadonlyTheme {
  isInverted?: boolean;
  semanticColors?: Record<string, string | undefined>;
}

// ── Version (from @microsoft/sp-core-library) ────────────────────────────

export class Version {
  static parse(_v: string): Version { return new Version(); }
}

// ── BaseClientSideWebPart (from @microsoft/sp-webpart-base) ──────────────

export abstract class BaseClientSideWebPart<T> {
  protected properties!: T;
  protected context!: WebPartContext;
  protected domElement!: HTMLElement;
  abstract render(): void;
  protected onInit(): Promise<void> { return Promise.resolve(); }
  protected onDispose(): void { /* stub */ }
  protected get dataVersion(): Version { return new Version(); }
}

// ── Property pane stubs (from @microsoft/sp-property-pane) ───────────────

export interface IPropertyPaneConfiguration { pages: unknown[] }
export const PropertyPaneTextField = (_key: string, _opts: unknown): unknown => null;
export const PropertyPaneCheckbox = (_key: string, _opts: unknown): unknown => null;
export const PropertyPaneDropdown = (_key: string, _opts: unknown): unknown => null;
export const PropertyPaneToggle = (_key: string, _opts: unknown): unknown => null;
