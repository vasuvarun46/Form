// Kronos Genie — Item Category / Product Master data API.
//
// The client supplied the base URL and API key below. They live here as module
// globals so the rest of the page imports them from one place (and so rotating
// the key is a one-line change).
//
// SECURITY NOTE: the API key is bundled into the browser build and sent as an
// "x-api-key" header straight from the client, so anyone with DevTools can read
// it. Per the current task this connection is left exactly as-is, but the key
// should really be proxied behind a server-side endpoint. (The rest of this app
// already uses the same browser-key pattern — ApprovalDecisionService.ts /
// TeamService.ts — so this is at least consistent, not a new exposure.)

import {
  ApiError,
  ApiRow,
  ListResponse,
  ProductFilters,
  ProductMasterRow,
  PRODUCT_FILTER_FIELDS,
} from "./itemCategoryTypes";

export const PRODUCT_API_BASE_URL = "https://bird-cron.suparagroup.com";
export const PRODUCT_API_KEY =
  "efp_facba71084f47c8e28a75d22670d5fdfb401b93f71c3a00d28653fdcf4052ded";

// Auth header. The rest of this app authenticates service calls with a
// lowercase "x-api-key" header (ApprovalDecisionService.ts / TeamService.ts).
export const PRODUCT_API_AUTH_HEADER = "x-api-key";

// The live catalog route. Returns { data: [...rows], meta: { has_next,
// next_cursor, filters, ... } } — cursor (keyset) paginated.
export const PRODUCT_API_SEARCH_PATH = "api/products/eform-catalog";

// Rows fetched per live search request. Small on purpose: the grid shows 20–100
// at a time, so one request covers a couple of grid pages; the next cursor page
// is fetched only when the user actually pages past the end.
export const SEARCH_PAGE_SIZE = 100;

// Export walks the whole matching set, so it uses the API's maximum page size
// (fewest round-trips, stays well under the rate limit).
export const EXPORT_PAGE_SIZE = 1000;

// The API's longest permitted value and most values per request.
const MAX_VALUE_LENGTH = 100;

// One live search result: mapped rows plus the paging state the page needs.
export interface SearchResult {
  rows: ProductMasterRow[];
  nextCursor: string | null;
  hasNext: boolean;
}

// A structured error carrying the API's error code (so callers can special-case
// invalid_cursor / rate_limited) and the Retry-After delay when present.
export class CatalogApiError extends Error {
  public readonly code: string;
  public readonly status: number;
  public readonly requestId: string;
  public readonly retryAfterMs: number | null;

  public constructor(
    message: string,
    code: string,
    status: number,
    requestId: string,
    retryAfterMs: number | null
  ) {
    super(message);
    this.name = "CatalogApiError";
    this.code = code;
    this.status = status;
    this.requestId = requestId;
    this.retryAfterMs = retryAfterMs;
  }
}

const str = (v: unknown): string => String(v ?? "").trim();

const sleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

// Map one raw catalog record to a table row. The "AnalNN" descriptors carry the
// merchandising hierarchy (confirmed against a live response).
const mapApiRow = (raw: ApiRow, index: number): ProductMasterRow => ({
  // SkuCode is the key; fall back to the variant's natural key (then a
  // sequential index) so ids stay unique even if a SkuCode is ever blank.
  id:
    str(raw.SkuCode) ||
    [str(raw.ProductCode), str(raw.ColourDesc), str(raw.SizeDesc)]
      .filter(Boolean)
      .join("|") ||
    `row-${index}`,
  sku: str(raw.SkuCode),
  productCode: str(raw.ProductCode),
  colour: str(raw.ColourDesc),
  size: str(raw.SizeDesc),
  subCollection: str(raw.Anal12Desc),
  brandFamily: str(raw.Anal15Desc),
  cluster: str(raw.Anal16Desc),
  category: str(raw.Anal17Desc),
  series: str(raw.Anal18Desc),
  subSeries: str(raw.Anal19Desc),
  productName: str(raw.ProductName),
  noffQty: Number(raw.Noff) || 0,
});

// Turn the active filters into query params. Each field maps to its exact API
// parameter name; blank values are skipped and every value is trimmed and
// capped at 100 chars (the API's limit). "match" is omitted, so the API's
// default case-insensitive "contains" applies. Different filters combine as AND.
const buildParams = (
  filters: ProductFilters,
  pageSize: number,
  cursor: string | null
): { params: URLSearchParams; sentParams: string[] } => {
  const params = new URLSearchParams();
  params.set("page_size", String(pageSize));

  const sentParams: string[] = [];
  PRODUCT_FILTER_FIELDS.forEach((f) => {
    const value = (filters[f.key] || "").trim();
    if (!value) return;
    params.set(f.filterParam, value.slice(0, MAX_VALUE_LENGTH));
    sentParams.push(f.filterParam);
  });

  // When paging, the same filters must be resent alongside the cursor — a cursor
  // used with different filters returns 422 invalid_cursor.
  if (cursor) params.set("cursor", cursor);

  return { params, sentParams };
};

// Perform one request and parse the JSON body, turning any non-2xx into a
// structured CatalogApiError (code + Retry-After).
const requestCatalog = async (
  params: URLSearchParams
): Promise<ListResponse> => {
  const url = `${PRODUCT_API_BASE_URL}/${PRODUCT_API_SEARCH_PATH}?${params.toString()}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      [PRODUCT_API_AUTH_HEADER]: PRODUCT_API_KEY,
      Accept: "application/json",
    },
  });

  if (!res.ok) {
    let body: Partial<ApiError> = {};
    try {
      body = (await res.json()) as Partial<ApiError>;
    } catch {
      /* non-JSON error body — fall back to status text */
    }
    const retryHeader = Number(res.headers.get("Retry-After"));
    throw new CatalogApiError(
      body.message || res.statusText || "Request failed",
      body.code || "http_error",
      res.status,
      body.request_id || "",
      Number.isFinite(retryHeader) && retryHeader > 0 ? retryHeader * 1000 : null
    );
  }

  return (await res.json()) as ListResponse;
};

// After each response, warn in the console about any filter we sent that the
// server did not echo back in meta.filters (i.e. it was ignored).
const warnDroppedFilters = (
  sentParams: string[],
  json: ListResponse
): void => {
  const applied = (json.meta && json.meta.filters) || {};
  sentParams.forEach((p) => {
    if (!(p in applied)) {
      // eslint-disable-next-line no-console
      console.warn(
        `[eform-catalog] filter "${p}" was ignored by the API ` +
          `(not present in meta.filters). request_id=${json.meta && json.meta.request_id}`
      );
    }
  });
};

// Run a live search against the whole catalog. Pass cursor=null for page 1, or a
// previous result's nextCursor (with the SAME filters) to continue.
export async function searchCatalog(
  filters: ProductFilters,
  cursor: string | null,
  pageSize: number = SEARCH_PAGE_SIZE
): Promise<SearchResult> {
  const { params, sentParams } = buildParams(filters, pageSize, cursor);
  const json = await requestCatalog(params);
  warnDroppedFilters(sentParams, json);

  const list: ApiRow[] = Array.isArray(json.data) ? json.data : [];
  const rows = list.map((raw, i) => mapApiRow(raw, i));
  const hasNext = !!(json.meta && json.meta.has_next);
  const nextCursor = hasNext ? json.meta.next_cursor ?? null : null;

  return { rows, nextCursor, hasNext };
}

// Fetch EVERY row matching the current filters by following the cursor to the
// end (for Export). Uses the max page size, paces requests to stay under the
// 120-req/min limit, and honours Retry-After on 429.
export async function exportAllMatching(
  filters: ProductFilters
): Promise<ProductMasterRow[]> {
  const all: ProductMasterRow[] = [];
  const seen = new Set<string>();
  let cursor: string | null = null;
  // Safety bound so a pathological response can't loop forever.
  const MAX_REQUESTS = 200;

  for (let i = 0; i < MAX_REQUESTS; i += 1) {
    const { params } = buildParams(filters, EXPORT_PAGE_SIZE, cursor);

    let json: ListResponse;
    try {
      json = await requestCatalog(params);
    } catch (e) {
      // On rate-limit, wait the server-requested time and retry the same page.
      if (e instanceof CatalogApiError && e.status === 429) {
        await sleep(e.retryAfterMs ?? 2000);
        i -= 1;
        continue;
      }
      throw e;
    }

    const list: ApiRow[] = Array.isArray(json.data) ? json.data : [];
    list.forEach((raw) => {
      const row = mapApiRow(raw, all.length);
      if (!seen.has(row.id)) {
        seen.add(row.id);
        all.push(row);
      }
    });

    if (!(json.meta && json.meta.has_next)) break;
    cursor = json.meta.next_cursor ?? null;
    if (!cursor) break;

    // ~0.6s between pages keeps us comfortably under 120 requests/minute.
    await sleep(600);
  }

  return all;
}
