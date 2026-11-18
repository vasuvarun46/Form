// Product Master (Item Category) — row shape + filter keys.
// Rows are fetched live from the Kronos Genie API (see itemCategoryApi.ts).

export type ProductMasterRow = {
  id: string;
  sku: string;
  productCode: string;
  colour: string;
  size: string;
  subCollection: string;
  brandFamily: string;
  cluster: string;
  category: string;
  series: string;
  subSeries: string;
  productName: string;
  noffQty: number;
};

export type ProductFilterKey =
  | "sku"
  | "productCode"
  | "colour"
  | "size"
  | "subCollection"
  | "brandFamily"
  | "cluster"
  | "category"
  | "series"
  | "subSeries";

export type ProductFilters = { [K in ProductFilterKey]: string };

export const EMPTY_PRODUCT_FILTERS: ProductFilters = {
  sku: "",
  productCode: "",
  colour: "",
  size: "",
  subCollection: "",
  brandFamily: "",
  cluster: "",
  category: "",
  series: "",
  subSeries: "",
};

// Each UI filter maps to an API query parameter (left) and a response column
// (documented right). Parameter names must match the API's contract exactly —
// unknown params are silently ignored, so e.g. SKU Code must be "barcode".
export const PRODUCT_FILTER_FIELDS: {
  key: ProductFilterKey;
  label: string;
  filterParam: string;
}[] = [
  { key: "sku", label: "SKU Code", filterParam: "barcode" }, //            -> SkuCode
  { key: "productCode", label: "Product Code", filterParam: "product_code" }, // -> ProductCode
  { key: "colour", label: "Colour", filterParam: "colour" }, //           -> ColourDesc
  { key: "size", label: "Size", filterParam: "size" }, //                 -> SizeDesc
  { key: "subCollection", label: "SubCollection", filterParam: "sub_collection" }, // -> Anal12Desc
  { key: "brandFamily", label: "Brand Family", filterParam: "brand_family" }, //      -> Anal15Desc
  { key: "cluster", label: "Product/Brand Cluster", filterParam: "brand_cluster" }, // -> Anal16Desc
  { key: "category", label: "Category", filterParam: "category" }, //     -> Anal17Desc
  { key: "series", label: "Series", filterParam: "series" }, //          -> Anal18Desc
  { key: "subSeries", label: "SubSeries", filterParam: "sub_series" }, // -> Anal19Desc
];

export type ProductSortKey = keyof Omit<ProductMasterRow, "id">;

// ── API response shapes (eform-catalog) ──

// A raw catalog record. String columns are "" when empty; Noff is an integer.
export interface ApiRow {
  SkuCode: string;
  ProductCode: string;
  ColourDesc: string;
  SizeDesc: string;
  Anal12Desc: string;
  Anal15Desc: string;
  Anal16Desc: string;
  Anal17Desc: string;
  Anal18Desc: string;
  Anal19Desc: string;
  ProductName: string;
  Noff: number;
}

export interface ListMeta {
  page_size: number;
  count: number;
  has_next: boolean;
  next_cursor: string | null;
  // The filters the server actually applied. A param we sent but that is absent
  // here was ignored (wrong name) — the search layer warns when that happens.
  filters: { [param: string]: unknown };
  match: string;
  request_id: string;
}

export interface ListResponse {
  data: ApiRow[];
  meta: ListMeta;
}

export interface ApiError {
  code: string;
  message: string;
  request_id: string;
}

