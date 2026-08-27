import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import * as strings from "EformWebPartStrings";
import { ValidationResult, CreateRequestPageProps, InventoryMode, PurposeOption, CostCenterOption, UsageFieldErrors, ItemRowErrors } from "./createRequestTypes";
import {
  IUsageState,
  ISuparaItem,
  IItemRow,
  FlowItemResponse,
  fetchProductFromAws,
  normalizeFlowToRow,
  getCurrentUserId,

  loadMasters,
  spCreateItem,
  spUpdateItem,
  makeTxnId,
  allocateUniqueTxnId,
  fetchCurrentUserFromGraph,
  AadUserProfile,
  readExcelToItemsMax300,
  validateExcelItemsWithFlowBulk,
  buildApprovalChainFromLists,
} from "./CreateRequestService";
import { IInventoryItemSP } from "./createRequestTypes";
import CreateRequestRender from "./CreateRequestRender";
import validateForm from "./Component/validateForm";
import Loading from "../../ui/Loading";
import { useSharePointTimezoneOffset } from "./sections/SpService"

type SPListResponse<T> = { value: T[] };

type InventoryRequestItemRow = {
  Id: number;
  RequestIDId?: number;
  RowStatus?: string;
  SKUCode?: string;
};

export default function CreateRequestPage({
  onSubmit,
  context,
  draftId,
  onNavigate,
  flowUrl,
}: CreateRequestPageProps): JSX.Element {
  const getApprovalAmountBucket = (amount: number): string => {
    if (amount <= 5000) return "LE_5000";
    if (amount <= 50000) return "5001_50000";
    if (amount <= 100000) return "50001_100000";
    if (amount <= 500000) return "100001_500000";
    return "GT_500000";
  };
  const buildBulkUploadFileName = (employeeId?: string): string => {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    const ss = String(d.getSeconds()).padStart(2, "0");
    const raw = String(employeeId || "").trim().toUpperCase();
    let code = "TSL0000";

    if (/^TSL\d{1,}$/.test(raw)) {
      // Keep exact Employee ID if already in TSL format
      code = raw;
    } else {
      const digits = raw.replace(/\D/g, "");
      const empSuffix = digits ? digits.slice(-4).padStart(4, "0") : "0000";
      code = `TSL${empSuffix}`;
    }

    return `BulkUpload_${code}_${yyyy}${mm}${dd}_${hh}${mi}${ss}`;
  };

  const initialUsage: IUsageState = {
    purpose: "",
    costCenter: "",
    expectedDelivery: "",
    placeOfDelivery: "",
    purposeDetail: "",
  };
  const [usage, setUsage] = React.useState<IUsageState>(initialUsage);
  const [inventoryMode, setInventoryMode] = React.useState<InventoryMode | "">("")
  const [inventoryModeError, setInventoryModeError] = React.useState(false);
  const userData: ISuparaItem[] = [];
  const [aadProfile, setAadProfile] = React.useState<AadUserProfile | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState<boolean>(false);
  const [submitError, setSubmitError] = React.useState<string>("");
  const [purposes, setPurposes] = React.useState<PurposeOption[]>([]);
  const [costCenters, setCostCenters] = React.useState<CostCenterOption[]>([]);
  const getInitialFieldErrors = (): UsageFieldErrors => ({
    purpose: "Purpose can't be blank",
    costCenter: "Cost Center can't be blank",
    expectedDelivery: "Expected Delivery Date can't be blank",
    placeOfDelivery: "Place of Delivery can't be blank",
    purposeDetail: "Purpose Detail can't be blank",
  });

  const [fieldErrors, setFieldErrors] = React.useState<UsageFieldErrors>(
    !draftId ? getInitialFieldErrors() : {}
  );

  const [itemErrors, setItemErrors] = React.useState<ItemRowErrors>({});
  const [searchSku, setSearchSku] = React.useState<string>("");
  const [isSearchingSku, setIsSearchingSku] = React.useState<boolean>(false);
  const [manualPopupOpen, setManualPopupOpen] = React.useState<boolean>(false);
  const [manualProduct, setManualProduct] = React.useState<FlowItemResponse | undefined>(undefined);
  const [manualQty, setManualQty] = React.useState<number>(1);
  const [manualQtyError, setManualQtyError] = React.useState<string>("");
  const [selectedExcelFile, setSelectedExcelFile] = React.useState<File | undefined>(undefined);
  const [excelFileName, setExcelFileName] = React.useState<string>("");
  const [isValidatingExcel, setIsValidatingExcel] = React.useState<boolean>(false);
  const [excelValidated, setExcelValidated] = React.useState<boolean>(false);
  const [excelErrors, setExcelErrors] = React.useState<string[]>([]);
  const [manualMaxStock, setManualMaxStock] = React.useState<number>(0);
  const [items, setItems] = React.useState<IItemRow[]>([]);
  const [errorOpen, setErrorOpen] = React.useState<boolean>(false);
  const [errorTitle, setErrorTitle] = React.useState<string>(strings.ErrorTitle);
  const [errorMessage, setErrorMessage] = React.useState<string>("");
  const [successOpen, setSuccessOpen] = React.useState<boolean>(false);
  const [successTxnId] = React.useState<string>("");
  const [successText] = React.useState<string>("");
  const [skuSearchError, setSkuSearchError] = React.useState<string>("");
  const [draftLoading, setDraftLoading] = React.useState<boolean>(!!draftId);
  const [removeConfirmOpen, setRemoveConfirmOpen] = React.useState(false);
  const [removeIndex, setRemoveIndex] = React.useState<number | null>(null);
  const [removeSku, setRemoveSku] = React.useState<string>("");
  const [excelHelpOpen, setExcelHelpOpen] = React.useState(false);
  const [maxRowsPopupOpen, setMaxRowsPopupOpen] = React.useState(false);
  const [maxRowsPopupMsg, setMaxRowsPopupMsg] = React.useState("");
  const [discardConfirmOpen, setDiscardConfirmOpen] = React.useState(false);
  const [excelUploadedAt, setExcelUploadedAt] = React.useState<string>("");
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [excelFailPopupOpen, setExcelFailPopupOpen] = React.useState(false);  // "Insufficient Stock Detected"
  const [manualStockPopupOpen, setManualStockPopupOpen] = React.useState(false);
  const [customErrorMessage, setCustomErrorMessage] = React.useState("");  // Error message for other errors
  const [customErrorOpen, setCustomErrorOpen] = React.useState(false);  // Show error message popup
  const [modeChangeConfirmOpen, setModeChangeConfirmOpen] = React.useState(false);
  const [pendingMode, setPendingMode] = React.useState<InventoryMode | "">("");
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const persistInFlightRef = React.useRef<boolean>(false);
  const excelOnlyErrorMessage = "Only Excel (.xlsx) files are allowed.";
  const isValidExcelUploadFile = React.useCallback((file?: File): boolean => {
    if (!file) return false;
    const fileName = String(file.name || "").trim().toLowerCase();
    return fileName.endsWith(".xlsx");
  }, []);
  const rejectInvalidExcelUploadFile = React.useCallback((message?: string): void => {
    setSelectedExcelFile(undefined);
    setExcelFileName("");
    setExcelValidated(false);
    setExcelErrors([]);
    setItemErrors({});
    setExcelUploadedAt("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setCustomErrorMessage(message || excelOnlyErrorMessage);
    setCustomErrorOpen(true);
  }, [excelOnlyErrorMessage]);
  const requestRemoveItem = (index: number): void => {
    setRemoveIndex(index);
    setRemoveConfirmOpen(true);
  };
  /** yyyy-MM-dd -> ISO */
  const spTimezoneOffset = useSharePointTimezoneOffset(context);
  const yyyyMmDdToIso = (yyyyMmDd: string): string => {
    const [yyyy, mm, dd] = (yyyyMmDd || "").trim().split("-");
    if (!yyyy || !mm || !dd) throw new Error("Invalid date. Use yyyy-MM-dd");

    // ✅ Use SP timezone offset instead of hardcoded +07:00
    const utcMidnight = Date.UTC(Number(yyyy), Number(mm) - 1, Number(dd));
    const adjustedMs = utcMidnight - spTimezoneOffset * 60 * 1000;
    return new Date(adjustedMs).toISOString();
  };
  const removeItem = (index: number): void => {
    setItems((prev) => prev.filter((_, i) => i !== index).map((x, j) => ({ ...x, no: j + 1 })));
    setItemErrors((prev) => {
      const nextErr: ItemRowErrors = {};
      const keys = Object.keys(prev);
      for (let k = 0; k < keys.length; k++) {
        const oldIndex = Number(keys[k]);
        const msg = prev[oldIndex];
        if (!msg) continue;
        if (oldIndex === index) continue; // removed row error removed
        if (oldIndex > index) nextErr[oldIndex - 1] = msg; // shift down
        else nextErr[oldIndex] = msg; // keep same
      }
      return nextErr;
    });
    // optional: in excel mode, re-check validation state
    if (inventoryMode === "excel") {
      setExcelValidated(false);
      setSelectedExcelFile(undefined);
      setExcelFileName("");
    }
  };
  const openRemoveConfirm = (idx: number, sku: string): void => {
    setRemoveIndex(idx);
    setRemoveSku(sku);
    setRemoveConfirmOpen(true);
  };

  const confirmRemoveItem = (): void => {
    if (removeIndex !== null) {
      removeItem(removeIndex); // ✅ YOUR EXISTING LOGIC
    }
    setRemoveConfirmOpen(false);
    setRemoveIndex(null);
    setRemoveSku("");
  };

  const cancelRemoveItem = (): void => {
    setRemoveConfirmOpen(false);
    setRemoveIndex(null);
    setRemoveSku("");
  };

  // const setUsageField = <K extends keyof IUsageState>(
  //   key: K,
  //   value: IUsageState[K]
  // ): void => {
  //   setUsage((prev) => ({
  //     ...prev,
  //     [key]: value,
  //   }));
  //   setFieldErrors((prev) => {
  //     if (!prev[key]) return prev;
  //     const next = { ...prev };
  //     delete next[key];
  //     return next;
  //   });
  // };
  const fieldErrorMessages: Record<keyof IUsageState, string> = {
    purpose: "Purpose can't be blank",
    costCenter: "Cost Center can't be blank",
    expectedDelivery: "Expected Delivery Date can't be blank",
    placeOfDelivery: "Place of Delivery can't be blank",
    purposeDetail: "Purpose Detail can't be blank",
  };

  const setUsageField = <K extends keyof IUsageState>(
    key: K,
    value: IUsageState[K]
  ): void => {
    setUsage((prev) => ({ ...prev, [key]: value }));

    setFieldErrors((prev) => {
      const isEmpty = !value || String(value).trim() === "";
      const next = { ...prev };
      if (isEmpty) {
        next[key] = fieldErrorMessages[key]; // re-set error if cleared
      } else {
        delete next[key]; // clear error when filled
      }
      return next;
    });
  };



  const closeErrorDialog = (): void => {
    setErrorOpen(false);
    setErrorTitle("Error");
    setErrorMessage("");
  };
  const totalQty = React.useMemo<number>(
    () => items.reduce((sum, it) => sum + Number(it.qty || 0), 0),
    [items]
  );
  const totalAmount = React.useMemo<number>(
    () => items.reduce((sum, it) => sum + Number(it.qty || 0) * Number(it.unitPrice || 0), 0),
    [items]
  );
  // ✅ Check if at least one field is filled
  const hasAnyData = React.useMemo(() => {
    const headerFilled =
      usage.purpose.trim() !== "" ||
      usage.costCenter.trim() !== "" ||
      usage.expectedDelivery.trim() !== "" ||
      usage.placeOfDelivery.trim() !== "" ||
      usage.purposeDetail.trim() !== "" ||
      items.length > 0;

    const hasItems = items.length > 0;

    return headerFilled || hasItems;
  }, [usage, items]);

  const isEditDraftMode = !!(draftId && Number(draftId) > 0);
  const initialDraftSnapshotRef = React.useRef<string | null>(null);

  const currentDraftSnapshot = React.useMemo(() => {
    const normalizedItems = items.map((it) => ({
      sku: String(it.sku || "").trim(),
      productCode: String(it.productCode || "").trim(),
      color: String(it.color || "").trim(),
      size: String(it.size || "").trim(),
      unitPrice: Number(it.unitPrice || 0),
      qty: Number(it.qty || 0),
    }));

    return JSON.stringify({
      usage: {
        purpose: String(usage.purpose || "").trim(),
        costCenter: String(usage.costCenter || "").trim(),
        expectedDelivery: String(usage.expectedDelivery || "").trim(),
        placeOfDelivery: String(usage.placeOfDelivery || "").trim(),
        purposeDetail: String(usage.purposeDetail || "").trim(),
      },
      inventoryMode: String(inventoryMode || ""),
      items: normalizedItems,
    });
  }, [usage, inventoryMode, items]);

  const hasDraftChanges = React.useMemo(() => {
    if (!isEditDraftMode) return hasAnyData;
    if (!hasAnyData) return false;
    if (initialDraftSnapshotRef.current === null) return false;
    return initialDraftSnapshotRef.current !== currentDraftSnapshot;
  }, [isEditDraftMode, hasAnyData, currentDraftSnapshot]);

  const updateQty = (index: number, next: string | number): void => {
    setItems((prev) => {
      const row = prev[index];
      if (!row) return prev;
      const rawInput =
        typeof next === "string" ? next.replace(/\D/g, "") : String(next ?? "");
      if (rawInput === "") {
        const fallbackQty = 1;
        setItemErrors((old) => {
          if (!old[index]) return old;
          const nextErr = { ...old };
          delete nextErr[index];
          return nextErr;
        });
        return prev.map((it, i) => (i === index ? { ...it, qty: fallbackQty } : it));
      }

      const wantedRaw = Number(rawInput);
      if (!Number.isFinite(wantedRaw)) {
        return prev;
      }
      const isZeroInput = wantedRaw === 0;
      if (isZeroInput) {
        const zeroMax = row.maxStock ? Number(row.maxStock) : 0;
        setItemErrors((old) => ({
          ...old,
          [index]: `Quantity must be 1–999 and not exceed available stock (${zeroMax} pcs).`,
        }));
      }

      // clamp 1..999
      const wanted = Math.min(999, Math.max(0, Math.trunc(wantedRaw)));


      const max = row.maxStock ? Number(row.maxStock) : 999999;

      if (wanted > max) {
        setItemErrors((old) => ({
          ...old,
          [index]: `Quantity must be 1–999 and not exceed available stock (${max} pcs).`,
        }));
      } else if (!isZeroInput) {
        setItemErrors((old) => {
          if (!old[index]) return old;
          const nextErr = { ...old };
          delete nextErr[index];
          return nextErr;
        });
      }

      return prev.map((it, i) => (i === index ? { ...it, qty: wanted } : it));
    });
  };
  const applyModeChange = (mode: InventoryMode): void => {
    setInventoryMode(mode);

    // 🔥 clear everything
    setItems([]);
    setItemErrors({});
    setSelectedExcelFile(undefined);
    setExcelFileName("");
    setExcelValidated(false);
    setExcelErrors([]);
    setManualProduct(undefined);
    setManualPopupOpen(false);
    setManualQty(1);
    setSearchSku("");

    if (fileInputRef.current) fileInputRef.current.value = "";
  };
  const requestModeChange = (mode: InventoryMode | ""): void => {
    if (mode === inventoryMode) return;
    const hasManualData = inventoryMode === "manual" && items.length > 0;
    const hasExcelData =
      inventoryMode === "excel" &&
      (!!selectedExcelFile || excelValidated || items.length > 0);

    if (hasManualData || hasExcelData) {
      setPendingMode(mode);
      setModeChangeConfirmOpen(true);
      return;
    }

    applyModeChange(mode as InventoryMode);
  };

  const handleManualSkuSearch = async (): Promise<void> => {
    // 🔄 Reset states
    setSubmitError("");
    setManualQtyError("");
    setManualPopupOpen(false);
    setManualProduct(undefined);
    setSkuSearchError("");


    if (inventoryMode !== "manual") {
      setInventoryMode("manual");
    }

    const sku = String(searchSku || "").trim();

    if (!sku) {
      setSkuSearchError(strings.SkuRequired);
      return;
    }
    const alreadyExists = items.some(
      (it) => String(it.sku).trim() === sku
    );
    const skuRegex = /^[a-zA-Z0-9\-_/]{3,20}$/;

    if (!skuRegex.test(sku)) {
      setSkuSearchError("Invalid SKU format. Please enter numbers only (8-20 digits).");
      return;
    }

    if (alreadyExists) {
      setSkuSearchError(`SKU "${sku}" was already added to the list`);
      return; //  STOP — DO NOT CALL FLOW
    }

    try {
      setIsSearchingSku(true);

      const json: unknown = await fetchProductFromAws(
        context.httpClient,
        sku,
        flowUrl
      );

      const flow: FlowItemResponse | undefined =
        normalizeFlowToRow(json);


      // ❌ SKU NOT FOUND (invalid response OR missing ProductCode)
      if (!flow || !flow.Barcode || String(flow.Barcode).trim() === "") {
        setSkuSearchError(`SKU "${sku}" not found.`);
        return;
      }

      const barcode = String(flow.Barcode || "").trim();
      const productCode = String(flow.ProductCode || "").trim();
      const stock = Number(flow.Warehouse_SOH_Qty ?? 0);

      // ❌ Invalid product data
      if (!barcode || !productCode) {
        setSkuSearchError(strings.InvalidProductData);
        return;
      }

      // ❌ Out of stock
      // ✅ Out of stock — still show popup, but with error and disabled Add
      if (stock <= 0) {
        setSkuSearchError("");
        setManualMaxStock(stock);
        setManualProduct(flow);
        setManualQty(1);
        setManualQtyError(`Quantity must be 1–999 and not exceed available stock (${stock} pcs).`);
        setManualPopupOpen(true);
        return;
      }

      // ✅ SUCCESS
      setSkuSearchError("");
      setManualMaxStock(stock);
      setManualProduct(flow);
      setManualQty(1);
      setManualQtyError("");
      setManualPopupOpen(true);

    } catch (error) {
      // 🔥 CRITICAL FIX: handle Flow 502 / network failure
      console.error("Flow error:", error);

      setManualPopupOpen(false);
      setManualProduct(undefined);
      // Show "not found" if Flow fails
      setSkuSearchError(`SKU "${sku}" not found.`);
    } finally {
      setIsSearchingSku(false);
    }
  };



  const handleAddManualItem = (): void => {
    setSubmitError("");
    setManualQtyError("");

    if (!manualProduct) return;

    const barcode = String(manualProduct.Barcode || "").trim();
    const productCode = String(manualProduct.ProductCode || "").trim();
    const color = String(manualProduct.ColourDesc || "-").trim();
    const size = String(manualProduct.SizeCode || "-").trim();
    const unitPrice = Number(manualProduct.Original_sp || 0);
    const stock = Number(manualProduct.Warehouse_SOH_Qty ?? 0);
    const qty = Number(manualQty || 0);

    const sku = barcode;

    if (!barcode) {
      setSubmitError("Invalid SKU returned from Flow.");
      return;
    }

    if (!productCode) {
      setSubmitError("Flow returned no ProductCode.");
      return;
    }

    // qty must be 1..999
    if (!Number.isFinite(qty) || qty < 1 || qty > 999) {
      setManualQtyError(
        `Quantity must be 1–999 and not exceed available stock (${stock} pcs).`
      );
      return;
    }

    // out of stock
    if (stock <= 0) {
      setManualPopupOpen(false);
      setManualProduct(undefined);
      setSkuSearchError(`SKU ${sku} is out of stock`);
      return;
    }

    // qty > stock
    if (qty > stock) {
      setManualQtyError(
        `Quantity must be 1–999 and not exceed available stock (${stock} pcs).`
      );
      return;
    }

    setItems((prev) => {
      let foundIndex = -1;

      for (let i = 0; i < prev.length; i++) {
        if (String(prev[i].sku).trim() === barcode) {
          foundIndex = i;
          break;
        }
      }

      // ❌ Duplicate SKU not allowed
      if (foundIndex >= 0) {
        setManualQtyError("");
        setSkuSearchError(`SKU "${sku}" was already added to the list`);
        return prev;
      }

      const next: IItemRow[] = [
        ...prev,
        {
          no: prev.length + 1,
          sku: barcode,
          productCode,
          color,
          size,
          unitPrice,
          qty,
          maxStock: stock,
        },
      ];

      return next.map((x, i) => ({ ...x, no: i + 1 }));
    });

    // clean exit
    setManualPopupOpen(false);
    setManualProduct(undefined);
    setManualQty(1);
    setSearchSku("");
  };







  const hardResetForm = (): void => {

    setUsage(initialUsage);
    setInventoryMode("");
    setItems([]);
    setSubmitError("");
    setFieldErrors(getInitialFieldErrors());
    setItemErrors({});
    setSelectedExcelFile(undefined);
    setExcelFileName("");
    setExcelValidated(false);
    setExcelErrors([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
    setManualPopupOpen(false);
    setManualProduct(undefined);
    setManualQty(1);
    setManualQtyError("");
    setSearchSku("");
  };

  const resetForm = (): void => {
    if (hasAnyData) {
      setDiscardConfirmOpen(true);
      return;
    }

    hardResetForm();
  };

  const siteUrl = context.pageContext.web.absoluteUrl;
  const getAllItems = async (): Promise<IInventoryItemSP[]> => {
    let allItems: IInventoryItemSP[] = [];

    let url =
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
      `?$filter=RequestIDId eq ${draftId} and RowStatus eq 'Active'` +
      `&$orderby=LineNumber asc&$top=5000`;

    while (url) {
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await res.json();

      allItems = [...allItems, ...(json.value || [])];


      url = json['@odata.nextLink'] || null;
    }

    return allItems;
  };
  const isoToLocalYyyyMmDd = (iso: string | undefined): string => {
    if (!iso) return "";
    const d = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Bangkok",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(new Date(iso));
    return d; // Returns "YYYY-MM-DD" in Bangkok time
  };


  const draftLoadedRef = React.useRef<number | null>(null);
  React.useEffect(() => {
    let isCancelled = false;

    const run = async (): Promise<void> => {
      try {
        if (!draftId) return;
        if (draftLoadedRef.current === draftId) return;

        draftLoadedRef.current = draftId;
        setDraftLoading(true);

        const siteUrl = context.pageContext.web.absoluteUrl;

        // 🔹 1. Get Header
        const res = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${draftId})` +
          `?$select=*,Purpose/Title,CostCenter/Title&$expand=Purpose,CostCenter`,
          SPHttpClient.configurations.v1
        );

        const header = await res.json();
        if (isCancelled) return;

        // 🔹 2. Get Items (pagination handled)
        const spRows: IInventoryItemSP[] = await getAllItems();
        if (isCancelled) return;

        // 🔹 3. Set Usage
        setUsage({
          purpose: header.Purpose?.Title || "",
          costCenter: header.CostCenter?.Title || "",
          expectedDelivery: isoToLocalYyyyMmDd(header.ExpectedDeliveryDate), // ✅ Bangkok date

          placeOfDelivery: header.PlaceOfDelivery || "",
          purposeDetail: header.PurposeDetails || "",
        });

        setFieldErrors({});

        // 🔹 4. Set Items
        const loadedItems = spRows.map((it, i) => ({
          no: i + 1,
          sku: String(it.SKUCode || "").trim(),
          productCode: it.ProductDescription || "",
          color: it.Color || "-",
          size: it.Size || "-",
          unitPrice: Number(it.UnitPrice || 0),
          qty: Number(it.Quantity || 0),
          maxStock: Number(it.StockAvailability ?? 0),
        }));
        setItems(loadedItems);

        // 🔹 5. Mode
        const draftMode: InventoryMode =
          header.EntryMode === "Excel" ? "excel" : "manual";

        setInventoryMode(draftMode);

        // 🔹 4b. Re-flag any rows saved with an invalid/out-of-stock quantity
        // (e.g. saved as draft with qty 0) so the error shows immediately on reopen
        // instead of only after the user touches the quantity stepper.
        if (draftMode === "manual") {
          const reloadedErrors: ItemRowErrors = {};
          loadedItems.forEach((row, idx) => {
            const qty = row.qty;
            const max = row.maxStock != null ? Number(row.maxStock) : 999999;
            if (qty < 1 || qty > 999) {
              reloadedErrors[idx] = `Quantity must be 1–999 and not exceed available stock (${max} pcs).`;
            } else if (max <= 0) {
              reloadedErrors[idx] = `Quantity must be 1–999 and not exceed available stock (${max} pcs).`;
            } else if (qty > max) {
              reloadedErrors[idx] = `Quantity must be 1–999 and not exceed available stock (${max} pcs).`;
            }
          });
          setItemErrors(reloadedErrors);
        }

        // 🔹 6. Excel attachment logic (FIXED)
        if (draftMode === "excel" && spRows.length > 0) {
          setExcelValidated(true);

          try {
            const attachUrl =
              `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestAttachments')/items` +
              `?$select=FileLeafRef,RequestID,OriginalFileName` +
              `&$filter=RequestID eq '${draftId}'&$orderby=Id desc&$top=1`;

            const attachRes = await context.spHttpClient.get(
              attachUrl,
              SPHttpClient.configurations.v1
            );

            const attachJson = await attachRes.json();
            const attachedFile = attachJson?.value?.[0];

            if (attachedFile?.FileLeafRef) {
              const originalFileName = attachedFile.OriginalFileName
                ? String(attachedFile.OriginalFileName).trim()
                : String(attachedFile.FileLeafRef);

              setExcelFileName(originalFileName);

              const fakeFile = new File([], originalFileName, {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });

              setSelectedExcelFile(fakeFile);
            }
          } catch (e) {
            console.error("Failed to load attachment:", e);
          }
        }
      } catch (e: unknown) {
        console.error("Draft load effect error:", e);
        setSubmitError("Failed to load draft. Please try again.");
      } finally {
        if (!isCancelled) setDraftLoading(false);
      }
    };

    run().catch(() => undefined);

    return () => {
      isCancelled = true;
    };
  }, [draftId, context]);



  React.useEffect(() => {
    if (!isEditDraftMode) {
      initialDraftSnapshotRef.current = null;
      return;
    }
    if (draftLoading) return;
    if (initialDraftSnapshotRef.current === null) {
      initialDraftSnapshotRef.current = currentDraftSnapshot;
    }
  }, [isEditDraftMode, draftLoading, currentDraftSnapshot]);

  const initRef = React.useRef(false);

React.useEffect(() => {
  let isCancelled = false;

  const run = async (): Promise<void> => {
    if (initRef.current) return;
    initRef.current = true;

    try {
      // Fetch Entra ID profile via Graph API (User.Read.All approved)
      const profile = await fetchCurrentUserFromGraph(context).catch(() => null);
      if (isCancelled) return;
      if (profile) setAadProfile(profile);

      // Use Entra ID department name to filter cost centres
      const deptName = profile?.department || null;
      await loadMasters(context, deptName, setPurposes, setCostCenters);

    } catch (e: unknown) {
      console.error("CreateRequestPage -> initial fetch error:", e);
    }
  };

  run().catch((e: unknown) => {
    console.error("CreateRequestPage -> initial fetch run() failed:", e);
  });

  return () => {
    isCancelled = true;
  };
}, [context]); // context is stable — this only fires once per mount
  const inventoryModeInitRef = React.useRef(false);
  React.useEffect(() => {
    // ✅ skip first run (when draft load sets inventoryMode)
    if (!inventoryModeInitRef.current) {
      inventoryModeInitRef.current = true;
      return;
    }
    // ✅ reset only when user switches mode manually
    if (inventoryMode !== "excel") {
      setSelectedExcelFile(undefined);
      setExcelFileName("");
      setExcelValidated(false);
      setExcelErrors([]);
      setItemErrors({});
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }, [inventoryMode]);
  const mapExcelError = (code: string): string => {
    if (code.indexOf("ERR_EXCEL_ONLY_AVAILABLE") === 0) {
      const qty = code.split(":")[1];
      return strings.ExcelOnlyAvailable.replace("{0}", qty);
    }

    switch (code) {
      case "ERR_EXCEL_SKU_MISSING":
        return strings.ExcelSkuMissing;
      case "ERR_EXCEL_QTY_INVALID":
        return strings.ExcelQtyInvalid;
      case "ERR_EXCEL_SKU_NOT_FOUND":
        return strings.ExcelSkuNotFound;
      case "ERR_EXCEL_DUPLICATE_SKU":
        return "Duplicate SKU – please remove this record.";

      case "ERR_EXCEL_OUT_OF_STOCK":
        return strings.ExcelOutOfStock;
      case "ERR_EXCEL_HAS_ERRORS":
        return strings.ExcelHasErrors;
      default:
        return code;
    }
  };
  // ✅ NEW — only fires when crossing the 500K threshold (false → true)
  const prevOverLimitRef = React.useRef<boolean>(false);
  React.useEffect(() => {
    const isOverLimit = totalAmount > 500000;
    if (isOverLimit && !prevOverLimitRef.current) {
      setErrorTitle("Something went wrong");
      setErrorMessage(
        `Total request amount (${totalAmount.toLocaleString()} THB) exceeds the 500,000 THB limit.`
      );
      setErrorOpen(true);
    }
    prevOverLimitRef.current = isOverLimit;
  }, [totalAmount]);
  React.useEffect(() => {
    if (!draftId) {
      hardResetForm();
      draftLoadedRef.current = null;
      initialDraftSnapshotRef.current = null;
    }
  }, [draftId]);
  const runExcelUploadAndValidate = async (): Promise<{
    ok: boolean;
    finalItems: IItemRow[];
    totalQty: number;
    totalAmount: number;
  }> => {
    setSubmitError("");
    setExcelErrors([]);
    setExcelValidated(false);
    setItemErrors({});
    setIsValidatingExcel(true);

    if (!selectedExcelFile) {
      setIsValidatingExcel(false);
      return { ok: false, finalItems: [], totalQty: 0, totalAmount: 0 };
    }
    if (!isValidExcelUploadFile(selectedExcelFile)) {
      setIsValidatingExcel(false);
      rejectInvalidExcelUploadFile();
      return { ok: false, finalItems: [], totalQty: 0, totalAmount: 0 };
    }

    try {
      const excelItems = await readExcelToItemsMax300(selectedExcelFile);
      const v = await validateExcelItemsWithFlowBulk(context, excelItems, flowUrl);

      const localizedRowErrors: ItemRowErrors = {};
      Object.keys(v.rowErrors).forEach((k) => {
        localizedRowErrors[Number(k)] = mapExcelError(v.rowErrors[Number(k)]);
      });

      setItemErrors(localizedRowErrors);
      setItems(v.finalItems);

      const hasErrors =
        v.messages.length > 0 ||
        Object.keys(localizedRowErrors).length > 0;
      const rawRowErrors = Object.keys(v.rowErrors).map((k) => String(v.rowErrors[Number(k)] || ""));
      const hasDuplicateSku = rawRowErrors.some((code) => code === "ERR_EXCEL_DUPLICATE_SKU");
      const hasQtyOutOfRange = rawRowErrors.some((code) => code === "ERR_EXCEL_QTY_INVALID");

      const totalQtyNext = v.finalItems.reduce((sum, it) => sum + Number(it.qty || 0), 0);
      const totalAmountNext = v.finalItems.reduce(
        (sum, it) => sum + Number(it.qty || 0) * Number(it.unitPrice || 0),
        0
      );

      setExcelErrors(v.messages.map(mapExcelError));

      if (hasErrors) {
        if (hasDuplicateSku || hasQtyOutOfRange) {
          const notices: string[] = [];
          if (hasDuplicateSku) {
            notices.push("Duplicate SKU codes are not allowed in Excel upload.");
          }
          if (hasQtyOutOfRange) {
            notices.push("Quantity must be between 1 and 999 for every row.");
          }
          setCustomErrorMessage(notices.join(" "));
          setCustomErrorOpen(true);
        }
        setExcelValidated(false);
        return {
          ok: false,
          finalItems: v.finalItems,
          totalQty: totalQtyNext,
          totalAmount: totalAmountNext,
        };
      }

      setExcelValidated(true);
      return {
        ok: true,
        finalItems: v.finalItems,
        totalQty: totalQtyNext,
        totalAmount: totalAmountNext,
      };

    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unable to complete the action. Please try again.";
      setCustomErrorMessage(message);
      setCustomErrorOpen(true);
      setExcelErrors([message]);
      return { ok: false, finalItems: [], totalQty: 0, totalAmount: 0 };
    } finally {
      setIsValidatingExcel(false);
    }
  };



  const upsertDraftHeaderAndItems = async (
    requestId: number,
    headerBody: Record<string, unknown>,
    fetchLiveStock: boolean = false
  ): Promise<void> => {
    // 1) Update header only
    await spUpdateItem(context, "InventoryRequests", requestId, headerBody);

    // 2) Fetch ALL existing item rows for this request — paginate to avoid 100-row cap
    const siteUrl = context.pageContext.web.absoluteUrl;
    const existingRows: InventoryRequestItemRow[] = [];

    let nextUrl: string | null =
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
      `?$select=Id,RequestIDId,RowStatus,SKUCode,StockAvailability` +
      `&$filter=RequestIDId eq ${requestId}` +
      `&$orderby=Id asc` +
      `&$top=5000`;

    while (nextUrl) {
      const res = await context.spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1
      );
      const json = (await res.json()) as SPListResponse<InventoryRequestItemRow> & {
        "@odata.nextLink"?: string;
      };
      if (json.value && json.value.length) {
        existingRows.push(...json.value);
      }
      nextUrl = json["@odata.nextLink"] || null;
    }

    // ───── Everything below is unchanged ─────
    const normalizeSku = (value?: string): string =>
      String(value || "").trim().toUpperCase();

    const rowsBySku = new Map<string, InventoryRequestItemRow[]>();
    for (let i = 0; i < existingRows.length; i++) {
      const row = existingRows[i];
      const key = normalizeSku(row.SKUCode);
      if (!key) continue;
      const bucket = rowsBySku.get(key) || [];
      bucket.push(row);
      rowsBySku.set(key, bucket);
    }

    const requestedBySku = new Map<string, {
      skuRaw: string;
      productCode: string;
      color: string;
      size: string;
      unitPrice: number;
      qty: number;
    }>();
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const key = normalizeSku(it.sku);
      if (!key) continue;
      if (!requestedBySku.has(key)) {
        requestedBySku.set(key, {
          skuRaw: String(it.sku || "").trim(),
          productCode: String(it.productCode || ""),
          color: String(it.color || ""),
          size: String(it.size || ""),
          unitPrice: Number(it.unitPrice),
          qty: Number(it.qty),
        });
      }
    }

    const keptRowIds = new Set<number>();
    const rowsToInactive = new Set<number>();
    const updatePromises: Promise<void>[] = [];
    const createPromises: Array<Promise<{ Id: number }>> = [];
    const requestedEntries = Array.from(requestedBySku.entries());

    for (let i = 0; i < requestedEntries.length; i++) {
      const [skuKey, reqItem] = requestedEntries[i];
      const matchedRows = rowsBySku.get(skuKey) || [];

      let selectedRowId = 0;
      for (let j = 0; j < matchedRows.length; j++) {
        const rowId = Number(matchedRows[j].Id || 0);
        if (rowId > 0 && !keptRowIds.has(rowId)) {
          selectedRowId = rowId;
          break;
        }
      }

      const rowTotal: number = Number(reqItem.qty) * Number(reqItem.unitPrice);
      const existingStockAvailability =
        matchedRows.length > 0
          ? Number((matchedRows[0] as any).StockAvailability ?? -1)
          : -1;

      let liveStockForUpsert: number =
        existingStockAvailability >= 0
          ? existingStockAvailability
          : Number(
            items.find(
              (x) =>
                String(x.sku || "").trim().toUpperCase() ===
                reqItem.skuRaw.toUpperCase()
            )?.maxStock ?? reqItem.qty
          );

      const isSubmittingNow =
        String(headerBody.WorkflowStatus || "") === "Submitted";

      if (isSubmittingNow && flowUrl && inventoryMode === "manual") {
        try {
          const json = await fetchProductFromAws(
            context.httpClient,
            reqItem.skuRaw,
            flowUrl
          );
          const flow = normalizeFlowToRow(json);
          if (flow) {
            liveStockForUpsert = Number(flow.Warehouse_SOH_Qty ?? 0);
          }
        } catch {
          // keep existing value on failure
        }
      }

      const itemBody: Record<string, unknown> = {
        Title: reqItem.skuRaw,
        LineNumber: i + 1,
        SKUCode: reqItem.skuRaw,
        ProductDescription: reqItem.productCode,
        Color: reqItem.color,
        Size: reqItem.size,
        UnitPrice: Number(reqItem.unitPrice),
        Quantity: Number(reqItem.qty),
        LineTotal: rowTotal,
        RequestIDId: requestId,
        LookupSource: inventoryMode === "excel" ? "Excel" : "Manual",
        RowStatus: "Active",
        StockAvailability: liveStockForUpsert,
      };

      if (selectedRowId > 0) {
        keptRowIds.add(selectedRowId);
        updatePromises.push(
          spUpdateItem(context, "InventoryRequestItems", selectedRowId, itemBody)
        );

        for (let j = 0; j < matchedRows.length; j++) {
          const rowId = Number(matchedRows[j].Id || 0);
          if (rowId > 0 && rowId !== selectedRowId) {
            rowsToInactive.add(rowId);
          }
        }
      } else {
        createPromises.push(
          spCreateItem(context, "InventoryRequestItems", itemBody)
        );
      }
    }

    for (let i = 0; i < existingRows.length; i++) {
      const rowId = Number(existingRows[i].Id || 0);
      if (rowId > 0 && !keptRowIds.has(rowId)) {
        rowsToInactive.add(rowId);
      }
    }

    const deactivatePromises: Promise<void>[] = [];
    const inactiveIds = Array.from(rowsToInactive.values());
    for (let i = 0; i < inactiveIds.length; i++) {
      deactivatePromises.push(
        spUpdateItem(context, "InventoryRequestItems", inactiveIds[i], {
          RowStatus: "Inactive",
        })
      );
    }

    await Promise.all(updatePromises);
    await Promise.all(deactivatePromises);
    await Promise.all(createPromises);
  };
  // Add near the top with other refs


  const runSubmitGuards = async (
    opts?: { allowExcelInsufficient?: boolean }
  ): Promise<boolean> => {
    setSubmitError("");
    setInventoryModeError(false);

    // ─────────────────────────────────────────────────────────────────
    // FIX F1: Excel "Insufficient Stock" → OK was clicked.
    // User has already acknowledged the warning and chose to proceed.
    // Bypass ALL re-validation (validateForm + stock loop) and submit.
    // This also kills the "Something went wrong" popup that was firing
    // from validateForm on the re-submit attempt.
    // ─────────────────────────────────────────────────────────────────
    if (opts?.allowExcelInsufficient) {
      return true;
    }

    // FIX F2: Clear stale stock errors before re-validating manual mode.
    // Without this, validateForm can detect the leftover itemErrors and
    // trigger the generic popup even after the user corrected the qty.
    if (inventoryMode === "manual") {
      setItemErrors({});
    }

    const vr: ValidationResult = validateForm(usage, items);
    if (!vr.ok) {
      setFieldErrors(vr.fieldErrors);
      setItemErrors(vr.itemErrors);

      if (vr.showAsPopup) {
        setErrorTitle("Something went wrong");
        setErrorMessage(vr.submitMessage);
        setErrorOpen(true);
      } else {
        setSubmitError(vr.submitMessage);
      }

      if (items.length === 0) {
        setInventoryModeError(true);
      }
      return false;
    }

    if (inventoryMode === "excel" && !excelValidated && !isEditDraftMode) {
      setSubmitError(strings.ExcelValidateBeforeSubmit);
      return false;
    }

    // Excel stock check (only runs when allowExcelInsufficient is NOT set)
    if (inventoryMode === "excel") {
      const skus = items.map((x) => String(x.sku || "").trim()).filter(Boolean);
      const insufficient: string[] = [];

      for (let i = 0; i < skus.length; i++) {
        const sku = skus[i];
        try {
          const json = await fetchProductFromAws(context.httpClient, sku, flowUrl);
          const flow = normalizeFlowToRow(json);
          const stock = Number(flow?.Warehouse_SOH_Qty ?? 0);
          const reqQty = Number(items.find((x) => x.sku === sku)?.qty || 0);
          if (flow && reqQty > stock) {
            insufficient.push(sku);
          }
        } catch {
          // Excel was already validated at upload time — don't re-block on transient errors
        }
      }

      if (insufficient.length > 0) {
        setExcelFailPopupOpen(true);
        return false;
      }
    }

    if (!inventoryMode) {
      setInventoryModeError(true);
      setSubmitError("Please select inventory detail mode before submit.");
      return false;
    }

    // Manual mode live stock validation
    if (inventoryMode === "manual") {
      const nextItemErrors: ItemRowErrors = {};
      const stockUpdates: Record<number, number> = {};
      let hasManualStockIssue = false;

      for (let i = 0; i < items.length; i++) {
        const sku = String(items[i].sku || "").trim();
        const reqQty = Number(items[i].qty || 0);
        if (!sku) continue;

        try {
          const json = await fetchProductFromAws(context.httpClient, sku, flowUrl);
          const flow = normalizeFlowToRow(json);

          if (!flow || !String(flow.Barcode || "").trim()) {
            nextItemErrors[i] = `SKU "${sku}" not found.`;
            hasManualStockIssue = true;
            continue;
          }

          const stock = Number(flow.Warehouse_SOH_Qty ?? 0);
          stockUpdates[i] = stock;

          if (stock <= 0) {
            nextItemErrors[i] = `Quantity must be 1–999 and not exceed available stock (${stock} pcs).`;
            hasManualStockIssue = true;
            continue;
          }

          if (reqQty < 1 || reqQty > 999 || reqQty > stock) {
            nextItemErrors[i] =
              `Quantity must be 1–999 and not exceed available stock (${stock} pcs).`;
            hasManualStockIssue = true;
          }
        } catch {
          // Live stock check unavailable (no flow URL or API unreachable) —
          // fall back to the stock value saved when the item was added.
          const savedStock = Number(items[i].maxStock ?? 0);
          stockUpdates[i] = savedStock;
          if (savedStock <= 0 || reqQty < 1 || reqQty > 999 || reqQty > savedStock) {
            nextItemErrors[i] = `Quantity must be 1–999 and not exceed available stock (${savedStock} pcs).`;
            hasManualStockIssue = true;
          }
        }
      }

      if (hasManualStockIssue) {
        setItemErrors((prev) => ({ ...prev, ...nextItemErrors }));
        if (Object.keys(stockUpdates).length > 0) {
          setItems((prev) =>
            prev.map((item, j) =>
              j in stockUpdates ? { ...item, maxStock: stockUpdates[j] } : item
            )
          );
        }
        setManualStockPopupOpen(true);
        setSubmitError("");
        return false;
      }
    }

    return true;
  };

  const handleSaveDraft = async (): Promise<void> => {


    if (persistInFlightRef.current) return;
    persistInFlightRef.current = true;
    setSubmitError("");

    if (items.length === 0) {
      setCustomErrorMessage("Please add at least 1 product before saving draft.");
      setCustomErrorOpen(true);
      const ref = persistInFlightRef;  // ✅ avoids race condition lint warning
      ref.current = false;
      return;
    }
    if (totalAmount > 500000) {
      setCustomErrorMessage("Total amount exceeds 500,000 THB. Draft cannot be saved above this limit.");
      setCustomErrorOpen(true);
      const ref = persistInFlightRef;
      ref.current = false;
      return;
    }

    try {
      setIsSubmitting(true);
      const requesterEmail: string = context.pageContext.user.email;
      const requesterUserId: number = await getCurrentUserId(context)
      const purposeMatches: PurposeOption[] = purposes.filter((p) => p.Title === usage.purpose);
      const costCenterMatches: CostCenterOption[] = costCenters.filter((c) => c.Title === usage.costCenter);
      const purposeId: number | undefined = purposeMatches.length ? purposeMatches[0].Id : undefined;
      const costCenterId: number | undefined = costCenterMatches.length ? costCenterMatches[0].Id : undefined;
      const nowIso: string = new Date().toISOString();
      const headerBody: Record<string, unknown> = {
        Title: `DRAFT`,
        RequestCreatedOn: nowIso,
        LastUpdatedOn: nowIso,
        EntryMode: inventoryMode === "excel" ? "Excel" : "Manual",
        RequesterId: requesterUserId,
        EmployeeID: aadProfile?.employeeId || "",
        RequesterEmail: requesterEmail,
        JobPosition: aadProfile?.jobTitle || "",
        Departments: aadProfile?.department || "",
        PurposeId: purposeId,
        CostCenterId: costCenterId,
        ExpectedDeliveryDate: usage.expectedDelivery ? yyyyMmDdToIso(usage.expectedDelivery) : undefined,
        PlaceOfDelivery: usage.placeOfDelivery,
        PurposeDetails: usage.purposeDetail,
        Currency: "THB",
        TotalQuantity: totalQty,
        TotalAmount: totalAmount,
        ApprovalAmountBucket: getApprovalAmountBucket(totalAmount),
        WorkflowStatus: "Draft",
        WorkflowStage: 0,
        ApprovalType: "Draft",
        IsEditableByRequester: true,
        HasAttachments: inventoryMode === "excel",
        ExcelUploadFileName: inventoryMode === "excel" ? excelFileName : "",
        ErrorFlag: false,
      };
      if (draftId && Number(draftId) > 0) {
        const requestId = Number(draftId);

        await upsertDraftHeaderAndItems(requestId, headerBody, true);


        const formattedId = makeTxnId(requestId, "draft");

        await spUpdateItem(context, "InventoryRequests", requestId, {
          Title: formattedId,
          RequestNumber: requestId,
        });

        // ❌ REMOVE THIS ENTIRE BLOCK — it duplicates what upsertDraftHeaderAndItems already does
        // if (items.length) {
        //   const createItemPromises = items.map((it, idx) => { ... });
        //   await Promise.all(createItemPromises);
        // }

        onSubmit();
        onNavigate("my", {
          toastType: "draft",
          toastMessage: "Your request is saved as draft.",
        });

        return;
      }
      // ✅ OLD: create new draft
      const createdHeader = await spCreateItem(context, "InventoryRequests", headerBody);
      const requestId: number = createdHeader.Id;

      const formattedId = makeTxnId(requestId, "draft");

      await spUpdateItem(context, "InventoryRequests", requestId, {
        Title: formattedId,
        RequestNumber: requestId,
      });
      if (items.length) {
        const createItemPromises: Array<Promise<{ Id: number }>> = items.map((it, idx) => {
          const rowTotal: number = Number(it.qty) * Number(it.unitPrice);

          const itemBody: Record<string, unknown> = {
            Title: it.sku,
            LineNumber: idx + 1,
            SKUCode: it.sku,
            ProductDescription: it.productCode,
            Color: it.color,
            Size: it.size,
            UnitPrice: Number(it.unitPrice),
            Quantity: Number(it.qty),
            LineTotal: rowTotal,
            RequestIDId: requestId,
            LookupSource: inventoryMode === "excel" ? "Excel" : "Manual",
            RowStatus: "Active",
            StockAvailability: it.maxStock ?? it.qty,  // ← ADD THIS LINE

          };

          return spCreateItem(context, "InventoryRequestItems", itemBody);
        });

        await Promise.all(createItemPromises);
      }

      onSubmit();
      onNavigate("my", {
        toastType: "draft",
        toastMessage: "Your request is saved as draft.",
      });
    } catch (e: unknown) {
      // eslint-disable-next-line no-console
      console.error(e);
      const message: string = e instanceof Error ? e.message : "Save draft failed.";
      setSubmitError(message);
    } finally {
      setIsSubmitting(false);
      const ref = persistInFlightRef;
      ref.current = false;
    }
  };

  const resolveUserIdByEmail = async (email: string): Promise<number | undefined> => {
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/ensureUser('${email.replace(/'/g, "''")}')`;
    const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { "Content-Type": "application/json;odata=verbose" },
      body: JSON.stringify({})
    });
    if (!res.ok) return undefined;
    const json = await res.json() as { Id?: number };
    return json.Id;
  };

  /* =========================================================
     SUBMIT
     ========================================================= */
  const handleSubmit = async (opts?: { allowExcelInsufficient?: boolean }): Promise<void> => {
    // When user clicks "OK, Got it" on insufficient-stock popup, skip the
    // in-flight guard and force a clean slate so the submit always proceeds.
    if (opts?.allowExcelInsufficient) {
      persistInFlightRef.current = false;
      setIsSubmitting(false);
    }
    if (persistInFlightRef.current) return;
    persistInFlightRef.current = true;
    setIsSubmitting(true);
    try {
      const canProceed = await runSubmitGuards(opts);
      if (!canProceed) return;

      const requesterEmail: string = context.pageContext.user.email;
      const purposeMatches: PurposeOption[] = purposes.filter((p) => p.Title === usage.purpose);
      const costCenterMatches: CostCenterOption[] = costCenters.filter((c) => c.Title === usage.costCenter);
      const purposeId: number | undefined = purposeMatches.length ? purposeMatches[0].Id : undefined;
      const costCenterId: number | undefined = costCenterMatches.length ? costCenterMatches[0].Id : undefined;
      if (!purposeId || !costCenterId) {
        setCustomErrorMessage(
          `Unable to find the selected ${!purposeId ? "Purpose" : "Cost Center"} in the master list. Please re-select it and try again.`
        );
        setCustomErrorOpen(true);
        return;
      }

      const chain = await buildApprovalChainFromLists(context, totalAmount);

      if (chain.stageCount === 0) {
        setCustomErrorMessage(
          "Unable to determine an approver for this request. Please check the LOA configuration or contact your administrator."
        );
        setCustomErrorOpen(true);
        return;
      }

      const requesterUserId: number = await getCurrentUserId(context);
      const nowIso: string = new Date().toISOString();
      const firstApproverEmail = chain.firstStageApproverEmails[0] || "";
      const firstApproverUserId = firstApproverEmail
        ? await resolveUserIdByEmail(firstApproverEmail)
        : undefined;

      // ✅ Shared approval fields — used in both paths
      const approvalFields: Record<string, unknown> = {
        RequestSubmittedOn: nowIso,
        LastUpdatedOn: nowIso,
        PendingSince: nowIso,
        EntryMode: inventoryMode === "excel" ? "Excel" : "Manual",
        RequesterId: requesterUserId,
        EmployeeID: aadProfile?.employeeId || "",
        RequesterEmail: requesterEmail,
        JobPosition: aadProfile?.jobTitle || "",
        Departments: aadProfile?.department || "",
        PurposeId: purposeId,
        CostCenterId: costCenterId,
        ExpectedDeliveryDate: usage.expectedDelivery
          ? yyyyMmDdToIso(usage.expectedDelivery)
          : undefined,
        PlaceOfDelivery: usage.placeOfDelivery,
        PurposeDetails: usage.purposeDetail,
        Currency: "THB",
        TotalQuantity: totalQty,
        TotalAmount: totalAmount,
        ApprovalAmountBucket: getApprovalAmountBucket(totalAmount),
        WorkflowStage: chain.stageCount > 0 ? 1 : 0,
        ApprovalType: chain.approvalLabel,
        HasMultipleApprovals: chain.hasMultipleApprovals,
        CurrentApproverId: firstApproverUserId || undefined,
        CurrentApproverEmail: firstApproverEmail || undefined,
        ApprovalChainJSON: chain.chainJSON,
        MultipleApproverData: chain.chainJSON,
        IsEditableByRequester: true,
        HasAttachments: inventoryMode === "excel",
        ExcelUploadFileName: inventoryMode === "excel" ? excelFileName : "",
        ErrorFlag: false,
      };

      // =========================================================
      // PATH A: Draft → Submit
      // ✅ Generate Transaction ID FIRST
      // ✅ Inject real Title into header BEFORE writing Submitted
      // ✅ PA receives correct Title on first Submitted write
      // =========================================================
      if (draftId && Number(draftId) > 0) {
        const requestId = Number(draftId);

        // STEP 1: Allocate a unique Transaction ID (writes Title, WorkflowStatus stays "Draft")
        const formattedId = await allocateUniqueTxnId(context, requestId);

        // STEP 2: Build header with confirmed Title + WorkflowStatus = Submitted
        const headerBody: Record<string, unknown> = {
          ...approvalFields,
          Title: formattedId,
          RequestNumber: requestId,
          RequestCreatedOn: nowIso,
          WorkflowStatus: "Submitted",
        };

        // STEP 3: Upsert — PA triggers here and reads the correct Transaction ID
        await upsertDraftHeaderAndItems(requestId, headerBody, true);

        await spUpdateItem(context, "InventoryRequests", requestId, {
          Title: formattedId,
          RequestNumber: requestId,
          WorkflowStatus: "Submitted",
        });

        onSubmit();
        onNavigate("my", {
          toastType: "success",
          toastMessage: "Your request is submitted to the approver(s).",
        });

        return;
      }

      // =========================================================
      // PATH B: New Record → Submit
      // ✅ Create item with WorkflowStatus = "Pending" (PA does NOT trigger)
      // ✅ Create all line items
      // ✅ Generate Transaction ID
      // ✅ Single final update sets real Title + WorkflowStatus = Submitted
      // ✅ PA triggers ONLY on this final write — always reads correct Title
      // =========================================================

      // ✅ STEP 1: Create the SP record with WorkflowStatus = "Pending"
      //           PA does NOT trigger on Pending — Title is irrelevant at this point
      const pendingHeaderBody: Record<string, unknown> = {
        ...approvalFields,
        Title: "PENDING",              // placeholder — will be replaced before Submitted
        RequestCreatedOn: nowIso,
        WorkflowStatus: "Pending",     // ✅ PA does not fire on Pending
      };

      const createdHeader = await spCreateItem(
        context,
        "InventoryRequests",
        pendingHeaderBody
      );

      const requestId: number = createdHeader.Id;

      // ✅ STEP 2: Create all line items while status is still "Pending"
      const createItemPromises = items.map(async (it, idx) => {
        const rowTotal = Number(it.qty) * Number(it.unitPrice);

        let liveStock: number = Number(it.maxStock ?? 0);

        if (flowUrl && inventoryMode === "manual") {
          try {
            const json = await fetchProductFromAws(
              context.httpClient,
              String(it.sku || "").trim(),
              flowUrl
            );
            const flow = normalizeFlowToRow(json);
            if (flow) {
              liveStock = Number(flow.Warehouse_SOH_Qty ?? 0);
            }
          } catch {
            liveStock = Number(it.maxStock ?? 0);
          }
        }

        const itemBody: Record<string, unknown> = {
          Title: it.sku,
          LineNumber: idx + 1,
          SKUCode: it.sku,
          ProductDescription: it.productCode,
          Color: it.color,
          Size: it.size,
          UnitPrice: Number(it.unitPrice),
          Quantity: Number(it.qty),
          LineTotal: rowTotal,
          RequestIDId: requestId,
          LookupSource: inventoryMode === "excel" ? "Excel" : "Manual",
          RowStatus: "Active",
          StockAvailability: liveStock,
        };

        return spCreateItem(context, "InventoryRequestItems", itemBody);
      });

      await Promise.all(createItemPromises);

      // STEP 3: Allocate unique Transaction ID (writes Title, WorkflowStatus stays "Pending")
      const formattedId = await allocateUniqueTxnId(context, requestId);

      // STEP 4: Flip WorkflowStatus to Submitted — PA triggers here with correct Title
      await spUpdateItem(context, "InventoryRequests", requestId, {
        Title: formattedId,
        RequestNumber: requestId,
        WorkflowStatus: "Submitted",
      });

      onSubmit();
      onNavigate("my", {
        toastType: "success",
        toastMessage: "Your request is submitted to the approver(s).",
      });

    } catch (e: unknown) {
      console.error(e);
      const message = e instanceof Error ? e.message : "Submit failed. Please try again.";
      setSubmitError(message);
      setCustomErrorMessage(message);
      setCustomErrorOpen(true);
    } finally {
      setIsSubmitting(false);
      const ref = persistInFlightRef;
      ref.current = false;
    }
  };

  const closeSuccessDialog = (): void => {
    setSuccessOpen(false);
    onSubmit();
  };
  const deleteDraft = async (id: number): Promise<void> => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const itemsRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
      `?$select=Id&$filter=RequestIDId eq ${id}`,
      SPHttpClient.configurations.v1
    );

    if (!itemsRes.ok) {
      throw new Error("Unable to load InventoryRequestItems for discard.");
    }

    const itemsJson = (await itemsRes.json()) as SPListResponse<InventoryRequestItemRow>;
    const linkedRows = itemsJson.value || [];

    if (linkedRows.length > 0) {
      await Promise.all(
        linkedRows.map((row) =>
          context.spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items(${row.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                "IF-MATCH": "*",
                "X-HTTP-Method": "DELETE",
              },
            }
          )
        )
      );
    }

    await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      }
    );
  };

  const handleDiscard = async (): Promise<void> => {
    try {
      if (draftId && draftId > 0) {
        await deleteDraft(draftId);
      }

      hardResetForm();
      onSubmit();
      onNavigate("my", {
        toastType: "error",
        toastMessage: "Your request is discarded.",
      });
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unable to discard request.";
      setSubmitError(message);
      setCustomErrorMessage(message);
      setCustomErrorOpen(true);
    }
  };

  const pageBg = "bg-[#f3f3f3]";
  const container = "w-full max-w-[1320px] mx-auto";
  const card = "bg-white rounded-xl border border-[#e5e5e5] shadow-sm";
  if (draftLoading) {
    return (
      <Loading />
    );
  }

  // Add this after the existing const declarations (around line where mandatoryCompleted is defined)


  return (

    <CreateRequestRender
      /* ================= Usage ================= */
      usage={usage}
      fieldErrors={fieldErrors}
      inventoryMode={inventoryMode}
      setInventoryMode={requestModeChange}
      inventoryModeError={inventoryModeError}

      /* ✅ ADD THESE 4 MISSING PROPS */
      modeChangeConfirmOpen={modeChangeConfirmOpen}
      setModeChangeConfirmOpen={setModeChangeConfirmOpen}
      pendingMode={pendingMode}
      applyModeChange={applyModeChange}

      items={items}
      itemErrors={itemErrors}
      updateQty={updateQty}
      removeItem={requestRemoveItem}
      removeConfirmOpen={removeConfirmOpen}
      confirmRemoveItem={confirmRemoveItem}
      cancelRemoveItem={cancelRemoveItem}
      removeSku={removeSku}
      openRemoveConfirm={openRemoveConfirm}
      totalQty={totalQty}
      totalAmount={totalAmount}
      isSubmitting={isSubmitting}
      isSearchingSku={isSearchingSku}
      isValidatingExcel={isValidatingExcel}
      excelValidated={excelValidated}
      excelErrors={excelErrors}
      excelFileName={excelFileName}
      fileInputRef={fileInputRef}

      onSelectExcelFile={(file) => {
        if (file && !isValidExcelUploadFile(file)) {
          rejectInvalidExcelUploadFile();
          return;
        }
        setSelectedExcelFile(file);
        setExcelFileName(file ? buildBulkUploadFileName(aadProfile?.employeeId) : "");
        setExcelValidated(false);
        setExcelErrors([]);
        setItemErrors({});
        if (!file) setExcelUploadedAt("");
      }}

      onRunExcelUploadAndValidate={runExcelUploadAndValidate}
      onSaveDraft={handleSaveDraft}
      onSubmit={handleSubmit}
      draftId={draftId}
      textFieldsDisabled={false}

      searchSku={searchSku}
      setSearchSku={setSearchSku}
      selectedExcelFile={selectedExcelFile}
      skuSearchError={skuSearchError}
      onManualSkuSearch={handleManualSkuSearch}
      manualPopupOpen={manualPopupOpen}
      setManualPopupOpen={setManualPopupOpen}
      manualQty={manualQty}
      setManualQty={setManualQty}
      manualQtyError={manualQtyError}
      manualMaxStock={manualMaxStock}
      manualProduct={manualProduct}
      setManualProduct={setManualProduct}
      onAddManualItem={handleAddManualItem}
      submitError={submitError}
      successOpen={successOpen}
      successTxnId={successTxnId}
      successText={successText}
      closeSuccessDialog={closeSuccessDialog}
      errorOpen={errorOpen}
      errorTitle={errorTitle}
      errorMessage={errorMessage}
      closeErrorDialog={closeErrorDialog}
      excelHelpOpen={excelHelpOpen}
      openExcelHelp={() => setExcelHelpOpen(true)}
      closeExcelHelp={() => setExcelHelpOpen(false)}
      userData={userData}
      aadProfile={aadProfile}
      pageBg={pageBg}
      container={container}
      card={card}
      onDiscard={handleDiscard}
      onNavigate={(route, state) => onNavigate(route, state)}
      setUsageField={setUsageField}
      excelUploadedAt={excelUploadedAt}
      setInventoryModeError={setInventoryModeError}
      runExcelUploadAndValidate={runExcelUploadAndValidate}
      setSkuSearchError={setSkuSearchError}
      handleManualSkuSearch={handleManualSkuSearch}
      handleAddManualItem={handleAddManualItem}
      maxRowsPopupOpen={maxRowsPopupOpen}
      maxRowsPopupMsg={maxRowsPopupMsg}
      closeMaxRowsPopup={() => {
        setMaxRowsPopupOpen(false);
        setMaxRowsPopupMsg("");
      }}
      setSelectedExcelFile={setSelectedExcelFile}
      setExcelFileName={setExcelFileName}
      setExcelValidated={setExcelValidated}
      setExcelErrors={setExcelErrors}
      setItemErrors={setItemErrors}
      setManualQtyError={setManualQtyError}
      draftLoading={draftLoading}
      purposes={purposes}
      costCenters={costCenters}
      context={context}
      draftDisabled={items.length === 0 || (isEditDraftMode ? !hasDraftChanges : false)}
      discardConfirmOpen={discardConfirmOpen}
      resetForm={resetForm}
      setDiscardConfirmOpen={setDiscardConfirmOpen}
      setExcelUploadedAt={setExcelUploadedAt}
      customErrorOpen={customErrorOpen}
      customErrorMessage={customErrorMessage}
      setCustomErrorOpen={setCustomErrorOpen}
      setCustomErrorMessage={setCustomErrorMessage}
      isProcessing={isProcessing}
      setIsProcessing={setIsProcessing}
      excelFailPopupOpen={excelFailPopupOpen}
      setExcelFailPopupOpen={setExcelFailPopupOpen}
      manualStockPopupOpen={manualStockPopupOpen}
      // In CreateRequestPage.tsx — replace the existing prop
      closeManualStockPopup={() => {
        setManualStockPopupOpen(false);
        // Defer scroll to allow popup to unmount and DOM to settle
        setTimeout(() => {
          const firstErrorRow =
            document.querySelector<HTMLTableRowElement>('tr[data-row-error="true"]');
          if (firstErrorRow) {
            firstErrorRow.scrollIntoView({ behavior: "smooth", block: "center" });
          }
        }, 50);
      }}
      setItems={setItems}
      setPendingMode={setPendingMode}
      onExcelFailOk={() => {
        setExcelFailPopupOpen(false);
        handleSubmit({ allowExcelInsufficient: true }).catch((e: unknown) => {
          console.error("Submit failed:", e);
          const message = e instanceof Error ? e.message : "Submit failed. Please try again.";
          setCustomErrorMessage(message);
          setCustomErrorOpen(true);
        });
      }}
    />
  );
}
