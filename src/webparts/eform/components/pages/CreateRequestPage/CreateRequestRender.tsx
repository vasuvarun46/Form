import * as React from "react";
import CreateRequestHeader from "../../Common/CreateRequestHeader";
import GeneralInfoSection from "./sections/GeneralInfoSection";
import UsageInfoSection from "./sections/UsageInfoSection";
import RadioRow from "./Component/RadioRow";
import QtyStepper from "./Component/QtyStepper";
import { IconSearch } from "../../ui/icons";
import { Route } from "../../Eform";
import { RouteState } from "../../spHelpers";
import { downloadTemplateFromLibrary, FlowItemResponse, ISuparaItem, IUsageState, AadUserProfile } from "./CreateRequestService";
import { CostCenterOption, IItemRow, InventoryMode, PurposeOption } from "./createRequestTypes";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Loading from "../../ui/Loading";
import { createExcelPendingRequest } from "./CreateRequestService";
import TagIcon from "../../../assets/Tagicon.png";
import { Icon } from "@fluentui/react";

type TdProps = {
  children: React.ReactNode;
  right?: boolean;
  center?: boolean;
};
type ThProps = {
  children: React.ReactNode;
  right?: boolean;
  center?: boolean;
};
type Props = {
  // layout
  pageBg: string;
  container: string;
  card: string;
  // header
  onDiscard: () => void;
  onSaveDraft: () => void;
  onSubmit: () => void;
  onNavigate: (route: Route, state?: RouteState) => void;
  purposes: PurposeOption[];        // ✅ correct
  costCenters: CostCenterOption[];
  // state flags
  isSubmitting: boolean;

  // general info
  userData: ISuparaItem[];
  aadProfile?: AadUserProfile | null;

  // usage
  usage: IUsageState;
  setUsageField: <K extends keyof IUsageState>(k: K, v: IUsageState[K]) => void;
  fieldErrors: Record<string, string>;

  // inventory mode
  inventoryMode: InventoryMode | "";
  inventoryModeError: boolean;
  setInventoryMode: (v: InventoryMode | "") => void;
  setInventoryModeError: (v: boolean) => void;

  // excel
  selectedExcelFile?: File;
  excelFileName: string;
  excelValidated: boolean;
  excelErrors: string[];
  isValidatingExcel: boolean;
  runExcelUploadAndValidate: () => Promise<{
    ok: boolean;
    finalItems: IItemRow[];
    totalQty: number;
    totalAmount: number;
  }>;

  // sku search
  searchSku: string;
  isSearchingSku: boolean;
  skuSearchError: string;
  setSkuSearchError: (v: string) => void;
  setSearchSku: (v: string) => void;
  handleManualSkuSearch: () => Promise<void>;

  // manual popup
  manualPopupOpen: boolean;
  manualProduct?: FlowItemResponse;
  manualQty: number;
  manualQtyError: string;
  manualMaxStock: number;
  handleAddManualItem: () => void;
  maxRowsPopupOpen: boolean;
  maxRowsPopupMsg: string;
  closeMaxRowsPopup: () => void;
  customErrorOpen: boolean;
  customErrorMessage: string;
  setCustomErrorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCustomErrorMessage: React.Dispatch<React.SetStateAction<string>>;
  totalQty: number;
  // items table
  items: IItemRow[];
  itemErrors: Record<number, string>;
  updateQty: (idx: number, v: string | number) => void;
  removeItem: (idx: number) => void;
  totalAmount: number;
  fileInputRef: React.RefObject<HTMLInputElement>;
  // dialogs
  successOpen: boolean;
  successTxnId: string;
  successText: string;
  closeSuccessDialog: () => void;
  excelUploadedAt: string;
  submitError: string;
  errorOpen: boolean;
  draftLoading: boolean;
  errorTitle: string;
  draftId?: number;
  textFieldsDisabled?: boolean; // ✅ NEW
  onManualSkuSearch: () => void;
  onAddManualItem: () => void;
  removeConfirmOpen: boolean;
  confirmRemoveItem: () => void;
  cancelRemoveItem: () => void;
  errorMessage: string;
  excelHelpOpen: boolean;
  openExcelHelp: () => void;
  closeExcelHelp: () => void;
  removeSku: string;
  openRemoveConfirm: (idx: number, sku: string) => void;
  context: WebPartContext;
  onSelectExcelFile: (file: File | undefined) => void;
  setExcelUploadedAt: (v: string) => void;
  onRunExcelUploadAndValidate: () => Promise<{
    ok: boolean;
    finalItems: IItemRow[];
    totalQty: number;
    totalAmount: number;
  }>;
  closeErrorDialog: () => void;
  setSelectedExcelFile: (f: File | undefined) => void;
  setExcelFileName: (v: string) => void;
  setExcelValidated: (v: boolean) => void;
  setExcelErrors: (v: string[]) => void;
  setItemErrors: React.Dispatch<React.SetStateAction<Record<number, string>>>;
  setManualQty: (v: number) => void;
  setManualQtyError: (v: string) => void;
  setManualPopupOpen: (v: boolean) => void;
  draftDisabled: boolean;
  setManualProduct: React.Dispatch<
    React.SetStateAction<FlowItemResponse | undefined>
  >;
  discardConfirmOpen: boolean;
  resetForm: () => void;
  setDiscardConfirmOpen: React.Dispatch<React.SetStateAction<boolean>>;
  isProcessing: boolean;  // Add this line to accept isProcessing as a prop
  setIsProcessing: React.Dispatch<React.SetStateAction<boolean>>;
  excelFailPopupOpen: boolean;
  setExcelFailPopupOpen: React.Dispatch<React.SetStateAction<boolean>>;
  onExcelFailOk: () => void;
  manualStockPopupOpen: boolean;
  closeManualStockPopup: () => void;
  setItems: React.Dispatch<React.SetStateAction<IItemRow[]>>;

  modeChangeConfirmOpen: boolean;
  setModeChangeConfirmOpen: React.Dispatch<React.SetStateAction<boolean>>;
  pendingMode: InventoryMode | "";
  applyModeChange: (mode: InventoryMode) => void;
  setPendingMode: React.Dispatch<React.SetStateAction<InventoryMode | "">>;


};
export default function CreateRequestRender(props: Props): JSX.Element {
  const {
    pageBg,
    container,
    card,
    onDiscard,
    onSaveDraft,
    onSubmit,
    onNavigate,
    isSubmitting,
    draftLoading,
    userData,
    aadProfile,
    usage,
    setUsageField,
    purposes,
    costCenters,
    fieldErrors,
    inventoryMode,
    inventoryModeError,
    setInventoryMode,
    setInventoryModeError,
    selectedExcelFile,
    excelFileName,
    excelValidated,
    excelErrors,
    isValidatingExcel,
    //runExcelUploadAndValidate,
    searchSku,
    isSearchingSku,
    skuSearchError,
    setSkuSearchError,
    setSearchSku,
    handleManualSkuSearch,
    manualPopupOpen,
    manualProduct,
    manualQty,
    manualQtyError,
    manualMaxStock,
    handleAddManualItem,
    maxRowsPopupOpen,
    maxRowsPopupMsg,
    closeMaxRowsPopup,
    items,
    itemErrors,
    updateQty,
    successOpen,
    successTxnId,
    successText,
    closeSuccessDialog,
    errorOpen,
    errorTitle,
    errorMessage,
    closeErrorDialog,
    setSelectedExcelFile,
    setExcelFileName,
    setExcelValidated,
    setExcelErrors,
    setManualQty,
    setManualQtyError,
    setManualPopupOpen,
    setManualProduct,
    fileInputRef,
    removeConfirmOpen,
    confirmRemoveItem,
    cancelRemoveItem,
    removeSku,
    openRemoveConfirm,
    excelHelpOpen,
    openExcelHelp,
    closeExcelHelp,
    setItemErrors,
    context,
    totalQty,
    totalAmount,
    discardConfirmOpen,
    setDiscardConfirmOpen,
    draftId,
    textFieldsDisabled,
    // excelUploadedAt,
    setExcelUploadedAt,
    onSelectExcelFile,
    customErrorOpen,
    customErrorMessage,
    setCustomErrorOpen,
    setCustomErrorMessage,
    isProcessing,
    setIsProcessing,
    excelFailPopupOpen,
    setExcelFailPopupOpen,
    onExcelFailOk,
    manualStockPopupOpen,
    closeManualStockPopup,
    setItems,
    modeChangeConfirmOpen,
    setModeChangeConfirmOpen,
    pendingMode,
    applyModeChange,
    setPendingMode,
    draftDisabled,

  } = props;

  const [removeFileConfirmOpen, setRemoveFileConfirmOpen] = React.useState(false);
  const isDownloadingTemplateRef = React.useRef(false);
  const [searchBarHovered, setSearchBarHovered] = React.useState(false);
  const [searchBtnHovered, setSearchBtnHovered] = React.useState(false);
  const productTableRef = React.useRef<HTMLTableElement>(null);
  const [fp, setFp] = React.useState<{ qL: number; qW: number; tL: number; tW: number } | null>(null);
  React.useEffect(() => {
    const u = () => {
      const t = productTableRef.current; if (!t) return; const h = t.querySelectorAll("thead tr th"); if (!h.length) return; const n = h.length; const qi = n >= 10 ? 7 : 6; const ti = n >= 10 ? 8 : 7; const q = h[qi].getBoundingClientRect(); const r = h[ti].getBoundingClientRect(); setFp(prev => {
        const next = { qL: q.left, qW: q.width, tL: r.left, tW: r.width };
        if (prev && prev.qL === next.qL && prev.qW === next.qW &&
          prev.tL === next.tL && prev.tW === next.tW) return prev;
        return next;
      });
    };
    u(); window.addEventListener("resize", u); const s = document.querySelector(".flex-1.min-h-0.overflow-auto"); if (s) s.addEventListener("scroll", u); const id = setInterval(u, 300);
    return () => { window.removeEventListener("resize", u); if (s) s.removeEventListener("scroll", u); clearInterval(id); };
  }, [items, inventoryMode]);
  const mandatoryCompleted =
    !!usage.purpose &&
    !!usage.costCenter &&
    !!usage.expectedDelivery &&
    !!usage.placeOfDelivery &&
    !!usage.purposeDetail;
  // Add this after the existing const declarations (around line where mandatoryCompleted is defined)
  const isExcelMode = inventoryMode === "excel" && items.length > 0;

  const onDownloadTemplate = async (): Promise<void> => {
    if (isDownloadingTemplateRef.current) return;
    isDownloadingTemplateRef.current = true;
    try {
      // ❌ remove global loader for download
      await downloadTemplateFromLibrary(context, "ExcelTemplate", { extensionFilter: ".xlsx" });
    } catch (e) {
      console.error(e);
    } finally {
      const ref = isDownloadingTemplateRef;  // ✅
      ref.current = false;
    }
  };



  // AFTER
  const Th = ({ children, right, center }: ThProps): JSX.Element => (
    <th
      className={[
        "px-2 py-2.5 sm:px-3 sm:py-3 text-[12px] sm:text-[13px] font-semibold leading-[1.35] text-gray-700",
        right ? "text-right whitespace-nowrap" : center ? "text-center" : "text-left",
      ].join(" ")}
    >
      <div className={right ? "text-right" : center ? "text-center" : "text-left"}>
        {children}
      </div>
    </th>
  );

  const Td = ({ children, right, center }: TdProps): JSX.Element => (
    <td
      className={[
        "px-2 py-2.5 sm:px-3 sm:py-3 text-[12px] sm:text-[13px] leading-[1.4] text-gray-700",
        right ? "text-right" : center ? "text-center" : "text-left",
      ].join(" ")}
    >
      {children}
    </td>
  );
  const [uploadConfirmOpen, setUploadConfirmOpen] = React.useState<boolean>(false);
  const [submitConfirmOpen, setSubmitConfirmOpen] = React.useState<boolean>(false);

  const showExcelNotice = inventoryMode === "excel" && excelValidated;
  const hasInvalidManualQty = inventoryMode === "manual" && items.some((it, idx) => {
    const qty = Number(it.qty ?? 0);
    if (!Number.isFinite(qty) || qty < 1 || qty > 999) return true;
    const maxStock = Number(it.maxStock ?? 0);
    if (maxStock > 0 && qty > maxStock) return true;
    return !!itemErrors[idx];
  });

  const clampQty = (v: number) => Math.max(0, v);
  const getManualQtyError = (qty: number, maxStock: number): string => {
    if (!Number.isFinite(qty) || qty < 1 || qty > 999 || qty > maxStock) {
      return `Quantity must be 1–999 and not exceed\navailable stock (${maxStock} pcs).`;
    }
    return "";
  };

  return (
    <div className={`${pageBg} pdf-thai flex flex-col h-full min-h-0`}>
      {/* HEADER */}
      <div className="sticky top-0 z-20 bg-white shrink-0">

        {/* HEADER */}
        <CreateRequestHeader
          onDiscard={() => setDiscardConfirmOpen(true)}
          onSaveDraft={onSaveDraft}
          draftDisabled={draftDisabled || hasInvalidManualQty}
          submitDisabled={items.length === 0 || isSubmitting || hasInvalidManualQty}
          hideSaveDraft={inventoryMode === "excel"}
          onSubmit={() => setSubmitConfirmOpen(true)}
          onNavigate={(route) => onNavigate(route)}
          showExcelNotice={showExcelNotice}
          excelNoticeText="This request was created by Excel upload and cannot be edited."
        />
      </div>
      {/* SCROLL AREA */}
      <div className="flex-1 min-h-0 overflow-auto">
        <div className={`${container} mx-auto w-full px-3 sm:px-6 py-4 space-y-4 pb-16`}>

          {/* SUBMITTING BAR */}
          {isSubmitting && (
            <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
              <div className="w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7]">
                <div className="px-6 py-8 text-center">

                  {/* Spinner — gray to match screenshot */}
                  <div className="flex justify-center mb-5">
                    <div className="h-12 w-12 border-4 border-gray-200 border-t-gray-500 rounded-full animate-spin" />
                  </div>

                  {/* Title */}
                  <div className="text-[16px] font-semibold text-gray-900 mb-3">
                    Please wait...
                  </div>

                  {/* Description — matches screenshot 2 exactly */}
                  <div className="text-[13px] text-gray-500 leading-relaxed">
                    Checking available stocks.<br />
                    This may take more than 5 minutes for 300 SKUs<br />
                    upload. Please do not close this page.
                  </div>

                </div>
              </div>
            </div>
          )}

          {modeChangeConfirmOpen && (
            <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/30">
              <div className="w-[360px] bg-white rounded-2xl shadow-xl border border-[#e8e8e8] relative">

                {/* Close icon */}
                <button
                  type="button"
                  className="absolute right-4 top-4 flex h-7 w-7 items-center justify-center text-[#b0b0b0] hover:text-[#888]"
                  onClick={() => {
                    setModeChangeConfirmOpen(false);
                    setPendingMode("");
                  }}
                >
                  <svg className="h-5 w-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                  </svg>
                </button>

                <div className="px-7 pt-8 pb-7 text-center">

                  {/* Title */}
                  <div className="text-[15px] font-semibold text-gray-900 leading-snug mb-3">
                    {pendingMode === "excel"
                      ? "Changing to bulk load using Excel template?"
                      : "Changing to manual key in?"}
                  </div>

                  {/* Description */}
                  <div className="text-[10px] text-gray-500 leading-relaxed mb-6">
                    {pendingMode === "excel"
                      ? "Items currently added will not be saved."
                      : "The selected file will not be saved."}
                    <br />
                    Please confirm if you want to proceed.
                    <br />
                    <br />
                  </div>

                  {/* Buttons */}
                  <div className="flex items-center justify-center gap-3">
                    <button
                      className="h-10 px-6 rounded-lg bg-[#E8E8E8] text-gray-700 text-[13px] font-medium hover:bg-[#EBEBEB]"
                      onClick={() => {
                        setModeChangeConfirmOpen(false);
                        setPendingMode("");
                      }}
                    >
                      Cancel
                    </button>

                    <button
                      className="h-10 px-6 rounded-lg bg-black text-white text-[13px] font-semibold whitespace-nowrap hover:bg-[#3B3B3B]"
                      onClick={() => {
                        setModeChangeConfirmOpen(false);
                        if (pendingMode) {
                          applyModeChange(pendingMode);
                          setPendingMode("");
                        }
                      }}
                    >
                      {pendingMode === "excel"
                        ? "Change to upload"
                        : "Change to Manual Key In"}
                    </button>
                  </div>

                </div>
              </div>
            </div>
          )}


          {draftLoading && (
            <Loading />
          )}
          <GeneralInfoSection userData={userData} card={card} aadProfile={aadProfile} />
          <UsageInfoSection
            usage={usage}
            setField={setUsageField}
            fieldErrors={fieldErrors}
            purposes={purposes}
            costCenters={costCenters}
            disabled={isExcelMode}
            context={context}
            textFieldsDisabled={textFieldsDisabled}
          />

          <section className={`${card} p-4 sm:p-5`} >
            <div className="flex items-center gap-4 mb-4">
              <h2
                className="text-[13px] font-bold whitespace-nowrap text-gray-900"

              >
                Inventory Detail
              </h2>

              {inventoryModeError || !mandatoryCompleted && (
                <div className="
      flex items-center gap-2
      rounded-md border border-[#F5D36E]
      bg-[#FFF1B9] px-3 py-1
      text-[11px] text-[#3E3C35]
      whitespace-nowrap
    ">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" className="shrink-0">
                    <circle cx="12" cy="12" r="10" stroke="#3E3C35" strokeWidth="2" fill="none" />
                    <line x1="12" y1="8" x2="12" y2="13" stroke="#3E3C35" strokeWidth="2" strokeLinecap="round" />
                    <circle cx="12" cy="16.5" r="1.2" fill="#3E3C35" />
                  </svg>
                  Please complete all required fields above to select how to add a product.

                </div>
              )}
            </div>

            <div className="space-y-2" >


              <RadioRow
                title="Manual Key In"
                sub="To manually added requested item one-by-one."
                checked={inventoryMode === "manual"}
                onChange={() => {
                  setInventoryMode("manual");
                  setInventoryModeError(false);
                }}
                disabled={!mandatoryCompleted || isExcelMode}


              />

              <RadioRow
                title="Excel File Upload"
                disabled={!mandatoryCompleted}
                sub={
                  <>
                    To add multiple products by uploading Excel Template.{" "}
                    <span
                      className="text-[#6099E3] cursor-pointer underline"
                      onClick={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onDownloadTemplate().catch((e) => console.error(e));
                      }}

                    >
                      Get Excel template.
                    </span>



                  </>
                }
                checked={inventoryMode === "excel"}
                onChange={() => {
                  setInventoryMode("excel");
                  setInventoryModeError(false);
                }}


              />

            </div>
          </section>

          <section className={`${card} p-4 sm:p-5`}>
            {/* ✅ Excel Upload Box */}
            {inventoryMode === "excel" && items.length === 0 && (

              <div className="mb-4 rounded-md border border-[#e7e7e7] bg-white p-4">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx"
                  className="hidden"
                  onChange={(e) => {
                    const f =
                      e.target.files && e.target.files.length
                        ? e.target.files[0]
                        : undefined;
                    onSelectExcelFile(f);
                  }}

                />

                <div className="flex items-center justify-between">
                  <div className="text-[12px] font-bold text-gray-900">Product</div>


                  {selectedExcelFile ? (
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setRemoveFileConfirmOpen(true)}

                        className="h-8 px-4 rounded-md border border-[#DA2128] bg-[#FBE9EA] text-[#DA2128] text-[12px] font-semibold hover:bg-[#DA2128] hover:text-white"
                      >
                        Remove file
                      </button>

                      <button
                        type="button"
                        disabled={isValidatingExcel}
                        onClick={() => {
                          setUploadConfirmOpen(true);
                        }}
                        className={[
                          "h-8 px-4 rounded-md text-[12px] font-semibold",
                          "bg-[#000000] text-white hover:bg-[#3B3B3B]",
                        ].join(" ")}
                      >
                        Upload file
                      </button>
                    </div>
                  ) : (
                    <button
                      type="button"
                      onClick={() => fileInputRef.current?.click()}
                      className="h-9 px-5 rounded-md border-[1.5px] border-[#478AE0] bg-[#2172DA] text-white text-[12px] font-bold hover:bg-[#F0F7FF] hover:text-[#2172DA]"
                    >
                      Select file
                    </button>
                  )}


                </div>

                <div className="mt-3 rounded-md bg-[#f7f7f7] p-3 text-[12px] text-gray-700">
                  <div className="text-[11px] text-gray-500">Selected file</div>
                  <div className="mt-1">
                    {selectedExcelFile ? (
                      <span className="inline-flex items-center gap-2">
                        <Icon
                          iconName="ExcelDocument"
                          style={{ fontSize: 18, color: "#217346" }}
                        />
                        <span className="text-[12px] text-gray-800 truncate max-w-[200px]">
                          {selectedExcelFile.name}
                        </span>

                      </span>
                    ) : (
                      <span className="text-gray-500">No file selected</span>
                    )}
                  </div>
                </div>

                <button
                  type="button"
                  onClick={openExcelHelp}
                  className="mt-2 text-[11px] text-[#6099E3] hover:underline"

                >
                  See how to upload the Excel file successfully
                </button>

                {isValidatingExcel && inventoryMode === "excel" && (
                  <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
                    <div className="w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7]">
                      <div className="px-6 py-6 text-center">
                        <div className="flex justify-center mb-4">
                          <div className="h-10 w-10 border-4 border-gray-200 border-t-blue-600 rounded-full animate-spin" />
                        </div>

                        <div className="text-[15px] font-semibold text-gray-900 mb-2">
                          Please wait...
                        </div>

                        <div className="text-[12px] text-gray-600">
                          Validating Excel (max 300 rows)...
                        </div>
                      </div>
                    </div>
                  </div>
                )}


                {excelErrors.length > 0 && (
                  <div className="mt-2 text-[11px] text-[#E56D71] space-y-1">
                    {excelErrors.map((x, i) => (
                      <div key={i}>• {x}</div>
                    ))}
                  </div>
                )}
              </div>
            )}
            {inventoryMode === "manual" && (
              <>
                <div className="text-[13px] font-bold text-gray-900 mb-4">Product</div>


                {/* ✅ Search bar */}
                <div className="w-[560px] max-w-full">
                  <div className={[
                    "h-10 w-full flex items-stretch overflow-hidden rounded-md border-2 bg-white transition-colors",
                    searchSku
                      ? "border-[#76A8E8]"
                      : searchBarHovered
                        ? "border-[#76A8E8]"
                        : "border-[#e6e6e6]",
                  ].join(" ")}>

                    <div
                      className="relative flex-1"
                      onMouseEnter={() => setSearchBarHovered(true)}
                      onMouseLeave={(e) => {
                        const related = e.relatedTarget as HTMLElement | null;
                        if (related && related.closest("[data-search-btn]")) return;
                        setSearchBarHovered(false);
                      }}
                    >

                      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#B1B1B1]">
                        <IconSearch className="h-5 w-5" />
                      </span>

                      <input
                        placeholder="Search by SKU e.g. 8851512119283"
                        value={searchSku}
                        maxLength={20}
                        disabled={isSearchingSku}
                        onChange={(e) => {
                          const value = e.target.value
                            .replace(/[^a-zA-Z0-9\-/_ ]/g, "")
                            .slice(0, 20);

                          setSearchSku(value);
                          setSkuSearchError("");
                        }}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            e.preventDefault();
                            handleManualSkuSearch().catch(() => undefined);
                          }
                        }}
                        className="h-full w-full bg-transparent pl-9 pr-8 text-[12px] text-[#7A7A7A] font-bold outline-none placeholder:text-[#8B8B8B] placeholder:font-normal"
                      />
                      {searchSku && !isSearchingSku && (
                        <button
                          type="button"
                          onClick={() => {
                            setSearchSku("");
                            setSkuSearchError("");
                          }}
                          className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center justify-center w-6 h-6 rounded-full bg-[#FFFFFF] hover:bg-[#F6F6F6] text-[#B1B1B1]"
                          aria-label="Clear search"
                        >
                          <svg className="w-4 h-5" viewBox="0 0 12 12" fill="none">
                            <path d="M2 2L10 10M10 2L2 10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
                          </svg>
                        </button>
                      )}

                    </div>
                    <button
                      type="button"
                      data-search-btn
                      disabled={isSearchingSku}
                      onClick={() => handleManualSkuSearch().catch(() => undefined)}
                      onMouseEnter={() => setSearchBtnHovered(true)}
                      onMouseLeave={() => { setSearchBtnHovered(false); setSearchBarHovered(false); }}
                      style={{
                        borderLeft: `2px solid ${(searchBarHovered || searchSku) ? "#76A8E8" : "transparent"}`,
                      }}
                      className={[
                        "h-full px-6 text-[12px] font-semibold transition-colors",
                        (!searchSku && searchBarHovered && !searchBtnHovered) || (searchSku && searchBtnHovered)
                          ? "bg-[#F0F7FF] text-[#2172DA]"
                          : "bg-[#2172DA] text-white",
                        isSearchingSku ? "opacity-60 cursor-not-allowed" : "",
                      ].join(" ")}
                    >

                      {isSearchingSku ? "Searching..." : "Search"}
                    </button>
                  </div>

                  {skuSearchError && (
                    <div className="relative w-full">
                      <div className="absolute left-0 top-0 z-30 w-full min-h-[70px] rounded-b-md border border-t-0 border-[#e7e7e7] bg-[#F8F8F8] shadow-md px-4 py-5 flex items-center">
                        <span className="text-[12px] text-[#7A7A7A]">
                          {skuSearchError.split(/(".*?")/).map((part, i) =>
                            part.startsWith('"') && part.endsWith('"')
                              ? <span key={i} className="font-bold">{part}</span>
                              : part
                          )}
                        </span>
                      </div>
                    </div>
                  )}

                </div>
              </>
            )}



            {/* ✅ Manual Add Popup */}
            {inventoryMode === "manual" && manualPopupOpen && manualProduct && (
              <div className="relative h-0 w-[560px] max-w-full">
                <div className="absolute left-0 top-0 z-20 w-full overflow-hidden rounded-md border border-[#e7e7e7] bg-[#F8F8F8] shadow-sm">
                  <div className="bg-[#F8F8F8] px-4 pt-3 pb-4">
                    <table className="w-full text-[12px]">
                      <thead className="bg-[#F8F8F8] text-gray-700">
                        <tr>
                          <th className="text-left text-black font-bold py-2">SKU</th>
                          <th className="text-center text-black font-bold py-2">In Stock</th>
                          <th className="text-center text-black font-bold py-2">
                            <span className="text-[#E56D71] ">*</span> Request Quantity

                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="border-t border-[#e6e6e6]" style={{ backgroundColor: "#F8F8F8" }}>
                          <td className="py-3 text-gray-800" style={{ backgroundColor: "#F8F8F8" }}>{String(manualProduct.Barcode || "-")}</td>

                          <td className="py-3 text-center text-gray-800" style={{ backgroundColor: "#F8F8F8" }}>
                            {Number(manualProduct.Warehouse_SOH_Qty ?? 0)}
                          </td>

                          <td className="py-3" style={{ backgroundColor: "#F8F8F8" }}>
                            <div className="mx-auto flex w-[120px] flex-col items-start bg-[#F8F8F8]">
                              <QtyStepper
                                value={manualQty}
                                onDec={() => {
                                  const next = clampQty(Number(manualQty || 0) - 1);
                                  setManualQty(next);
                                  setManualQtyError(getManualQtyError(next, manualMaxStock));
                                }}
                                onInc={() => {
                                  const next = clampQty(Number(manualQty || 0) + 1);
                                  setManualQty(next);
                                  setManualQtyError(getManualQtyError(next, manualMaxStock));
                                }}
                                onChange={(v) => {
                                  const next = clampQty(Number(v || 0));
                                  setManualQty(next);
                                  setManualQtyError(getManualQtyError(next, manualMaxStock));
                                }}
                                emptyValue={0}
                                maxStock={manualMaxStock}
                              />

                              {manualQtyError && (
                                <div className="mt-2 w-[170px] max-w-none text-left text-[9px] leading-[1.35] text-[#E56D71] whitespace-pre-line">

                                  {manualQtyError}
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <div className="border-t border-[#e6e6e6] bg-white px-4 py-3">
                    <div className="flex justify-end gap-3">
                      <button
                        type="button"
                        onClick={() => {
                          setManualPopupOpen(false);
                          setManualProduct(undefined);
                          setManualQty(1);
                          setManualQtyError("");
                        }}
                        className="h-8 min-w-[88px] rounded-md bg-[#E8E8E8] px-4 text-[12px] font-semibold text-[#222222] hover:bg-[#EBEBEB]"

                      >
                        Cancel
                      </button>

                      <button
                        type="button"
                        onClick={handleAddManualItem}
                        disabled={
                          Number(manualProduct.Warehouse_SOH_Qty ?? 0) <= 0 ||
                          manualQty < 1 ||
                          manualQty > 999 ||
                          manualQty > manualMaxStock ||
                          !!manualQtyError
                        }
                        className={[
                          "h-8 min-w-[76px] rounded-md px-4 text-[12px] font-semibold",
                          Number(manualProduct.Warehouse_SOH_Qty ?? 0) <= 0 ||
                            manualQty < 1 ||
                            manualQty > 999 ||
                            manualQty > manualMaxStock ||
                            !!manualQtyError
                            ? "bg-[#d9d9d9] text-[#888888] cursor-not-allowed"
                            : "bg-black text-white hover:bg-[#1f1f1f]",
                        ].join(" ")}
                      >
                        Add
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}


            {!isProcessing &&
              items.length === 0 &&
              (inventoryMode === "" || inventoryMode === "manual") && (
                <div>
                  {/* ✅ Product header shown above empty state */}
                  {inventoryMode === "" && (
                    <div className="text-[13px] font-bold text-gray-900 mb-4">Product</div>

                  )}
                  <div className="flex flex-col items-center justify-center py-16 text-center">
                    <div className="mb-4 flex justify-center">
                      <img
                        src={TagIcon}
                        alt=""
                        aria-hidden="true"
                        className="h-7 w-7 object-contain scale-[4] origin-center"
                      />
                    </div>


                    <div className="text-sm font-bold text-gray-900">

                      No products added in this request
                    </div>

                    <div className="mt-1 text-sm font-semibold text-[#8C8C8C] max-w-md">

                      The request must contain at least one SKU with requested quantity to
                      Save as Draft or Submit.
                    </div>
                  </div>
                </div>
              )}
            {items.length > 0 && inventoryMode === "manual" && (
              <div className="mt-4 overflow-hidden rounded-lg border border-[#e7e7e7] bg-white mx-auto">
                <table ref={productTableRef} className="w-full table-auto text-[12px] sm:text-[13px] text-left">
                  <thead className="bg-[#f7f7f7] text-gray-600">
                    <tr>
                      <Th><div>ลำดับ</div><div>No.</div></Th>
                      <Th><div>รหัสสินค้า</div><div>SKU Code</div></Th>
                      <Th><div>กลุ่ม</div><div>Product Code</div></Th>
                      <Th><div>สี</div><div>Color</div></Th>
                      <Th><div>ขนาด</div><div>Size</div></Th>
                      <Th right>
                        <div>ราคาป้ายต่อหน่วย</div>
                        <div>Tag Price per Unit</div>
                      </Th>
                      <Th center><div>จำนวน</div><div>Quantity</div></Th>
                      <Th right>
                        <div>ราคาป้ายทั้งหมด</div>
                        <div>Total Tag Price</div>
                      </Th>
                      {inventoryMode === "manual" && <Th>Action</Th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#eeeeee] bg-white">
                    {items.map((it, idx) => {
                      const rowTotal: number = Number(it.qty) * Number(it.unitPrice);
                      const rowErr: string | undefined = itemErrors[idx];
                      return (
                        // AFTER
                        <tr key={`${it.sku}-${idx}`} className="">

                          <Td>{it.no}.</Td>
                          <Td>
                            <span className="block truncate" title={it.sku}>{it.sku}</span>
                          </Td>
                          <Td>
                            <span className="block truncate" title={it.productCode}>{it.productCode}</span>
                          </Td>
                          <Td>
                            <span className="block truncate" title={it.color}>{it.color}</span>
                          </Td>
                          <Td>{it.size}</Td>
                          <Td right>{Number(it.unitPrice).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</Td>
                          <Td center>
                            <div className="origin-center">
                              <QtyStepper
                                value={it.qty}
                                onDec={() => updateQty(idx, Number(it.qty) - 1)}
                                onInc={() => updateQty(idx, Number(it.qty) + 1)}
                                onChange={(v) => updateQty(idx, v)}
                                maxStock={it.maxStock}
                              />
                            </div>
                            {rowErr && (
                              <div className="mt-1 mx-auto w-[136px] text-[11px] leading-[1.25] text-[#E56D71] text-left whitespace-normal break-words">

                                {rowErr}
                              </div>
                            )}
                          </Td>
                          <Td right>{rowTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</Td>
                          {inventoryMode === "manual" && (
                            <Td>
                              <button
                                type="button"
                                onClick={() => openRemoveConfirm(idx, it.sku)}
                                className="h-8 sm:h-9 px-3 sm:px-4 whitespace-nowrap rounded-md text-[11px] sm:text-[12px] font-semibold transition-all duration-200 !border !border-[#DA2128] !bg-[#FBE9EA] !text-[#DA2128] hover:!bg-[#DA2128] hover:!text-white hover:!border-[#DA2128]"
                              >
                                <span className="inline-flex items-center gap-1">
                                  <Icon iconName="Delete" className="text-[11px] leading-none" />
                                  <span>Remove</span>
                                </span>
                              </button>
                            </Td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {inventoryMode === "excel" && items.length > 0 && draftId && (
              <div className="mb-4 rounded-md border border-[#e7e7e7] bg-white p-4">
                <div className="mt-3 rounded-md bg-[#f7f7f7] p-3 text-[12px] text-gray-700">
                  <div className="text-[11px] text-gray-500">Selected file</div>
                  <div className="mt-1">
                    {selectedExcelFile ? (
                      <span className="inline-flex items-center gap-2">
                        <Icon
                          iconName="ExcelDocument"
                          style={{ fontSize: 18, color: "#217346" }}
                        />
                        <span className="text-[12px] text-gray-800 truncate max-w-[200px]">
                          {selectedExcelFile.name}
                        </span>
                      </span>
                    ) : (
                      <span className="text-gray-500">No file selected</span>
                    )}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={openExcelHelp}
                  className="mt-2 text-[11px] text-[#6099E3] hover:underline"
                >
                  See how to upload the Excel file successfully
                </button>
              </div>
            )}

            {items.length > 0 && inventoryMode === "excel" && draftId && (
              <div className="mt-4 overflow-hidden rounded-lg border border-[#e7e7e7] bg-white mx-auto">
                <table ref={productTableRef} className="w-full table-auto text-[12px] sm:text-[13px] text-left">
                  <colgroup>
                    <col style={{ width: "5%" }} />   {/* No. */}
                    <col style={{ width: "16%" }} />  {/* SKU Code */}
                    <col style={{ width: "15%" }} />  {/* Product Code */}
                    <col style={{ width: "10%" }} />  {/* Color */}
                    <col style={{ width: "8%" }} />   {/* Size */}
                    <col style={{ width: "14%" }} />  {/* Tag Price per Unit */}
                    <col style={{ width: "14%" }} />  {/* Quantity */}
                    <col style={{ width: "18%" }} />  {/* Total Tag Price */}
                  </colgroup>
                  <thead className="bg-[#f7f7f7] text-gray-600">
                    <tr>
                      <Th>
                        <div>ลำดับ</div>
                        <div>No.</div>
                      </Th>
                      <Th>
                        <div>รหัสสินค้า</div>
                        <div>SKU Code</div>
                      </Th>
                      <Th>
                        <div>กลุ่ม</div>
                        <div>Product Code</div>
                      </Th>
                      <Th>
                        <div>สี</div>
                        <div>Color</div>
                      </Th>
                      <Th>
                        <div>ขนาด</div>
                        <div>Size</div>
                      </Th>
                      <Th right>
                        <div>ราคาป้ายต่อหน่วย</div>
                        <div>Tag Price per Unit</div>
                      </Th>
                      <Th center>
                        <div>จำนวน</div>
                        <div>Quantity</div>
                      </Th>
                      <Th right>
                        <div>ราคาป้ายทั้งหมด</div>
                        <div>Total Tag Price</div>
                      </Th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-[#eeeeee] bg-white">
                    {items.map((it, idx) => {
                      const rowTotal: number = Number(it.qty) * Number(it.unitPrice);
                    //  const rowErr: string | undefined = itemErrors[idx];

                      return (
                        <tr key={`${it.sku}-${idx}`} className="">
                          <Td>{it.no}.</Td>
                          <Td>
                            <span className="block truncate" title={it.sku}>{it.sku}</span>
                          </Td>
                          <Td>
                            <span className="block truncate" title={it.productCode}>{it.productCode}</span>
                          </Td>
                          <Td>
                            <span className="block truncate" title={it.color}>{it.color}</span>
                          </Td>
                          <Td>{it.size}</Td>
                          <Td right>{Number(it.unitPrice).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</Td>
                          <Td center>{Number(it.qty)}</Td>
                          <Td right>{rowTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</Td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}


          </section>
        </div>

        {items.length > 0 && (inventoryMode === "manual" || inventoryMode === "excel") && !isProcessing && !isSubmitting && fp && (
          <div className="fixed left-0 right-0 bottom-0 z-40 border-t border-[#d1d5db]" style={{ height: 36, backgroundColor: "#BCC8D6" }}>
            {/* Label */}
            <div
              style={{
                position: "absolute", left: 0, top: 0, bottom: 0,
                width: fp.qL,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontWeight: 600, fontSize: 12, color: "#374151",
                backgroundColor: "#BCC8D6",
              }}
            >
              ยอดรวมทั้งหมด / Grand Total
            </div>
            {/* Total Tag Price + Action background */}
            <div
              style={{
                position: "absolute", top: 0, bottom: 0,
                left: fp.tL, right: 0, zIndex: 1,
                backgroundColor: "#E0EEFF",
              }}
            />
            {/* Quantity */}
            <div
              style={{
                position: "absolute", top: 0, bottom: 0,
                left: fp.qL, width: fp.qW, zIndex: 2,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontWeight: 600, fontSize: 13, color: "#111827", backgroundColor: "#CAD6E6",
              }}
            >
              {totalQty.toLocaleString()}
            </div>
            {/* Total Tag Price value */}
            <div
              style={{
                position: "absolute", top: 0, bottom: 0,
                left: fp.tL, width: fp.tW, zIndex: 3,
                display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 12,
                fontWeight: 600, fontSize: 13, color: "#111827",
              }}
            >
              {totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </div>
          </div>
        )}

        {successOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[520px] rounded-xl bg-white shadow-xl border border-[#e7e7e7] overflow-hidden">
              <div className="px-5 py-4 bg-green-600 text-white">
                <div className="text-[14px] font-semibold">Success</div>
                <div className="text-[12px] opacity-95 mt-1">
                  Transaction ID: <span className="font-bold">{successTxnId}</span>
                </div>
              </div>

              <div className="px-5 py-4 text-[12px] text-gray-700">{successText}</div>

              <div className="px-5 pb-4 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={closeSuccessDialog}
                  className="h-9 px-5 rounded-md bg-[#0a66d8] text-white text-[12px] font-semibold hover:bg-[#095ec7]"
                >
                  OK
                </button>
              </div>
            </div>
          </div>
        )}



        {errorOpen && !customErrorOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="relative w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7] p-6 text-center">

              {/* Close Button */}
              <button
                type="button"
                onClick={closeErrorDialog}
                className="absolute top-3 right-3 text-gray-400 hover:text-gray-600 text-[18px]"
              >
                ✕
              </button>

              {/* Title */}
              <div className="text-[16px] font-semibold text-gray-900 mb-3">
                {errorTitle}
              </div>

              {/* Message */}
              <div className="text-[13px] text-gray-600 leading-relaxed whitespace-pre-line mb-6">
                {errorMessage}
              </div>

              {/* Button */}
              <button
                type="button"
                onClick={closeErrorDialog}
                className="h-10 px-8 rounded-md bg-[#000000] text-white text-[12px] font-semibold hover:bg-[#222222]"

              >
                OK, Got it
              </button>

            </div>
          </div>
        )}

        {removeFileConfirmOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/35 px-4">
            <div className="relative w-[365px] max-w-[92vw] min-h-[240px] rounded-[26px] border border-[#ececec] bg-white shadow-[0_14px_34px_rgba(0,0,0,0.2)]">
              <button
                type="button"
                onClick={() => setRemoveFileConfirmOpen(false)}
                aria-label="Close remove file confirmation"
                className="absolute right-4 top-4 flex h-8 w-8 items-center justify-center text-[#a0a0a0] hover:text-[#7a7a7a]"
              >
                <svg
                  viewBox="0 0 24 24"
                  className="h-7 w-7"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                >
                  <path d="M6 6l12 12M18 6L6 18" />
                </svg>
              </button>

              <div className="px-5 pt-12 pb-8 text-center">
                <div
                  className="mx-auto max-w-[400px] text-[15px] font-semibold leading-[1.45] text-[#111111]"
                  style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                >
                  Do you want to remove the selected file?
                </div>

                <div className="mt-3 text-[13px] text-[#5f5f5f]">
                  Please confirm if you want to proceed.
                </div>

                <div className="mt-9 flex justify-center gap-3">
                  <button
                    type="button"
                    onClick={() => setRemoveFileConfirmOpen(false)}
                    className="h-11 min-w-[110px] rounded-[8px] bg-[#E8E8E8] px-6 text-[14px] font-medium text-[#2f2f2f] hover:bg-[#EBEBEB]"

                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setRemoveFileConfirmOpen(false);
                      setSelectedExcelFile(undefined);
                      setExcelFileName("");
                      setExcelValidated(false);
                      setExcelErrors([]);
                      setItemErrors({});
                      setItems([]);

                      setExcelUploadedAt("");

                      if (fileInputRef.current) fileInputRef.current.value = "";
                    }}
                    className="h-11 min-w-[140px] rounded-[8px] bg-[#DA2128] px-6 text-[15px] font-semibold text-white hover:bg-[#AF0007]"

                  >
                    Remove
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}





        {removeConfirmOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/35 px-4">
            <div className="relative w-[430px] max-w-[92vw] min-h-[256px] rounded-[26px] border border-[#ececec] bg-white shadow-[0_14px_34px_rgba(0,0,0,0.2)]">

              {/* ❌ Close icon */}
              <button
                type="button"
                onClick={cancelRemoveItem}
                aria-label="Close remove confirmation"
                className="absolute right-4 top-4 flex h-8 w-8 items-center justify-center text-[#a0a0a0] hover:text-[#7a7a7a]"
              >
                <svg
                  viewBox="0 0 24 24"
                  className="h-7 w-7"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                >
                  <path d="M6 6l12 12M18 6L6 18" />
                </svg>
              </button>

              {/* Content */}
              <div className="px-8 pt-12 pb-8 text-center">
                <div className="mx-auto max-w-[320px] text-[15px] font-semibold leading-[1.45] text-[#111111]">
                  Are you sure do you want to remove
                  <br />
                  SKU &quot;{removeSku}&quot; ?
                </div>

                <div className="mt-3 text-[13px] text-[#5f5f5f]">
                  Please confirm if you want to proceed.
                </div>

                <div className="mt-9 flex justify-center gap-3">
                  <button
                    type="button"
                    onClick={cancelRemoveItem}
                    className="h-11 min-w-[110px] rounded-[8px] bg-[#E8E8E8] px-6 text-[14px] font-medium text-[#2f2f2f] hover:bg-[#EBEBEB]"
                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    onClick={confirmRemoveItem}
                    className="h-11 min-w-[140px] rounded-[8px] bg-[#DA2128] px-6 text-[15px] font-semibold text-white hover:bg-[#AF0007]"
                  >
                    Remove
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {uploadConfirmOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/35 px-4">
            <div className="relative w-[400px] max-w-[92vw] rounded-[26px] border border-[#ececec] bg-white shadow-[0_14px_34px_rgba(0,0,0,0.2)]">

              {/* Close icon */}
              <button
                type="button"
                onClick={() => setUploadConfirmOpen(false)}
                aria-label="Close upload confirmation"
                className="absolute right-4 top-4 flex h-8 w-8 items-center justify-center text-[#a0a0a0] hover:text-[#7a7a7a]"
              >
                ✕
              </button>

              <div className="px-8 pt-12 pb-8 text-center">
                <div className="text-[16px] font-semibold leading-[1.45] text-[#111111]">
                  Do you want to upload the file?
                </div>

                <div className="mt-3 text-[13px] leading-[1.55] text-[#5f5f5f]">
                  Your request will display the status “Pending” during processing.
                  Please confirm if you want to proceed.
                </div>

                <div className="mt-8 flex justify-center gap-3">
                  <button
                    type="button"
                    onClick={() => setUploadConfirmOpen(false)}
                    className="h-11 min-w-[110px] rounded-[8px] bg-[#E8E8E8] px-6 text-[14px] font-medium text-[#2f2f2f] hover:bg-[#EBEBEB]"

                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    onClick={async () => {
                      try {
                        setIsProcessing(true);
                        setUploadConfirmOpen(false);

                        if (!selectedExcelFile) {
                          throw new Error("Please select an Excel file.");
                        }



                        await createExcelPendingRequest(
                          context,
                          selectedExcelFile,
                          userData,
                          usage,
                          0,
                          0,
                          excelFileName
                        );

                        onNavigate("my", {
                          toastType: "pending",
                          toastMessage: "Your request status will be shown as “Pending” while it is being processed.",
                        });
                      } catch (err) {
                        console.error(err);
                        const message = err instanceof Error ? err.message : "Upload failed.";
                        setCustomErrorMessage(message);
                        setCustomErrorOpen(true);
                      } finally {
                        setIsProcessing(false);
                      }
                    }}
                    className="h-11 min-w-[170px] rounded-[8px] bg-[#000000] px-7 text-[14px] font-semibold text-white hover:bg-[#3B3B3B]"

                  >
                    Upload file
                  </button>

                </div>
              </div>
            </div>
          </div>
        )}
        {excelHelpOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/30 px-4">
            <div className="relative w-[520px] max-w-[92vw] min-h-[360px] rounded-[30px] border border-[#e6e6e6] bg-white shadow-[0_16px_36px_rgba(0,0,0,0.22)]">

              {/* Close icon */}
              <button
                type="button"
                onClick={closeExcelHelp}
                aria-label="Close excel help"
                className="group absolute right-3 top-2 flex h-12 w-12 items-center justify-center"
              >
                <span aria-hidden="true" className="relative block h-7 w-7">
                  <span className="absolute left-1/2 top-0 h-full w-[1.5px] -translate-x-1/2 rotate-45 bg-[#b3b3b3] transition-colors group-hover:bg-[#8f8f8f]" />
                  <span className="absolute left-1/2 top-0 h-full w-[1.5px] -translate-x-1/2 -rotate-45 bg-[#b3b3b3] transition-colors group-hover:bg-[#8f8f8f]" />
                </span>
              </button>

              <div className="px-8 pt-10 pb-9">
                <div className="mb-4 text-center text-[16px] font-bold text-[#111111]">
                  How to upload the Excel file successfully
                </div>

                <ol className="list-decimal pl-6 space-y-0.5 text-[12px] leading-[1.45] text-[#5a5a5a]">
                  <li>Each SKU code must exactly match a record in the Product Master;
                    otherwise, the upload will not be completed.</li>
                  <li>Only the .xlsx file format is supported.</li>
                  <li>Please keep your file within 300 rows.</li>
                  <li>Duplicate SKU codes are not allowed.</li>
                  <li>
                    After uploading, the system will take a few minutes to verify your file.
                    During this time, the status will display as &quot;Pending&quot;.
                  </li>
                  <li>
                    Once verification is complete, the status will change to &quot;Draft&quot;, and you
                    will receive an email notification informing you whether the verification
                    succeeded or failed.
                  </li>
                  <li>Once uploaded, requested items cannot be edited.</li>
                </ol>
              </div>
            </div>
          </div>
        )}
        {submitConfirmOpen && (
          <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/40 px-4">
            <div className="relative w-[420px] max-w-[92vw] rounded-2xl bg-white shadow-xl border border-[#e7e7e7]">

              {/* Close icon */}
              <button
                type="button"
                onClick={() => setSubmitConfirmOpen(false)}
                className="absolute top-3 right-4 text-gray-400 hover:text-gray-600 text-[20px] leading-none"
              >
                ✕
              </button>

              <div className="px-8 py-8 text-center">

                {/* Title — matches screenshot 1 */}
                <div className="text-[16px] font-semibold text-gray-900 mb-4">
                  Do you want to submit the request?
                </div>

                {/* Bold warning — matches screenshot 1 */}
                <div className="text-[13px] font-bold text-gray-800 mb-2">
                  The request cannot be edited once it is submitted.
                </div>

                {/* Note text — matches screenshot 1 */}
                <div className="text-[12px] text-gray-500 leading-relaxed mb-7">
                  Note: The system will verify the requested quantity
                  against available stocks. Any unavailable items will be
                  flagged. Please confirm if you want to proceed.
                </div>

                {/* Buttons — matches screenshot 1 */}
                <div className="flex justify-center gap-4">
                  <button
                    type="button"
                    onClick={() => setSubmitConfirmOpen(false)}
                    className="h-10 min-w-[100px] px-6 rounded-lg bg-[#E8E8E8] text-gray-700 text-[13px] font-medium hover:bg-[#EBEBEB]"

                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setSubmitConfirmOpen(false);
                      onSubmit();
                    }}
                    className="h-10 min-w-[100px] px-6 rounded-lg bg-green-600 text-white text-[13px] font-semibold hover:bg-green-700"
                  >
                    Submit
                  </button>
                </div>

              </div>
            </div>
          </div>
        )}

        {maxRowsPopupOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7] overflow-hidden">
              <div className="px-5 py-4 bg-red-600 text-white">
                <div className="text-[14px] font-semibold">Limit reached</div>
              </div>

              <div className="px-5 py-4 text-[12px] text-gray-700">
                {maxRowsPopupMsg}
              </div>

              <div className="px-5 pb-4 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={closeMaxRowsPopup}
                  className="h-9 px-5 rounded-md bg-[#0a66d8] text-white text-[12px] font-semibold hover:bg-[#095ec7]"
                >
                  OK
                </button>
              </div>
            </div>
          </div>
        )}



        {discardConfirmOpen && (
          <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/35 px-4">
            <div className="relative w-[400px] max-w-[92vw] min-h-[250px] rounded-[26px] border border-[#ececec] bg-white shadow-[0_14px_34px_rgba(0,0,0,0.2)]">
              {/* Close X */}
              <button
                type="button"
                className="absolute right-4 top-4 flex h-8 w-8 items-center justify-center text-[#a0a0a0] hover:text-[#7a7a7a]"
                onClick={() => setDiscardConfirmOpen(false)}
              >
                <svg className="h-7 w-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>

              <div className="px-8 pt-12 pb-8 text-center">
                {/* Title */}
                <h3 className="text-[16px] font-semibold leading-[1.45] text-[#111111]">
                  Are you sure you want to discard?
                </h3>

                {/* Body */}
                <p className="mt-3 text-[13px] leading-[1.55] text-[#5f5f5f]">
                  The request will not be saved.
                  <br />
                  Please confirm if you want to proceed.
                </p>

                {/* Buttons */}
                <div className="mt-8 flex justify-center gap-3">
                  <button
                    type="button"
                    className="h-11 min-w-[110px] rounded-[8px] bg-[#E8E8E8] px-6 text-[14px] font-medium text-[#2f2f2f] hover:bg-[#EBEBEB]"
                    onClick={() => setDiscardConfirmOpen(false)}
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    className="h-11 min-w-[126px] rounded-[8px] bg-[#e8242a] px-6 text-[15px] font-semibold text-white hover:bg-[#d61b22]"
                    onClick={() => {
                      setDiscardConfirmOpen(false);
                      onDiscard();
                    }}
                  >
                    Discard
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
        {isProcessing && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7]">
              <div className="px-6 py-6 text-center">
                <div className="flex justify-center mb-4">
                  <div className="h-10 w-10 border-4 border-gray-200 border-t-blue-600 rounded-full animate-spin" />
                </div>

                <div className="text-[15px] font-semibold text-gray-900 mb-2">
                  Please wait...
                </div>

                <div className="text-[12px] text-gray-600">
                  Checking available stocks.
                  This may taken more than 5 minutes for 300 SKUs
                  upload. Please do not close the page.
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Insufficient Stock Detected Popup */}
        {excelFailPopupOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[430px] rounded-2xl bg-white shadow-xl overflow-hidden relative">
              <button
                type="button"
                onClick={() => setExcelFailPopupOpen(false)}
                className="absolute top-4 right-4 text-gray-400 hover:text-gray-600 text-base leading-none"
              >
                ✕
              </button>
              <div className="px-10 pt-10 pb-8 text-center">
                <div className="text-[18px] font-bold text-gray-900 mb-3">
                  Insufficient Stock Detected
                </div>
                <div className="text-[13px] text-gray-600 leading-6 text-center">
                  <p className="font-bold text-gray-900 mb-2">Some items currently have insufficient stock.</p>
                  The system will recheck quantities again at approval.{" "}
                  If any items remain unavailable, approver(s), requestor, and admin will be notified by email.{" "}
                  You may create a new request for those items later.
                </div>
                <div className="mt-7 flex justify-center">
                  <button
                    type="button"
                    onClick={onExcelFailOk}
                    className="h-11 w-[190px] rounded-lg bg-[#000000] text-white text-[13px] font-semibold hover:bg-[#222222]"
                  >
                    OK, Got it
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {manualStockPopupOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[520px] rounded-xl bg-white shadow-xl border border-[#e7e7e7] overflow-hidden relative">
              <button
                type="button"
                onClick={closeManualStockPopup}
                className="absolute top-3 right-3 text-gray-500 hover:text-gray-600"
              >
                ✕
              </button>
              <div className="px-6 py-6 text-center">
                <div className="text-[19px] font-bold text-gray-900 mb-2">
                  Unavailable items flagged.
                </div>
                <div className="text-[15px] text-gray-600 leading-relaxed">
                  Some items are not in stock.
                  <br />
                  Please revise and resubmit the request.
                </div>
                <div className="mt-6 flex justify-center">
                  <button
                    type="button"
                    onClick={closeManualStockPopup}
                    className="h-10 px-8 rounded-md bg-[#000000] text-white text-[12px] font-semibold hover:bg-[#222222]"
                  >
                    Edit quantity
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Custom Error Message Popup */}
        {customErrorOpen && (
          <div className="fixed inset-0 z-[999] flex items-center justify-center bg-black/40 px-4">
            <div className="w-full max-w-[420px] rounded-xl bg-white shadow-xl border border-[#e7e7e7] relative">
              <button
                type="button"
                onClick={() => setCustomErrorOpen(false)}  // Close the popup
                className="absolute top-3 right-3 text-gray-500 hover:text-gray-600"
              >
                ✕
              </button>
              <div className="px-6 py-6 text-center">
                <div className="text-[16px] font-bold text-gray-900 mb-2">
                  Something went wrong
                </div>
                <div className="text-[12px] text-gray-600 leading-relaxed">
                  {customErrorMessage || "Unable to complete the action. Please try submit again."}
                </div>
                <div className="mt-6 flex justify-center">
                  <button
                    type="button"
                    onClick={() => setCustomErrorOpen(false)}  // Close the popup
                    className="h-10 px-8 rounded-md bg-[#000000] text-white text-[12px] font-semibold hover:bg-[#222222]"
                  >
                    Try again
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
