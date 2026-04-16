import * as React from "react";
import { IconChevronDown, IconSearch } from "../../../ui/icons";
import QtyStepper from "../Component/QtyStepper";

export interface ManualProduct {
  Barcode: string;
  Warehouse_SOH_Qty: number;
  ProductCode?: string;
  ColourDesc?: string;
  SizeDesc?: string;
  Original_sp?: number;
}

export interface ItemRow {
  no: number;
  sku: string;
  productCode: string;
  color: string;
  size: string;
  unitPrice: number;
  qty: number;
  maxStock?: number;
}


type InventoryItemsSectionProps = {
  card: string;
  inventoryMode: "manual" | "excel";

  // Manual search
  searchSku: string;
  setSearchSku: (v: string) => void;
  isSearchingSku: boolean;
  handleManualSkuSearch: () => Promise<void>;

  // Manual popup
  manualPopupOpen: boolean;
  manualProduct?: ManualProduct;
  manualQty: number;
  manualQtyError: string;
  manualMaxStock: number;

  setManualQty: (v: number) => void;
  setManualQtyError: (v: string) => void; // ✅ ADD THIS

  setManualPopupOpen: (v: boolean) => void;
  setManualProduct: (v: ManualProduct | undefined) => void;
  handleAddManualItem: () => void;

  // Items table
  items: ItemRow[];
  itemErrors: Record<number, string>;
  updateQty: (index: number, next: string | number) => void;
  removeItem: (index: number) => void;
};


export default function InventoryItemsSection(
  props: InventoryItemsSectionProps
): JSX.Element {
  const {
    card,
    inventoryMode,
    searchSku,
    setSearchSku,
    isSearchingSku,
    handleManualSkuSearch,

    manualPopupOpen,
    manualProduct,
    manualQty,
    manualQtyError,
    manualMaxStock,

    setManualQty,
    setManualQtyError, // ✅ ADD THIS

    setManualPopupOpen,
    setManualProduct,
    handleAddManualItem,

    items,
    itemErrors,
    updateQty,
    removeItem,
  } = props;

  const handleManualQtyChange = (value: number): void => {
    setManualQty(value);

    if (value < 0) {
      setManualQtyError("Quantity cannot be negative.");
    } else if (value > manualMaxStock) {
      setManualQtyError(`Only ${manualMaxStock} available.`);
    } else if (value === 0) {
      // ⚠️ allowed for DRAFT only
      setManualQtyError("Quantity is 0. You can save draft but cannot submit.");
    } else {
      // ✅ THIS IS WHERE YOUR LINE GOES
      setManualQtyError("");
    }
  };

  return (
    <section className={`${card} p-4 sm:p-5 pdf-thai`}>
      <div className="text-[13px] font-semibold text-gray-900 mb-4">
        Product
      </div>

      {/* 🔍 SKU Search */}
      <div className="w-[560px] max-w-full">
        <div className="h-10 w-full flex items-stretch overflow-hidden rounded-md border border-[#e6e6e6] bg-white">
          <div className="relative w-[160px] shrink-0">
            <select
              disabled
              className="h-full w-full bg-transparent px-3 pr-9 text-[12px] text-gray-700 appearance-none"
            >
              <option>SKU Code</option>
            </select>

            <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">
              <IconChevronDown />
            </span>

            <span className="absolute right-0 top-0 bottom-0 w-px bg-[#eeeeee]" />
          </div>

          <div className="relative flex-1">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500">
              <IconSearch />
            </span>

            <input
              placeholder="Enter SKU Code"
              value={searchSku}
              disabled={inventoryMode !== "manual" || isSearchingSku}
              onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
                setSearchSku(e.target.value)
              }
              onKeyDown={async (e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  await handleManualSkuSearch();
                }
              }}
              className="h-full w-full bg-transparent pl-9 pr-3 text-[12px] text-gray-700 focus:outline-none focus-visible:outline focus-visible:outline-2 focus-visible:outline-[#2563eb] focus-visible:outline-offset-2"
            />

          </div>

          <button
            type="button"
            disabled={inventoryMode !== "manual" || isSearchingSku}
            onClick={handleManualSkuSearch}
            className={[
              "h-full px-6 bg-[#0a66d8] text-white text-[12px] font-semibold",
              inventoryMode !== "manual" || isSearchingSku
                ? "opacity-60 cursor-not-allowed"
                : "hover:bg-[#095ec7]",
            ].join(" ")}
          >
            {isSearchingSku ? "Searching..." : "Search"}
          </button>
        </div>
      </div>

      {/* ➕ Manual Add Popup */}
      {inventoryMode === "manual" && manualPopupOpen && manualProduct && (
        <div className="mt-3 w-[520px] max-w-full rounded-lg border border-[#e7e7e7] bg-white shadow-sm">

          {/* Header */}
          <div className="px-4 py-3 text-[12px] font-semibold border-b">
            Product
          </div>

          {/* Content */}
          <div className="px-4 py-4">
            <table className="w-full text-[12px]">
              <thead className="text-gray-500">
                <tr>
                  <th className="text-left font-semibold">SKU</th>
                  <th className="text-center font-semibold">In Stock</th>
                  <th className="text-center font-semibold">
                    <span className="text-red-500">*</span> Request Quantity
                  </th>
                </tr>
              </thead>

              <tbody>
                <tr className="border-t">
                  {/* SKU */}
                  <td className="py-3">
                    {String(manualProduct.Barcode || "-")}
                  </td>

                  {/* Stock */}
                  <td className="py-3 text-center">
                    {Number(manualProduct.Warehouse_SOH_Qty ?? 0)}
                  </td>

                  {/* Qty */}
                  <td className="py-3">
                    <div className="flex justify-center">
                      <div className="flex items-center w-[120px] border border-[#0a66d8] rounded-md overflow-hidden">
                        <button
                          type="button"
                          className="w-9 h-8 border-r hover:bg-[#f2f7ff]"
                          onClick={() => handleManualQtyChange(manualQty - 1)}
                        >
                          –
                        </button>

                        <input
                          value={manualQty}
                          onChange={(e) =>
                            handleManualQtyChange(Number(e.target.value || 0))
                          }
                          className="w-full h-8 text-center text-[12px] focus:outline-none focus-visible:outline focus-visible:outline-2 focus-visible:outline-[#2563eb] focus-visible:outline-offset-2"
                        />

                        <button
                          type="button"
                          className="w-9 h-8 border-l hover:bg-[#f2f7ff]"
                          onClick={() => handleManualQtyChange(manualQty + 1)}
                        >
                          +
                        </button>
                      </div>
                    </div>

                    {manualQtyError && (
                      <div className="mt-1 text-[10px] text-red-600 text-center">
                        {manualQtyError}
                      </div>
                    )}
                  </td>
                </tr>
              </tbody>
            </table>

            {/* Actions */}
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setManualPopupOpen(false);
                  setManualProduct(undefined);
                  setManualQty(1);
                  setManualQtyError("");
                }}
                className="h-8 px-4 rounded-md border text-[12px]"
              >
                Cancel
              </button>

              <button
                type="button"
                onClick={handleAddManualItem}
                className="h-8 px-5 rounded-md bg-[#0a66d8] text-white text-[12px] font-semibold hover:bg-[#095ec7]"
              >
                Add
              </button>
            </div>
          </div>
        </div>
      )}


      {/* 📦 Items Table */}
      {items.length === 0 ? (
        <div className="mt-4 text-[12px] text-gray-500">
          No items added yet.
        </div>
      ) : (
        <div className="mt-4 overflow-x-auto rounded-lg border bg-white">
          <table className="min-w-[900px] w-full text-[12px]">
            <thead className="bg-[#f7f7f7] text-gray-600">
              <tr>
                <th className="px-4 py-3">No</th>
                <th className="px-4 py-3">SKU</th>
                <th className="px-4 py-3">Product</th>
                <th className="px-4 py-3 text-center">Qty</th>
                {inventoryMode === "manual" && <th className="px-4 py-3 text-center">Action</th>}
              </tr>
            </thead>

            <tbody>
              {items.map((it, idx) => (
                <tr key={idx} className="border-t">
                  <td className="px-4 py-3">{it.no}</td>
                  <td className="px-4 py-3">{it.sku}</td>
                  <td className="px-4 py-3">{it.productCode}</td>
                  <td className="px-4 py-3 text-center">
                    <QtyStepper
                      value={it.qty}
                      onDec={() => updateQty(idx, it.qty - 1)}
                      onInc={() => updateQty(idx, it.qty + 1)}
                      onChange={(v: string | number) => updateQty(idx, v)}
                        disabled={inventoryMode === "excel"}
                    />
                  </td>
                  <td className="px-4 py-3 text-center">
                    {inventoryMode === "manual" && (<button
                      onClick={() => removeItem(idx)}
                      className="h-7 px-3 rounded-md border text-[11px]"
                    >
                      Remove
                    </button>)}


                    {itemErrors[idx] && (
                      <div className="mt-1 text-[10px] text-red-600">
                        {itemErrors[idx]}
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}
