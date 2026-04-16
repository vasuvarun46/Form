import * as React from "react";
import { SecurityPermissions } from "../spHelpers";

interface Props {
  onCreateRequest: () => void;
  onGoToForm: () => void;
  permissions: SecurityPermissions;
}

const forms = [
  { titleTH: "แบบเบิกสินค้าคงเหลือเพื่อใช้ภายใน", titleEN: "Internal Usage Inventory Request Form", active: true },
  { titleTH: "แบบเบิกค่าใช้จ่ายในการเดินทาง", titleEN: "Travelling Reimbursement Form", active: false },
  { titleTH: "แบบเบิกค่าใช้จ่าย", titleEN: "Expense Reimbursement Form", active: false },
  { titleTH: "แบบแจ้งชำระเงินผู้ขาย", titleEN: "Supplier Request Form", active: false },
  { titleTH: "แบบเบิกเงินทดรองจ่าย", titleEN: "Cash Advance Request Form", active: false },
  { titleTH: "แบบคืนเงินทดรองจ่าย", titleEN: "Cash Advance Clearing Form", active: false },
  { titleTH: "แบบเบิกเงินสดย่อย", titleEN: "Petty Cash Form", active: false },
];

export default function HomePage({ onCreateRequest, onGoToForm, permissions }: Props): JSX.Element {
  const [isGoTouched, setIsGoTouched] = React.useState<boolean>(false);
  const [isCreateTouched, setIsCreateTouched] = React.useState<boolean>(false);

  return (
    <div className="w-full bg-[#f3f3f3]">
      <div className="w-full px-4 py-6">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-center text-lg font-bold text-gray-900 mb-6">
            E-Form
          </h1>

          <div className="w-full space-y-4">
            {forms.map((f, idx) => (
              <div
                key={idx}
                className="
                  bg-white rounded-xl shadow-sm border border-gray-200
                  px-4 sm:px-5 py-4
                  flex items-center justify-between flex-wrap gap-4
                "
              >
                <div className="flex items-center gap-4 min-w-0">
                  <div className="h-12 w-12 rounded-lg overflow-hidden bg-gray-100 flex-shrink-0">
                    <img
                      src={require("../../assets/FileIcon.png")}
                      alt="Form Icon"
                      className="h-full w-full object-cover"
                    />
                  </div>

                  <div className="min-w-0">
                    <div className="text-sm font-bold text-gray-900 truncate">
                      {f.titleTH}
                    </div>
                    <div className="text-xs truncate">
                      <span className="text-[#E60023]">/</span>{" "}
                      <span className="text-black">{f.titleEN}</span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center ml-auto">
                  {f.active ? (
                    <div className="flex gap-2">
                      <button
                        type="button"
                        className={[
                          "h-9 px-4 rounded-md border text-sm font-semibold transition-colors duration-150",
                          isGoTouched
                            ? "border-black bg-black text-white"
                            : "border-gray-300 bg-white text-black",
                        ].join(" ")}
                        onClick={onGoToForm}
                        onMouseEnter={() => setIsGoTouched(true)}
                        onMouseLeave={() => setIsGoTouched(false)}
                        onFocus={() => setIsGoTouched(true)}
                        onBlur={() => setIsGoTouched(false)}
                        onTouchStart={() => setIsGoTouched(true)}
                        onTouchEnd={() => setIsGoTouched(false)}
                      >
                        Go to Form
                      </button>

                      {permissions.submit && (
                        <button
                          type="button"
                          className={[
                            "h-9 px-4 rounded-md text-sm font-semibold transition-colors duration-150",
                            isCreateTouched
                              ? "bg-[#2f2f2f] text-[#f3f4f6]"
                              : "bg-black text-white",
                          ].join(" ")}
                          onClick={onCreateRequest}
                          onMouseEnter={() => setIsCreateTouched(true)}
                          onMouseLeave={() => setIsCreateTouched(false)}
                          onFocus={() => setIsCreateTouched(true)}
                          onBlur={() => setIsCreateTouched(false)}
                          onTouchStart={() => setIsCreateTouched(true)}
                          onTouchEnd={() => setIsCreateTouched(false)}
                        >
                          Create request
                        </button>
                      )}
                    </div>
                  ) : (
                    <div className="h-10 px-6 rounded-md bg-gray-100 text-gray-500 text-sm font-semibold flex items-center">
                      Coming Soon
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="mt-16">
        <div className="relative h-[240px] w-full">
          <div className="absolute inset-x-0 top-0 h-[230px] bg-black rounded-t-[120px]" />
          <div className="absolute inset-x-0 top-[10px] h-[230px] bg-[#8B1418] rounded-t-[110px] flex items-center justify-center text-center text-white px-4">
            <div>
              <h2 className="text-xl font-bold mb-2">
                Bring everyone together with GQ Standards
              </h2>
              <p className="text-sm text-white/80 mb-6">
                Explore more guides and tools
              </p>
              <button
                type="button"
                className="bg-white text-black px-6 py-2 rounded-md text-sm font-semibold hover:bg-black hover:text-white transition-colors"
              >
                Let&apos;s connect
              </button>
            </div>
          </div>
        </div>

        <div className="bg-black py-3 text-center">
          <p className="text-xs text-white">
            GQ E-Form (c) Copyright 2025. All rights reserved Terms & Privacy
          </p>
        </div>
      </div>
    </div>
  );
}
