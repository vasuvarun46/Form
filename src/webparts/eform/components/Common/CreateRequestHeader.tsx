import * as React from 'react';
import { Route } from '../Eform';
import BaseInnerHeader from './PageHeader/BaseInnerHeader';
import { useIsDesktop } from '../hooks/useIsDesktop';

export interface ICreateRequestHeaderProps {
  onDiscard: () => void;
  onSaveDraft: () => void;
  onSubmit: () => void;
  onNavigate: (route: Route) => void;
  draftDisabled: boolean;

  showExcelNotice?: boolean;
  excelNoticeText?: string;
  submitDisabled: boolean;
  hideSaveDraft?: boolean;
}

export default function CreateRequestHeader({
  onDiscard,
  onSaveDraft,
  onSubmit,
  onNavigate,
  draftDisabled,
  submitDisabled,
  showExcelNotice = false,
  excelNoticeText = "This request was created by Excel upload and cannot be edited.",
  hideSaveDraft = false,
}: ICreateRequestHeaderProps): JSX.Element {
  const isDesktop = useIsDesktop(900);
  return (
    <BaseInnerHeader
      title="Create request"
      subtitle={
        <div className={isDesktop ? "flex items-center gap-2 flex-nowrap" : "flex items-center gap-1 flex-wrap"}>
          <span
            className="text-blue-600 cursor-pointer hover:underline text-xs"
            onClick={() => onNavigate("my")}
          >
            My requests
          </span>
          <span className="text-xs text-gray-500">
            / Create request details
          </span>
          {showExcelNotice && (
            <div
              className="
                inline-flex items-center
                rounded-md
                border border-[#f59e0b]
                bg-[#fff7ed]
                px-3 py-1
                text-[11px] font-semibold
                text-[#c2410c]
              "
            >
              {excelNoticeText}
            </div>
          )}
        </div>
      }
      right={
        <div className={isDesktop ? "ml-auto flex items-center justify-end gap-2" : "mt-2 grid w-full grid-cols-3 gap-2"}>
          <button
            type="button"
            onClick={onDiscard}
            className={isDesktop
              ? "h-10 w-[106px] rounded-[8px] border border-[#e5e7eb] bg-[#e5e7eb] text-sm font-semibold text-[#2f2f2f] hover:bg-[#dfe3e6]"
              : "h-11 min-w-0 w-full px-2 rounded-lg border border-[#e5e7eb] bg-[#e5e7eb] text-[13px] leading-none font-semibold text-[#2f2f2f] hover:bg-[#dfe3e6]"}
          >
            Discard
          </button>

          {!hideSaveDraft && (
            <button
              type="button"
              onClick={onSaveDraft}
              disabled={draftDisabled}
              className={`
              ${isDesktop ? "h-10 w-[170px] rounded-[8px] text-sm" : "h-11 min-w-0 w-full px-2 rounded-lg text-[13px] leading-none"}
              font-semibold transition-colors
              ${draftDisabled
                  ? "bg-[#cfcfcf] text-[#7f7f7f] cursor-not-allowed"
                  : "bg-black text-white hover:opacity-90"
                }
            `}
            >
              Save draft
            </button>
          )}

          <button
            type="button"
            disabled={submitDisabled}
            onClick={submitDisabled ? undefined : onSubmit}
            className={[
              isDesktop
                ? "h-10 w-[170px] rounded-[8px] border text-sm font-semibold"
                : "h-11 min-w-0 w-full px-2 rounded-lg border text-[13px] leading-none font-semibold",
              submitDisabled
                ? "border-[#cfcfcf] bg-[#cfcfcf] text-[#7f7f7f] cursor-not-allowed"
                : "border-[#16a34a] bg-[#eaf7ea] text-[#0a8f2f] hover:bg-[#ddf1de]"
            ].join(" ")}
          >
            Submit
          </button>
        </div>
      }
    >
    </BaseInnerHeader>

  );
}
