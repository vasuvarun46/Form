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
      title={
        <div className="flex items-center gap-2 flex-wrap">
          <span className={isDesktop ? "text-[18px] font-semibold text-gray-900" : "text-[16px] font-semibold text-gray-900"}>
            Create request
          </span>
          {showExcelNotice && (
            <div
              className="
                inline-flex items-center
                rounded-md
                border border-[#FF8040]
                bg-[#FFF2EC]
                px-3 py-1
                text-[11px] font-semibold
                text-[#FF9058]
              "
            >
              {excelNoticeText}
            </div>
          )}

        </div>
      }
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
        </div>
      }
      right={
        <div className={isDesktop ? "ml-auto flex items-center justify-end gap-2" : "mt-2 grid w-full grid-cols-3 gap-2"}>
          <button
            type="button"
            onClick={onDiscard}
            className={isDesktop
              ? "h-10 w-[106px] rounded-[8px] border border-[#EBEBEB] bg-[#EBEBEB] text-sm font-semibold text-[#2f2f2f] hover:bg-[#E8E8E8]"
              : "h-11 min-w-0 w-full px-2 rounded-lg border border-[#EBEBEB] bg-[#EBEBEB] text-[13px] leading-none font-semibold text-[#2f2f2f] hover:bg-[#E8E8E8]"}
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
                  ? "bg-[#D9D9D9] text-[#7f7f7f] cursor-not-allowed"
                  : "bg-[#000000] text-white hover:bg-[#3B3B3B]"
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
                ? "h-10 w-[170px] rounded-[8px] border-2 text-sm font-semibold"
                : "h-11 min-w-0 w-full px-2 rounded-lg border-2 text-[13px] leading-none font-semibold",
              submitDisabled
                ? "border-[#D9D9D9] bg-[#D9D9D9] text-[#7f7f7f] cursor-not-allowed"
                : "border-[#009A11] bg-[#E6F5E7] text-[#009A11] hover:bg-[#009A11] hover:text-white"
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
