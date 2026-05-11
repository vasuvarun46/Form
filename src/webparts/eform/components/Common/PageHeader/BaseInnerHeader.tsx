import * as React from "react";
import { useIsDesktop } from "../../hooks/useIsDesktop";

type BaseInnerHeaderProps = {
  title: React.ReactNode;
  subtitle?: React.ReactNode;
  right?: React.ReactNode;
  children?: React.ReactNode;
  fullWidth?: boolean;
}

export default function BaseInnerHeader(props: BaseInnerHeaderProps): JSX.Element {
  const { title, subtitle, right,  } = props;
  const isDesktop = useIsDesktop(900);

  return (
    <div className="bg-[#f3f3f3] border-b border-[#e5e5e5] pdf-thai">
      <div
  className={
    isDesktop
      ? "w-full px-6 py-4 flex items-start justify-between gap-6"
      : "w-full px-3 py-3 flex flex-col items-start gap-3"
  }
>
        {/* LEFT */}
        <div className="w-full sm:w-auto">
          <div className={isDesktop ? "text-[18px] font-semibold text-gray-900" : "text-[16px] font-semibold text-gray-900"}>
            {title}
          </div>
          {subtitle && (
            <div className={isDesktop ? "text-[12px] text-gray-500" : "text-[11px] text-gray-500"}>
              {subtitle}
            </div>
          )}
        </div>

        {/* RIGHT */}
        <div
          className={
            isDesktop
              ? "w-auto ml-auto shrink-0 flex items-center justify-end gap-2"
              : "w-full flex items-center justify-start gap-2"
          }
        >
          {right}
        </div>
      </div>
    </div>
  );
}
