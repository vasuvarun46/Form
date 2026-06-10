import * as React from "react";

export interface NotificationBarProps {
  type: "success" | "error";
  message: string;
  onClose: () => void;
  inline?: boolean;
}

export default function NotificationBar(
  props: NotificationBarProps
): JSX.Element {
  const bgColor = props.type === "success" ? "bg-green-600" : "bg-red-600";
  const layoutClass = props.inline
    ? "w-full text-white rounded-md shadow-lg"
    : "fixed top-[72px] left-1/2 -translate-x-1/2 z-[1200] w-[calc(100vw-24px)] max-w-[1320px] text-white rounded-md shadow-lg";

  return (
    <div
      className={[
        layoutClass,
        bgColor
      ].join(" ")}
    >
      <div className="px-4 py-2.5 flex justify-between items-center gap-3">
        <div className="flex items-center gap-3 min-w-0">
          {props.type === "success" ? (
            <svg viewBox="0 0 24 24" className="h-5 w-5 shrink-0" fill="none" aria-hidden="true">
              <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
              <path d="M7 12.5L10.5 16L17 9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          ) : (
            <svg viewBox="0 0 24 24" className="h-5 w-5 shrink-0" fill="none" aria-hidden="true">
              <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
              <path d="M9 9L15 15M15 9L9 15" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            </svg>
          )}
          <span className="text-sm font-medium truncate">{props.message}</span>
        </div>
        <button
          onClick={props.onClose}
          className="text-white opacity-90 hover:opacity-100 shrink-0"
          aria-label="Close"
        >
          <svg className="h-5 w-5" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path d="M6 6L18 18M18 6L6 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          </svg>
        </button>
      </div>
    </div>
  );
}
