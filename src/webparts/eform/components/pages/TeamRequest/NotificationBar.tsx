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
          <span className="inline-flex h-4 w-4 items-center justify-center shrink-0 rounded-full border border-white/70">
            {props.type === "success" ? (
              <svg viewBox="0 0 24 24" className="h-3 w-3" fill="none" aria-hidden="true">
                <path d="M5 12.5L10 17.5L19 8.5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            ) : (
              <svg viewBox="0 0 24 24" className="h-3 w-3" fill="none" aria-hidden="true">
                <path d="M7 7L17 17M17 7L7 17" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
            )}
          </span>
          <span className="text-sm font-medium truncate">{props.message}</span>
        </div>
        <button
          onClick={props.onClose}
          className="text-white text-base leading-none opacity-90 hover:opacity-100"
          aria-label="Close"
        >
          x
        </button>
      </div>
    </div>
  );
}
