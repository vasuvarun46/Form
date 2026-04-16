import * as React from "react";

interface RequestEmptyStateProps {
  title?: string;
  description?: string;
  actionLabel?: string;
  onAction?: () => void;
  showAction?: boolean;
}

export default function RequestEmptyState({
  title = "Empty team’s requests",
  description = "There is no any submitted requests from the team yet",
  actionLabel,
  onAction,
  showAction = true,
}: RequestEmptyStateProps): JSX.Element {
  return (
    <div className="pdf-thai flex justify-center px-3 sm:px-6 py-10">
      {/* CARD */}
      <div
        className="
          w-full max-w-[900px]
          bg-white
          rounded-xl
          border border-[#e5e5e5]
          shadow-sm
          px-6 py-10
          text-center
        "
        style={{
          fontFamily:
            "Noto Sans Thai, Segoe UI, -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif",
        }}
      >
        {/* ICON */}
        <div className="flex justify-center mb-4">
          <svg width="40" height="40" viewBox="0 0 24 24" fill="none">
            <path
              d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"
              stroke="#E60023"
              strokeWidth="1.5"
            />
          </svg>
        </div>

        {/* TITLE */}
        <div className="text-[14px] font-semibold text-gray-900">
          {title}
        </div>

        {/* DESCRIPTION */}
        <div className="mt-1 text-[12px] text-gray-500">
          {description}
        </div>

        {/* OPTIONAL ACTION */}
    {/* OPTIONAL ACTION */}
{showAction && actionLabel && onAction && (
  <div className="mt-6">
    <button
      type="button"
      onClick={onAction}
      className="
        h-9 px-5
        rounded-md
        bg-[#e6ecf1]
        text-black
        text-[12px]
        font-semibold
        hover:text-[#095ec7]
      "
    >
      {actionLabel}
    </button>
  </div>
)}


      </div>
    </div>
  );
}
