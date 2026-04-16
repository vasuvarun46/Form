import * as React from "react";

interface LoadingProps {
  title?: string;
  description?: string;
  className?: string;
}

export default function Loading({
  title = "Loading...",
  description = "Please wait a moment.",
  className = "",
}: LoadingProps): JSX.Element {
  return (
    <div
      className={[
        "rounded-lg border border-[#e7e7e7] bg-white py-10 text-center",
        className,
      ].join(" ")}
    >
      <div className="flex justify-center">
        <div className="h-8 w-8 rounded-full border-2 border-gray-200 border-t-[#0a66d8] animate-spin" />
      </div>

      <div className="mt-3 text-[13px] font-semibold text-gray-900">{title}</div>
      <div className="mt-1 text-[11px] text-gray-500">{description}</div>
    </div>
  );
}
