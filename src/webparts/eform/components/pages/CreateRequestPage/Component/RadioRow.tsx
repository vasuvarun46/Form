import * as React from "react";

type RadioRowProps = {
  title: string;
  sub: React.ReactNode;
  checked: boolean;
  onChange: () => void;
  disabled: boolean;

};

export default function RadioRow({
  title,
  sub,
  checked,
  onChange,
  disabled
}: RadioRowProps): JSX.Element {
  const handleClick = () => {
    if (disabled) return;   // ✅ block click
    onChange();
  };

  return (
    <button
      type="button"
      onClick={handleClick}
      disabled={disabled}

      className={[
        "w-full text-left flex items-start gap-3 rounded-md px-2 py-3 transition",
        disabled
          ? "opacity-50 cursor-not-allowed"
          : "hover:bg-[#f7f7f7]"
      ].join(" ")}

    >
      <span className="mt-1.5">
        <span
          className={[
            "h-4 w-4 rounded-full border flex items-center justify-center",
            checked ? "border-gray-900" : "border-gray-300",
          ].join(" ")}
        >
          {checked && <span className="h-2 w-2 rounded-full bg-gray-900" />}
        </span>
      </span>

      <div className="min-w-0">
        <div className="text-[12px] font-semibold text-gray-900">{title}</div>
        <div className="text-[11px] text-gray-500 mt-0.5">{sub}</div>
      </div>
    </button>
  );
}
