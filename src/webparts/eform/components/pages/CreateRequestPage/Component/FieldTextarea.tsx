import * as React from "react";

export default function FieldTextarea({
  label,
  value,
  onChange,
  error,
  maxLength,
  placeholder,
  disabled,
  grayBackground = false,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  error?: string;
  maxLength: number;
  placeholder?: string;
  disabled?: boolean;
  grayBackground?: boolean;
}): JSX.Element {
  const currentLength = value ? value.length : 0;
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);

  const bgColor = disabled && grayBackground ? "#f3f4f6" : "#ffffff";

  return (
    <div>
      <div className="relative">
        <textarea
          ref={textareaRef}
          value={value || ""}
          placeholder={placeholder}
          maxLength={maxLength}
          disabled={disabled}
          onChange={(e) => {
            const text = e.target.value;
            const lines = text.split("\n").slice(0, 2);
            onChange(lines.join("\n"));
          }}
          onKeyDown={(e) => {
            const lines = value.split("\n");
            if (e.key === "Enter" && lines.length >= 2) {
              e.preventDefault();
            }
          }}
          onPaste={(e) => {
            e.preventDefault();
            const el = textareaRef.current;
            if (!el) return;

            const pasted = e.clipboardData.getData("text");
            const start = el.selectionStart ?? value.length;
            const end = el.selectionEnd ?? value.length;

            const merged = value.slice(0, start) + pasted + value.slice(end);

            const lines = merged.split("\n").slice(0, 2);
            const result = lines.join("\n").slice(0, maxLength);

            onChange(result);

            const newCursor = Math.min(start + pasted.length, result.length);
            requestAnimationFrame(() => {
              el.selectionStart = newCursor;
              el.selectionEnd = newCursor;
            });
          }}
          className={[
            "w-full rounded-md",
            "px-3 pt-[18px] pb-3",
            "text-[12px]",
            "resize-none",
            "min-h-[90px]",
            "overflow-hidden",
            disabled && grayBackground ? "text-gray-500" : "text-gray-700",
            "disabled:opacity-100",
            "disabled:cursor-not-allowed",
            "border-2 border-[#e7e7e7] focus:outline-none focus:border-blue-500 focus:ring-0 hover:!border-[#76A8E8]",

          ].join(" ")}
          style={{ backgroundColor: bgColor }}
        />

        <div
          className="absolute left-3 top-1.5 z-10 text-[10px] text-gray-500 pointer-events-none"
          style={{ backgroundColor: bgColor }}
        >
          <span className="text-[#E56D71]">*</span>{label} ({currentLength}/{maxLength})
        </div>

      </div>

      {error && <div className="mt-1 text-[10px] text-[#E56D71]">{error}</div>}

    </div>
  );
}