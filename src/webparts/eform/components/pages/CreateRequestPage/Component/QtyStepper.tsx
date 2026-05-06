import * as React from "react";

type QtyStepperProps = {
  value: number;
  onDec: () => void;
  onInc: () => void;
  onChange: (v: string | number) => void;
  disabled?: boolean;
  emptyValue?: number;
  maxStock?: number;
};

export default function QtyStepper({
  value,
  onDec,
  onInc,
  onChange,
  disabled,
  emptyValue = 1,
  maxStock

}: QtyStepperProps): JSX.Element {
  const [draftValue, setDraftValue] = React.useState<string>(String(value ?? 1));

  React.useEffect(() => {
    setDraftValue(String(value ?? 1));
  }, [value]);

  const handleTextChange = (raw: string): void => {
    const digitsOnly = raw.replace(/\D/g, "").slice(0, 5);
    setDraftValue(digitsOnly);
  };

  const commitDraftValue = (): void => {
    if (draftValue === "") {
      onChange(emptyValue);
      return;
    }
    onChange(draftValue);
  };
  const currentQty = Number(draftValue) || value;

  return (
    <div className="pdf-thai mx-auto w-[120px] flex items-center rounded-md overflow-hidden bg-white border-2 border-[#C4C4C4]">

      <button
        type="button"
        style={{ borderRight: "2px solid #C4C4C4" }}
        className={`w-9 h-8 text-gray-700 ${disabled || value <= 1 || (maxStock !== undefined && maxStock > 0 && currentQty > maxStock) ? "bg-[#D9D9D9]" : "bg-white hover:bg-[#E8E8E8]"}`}

        onClick={onDec}
        disabled={disabled || value <= 1}
        aria-label="Decrease quantity"
      >
        -
      </button>

      <input
        type="text"
        inputMode="numeric"
        pattern="[0-9]*"
        value={draftValue}
        onChange={(e) => handleTextChange(e.target.value)}
        onFocus={(e) => e.currentTarget.select()}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            commitDraftValue();
          }
        }}
        onBlur={commitDraftValue}
        className="w-full h-8 text-center text-[12px] focus:outline-none focus-visible:outline-none"
        disabled={disabled}
      />

      <button
        type="button"
        style={{ borderLeft: "2px solid #C4C4C4", borderRight: "none", borderTop: "none", borderBottom: "none" }}
        className={`w-9 h-8 text-gray-700 ${disabled || (maxStock !== undefined && maxStock > 0 && currentQty >= maxStock) ? "bg-[#D9D9D9] cursor-not-allowed" : "bg-white hover:bg-[#E8E8E8]"}`}

        onClick={onInc}
        disabled={disabled || (maxStock !== undefined && maxStock > 0 && currentQty >= maxStock)}

        aria-label="Increase quantity"
      >
        +
      </button>
    </div>
  );
}
