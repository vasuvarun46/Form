import * as React from "react";

type QtyStepperProps = {
  value: number;
  onDec: () => void;
  onInc: () => void;
  onChange: (v: string | number) => void;
  disabled?: boolean;
  emptyValue?: number;
};

export default function QtyStepper({
  value,
  onDec,
  onInc,
  onChange,
  disabled,
  emptyValue = 1
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

  return (
    <div className="pdf-thai mx-auto w-[120px] flex items-center rounded-md overflow-hidden bg-white border border-[#0a66d8]">
      <button
        type="button"
        className="w-9 h-8 bg-white hover:bg-[#f2f7ff] text-gray-700 border-r border-[#dbe8ff]"
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
        className="w-9 h-8 bg-white hover:bg-[#f2f7ff] text-gray-700 border-l border-[#dbe8ff]"
        onClick={onInc}
        disabled={disabled}
        aria-label="Increase quantity"
      >
        +
      </button>
    </div>
  );
}
