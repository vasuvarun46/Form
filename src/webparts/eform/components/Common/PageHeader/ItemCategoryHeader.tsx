import * as React from "react";
import BaseInnerHeader from "./BaseInnerHeader";

type Props = {
  onExport: () => Promise<void> | void;
  disabled?: boolean;
};

export default function ItemCategoryHeader(props: Props): JSX.Element {
  const { onExport, disabled } = props;
  const [isExporting, setIsExporting] = React.useState(false);

  const exportDisabled = !!disabled || isExporting;

  const handleExport = async (): Promise<void> => {
    if (exportDisabled) return;
    setIsExporting(true);
    try {
      await onExport();
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <BaseInnerHeader
      title="E-Form Product Master Export"
      subtitle="Item Category"
      right={
        <div className="flex items-center gap-2">
          <button
            type="button"
            disabled={exportDisabled}
            onClick={handleExport}
            className={[
              "h-8 px-3 rounded-md text-[12px] font-medium transition",
              exportDisabled
                ? "bg-gray-200 text-gray-500 cursor-not-allowed"
                : "bg-black text-white hover:bg-[#2a2a2a]",
            ].join(" ")}
          >
            {isExporting ? "Exporting..." : "Export for E-Form"}
          </button>
        </div>
      }
    />
  );
}
