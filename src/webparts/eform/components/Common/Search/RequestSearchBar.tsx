// RequestSearchBar.tsx
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IconSearch, IconChevronDown } from "../../ui/icons";

export type SearchFieldKey = "TransactionID" | "RequesterName";

type RequestSearchBarProps = {
  searchDraft: string;
  setSearchDraft: (v: string) => void;
  onSearch: () => void;
  onSearchWithField?: (args: { field: SearchFieldKey; text: string }) => void;
  onClear?: () => void;
};

const FIELDS: { key: SearchFieldKey; label: string }[] = [
  { key: "TransactionID", label: "Transaction ID" },
  { key: "RequesterName", label: "Requestor name" },
];

function CheckIcon(): JSX.Element {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" className="text-[#0a66d8]">
      <path
        fill="currentColor"
        d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"
      />
    </svg>
  );
}

function ClearIcon(): JSX.Element {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M18 6L6 18M6 6l12 12" />
    </svg>
  );
}

type MenuPos = { top: number; left: number; width: number };

export default function RequestSearchBar(props: RequestSearchBarProps): JSX.Element {
  const { searchDraft, setSearchDraft, onSearch, onSearchWithField, onClear } = props;

  const [field, setField] = React.useState<SearchFieldKey>("TransactionID");

  // dropdown open + portal positioning
  const [open, setOpen] = React.useState(false);
  const [pos, setPos] = React.useState<MenuPos>({ top: 0, left: 0, width: 150 });
  const wrapRef = React.useRef<HTMLDivElement | null>(null);
  const btnRef = React.useRef<HTMLButtonElement | null>(null);

  const label = FIELDS.find((f) => f.key === field)?.label ?? "Transaction ID";
  const hasValue = searchDraft.trim().length > 0;

  const doSearch = (): void => {
    onSearch();
    if (onSearchWithField) {
      onSearchWithField({ field, text: searchDraft.trim() });
    }
  };

  const handleClear = (): void => {
    setSearchDraft("");
    if (onClear) onClear();
  };

  const updateMenuPosition = React.useCallback(() => {
    if (!btnRef.current) return;
    const r = btnRef.current.getBoundingClientRect();
    setPos({
      left: Math.round(r.left),
      top: Math.round(r.bottom + 6),
      width: Math.round(r.width),
    });
  }, []);

  // close outside click (portal-safe)
  React.useEffect(() => {
    const onDocClick = (e: MouseEvent): void => {
      const t = e.target as Node;
      const inWrap = !!wrapRef.current && wrapRef.current.contains(t);
      const inMenu = (e.target as HTMLElement | null)?.closest?.("[data-rsb-menu='1']");
      if (!inWrap && !inMenu) setOpen(false);
    };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  // reposition menu on open / scroll / resize
  React.useEffect(() => {
    if (!open) return;

    updateMenuPosition();
    const onScroll = (): void => updateMenuPosition();
    const onResize = (): void => updateMenuPosition();

    window.addEventListener("scroll", onScroll, true);
    window.addEventListener("resize", onResize);

    return () => {
      window.removeEventListener("scroll", onScroll, true);
      window.removeEventListener("resize", onResize);
    };
  }, [open, updateMenuPosition]);

  const menu = open
    ? ReactDOM.createPortal(
      <div
        data-rsb-menu="1"
        style={{
          position: "fixed",
          top: pos.top,
          left: pos.left,
          width: pos.width,
          zIndex: 2147483647,
        }}
        className="rounded-md border border-[#e6e6e6] bg-white shadow-lg py-1"
      >
        {FIELDS.map((f) => {
          const active = f.key === field;
          return (
            <button
              key={f.key}
              type="button"
              onClick={() => {
                setField(f.key);
                setOpen(false);
              }}
              className={`
                  w-full flex items-center justify-between
                  px-3 py-2 text-left text-[11px]
                  ${active ? "bg-[#eef5ff] text-gray-900" : "text-gray-700"}
                  hover:bg-gray-50
                `}
            >
              <span className="truncate">{f.label}</span>
              {active ? <CheckIcon /> : <span className="w-4" />}
            </button>
          );
        })}
      </div>,
      document.body
    )
    : null;

  return (
    <div className="w-full max-w-[640px]" ref={wrapRef}>
      <div className="h-10 w-full flex items-stretch rounded-md border border-[#ececec] bg-white">
        {/* ✅ Custom dropdown button (looks like image) */}
        <button
          ref={btnRef}
          type="button"
          onClick={() => setOpen((s) => !s)}
          className="
            relative w-[150px] shrink-0
            bg-[#f7f7f7]
            px-3 pr-9
            text-[11px] text-gray-800
            flex items-center justify-between
            hover:!bg-[#EBEBEB]
            focus:outline-none focus-visible:outline-none
            rounded-l-md
          "
        >
          <span className="truncate">{label}</span>

          <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">
            <IconChevronDown />
          </span>

          <span className="absolute right-0 top-2 bottom-2 w-px bg-[#e5e5e5]" />
        </button>

        {/* Input with X clear button */}
        <div className="relative flex-1 hover:!border-2 hover:!border-[#76A8E8]">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500">
            <IconSearch />
          </span>

          <input
            value={searchDraft}
            onChange={(e) => setSearchDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") doSearch();
            }}
            placeholder="Search"
            className={`h-full w-full bg-white pl-9 text-[11px] text-gray-700 focus:outline-none focus-visible:outline-none ${hasValue ? "pr-8" : "pr-3"
              }`}
          />


          {/* X button — only visible when input has a value */}
          {hasValue && (
            <button
              type="button"
              onClick={handleClear}
              aria-label="Clear search"
              className="absolute right-2 top-1/2 -translate-y-1/2 min-w-[20px] min-h-[20px] flex items-center justify-center text-gray-400 hover:text-gray-600 focus:outline-none"
            >
              <ClearIcon />
            </button>
          )}
        </div>

        {/* Search button (unchanged) */}
        <button
          type="button"
          onClick={doSearch}
          className="h-full px-6 bg-[#2172DA] text-white text-[11px] font-semibold rounded-r-md border-2 border-transparent hover:!bg-[#F0F7FF] hover:!text-[#81AFEB] hover:!border-[#76A8E8]"
        >
          Search
        </button>
      </div>

      {menu}
    </div>
  );
}