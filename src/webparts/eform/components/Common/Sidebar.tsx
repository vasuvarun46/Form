import * as React from "react";
import { Route } from "../Eform";
import { IconChevronDown } from "../ui/icons";
import { SecurityPermissions } from "../spHelpers";

interface SidebarProps {
  activeRoute: Route;
  onNavigate: (r: Route) => void;
  isDrawer?: boolean;
  collapsed?: boolean;
  onToggleCollapse?: () => void;
  exportSource?: Route;
  permissions?: SecurityPermissions;
}

type LocationItem = {
  title: string;
  subtitle: string;
  active?: boolean;
  comingSoon?: boolean;
};

const LOCATION_ITEMS: LocationItem[] = [
  {
    title: "แบบเบิกสินค้าคงเหลือเพื่อใช้ภายใน",
    subtitle: "Internal Usage Inventory Request",
    active: true,
  },
  {
    title: "แบบเบิกค่าใช้จ่ายในการเดินทาง",
    subtitle: "Travelling Reimbursement Form",
    comingSoon: true,
  },
  {
    title: "แบบเบิกค่าใช้จ่าย",
    subtitle: "Expense Reimbursement Form",
    comingSoon: true,
  },
  {
    title: "แบบแจ้งชำระเบิกล่วงจ่าย",
    subtitle: "Supplier Form",
    comingSoon: true,
  },
  {
    title: "แบบเงินทดรองจ่าย",
    subtitle: "Cash Advance Request Form",
    comingSoon: true,
  },
  {
    title: "แบบคืนเงินทดรองจ่าย",
    subtitle: "Cash Advance Clearing Form",
    comingSoon: true,
  },
  {
    title: "แบบเบิกเงินสดย่อย",
    subtitle: "Petty Cash Form",
    comingSoon: true,
  },
];

export default function Sidebar({
  activeRoute,
  onNavigate,
  isDrawer = false,
  collapsed = false,
  onToggleCollapse,
  exportSource,
  permissions,
}: SidebarProps): JSX.Element {
  const [locationExpanded, setLocationExpanded] = React.useState<boolean>(false);

  const itemBase = "relative w-full text-left transition-colors";
  const activeItem = "font-semibold text-gray-900 bg-transparent";
  const normalItem = "text-[#3C3C3C] hover:bg-[#F6F6F6] hover:text-[#222222] hover:font-bold";


  const isActive = (r: Route): boolean => activeRoute === r;
  const isMySectionActive =
    activeRoute === "my" ||
    activeRoute === "create" ||
    (activeRoute === "exportPdf" && exportSource === "my");

  const isTeamSectionActive =
    activeRoute === "team" ||
    (activeRoute === "exportPdf" && exportSource === "team");

  const isAllSectionActive =
    activeRoute === "all" ||
    (activeRoute === "exportPdf" && exportSource === "all");

  const Item = ({
    route,
    label,
    showActiveBar = true,
    forceActive = false,
  }: {
    route: Route;
    label: string;
    showActiveBar?: boolean;
    forceActive?: boolean;
  }): JSX.Element => {
    const active = forceActive || isActive(route);

    return (
      <button
        type="button"
        onClick={() => onNavigate(route)}
        className={[itemBase, "pl-5 pr-3 py-1.5 rounded-md", active ? activeItem : normalItem].join(" ")}
        title={label}
      >
        {showActiveBar && active && !collapsed && (
          <span className="absolute left-0 top-1 bottom-1 w-[3px] bg-[#E60023]" />

        )}

        <div className="flex items-center gap-3">
          {!collapsed && <span className="text-[12px]">{label}</span>}
        </div>
      </button>
    );
  };

  return (
    <aside
      className={[
        "h-full flex flex-col pdf-thai relative",
        isDrawer ? "bg-white" : "bg-[#f2f2f2]",
        collapsed ? "w-[76px]" : "w-full",
      ].join(" ")}
    >
      {!isDrawer && onToggleCollapse && (
        <button
          type="button"
          onClick={onToggleCollapse}
          className="absolute right-2 top-1/2 -translate-y-1/2 z-20 h-8 w-8 rounded-md grid place-items-center hover:bg-black/5 active:bg-black/10"
          aria-label={collapsed ? "Expand sidebar" : "Collapse sidebar"}
          title={collapsed ? "Expand" : "Collapse"}
        >
          <span
            className={[
              "inline-block transition-transform text-gray-600",
              collapsed ? "rotate-180" : "",
            ].join(" ")}
          >
            <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true">
              <path fill="currentColor" d="M15.41 7.41 14 6l-6 6 6 6 1.41-1.41L10.83 12z" />
            </svg>
          </span>
        </button>
      )}

      <div className={["shrink-0", collapsed ? "px-3 pt-4" : "px-5 pt-5"].join(" ")}>
        <div className="flex items-center justify-between">

        </div>

        {!collapsed && (
          <div className="relative w-full rounded-xl bg-white border border-gray-200 shadow-sm px-3 py-3">
            <button
              type="button"
              onClick={() => setLocationExpanded((p) => !p)}
              className="w-full text-left"
            >
              <div className="flex items-center justify-between">
                <div className="text-[11px] font-semibold text-gray-900">Current location</div>
                <span
                  className={[
                    "shrink-0 h-5 w-5 rounded-full bg-[#FBE9EA] flex items-center justify-center text-[#E60023] transition-all cursor-pointer",
                    locationExpanded ? "rotate-180" : "",
                  ].join(" ")}
                  style={{ pointerEvents: "all" }}
                  onMouseOver={(e) => { e.currentTarget.style.backgroundColor = "#E8E8E8"; }}
                  onMouseOut={(e) => { e.currentTarget.style.backgroundColor = "#FBE9EA"; }}
                >
                  <IconChevronDown />
                </span>
              </div>

              {!locationExpanded && (
                <div className="mt-2 min-w-0">
                  <div className="text-[11px] text-gray-700 truncate">{LOCATION_ITEMS[0].title}</div>
                  {/* ✅ CHANGED: "/" in red */}
                  <div className="text-[11px] text-gray-500 truncate">
                    <span className="text-[#E60023]">/</span>
                    <span className="ml-1">{LOCATION_ITEMS[0].subtitle}</span>
                  </div>
                </div>
              )}
            </button>

            {locationExpanded && (
              <div className="mt-3 border-t border-gray-100 pt-2 space-y-2">
                {LOCATION_ITEMS.map((item, idx) => (
                  <div
                    key={`${item.title}-${idx}`}
                    className={[
                      "rounded-md px-2 py-2",
                      item.active ? "bg-[#f7f7f7] border border-[#ededed]" : "",
                    ].join(" ")}
                  >
                    <div className="flex items-start justify-between gap-2">
                      <div className="min-w-0">
                        <div className="text-[10px] text-gray-700 truncate">{item.title}</div>
                        <div className="text-[10px] text-gray-500 truncate">
                          <span className="text-[#E60023]">/</span>
                          <span className="ml-1">{item.subtitle}</span>
                        </div>
                      </div>
                      {item.active && (
                        <span className="text-[#0a66d8] text-[12px] font-semibold">✓</span>
                      )}
                    </div>
                    {item.comingSoon && (
                      <div className="mt-1 inline-flex h-5 items-center rounded-md bg-[#f3f3f3] px-2 text-[10px] text-gray-600">
                        Coming Soon
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      {!locationExpanded && (
        <nav
          className={[
            "flex-1",
            collapsed ? "overflow-hidden" : "overflow-y-auto",
            "overflow-x-hidden",
            collapsed ? "px-2 pt-4" : "px-4 pt-5",
            "pb-6",
          ].join(" ")}
        >
          <div className="space-y-1">
            {permissions?.viewOwn && <Item route="my" label="My requests" forceActive={isMySectionActive} />}
            {permissions?.viewTeam && <Item route="team" label="Team's requests" forceActive={isTeamSectionActive} />}
            {permissions?.viewAll && <Item route="all" label="All requests" forceActive={isAllSectionActive} />}
          </div>
        </nav>
      )}
    </aside>
  );
}
