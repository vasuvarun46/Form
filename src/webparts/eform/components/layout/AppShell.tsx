
import * as React from "react";
import { Route, } from "../Eform"; // ✅ add UserRole
import { useIsDesktop } from "../hooks/useIsDesktop";
import AppHeader from "../Common/AppHeader";
import Sidebar from "../Common/Sidebar";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SecurityPermissions } from "../spHelpers";

interface Props {
  children: React.ReactNode;
  activeRoute: Route;
  onNavigate: (r: Route) => void;
  onGoForm: () => void;
  context: WebPartContext;         // ✅ add
  roleReady: boolean;
  exportSource?: Route;
  setRuntimeFlowUrl?: (url: string) => void;
  permissions: SecurityPermissions;
  // ✅ add (optional)
}


export default function AppShell({
  children, activeRoute, onNavigate, onGoForm, context, exportSource, setRuntimeFlowUrl, permissions,
}: Props): JSX.Element {
  const isDesktop = useIsDesktop(1285);
  const [drawerOpen, setDrawerOpen] = React.useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = React.useState(false);


  const isHome = activeRoute === "home";

  React.useEffect(() => {
    if (isDesktop) setDrawerOpen(false);
  }, [isDesktop]);

  const navTo = (r: Route): void => {
    onNavigate(r);
    setDrawerOpen(false);
  };

  const userName = context.pageContext.user.displayName || "User";
  const userEmail = context.pageContext.user.email || "";

  return (
    // lock to viewport, app controls scroll
    <div className="fixed inset-0 bg-[#F6F6F6] flex flex-col overflow-hidden">

      {/* Header */}
      <div className="shrink-0 sticky top-0 z-40">
        <AppHeader
          onOpenSidebar={() => setDrawerOpen(true)}
          onGoForm={onGoForm}
          showHamburger={!isDesktop && !isHome}
          userName={userName}
          userEmail={userEmail}
          spContext={context}
          role={permissions.viewAll}
          setRuntimeFlowUrl={setRuntimeFlowUrl}
        />
      </div>

      {/* Body (sidebar + content) */}
      <div className="flex-1 min-h-0 flex">
        {/* ✅ Desktop sidebar should look like screenshot-1 */}
        {isDesktop && !isHome && (
          <aside
            className={[
              "shrink-0 min-w-0 border-r border-gray-200 bg-[#F2F2F2]",
              "sticky top-16 h-[calc(100dvh-64px)]",
              "overflow-hidden",
              sidebarCollapsed ? "w-[72px]" : "w-[236px]",
              "transition-[width] duration-200",
            ].join(" ")}
          >
            <Sidebar
              activeRoute={activeRoute}
              onNavigate={navTo}
              collapsed={sidebarCollapsed}
              onToggleCollapse={() => setSidebarCollapsed((p) => !p)}

              exportSource={exportSource}
              permissions={permissions}
            />

          </aside>
        )}

        {/* ✅ Main content scrolls (this is your vertical scrollbar) */}
        <main className="flex-1 min-w-0 min-h-0 flex flex-col overflow-hidden">


          {isHome ? (
            <div className="flex-1 overflow-y-auto overflow-x-hidden">
              {children}
            </div>
          ) : (
            <div className="flex-1 min-h-0 w-full max-w-full px-0.1 sm:px-6 lg:px-1 py-0.1 sm:py-1 overflow-y-auto overflow-x-hidden">
              {children}
            </div>
          )}



        </main>

        {/* Mobile drawer */}
        {!isDesktop && !isHome && drawerOpen && (
          <div className="fixed inset-0 z-50">
            <div
              className="absolute inset-0 bg-black/40"
              onClick={() => setDrawerOpen(false)}
            />
            <div className="absolute left-0 top-0 h-full w-[236px] bg-white shadow-xl">
              <div className="h-16 px-3 flex items-center justify-end border-b border-gray-200">
                <button
                  className="h-9 w-9 rounded-md hover:bg-gray-100"
                  onClick={() => setDrawerOpen(false)}
                  aria-label="Close sidebar"
                  type="button"
                >
                  ✕
                </button>
              </div>
              <Sidebar
                activeRoute={activeRoute}
                onNavigate={navTo}
                isDrawer
                permissions={permissions}
                exportSource={exportSource}
              />


            </div>
          </div>
        )}
      </div>
    </div>
  );
}
