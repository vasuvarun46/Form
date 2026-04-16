// ✅ DEFINE TYPES LOCALLY (no SPFx dependency)
export type Route = "home" | "create" | "exportPdf" | "team" | "all" | "my";
export type UserRole = "admin" | "member" | "visitor";

export const getGoToFormRoute = (role: UserRole): Route => {
  return role === "member" ? "team" : "home";
};

export const canAccessRoute = (role: UserRole, route: Route): boolean => {
  if (route === "home" || route === "exportPdf") return true;

  if (role === "admin") return route === "all" || route === "create" || route === "my";
  if (role === "member") return route === "team";

  return route === "create" || route === "my";
};
