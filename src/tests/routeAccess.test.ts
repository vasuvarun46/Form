
import { canAccessRoute, getGoToFormRoute, Route, UserRole } from "../tests/utils/routeAccess";

describe("Route access rules", () => {
  test("admin access", () => {
    const role: UserRole = "admin";

    expect(canAccessRoute(role, "all" as Route)).toBe(true);
    expect(canAccessRoute(role, "create" as Route)).toBe(true);
    expect(canAccessRoute(role, "team" as Route)).toBe(false);
  });

  test("member access", () => {
    const role: UserRole = "member";

    expect(canAccessRoute(role, "team" as Route)).toBe(true);
    expect(canAccessRoute(role, "my" as Route)).toBe(true);
    expect(canAccessRoute(role, "create")).toBe(true)
  });

  test("visitor access", () => {
    const role: UserRole = "visitor";

    expect(canAccessRoute(role, "my" as Route)).toBe(true);
    expect(canAccessRoute(role, "all" as Route)).toBe(false);
  });
});

describe("Go-to-form routing", () => {
  test("member goes to team", () => {
    const role: UserRole = "member";
    expect(getGoToFormRoute(role)).toBe("team");
  });

  test("visitor goes home", () => {
    const role: UserRole = "visitor";
    expect(getGoToFormRoute(role)).toBe("home");
  });
});
