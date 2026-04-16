// ✅ UI-only mocks (no SharePoint runtime)
jest.mock("@microsoft/sp-core-library", () => ({}));
jest.mock("@microsoft/sp-http", () => ({}));
jest.mock("@microsoft/sp-webpart-base", () => ({}));

import * as React from "react";
import { render, screen } from "@testing-library/react";
import Sidebar from "../webparts/eform/components/Common/Sidebar";
import { makePerms } from "../webparts/eform/components/spHelpers";

describe("Sidebar rendering", () => {
  test("visitor sees Requests menu", () => {
    render(
      <Sidebar
        activeRoute="my"
        onNavigate={jest.fn()}
        collapsed={false}
        onToggleCollapse={jest.fn()}
        exportSource="create"
        permissions={makePerms({ viewOwn: true })}
      />
    );

    expect(screen.getByText("Requests")).toBeInTheDocument();
  });

  test("admin sees Company-wide menu", () => {
    render(
      <Sidebar
        activeRoute="all"
        onNavigate={jest.fn()}
        collapsed={false}
        onToggleCollapse={jest.fn()}
        exportSource="create"
        permissions={makePerms({ viewAll: true })}
      />
    );

    expect(screen.getByText("Company-wide")).toBeInTheDocument();
  });
});