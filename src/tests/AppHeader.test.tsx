// ✅ UI-only mocks (no SharePoint runtime)
jest.mock("@microsoft/sp-core-library", () => ({}));
jest.mock("@microsoft/sp-http", () => ({}));
jest.mock("@microsoft/sp-webpart-base", () => ({}));

import * as React from "react";
import { render, screen } from "@testing-library/react";



import AppHeader from "../webparts/eform/components/Common/AppHeader";

describe("AppHeader UI", () => {
  test("renders user name and email", () => {
    render(
      <AppHeader
        onOpenSidebar={jest.fn()}   // ✅ FIXED
        onGoForm={jest.fn()}       // ✅ FIXED
        showHamburger={false}
        userName="Test User"
        userEmail="test@example.com"
        role={true}
      />
    );

    expect(screen.getByText("Hi, Test User")).toBeInTheDocument();
    expect(screen.getByText("test@example.com")).toBeInTheDocument();
  });
});
