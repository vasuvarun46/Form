import {
  compileFilterSets,
  matchesCompiledFilters,
} from "../webparts/eform/components/pages/AllRequestPage/allRequestFilterUtils";

describe("allRequestFilterUtils", () => {
  test("matches purpose filter case-insensitively", () => {
    const filters = compileFilterSets({
      statusFilter: [],
      purposeFilter: ["it equipment"],
      costCenterFilter: [],
    });

    expect(
      matchesCompiledFilters(
        { status: "Approved", purpose: "IT Equipment", costCenter: "CC-100" },
        filters
      )
    ).toBe(true);
  });

  test("matches cost center filter with trimmed values", () => {
    const filters = compileFilterSets({
      statusFilter: [],
      purposeFilter: [],
      costCenterFilter: [" CC-100 "],
    });

    expect(
      matchesCompiledFilters(
        { status: "Approved", purpose: "Hardware", costCenter: "CC-100" },
        filters
      )
    ).toBe(true);
  });

  test("matches Pending filter against In Approval row status", () => {
    const filters = compileFilterSets({
      statusFilter: ["Pending"],
      purposeFilter: [],
      costCenterFilter: [],
    });

    expect(
      matchesCompiledFilters(
        { status: "In Approval", purpose: "Hardware", costCenter: "CC-100" },
        filters
      )
    ).toBe(true);
  });

  test("supports OR semantics within status filter", () => {
    const filters = compileFilterSets({
      statusFilter: ["Approved", "Rejected"],
      purposeFilter: [],
      costCenterFilter: [],
    });

    expect(
      matchesCompiledFilters(
        { status: "Rejected", purpose: "Hardware", costCenter: "CC-100" },
        filters
      )
    ).toBe(true);
  });

  test("supports AND semantics across different filter groups", () => {
    const filters = compileFilterSets({
      statusFilter: ["Approved"],
      purposeFilter: ["IT Equipment"],
      costCenterFilter: ["CC-100"],
    });

    expect(
      matchesCompiledFilters(
        { status: "Approved", purpose: "IT Equipment", costCenter: "CC-100" },
        filters
      )
    ).toBe(true);

    expect(
      matchesCompiledFilters(
        { status: "Approved", purpose: "IT Equipment", costCenter: "CC-999" },
        filters
      )
    ).toBe(false);
  });

  test("does not falsely match empty or placeholder row values", () => {
    const filters = compileFilterSets({
      statusFilter: [],
      purposeFilter: ["IT Equipment"],
      costCenterFilter: ["CC-100"],
    });

    expect(
      matchesCompiledFilters(
        { status: "Approved", purpose: "-", costCenter: "" },
        filters
      )
    ).toBe(false);
  });
});
