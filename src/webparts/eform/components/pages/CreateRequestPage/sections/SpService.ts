import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Fetches SharePoint regional timezone offset once.
 * Returns offset in minutes east of UTC.
 * e.g. Bangkok UTC+7 → returns 420
 */
export function useSharePointTimezoneOffset(context: WebPartContext): number {
  const [offsetMinutes, setOffsetMinutes] = React.useState<number>(0);

  React.useEffect(() => {
    const fetch = async (): Promise<void> => {
      try {
        const siteUrl = context.pageContext.web.absoluteUrl;
        const res = await context.spHttpClient.get(
          `${siteUrl}/_api/web/RegionalSettings/TimeZone`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "odata-version": "",
            },
          }
        );
        if (!res.ok) return;
        const json = await res.json() as { Bias?: number; DaylightBias?: number };
        // SP Bias is minutes west of UTC (opposite sign to JS)
        // e.g. Bangkok UTC+7 → Bias = -420
        const bias = Number(json.Bias ?? 0);
        const daylightBias = Number(json.DaylightBias ?? 0);
        setOffsetMinutes(-(bias + daylightBias));
      } catch (e) {
        console.error("Failed to fetch SP timezone:", e);
      }
    };
    fetch().catch(console.error);
  }, [context]);

  return offsetMinutes;
}