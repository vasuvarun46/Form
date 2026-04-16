import { HttpClient } from "@microsoft/sp-http";

export async function postApprovalAction(
  httpClient: HttpClient,
  flowUrl: string,
  payload: {
    requestId: number;
    action: "Approve" | "Reject" | "Bypass";
    userEmail: string;
    comments?: string;
  }
): Promise<unknown> {

  const response = await httpClient.post(
    flowUrl,
    HttpClient.configurations.v1,
    {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || response.statusText);
  }

  return await response.json();
}
