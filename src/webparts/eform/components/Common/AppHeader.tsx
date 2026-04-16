// AppHeader.tsx
import * as React from "react";
import { IconHamburger } from "../ui/icons";
import {
  Callout,
  DirectionalHint,
  Icon,
  IconButton,
  ICalloutContentStyles,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import GqEformLogo from "../../assets/EformLogo.png";
import * as strings from "EformWebPartStrings";




interface Props {
  onOpenSidebar: () => void;
  onGoForm: () => void;
  showHamburger: boolean;

  userName: string;
  userEmail: string;

  spContext?: WebPartContext; // SPFx WebPartContext
  role: boolean; // ✅ add role prop (e.g. "admin", "user", etc.)
  setRuntimeFlowUrl?: (url: string) => void;
}
type PageContextUserLike = {
  loginName?: string;
  userPrincipalName?: string;
};


export default function AppHeader({
  onOpenSidebar,
  onGoForm,
  showHamburger,
  userName,
  userEmail,
  spContext,
  role
}: Props): JSX.Element {
  const [photoUrl, setPhotoUrl] = React.useState<string>("");

  // ✅ Profile callout (like code-2)
  const [isProfileOpen, setIsProfileOpen] = React.useState(false);
  const profileBtnRef = React.useRef<HTMLButtonElement | null>(null);
  const companyName = spContext?.pageContext?.web?.title ?? "Company";
  function getInitials(name?: string, email?: string): string {

    const src = (name || "").trim();
    if (src) {
      const parts = src.split(/\s+/).filter(Boolean);
      const first = parts[0]?.[0] ?? "";
      const last = (parts.length > 1 ? parts[parts.length - 1][0] : "") ?? "";
      const ini = (first + last).toUpperCase();
      if (ini) return ini;
    }
    if (email) return email.slice(0, 2).toUpperCase();
    return "NA";
  }

  const initials = getInitials(userName, userEmail);
  const user = spContext?.pageContext?.user as PageContextUserLike | undefined;

  const resolvedAccount =
    userEmail ||
    user?.loginName ||
    user?.userPrincipalName ||
    "";


  React.useEffect(() => {
    let cancelled = false;

    const loadPhoto = async (): Promise<void> => {
      try {
        const webUrl: string =
          spContext?.pageContext?.web?.absoluteUrl ||
          spContext?.pageContext?.site?.absoluteUrl ||
          window.location.origin;

        const apiUrl = `${webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=PictureUrl`;

        const res = await spContext?.spHttpClient?.get(apiUrl, SPHttpClient.configurations.v1);
        if (!res || !res.ok) return;

        const json = await res.json();
        const pic = (json?.PictureUrl || "") as string;

        if (!cancelled && pic) setPhotoUrl(pic);
      } catch {
        // ignore (fallback to initials)
      }
    };

    if (spContext?.spHttpClient) loadPhoto().catch((e) => {
      console.error(e, "Error in loadPhoto()");
    })

    return () => {
      cancelled = true;
    };
  }, [spContext]);


  /** ✅ Open real Microsoft 365 App Launcher (9 dots) */
  const tryOpenO365AppLauncher = React.useCallback((): void => {
    const clickIfFound = (): boolean => {
      const btn = document.getElementById("O365_MainLink_NavMenu") as HTMLButtonElement | null;
      if (btn) {
        btn.click();
        return true;
      }
      return false;
    };

    if (clickIfFound()) return;

    let tries = 0;
    const maxTries = 12;
    const interval = window.setInterval(() => {
      tries++;
      if (clickIfFound() || tries >= maxTries) window.clearInterval(interval);
    }, 250);
  }, []);

  /** ✅ Trigger REAL suite-bar Settings (gear) */
  const tryOpenO365Settings = React.useCallback((): void => {
    const clickIfFound = (): boolean => {
      const byId =
        (document.getElementById("O365_MainLink_Settings") as HTMLElement | null) ||
        (document.getElementById("O365_MainLink_SettingsMenu") as HTMLElement | null);

      if (byId) {
        byId.click();
        return true;
      }

      const iconSpan = document.querySelector("div.o365cs-base span.ms-Icon--Settings") as HTMLElement | null;
      if (iconSpan) {
        const clickable = iconSpan.closest('button, a, div[role="button"], span[role="button"]') as HTMLElement | null;
        (clickable ?? iconSpan).click();
        return true;
      }

      const ariaBtn =
        (document.querySelector('#SuiteNavWrapper button[aria-label*="Settings"]') as HTMLElement | null) ||
        (document.querySelector('#SuiteNavWrapper a[aria-label*="Settings"]') as HTMLElement | null);

      if (ariaBtn) {
        ariaBtn.click();
        return true;
      }

      return false;
    };

    if (clickIfFound()) return;

    let tries = 0;
    const maxTries = 12;
    const interval = window.setInterval(() => {
      tries++;
      if (clickIfFound() || tries >= maxTries) window.clearInterval(interval);
    }, 250);
  }, []);

  const signOut = (): void => {
    window.location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/logout";
  };

  const calloutStyles = React.useMemo<Partial<ICalloutContentStyles>>(
    () => ({
      root: { zIndex: 99999 },
      calloutMain: {
        borderRadius: 0,
        overflow: "hidden" as const,
        boxShadow: "0 12px 32px rgba(0, 0, 0, 0.14)",
      },
    }),
    []
  );


  return (
    <header className="pdf-thai z-40 bg-white border-b border-gray-200">
      <div className="relative h-16 px-4 lg:px-6 flex items-center justify-between">
        {/* LEFT */}
        <div className="flex items-center gap-3">
          {showHamburger && (
            <button
              type="button"
              onClick={onOpenSidebar}
              className="h-10 w-10 rounded-md hover:bg-gray-100 flex items-center justify-center"
              aria-label="Open sidebar"
              title="Open sidebar"
            >
              <IconHamburger />
            </button>
          )}

          {/* 9-dots */}
          <button
            type="button"
            onClick={tryOpenO365AppLauncher}
            className="h-10 w-10 rounded-md hover:bg-gray-100 flex items-center justify-center"
            aria-label="Microsoft 365 apps"
            title="Microsoft 365 apps"
          >
            <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
              {[0, 7, 14].map((x) =>
                [0, 7, 14].map((y) => (
                  <rect key={`${x}-${y}`} x={x} y={y} width="3" height="3" rx="1" fill="currentColor" />
                ))
              )}
            </svg>
          </button>

          {/* ✅ Logo (IMAGE) - replaces GQ + E-FORM text */}
          <button
            type="button"
            onClick={() => {
              setIsProfileOpen(false);
              onGoForm();
            }}
            className="flex items-center select-none"
            aria-label="Go to Home"
            title="Go to Home"
          >
            <img
              src={GqEformLogo}
              alt="GQ E-Form"
              className="h-12 w-auto object-contain"
              draggable={false}
            />
          </button>
        </div>

        {/* RIGHT */}
        <div className="flex items-center h-full">
          {role && (
            <div className="pr-2">
              <IconButton
                iconProps={{ iconName: "Settings" }}
                title={strings.Settings}
                ariaLabel={strings.Settings}
                onClick={tryOpenO365Settings}
                styles={{ root: { width: 40, height: 40, borderRadius: 8 } }}
              />
            </div>
          )}

          <div className="h-full">
            <button
              type="button"
              onClick={() => setIsProfileOpen((v) => !v)}
              ref={profileBtnRef}
              className={[
                "h-full flex w-[320px] lg:w-[380px] items-center gap-3 px-4 pr-5 active:bg-gray-100 transition-colors",
                isProfileOpen ? "bg-[#f3f3f3]" : "bg-white hover:bg-[#f3f3f3]",
              ].join(" ")}
              aria-label="User profile"
              title={`${userName} (${userEmail})`}
            >
              {photoUrl ? (
                <img
                  src={photoUrl}
                  alt={userName}
                  className="h-12 w-12 rounded-full border border-gray-300 object-cover"
                  referrerPolicy="no-referrer"
                />
              ) : (
                <div className="flex h-12 w-12 items-center justify-center rounded-full border border-gray-300 bg-[#f5f5f5] text-[20px] leading-none font-normal text-[#1f1f1f]">
                  {initials}
                </div>
              )}

              <div className="profileText flex flex-col text-left leading-tight min-w-0">
                <div className="truncate text-[15px] font-semibold text-gray-900">
                  Hi,{userName}
                </div>
                <div className="mt-1 truncate text-[12px] text-[#5f5f5f]">
                  {userEmail}
                </div>
              </div>
            </button>
          </div>
          
          {isProfileOpen && profileBtnRef.current && (
            <Callout
              target={profileBtnRef.current}
              onDismiss={() => setIsProfileOpen(false)}
              directionalHint={DirectionalHint.bottomRightEdge}
              alignTargetEdge
              gapSpace={0}
              isBeakVisible={false}
              setInitialFocus={false}
              className="z-[99999]"
              styles={calloutStyles}
            >
              <div className="w-[320px] lg:w-[380px] max-w-[calc(100vw-12px)] overflow-hidden border border-gray-200 bg-white">
                <div className="flex items-center justify-between px-5 py-4 text-[14px] text-[#4a4a4a]">
                  <div className="truncate pr-3">{companyName}</div>
                  <button
                    type="button"
                    onClick={() => signOut()}
                    className="shrink-0 text-[14px] text-[#4a4a4a] hover:underline"
                  >
                    Sign out
                  </button>
                </div>

                <div className="flex gap-4 bg-white px-5 py-5">
                  <div className="flex h-[82px] w-[82px] shrink-0 items-center justify-center overflow-hidden rounded-full bg-[#082540]">
                    {photoUrl ? (
                      <img
                        src={photoUrl}
                        alt={userName}
                        className="h-full w-full object-cover"
                        referrerPolicy="no-referrer"
                      />
                    ) : (
                      <span className="text-[34px] leading-none text-white font-normal">{initials}</span>
                    )}
                  </div>

                  <div className="min-w-0 pt-1">
                    <div className="truncate text-[16px] leading-[1.2] font-semibold text-[#2c2c2c]">
                      {userName}
                    </div>
                    <div className="mt-1 truncate text-[12px] leading-[1.3] text-[#4f4f4f]">
                      {resolvedAccount}
                    </div>

                    <div className="mt-2 flex flex-col gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          window.open("https://myaccount.microsoft.com/?ref=MeControl", "_blank", "noopener,noreferrer")
                        }
                        className="text-left text-[12px] leading-[1.2] text-[#2563eb] hover:underline"
                      >
                        View account
                      </button>

                      <button
                        type="button"
                        onClick={() =>
                          window.open("https://myaccount.microsoft.com/profile/", "_blank", "noopener,noreferrer")
                        }
                        className="text-left text-[12px] leading-[1.2] text-[#2563eb] hover:underline"
                      >
                        Open another mailbox
                      </button>
                    </div>
                  </div>
                </div>

                <div className="h-px bg-gray-200" />

                <button
                  type="button"
                  onClick={() => {
                    window.location.href =
                      "https://www.office.com/login?prompt=select_account&ru=/launch/sharepoint.";
                  }}
                  className="flex w-full items-center gap-4 bg-[#f3f3f3] px-5 py-3 hover:bg-[#ececec]"
                >
                  <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full border border-[#9d9d9d] bg-transparent text-[#666]">
                    <Icon iconName="AddFriend" className="text-[18px]" />
                  </div>
                  <div className="text-left text-[14px] leading-[1.25] text-[#4a4a4a]">
                    Sign in with a different account
                  </div>
                </button>
              </div>
            </Callout>
          )}
        </div>
      </div>
    </header>
  );
}
