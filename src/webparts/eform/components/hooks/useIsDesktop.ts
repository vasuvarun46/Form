// hooks/useIsDesktop.ts
import * as React from "react";

export function useIsDesktop(bp: number = 1024): boolean {
  const [isDesktop, setIsDesktop] = React.useState(() => {
    if (typeof window === "undefined") return false;
    return window.matchMedia(`(min-width:${bp}px)`).matches;
  });

  React.useEffect(() => {
    const mq = window.matchMedia(`(min-width:${bp}px)`);
    const onChange = (): void => setIsDesktop(mq.matches);

    onChange();

    if (mq.addEventListener) {
      mq.addEventListener("change", onChange);
      return () => mq.removeEventListener("change", onChange);
    } else {
      mq.addListener(onChange);
      return () => mq.removeListener(onChange);
    }
  }, [bp]);

  return isDesktop;
}
