import * as React from "react";

type IconProps = {
  className?: string;
};

export function IconHamburger({ className = "h-5 w-5" }: IconProps): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <path
        fill="currentColor"
        d="M4 6h16v2H4V6zm0 5h16v2H4v-2zm0 5h16v2H4v-2z"
      />
    </svg>
  );
}

export function IconChevronDown({ className = "h-4 w-4" }: IconProps): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <path
        fill="currentColor"
        d="M7.41 8.59 12 13.17l4.59-4.58L18 10l-6 6-6-6z"
      />
    </svg>
  );
}

export function IconSearch({ className = "h-4 w-4" }: IconProps): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <path
        fill="currentColor"
        d="M10 4a6 6 0 104.472 10.03l3.749 3.75 1.414-1.415-3.75-3.749A6 6 0 0010 4zm0 2a4 4 0 110 8 4 4 0 010-8z"
      />
    </svg>
  );
}

/** Red document icon used in Home page cards (like screenshot) */
export function IconDocRed({ className = "h-6 w-6" }: IconProps): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <path
        fill="currentColor"
        d="M6 2h9l5 5v15a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z"
        opacity="0.18"
      />
      <path
        fill="currentColor"
        d="M15 2v5h5v2h-7V2h2z"
      />
      <path
        fill="currentColor"
        d="M7 11h10v1.8H7V11zm0 4h10v1.8H7V15zm0-8h7v1.8H7V7z"
      />
    </svg>
  );
}

export function IconClose({ className = "h-5 w-5" }: IconProps): JSX.Element {
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <path
        fill="currentColor"
        d="M18.3 5.71 12 12l6.3 6.29-1.41 1.42L10.59 13.4 4.29 19.71 2.88 18.3 9.17 12 2.88 5.71 4.29 4.29 10.59 10.6l6.3-6.31 1.41 1.42z"
      />
    </svg>
  );
}
