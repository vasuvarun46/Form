import * as React from "react";

interface Props {
  children: React.ReactNode;
}

interface State {
  hasError: boolean;
  error?: Error;
}

export default class ErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error("[ErrorBoundary] Uncaught error:", error, info);
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div className="flex flex-col items-center justify-center min-h-[300px] p-8 text-center">
          <div className="text-red-500 text-4xl mb-4">⚠️</div>
          <h2 className="text-lg font-semibold text-gray-900 mb-2">
            Something went wrong
          </h2>
          <p className="text-sm text-gray-600 mb-4 max-w-md">
            An unexpected error occurred. Please refresh the page and try again.
          </p>
          <p className="text-xs text-gray-500 font-mono">
            {this.state.error?.message || "Unknown error"}
          </p>
          <button
            type="button"
            onClick={() => window.location.reload()}
            className="mt-4 h-9 px-5 rounded-md bg-[#0a66d8] text-white text-sm font-semibold hover:bg-[#095ec7]"
          >
            Reload Page
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
