import { defineConfig, Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

// Transform CommonJS require() for static assets into Vite-compatible new URL() calls.
// Vite statically analyzes `new URL(path, import.meta.url)` and correctly hashes/bundles assets.
function requireAssets(): Plugin {
  return {
    name: 'require-assets',
    enforce: 'pre',
    transform(code: string, id: string) {
      if (!/\.[tj]sx?$/.test(id)) return null;
      if (!code.includes('require(')) return null;
      const transformed = code.replace(
        /require\((['"`])((?:\.\.?\/)[^'"`)]+)\1\)/g,
        (_m, q, p) => `(new URL(${q}${p}${q}, import.meta.url).href)`
      );
      return transformed !== code ? { code: transformed, map: null } : null;
    },
  };
}

export default defineConfig({
  plugins: [requireAssets(), react()],
  resolve: {
    alias: {
      '@microsoft/sp-webpart-base': path.resolve(__dirname, 'src/lib/spShim.ts'),
      '@microsoft/sp-http': path.resolve(__dirname, 'src/lib/spShim.ts'),
      '@microsoft/sp-core-library': path.resolve(__dirname, 'src/lib/spShim.ts'),
      '@microsoft/sp-component-base': path.resolve(__dirname, 'src/lib/spShim.ts'),
      '@microsoft/sp-property-pane': path.resolve(__dirname, 'src/lib/spShim.ts'),
      'EformWebPartStrings': path.resolve(__dirname, 'src/lib/strings.ts'),
    },
  },
  css: {
    modules: {
      localsConvention: 'camelCase',
    },
  },
  server: {
    port: 3000,
    open: true,
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
  },
});
