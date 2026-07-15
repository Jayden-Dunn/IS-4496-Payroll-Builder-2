import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

// This project lives in a folder whose name contains a literal "%" ("…for %'s
// Attempt 1"). Vite's dev server runs decodeURIComponent() over paths built
// from the project root, and "%'s" is an invalid percent-escape, so every
// request dies with "URI malformed" (vitejs/vite#9583). The build is fine —
// only the dev server hits this. Guard both decoders so an invalid escape
// falls back to the raw string, which IS the correct filesystem path here.
for (const key of ['decodeURIComponent', 'decodeURI'] as const) {
  const original = globalThis[key];
  globalThis[key] = ((value: string): string => {
    try {
      return original(value);
    } catch {
      return value;
    }
  }) as typeof original;
}

// Local-first app: everything is bundled, nothing is fetched at runtime.
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    chunkSizeWarningLimit: 2000,
  },
});
