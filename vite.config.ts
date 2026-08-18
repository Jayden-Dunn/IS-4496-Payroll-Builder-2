import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Plugin } from 'vite';
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

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EMPLOYEE_DIR_FILE = path.join(__dirname, 'public', 'employee-directory.json');
const SETTINGS_FILE = path.join(__dirname, 'data', 'app-settings.json');

function jsonFileApi(filePath: string): (req: import('node:http').IncomingMessage, res: import('node:http').ServerResponse, next: () => void) => void {
  return (req, res, next) => {
    if (req.method === 'GET') {
      try {
        if (!fs.existsSync(filePath)) {
          res.statusCode = 404;
          res.end('Not found');
          return;
        }
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Cache-Control', 'no-store');
        res.end(fs.readFileSync(filePath, 'utf8'));
      } catch (err) {
        res.statusCode = 500;
        res.end(err instanceof Error ? err.message : 'Read failed');
      }
      return;
    }
    if (req.method === 'PUT' || req.method === 'POST') {
      const chunks: Buffer[] = [];
      req.on('data', (c: Buffer) => chunks.push(c));
      req.on('end', () => {
        try {
          const body = Buffer.concat(chunks).toString('utf8');
          JSON.parse(body);
          fs.mkdirSync(path.dirname(filePath), { recursive: true });
          fs.writeFileSync(filePath, body, 'utf8');
          res.statusCode = 204;
          res.end();
        } catch (err) {
          res.statusCode = 400;
          res.end(err instanceof Error ? err.message : 'Write failed');
        }
      });
      return;
    }
    if (req.method === 'DELETE') {
      try {
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        res.statusCode = 204;
        res.end();
      } catch (err) {
        res.statusCode = 500;
        res.end(err instanceof Error ? err.message : 'Delete failed');
      }
      return;
    }
    next();
  };
}

/**
 * Persist the Employee Name List into the app itself (`public/employee-directory.json`)
 * so every browser / phone hitting this same dev server (or the same built
 * dist) loads the same current list — not just the one machine's localStorage.
 */
function employeeDirectoryPlugin(): Plugin {
  return {
    name: 'employee-directory-persist',
    configureServer(server) {
      server.middlewares.use('/api/employee-directory', jsonFileApi(EMPLOYEE_DIR_FILE));
    },
  };
}

/**
 * Round 8 §1 — shared Settings JSON on this machine (`data/app-settings.json`).
 * Any browser hitting this same Vite server (including incognito) reads the
 * same saved Settings. Vercel production uses `api/settings.ts` + Blob instead.
 */
function settingsStorePlugin(): Plugin {
  return {
    name: 'settings-store-persist',
    configureServer(server) {
      server.middlewares.use('/api/settings', jsonFileApi(SETTINGS_FILE));
    },
  };
}

// Local-first app: everything is bundled; the employee directory is also
// written into public/ so LAN clients share one current list.
export default defineConfig({
  plugins: [react(), tailwindcss(), employeeDirectoryPlugin(), settingsStorePlugin()],
  server: {
    host: true, // phone / other devices on the LAN can open the app
    port: 5173,
  },
  build: {
    chunkSizeWarningLimit: 2000,
  },
});
