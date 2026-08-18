/**
 * Shared Settings store for the Vercel deployment (Round 8 §1).
 *
 * Browser localStorage is per-device and disappears when storage is cleared or
 * an incognito window closes. This API keeps one current Settings blob in
 * Vercel Blob so every client on the same deployment loads the same saved
 * Settings. Last write wins.
 *
 * Setup (one time, optional): Vercel project → Storage → Create Blob store →
 * connect to this project (adds BLOB_READ_WRITE_TOKEN) → redeploy. Without it
 * the app silently falls back to localStorage.
 */
import { del, list, put } from '@vercel/blob';
import type { VercelRequest, VercelResponse } from '@vercel/node';

const PATH = 'pct-completion/settings.json';

function noStore(res: VercelResponse): void {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
}

export default async function handler(req: VercelRequest, res: VercelResponse): Promise<void> {
  noStore(res);

  if (!process.env.BLOB_READ_WRITE_TOKEN) {
    res.status(503).json({
      error: 'missing_blob_token',
      message:
        'Shared Settings are optional. This browser still keeps your Settings in localStorage. To also share them across browsers/devices, enable Vercel Blob (Storage → Blob → Create) and redeploy.',
    });
    return;
  }

  try {
    if (req.method === 'GET') {
      const { blobs } = await list({ prefix: PATH, limit: 20 });
      const blob = blobs.find((b) => b.pathname === PATH);
      if (!blob) {
        res.status(404).json({ error: 'not_found' });
        return;
      }
      const upstream = await fetch(blob.url, { cache: 'no-store' });
      if (!upstream.ok) {
        res.status(404).json({ error: 'not_found' });
        return;
      }
      const data: unknown = await upstream.json();
      res.status(200).json(data);
      return;
    }

    if (req.method === 'PUT' || req.method === 'POST') {
      const raw =
        typeof req.body === 'string'
          ? req.body
          : req.body !== undefined && req.body !== null
            ? JSON.stringify(req.body)
            : '';
      if (!raw) {
        res.status(400).json({ error: 'empty_body' });
        return;
      }
      JSON.parse(raw); // validate
      await put(PATH, raw, {
        access: 'public',
        contentType: 'application/json',
        addRandomSuffix: false,
        allowOverwrite: true,
        cacheControlMaxAge: 60,
      });
      res.status(204).end();
      return;
    }

    if (req.method === 'DELETE') {
      try {
        await del(PATH);
      } catch {
        // already missing
      }
      res.status(204).end();
      return;
    }

    res.setHeader('Allow', 'GET, PUT, POST, DELETE');
    res.status(405).json({ error: 'method_not_allowed' });
  } catch (err) {
    console.error('[settings]', err);
    res.status(500).json({
      error: 'server_error',
      message: err instanceof Error ? err.message : 'Unexpected error',
    });
  }
}
