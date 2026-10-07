/**
 * HTTP Basic Auth across the whole deployment.
 *
 * This is Vercel Edge Middleware, not Next.js middleware — this project is a
 * Vite SPA with serverless functions under `api/` (see `vercel.json`), so the
 * handler takes a Web `Request` and returning nothing lets the request through
 * to its destination. If this project ever moves to Next.js, the signature
 * changes to `export function middleware(req: NextRequest)` returning
 * `NextResponse.next()`.
 *
 * The app holds timesheets, wage rates and P&L figures, so the gate is
 * deliberately fail-closed: anything other than a correct credential gets a
 * 401, including the case where the environment variables are missing. A
 * misconfigured deployment locks everyone out rather than quietly serving
 * payroll data to the open internet.
 *
 * Setup: Vercel project → Settings → Environment Variables → add
 * `BASIC_AUTH_USER` and `BASIC_AUTH_PASSWORD` (all environments you want
 * protected), then redeploy. Locally, `vercel dev` reads them from `.env`.
 *
 * Note that Basic Auth sends the credential on every request, base64-encoded
 * rather than encrypted. That is fine over HTTPS, which Vercel enforces, but
 * it is a deployment gate — not a substitute for per-user accounts.
 */

export const config = {
  /**
   * Every path, including `/api/*` and static assets. Browsers resend the
   * credential automatically for the rest of the origin once the challenge is
   * answered, so the app's own `fetch('/api/settings')` calls keep working.
   */
  matcher: '/:path*',
};

/**
 * ASCII only. Header values are ByteStrings: a non-Latin-1 character here (an
 * em dash, a curly quote) makes `new Response(...)` throw, and a throw in
 * middleware is a 500 on every request — the whole site down, instead of a
 * password prompt.
 */
const REALM = 'AllPro Painters - Percent of Completion Builder';

function unauthorized(): Response {
  return new Response('Authentication required.', {
    status: 401,
    headers: {
      // `charset="UTF-8"` tells the browser to encode non-ASCII credentials as
      // UTF-8, which is what `decodeCredentials` below expects.
      'WWW-Authenticate': `Basic realm="${REALM}", charset="UTF-8"`,
      'Content-Type': 'text/plain; charset=utf-8',
      // Never let a proxy or the browser cache a challenge or a protected page.
      'Cache-Control': 'no-store',
    },
  });
}

/**
 * Compare without returning early on the first differing byte.
 *
 * A plain `===` leaks, through response timing, how much of a guess was
 * correct, which is enough to walk a secret out one character at a time. This
 * always walks the full width of the longer input. The length itself is still
 * observable; that is a far weaker signal than the contents, and the Edge
 * runtime has no `timingSafeEqual` to do better synchronously.
 */
function timingSafeEqual(a: string, b: string): boolean {
  const encoder = new TextEncoder();
  const aBytes = encoder.encode(a);
  const bBytes = encoder.encode(b);
  let diff = aBytes.length ^ bBytes.length;
  const width = Math.max(aBytes.length, bBytes.length);
  for (let i = 0; i < width; i++) {
    diff |= (aBytes[i] ?? 0) ^ (bBytes[i] ?? 0);
  }
  return diff === 0;
}

/**
 * `Authorization: Basic <base64>` → the user and password it carries.
 *
 * Returns null for anything malformed rather than throwing: `atob` rejects
 * invalid base64, and an unhandled throw in middleware is a 500 on every
 * request, which would take the site down instead of asking for a password.
 */
function decodeCredentials(header: string | null): { user: string; password: string } | null {
  if (!header) return null;

  const space = header.indexOf(' ');
  if (space === -1) return null;
  if (header.slice(0, space).toLowerCase() !== 'basic') return null;

  const encoded = header.slice(space + 1).trim();
  if (encoded === '') return null;

  let decoded: string;
  try {
    // `atob` yields one character per byte; reassemble those bytes and decode
    // them as UTF-8 so non-ASCII passwords survive the round trip.
    const bytes = Uint8Array.from(atob(encoded), (char) => char.charCodeAt(0));
    decoded = new TextDecoder().decode(bytes);
  } catch {
    return null;
  }

  // Split on the FIRST colon only — a password may contain colons, a username
  // may not (RFC 7617).
  const separator = decoded.indexOf(':');
  if (separator === -1) return null;

  return { user: decoded.slice(0, separator), password: decoded.slice(separator + 1) };
}

export default function middleware(request: Request): Response | undefined {
  const expectedUser = process.env.BASIC_AUTH_USER;
  const expectedPassword = process.env.BASIC_AUTH_PASSWORD;

  // Fail closed. Without this, an unset variable would make the comparison
  // below `undefined === undefined` for anyone who sent no credential at all.
  if (!expectedUser || !expectedPassword) {
    console.error(
      'middleware: BASIC_AUTH_USER / BASIC_AUTH_PASSWORD are not set; refusing every request.',
    );
    return unauthorized();
  }

  const credentials = decodeCredentials(request.headers.get('authorization'));
  if (!credentials) return unauthorized();

  // Both halves are always compared, so a wrong username costs the same time
  // as a wrong password.
  const userOk = timingSafeEqual(credentials.user, expectedUser);
  const passwordOk = timingSafeEqual(credentials.password, expectedPassword);
  if (!userOk || !passwordOk) return unauthorized();

  // Authenticated — fall through to the SPA or the `api/` function.
  return undefined;
}
