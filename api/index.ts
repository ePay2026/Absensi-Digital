import type { IncomingMessage, ServerResponse } from 'http';
import app from '../src/server/app.js';

function normalizeReqUrl(req: IncomingMessage): void {
  if (!req.url) return;
  try {
    const parsed = new URL(req.url, 'http://localhost');
    const routeParam = parsed.searchParams.get('__route') || parsed.searchParams.get('path');
    
    if (routeParam) {
      parsed.searchParams.delete('__route');
      parsed.searchParams.delete('path');
      const search = parsed.searchParams.toString();
      const cleanRoute = routeParam.startsWith('/') ? routeParam : `/${routeParam}`;
      req.url = `/api${cleanRoute}${search ? `?${search}` : ''}`;
      return;
    }
    // Check Vercel headers
    const matchedPath = req.headers['x-matched-path'] || req.headers['x-vercel-matched-path'];
    if (typeof matchedPath === 'string' && matchedPath.startsWith('/api/') && (req.url === '/api' || req.url === '/api/' || req.url === '/')) {
      const search = parsed.searchParams.toString();
      req.url = `${matchedPath}${search ? `?${search}` : ''}`;
      return;
    }
    if (!req.url.startsWith('/api')) {
      req.url = `/api${req.url.startsWith('/') ? '' : '/'}${req.url}`;
    }
  } catch (e) {
    // If URL parsing fails, continue with original req.url
  }
}

// Vercel Serverless Function entry point
export default function handler(req: IncomingMessage, res: ServerResponse) {
  try {
    normalizeReqUrl(req);
    // Execute Express app
    return app(req as any, res as any);
  } catch (err: any) {
    console.error('CRITICAL ERROR IN HANDLER:', err);
    res.statusCode = 500;
    res.end(JSON.stringify({ error: 'Handler Error', message: err.message }));
  }
}
