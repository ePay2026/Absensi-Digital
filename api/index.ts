import type { IncomingMessage, ServerResponse } from 'http';
import app from '../src/server/app';

// Vercel Serverless Function entry point
export default function handler(req: IncomingMessage, res: ServerResponse) {
  return app(req as any, res as any);
}
