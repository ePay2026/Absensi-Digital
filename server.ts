import path from 'path';
import app from './src/server/app';

const PORT = parseInt(process.env.PORT || '3000', 10);

async function startServer() {
  // Vite middleware for local development
  if (process.env.NODE_ENV !== 'production') {
    try {
      const { createServer: createViteServer } = await import('vite');
      const vite = await createViteServer({
        server: { middlewareMode: true },
        appType: 'spa',
      });
      app.use(vite.middlewares);
    } catch (e) {
      console.error('Failed to load Vite middleware:', e);
    }
  } else {
    // Production static serving (Docker / Railway / Local Container)
    const distPath = path.join(process.cwd(), 'dist');
    const express = (await import('express')).default;
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

// Automatically start server in standalone node environments (Railway, Docker, AI Studio container, Local)
const isServerless = process.env.VERCEL === '1' || !!process.env.AWS_LAMBDA_FUNCTION_NAME;
if (!isServerless) {
  startServer().catch(err => {
    console.error('Fatal error starting server:', err);
  });
}

export { app, startServer };
export default app;
