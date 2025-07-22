import express from 'express';
import cors from 'cors';
import { createProxyMiddleware } from 'http-proxy-middleware';

const app = express();
app.use(cors());

app.use('/proxy', createProxyMiddleware({
  target: '', // Not used, dynamic routing
  changeOrigin: true,
  pathRewrite: (path, req) => path.replace(/^\/proxy\//, ''),
  router: (req) => new URL(req.url.replace(/^\/proxy\//, '')),
  logger: console
}));

export default app;
