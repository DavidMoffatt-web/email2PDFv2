export default async function handler(req, res) {
  const url = req.query.url;
  if (!url || !/^https?:\/\//.test(url)) {
    res.status(400).json({ error: 'Missing or invalid url parameter', url });
    return;
  }
  try {
    const response = await fetch(url);
    if (!response.ok) {
      res.status(response.status).json({ error: 'Failed to fetch image', url, status: response.status });
      return;
    }
    const contentType = response.headers.get('content-type');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Content-Type', contentType);
    const buffer = await response.arrayBuffer();
    res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: 'Proxy error', message: err.message, url });
  }
}
