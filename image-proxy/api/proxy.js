export default async function handler(req, res) {
  const url = req.query.url;
  if (!url) {
    res.status(400).send('Missing url parameter');
    return;
  }
  try {
    const response = await fetch(url);
    if (!response.ok) {
      res.status(response.status).send('Failed to fetch image');
      return;
    }
    const contentType = response.headers.get('content-type');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Content-Type', contentType);
    const buffer = await response.arrayBuffer();
    res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).send('Proxy error: ' + err.message);
  }
}
