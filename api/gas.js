const GAS_BASE = 'https://script.google.com/macros/s/AKfycbyrNp04DqrGf1rahuiNNSyRjorfcstjWcDQa2GXU4nMvfF1QJcW8ucYjWhfx4_WlLOT/exec';

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks).toString('utf-8');
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const params = new URLSearchParams(req.query || {}).toString();
    const targetUrl = params ? `${GAS_BASE}?${params}` : GAS_BASE;

    const fetchOpts = { method: req.method };

    if (req.method === 'POST') {
      const body = await readBody(req);
      fetchOpts.body = body;
      const ct = req.headers['content-type'];
      if (ct) fetchOpts.headers = { 'Content-Type': ct };
    }

    const gasRes = await fetch(targetUrl, fetchOpts);
    const text = await gasRes.text();

    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.status(gasRes.status).send(text);
  } catch (e) {
    console.error('[api/gas] プロキシエラー:', e.message);
    return res.status(502).json({ error: 'GAS接続失敗: ' + (e.message || String(e)) });
  }
};
