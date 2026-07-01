// Vercel Edge Runtime（タイムアウト30秒、Hobbyプランでも有効）
export const config = { runtime: 'edge' };

const GAS_BASE = 'https://script.google.com/macros/s/AKfycbyrNp04DqrGf1rahuiNNSyRjorfcstjWcDQa2GXU4nMvfF1QJcW8ucYjWhfx4_WlLOT/exec';

// GASへのfetch（リダイレクトを手動追跡してPOSTボディを保持する）
async function fetchGas(url, method, body, contentType, hops) {
  const opts = { method, redirect: 'manual' };
  if (body !== null && body !== undefined) {
    opts.body = body;
    opts.headers = { 'Content-Type': contentType || 'text/plain;charset=utf-8' };
  }
  const r = await fetch(url, opts);
  const isRedirect = r.status >= 301 && r.status <= 308;
  if (isRedirect && hops < 5) {
    const loc = r.headers.get('location');
    if (!loc) return r;
    // 307/308はPOSTを維持、301/302/303はGETへ変換（RFC標準）
    const keepPost = r.status === 307 || r.status === 308;
    const nextMethod   = keepPost ? method : 'GET';
    const nextBody     = nextMethod === 'POST' ? body : null;
    const nextCt       = nextMethod === 'POST' ? contentType : null;
    console.log(`[api/gas] redirect ${r.status} ${method}->${nextMethod} hops=${hops}`);
    return fetchGas(loc, nextMethod, nextBody, nextCt, hops + 1);
  }
  return r;
}

export default async function handler(req) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  const url = new URL(req.url);
  const params = url.searchParams.toString();
  const targetUrl = params ? `${GAS_BASE}?${params}` : GAS_BASE;

  console.log(`[api/gas] ${req.method} ${params || '(no params)'}`);

  try {
    let body = null;
    let contentType = null;

    if (req.method === 'POST') {
      body = await req.text();
      contentType = req.headers.get('content-type') || 'text/plain;charset=utf-8';
      console.log(`[api/gas] POST ct:${contentType} bodyLen:${body.length} body0:${body.slice(0, 120)}`);
    }

    const gasRes = await fetchGas(targetUrl, req.method, body, contentType, 0);
    const text = await gasRes.text();

    console.log(`[api/gas] done status:${gasRes.status} resLen:${text.length} res0:${text.slice(0, 120)}`);

    return new Response(text || '{}', {
      status: 200,
      headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
    });
  } catch (e) {
    console.error(`[api/gas] ERROR: ${e.name} ${e.message}`);
    return new Response(JSON.stringify({ error: 'GAS接続失敗: ' + (e.message || String(e)) }), {
      status: 502,
      headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
    });
  }
}
