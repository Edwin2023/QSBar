export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    if (request.method === 'POST' && path === '/ping') {
      return handlePing(request, env);
    }
    if (path === '/stats') {
      return handleStats(env);
    }
    if (path === '/') {
      return handleDashboard();
    }
    return new Response('Not Found', { status: 404 });
  }
};

async function handlePing(request, env) {
  try {
    const body = await request.json();
    const { client_id, version, language, app, os_version } = body;
    if (!client_id) return new Response('{}', { status: 200 });

    const today = new Date().toISOString().slice(0, 10);

    const existingRaw = await env.CLIENTS.get(client_id);
    let record = existingRaw ? JSON.parse(existingRaw) : { first_seen: today, ping_count: 0 };
    record.last_seen = today;
    record.ping_count = (record.ping_count || 0) + 1;
    record.version = version;
    record.language = language;
    record.app = app;
    record.os_version = os_version;
    await env.CLIENTS.put(client_id, JSON.stringify(record));

    const countRaw = await env.COUNTERS.get(today);
    const count = countRaw ? parseInt(countRaw) + 1 : 1;
    await env.COUNTERS.put(today, String(count));

    return new Response(JSON.stringify({ status: 'ok' }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
  } catch {
    return new Response(JSON.stringify({ status: 'error' }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleStats(env) {
  const total_installs = (await env.CLIENTS.list()).keys.length;
  const today = new Date().toISOString().slice(0, 10);
  const daily_active = parseInt((await env.COUNTERS.get(today)) || '0');

  const last_7_days = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(Date.now() - i * 86400000).toISOString().slice(0, 10);
    const c = parseInt((await env.COUNTERS.get(d)) || '0');
    last_7_days.push({ date: d, count: c });
  }

  return new Response(JSON.stringify({ total_installs, daily_active, last_7_days }), {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }
  });
}

function handleDashboard() {
  const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>QSBar Analytics</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;color:#333;padding:40px 20px}
h1{text-align:center;font-size:24px;margin-bottom:30px;color:#0078d7}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:20px;max-width:800px;margin:0 auto}
.card{background:#fff;border-radius:8px;padding:24px;box-shadow:0 1px 3px rgba(0,0,0,.1);text-align:center}
.card .number{font-size:48px;font-weight:700;color:#0078d7}
.card .label{font-size:14px;color:#666;margin-top:8px}
.chart{max-width:800px;margin:30px auto;background:#fff;border-radius:8px;padding:24px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.chart h2{font-size:16px;color:#666;margin-bottom:16px}
.bars{display:flex;align-items:flex-end;gap:12px;height:160px}
.bar{flex:1;text-align:center}
.bar .fill{background:#0078d7;border-radius:4px 4px 0 0;min-height:4px;transition:height .3s}
.bar .val{font-size:12px;margin-top:6px;color:#333}
.bar .dat{font-size:11px;color:#999;margin-top:2px}
.footer{text-align:center;color:#999;font-size:12px;margin-top:30px}
</style>
</head>
<body>
<h1>QSBar 使用统计</h1>
<div class="grid">
  <div class="card"><div class="number" id="total">-</div><div class="label">总安装量</div></div>
  <div class="card"><div class="number" id="dau">-</div><div class="label">今日活跃</div></div>
</div>
<div class="chart">
  <h2>最近 7 天日活</h2>
  <div class="bars" id="bars"></div>
</div>
<p class="footer">QSBar Analytics</p>
<script>
fetch('/stats').then(r=>r.json()).then(d=>{
  document.getElementById('total').textContent = d.total_installs;
  document.getElementById('dau').textContent = d.daily_active;
  const max = Math.max(...d.last_7_days.map(x=>x.count), 1);
  const bars = document.getElementById('bars');
  d.last_7_days.forEach(day => {
    const h = (day.count / max) * 150;
    const div = document.createElement('div'); div.className='bar';
    div.innerHTML = '<div class="fill" style="height:'+h+'px"></div><div class="val">'+day.count+'</div><div class="dat">'+day.date.slice(5)+'</div>';
    bars.appendChild(div);
  });
});
</script>
</body>
</html>`;

  return new Response(html, {
    status: 200,
    headers: { 'Content-Type': 'text/html; charset=utf-8' }
  });
}
