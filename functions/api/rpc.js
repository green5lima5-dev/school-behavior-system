/** Cloudflare Pages Functions proxy. Secrets exist only on Cloudflare. */
export async function onRequestPost({request, env}) {
  const json = (body, status=200) => new Response(JSON.stringify(body), {
    status, headers: {'content-type':'application/json; charset=utf-8', 'cache-control':'no-store',
      'x-content-type-options':'nosniff'}
  });
  if (!env.GAS_WEBAPP_URL || !env.RPC_SHARED_SECRET) {
    return json({ok:false,error:'API ยังไม่ได้ตั้งค่า Secrets'}, 503);
  }
  let body;
  try { body = await request.json(); }
  catch { return json({ok:false,error:'Invalid JSON'},400); }
  if (!body || typeof body.method !== 'string' || !Array.isArray(body.args)) {
    return json({ok:false,error:'Invalid request'},400);
  }
  const methods = new Set([
    'login','registerStudent','getSession','getDashboardData','getStudentPortalData',
    'testConnection','saveTransaction','deleteTransaction','listUsers','createUser',
    'resetUserPassword','deleteUser','setUserActive','resetAllStudentScores',
    'getMyProfile','updateMyProfile'
  ]);
  if (!methods.has(body.method) || body.args.length > 4) {
    return json({ok:false,error:'Method not permitted'},403);
  }
  const url = String(env.GAS_WEBAPP_URL);
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url)) {
    return json({ok:false,error:'Invalid backend URL configuration'},503);
  }
  try {
    const upstream = await fetch(url, {
      method:'POST',
      headers:{'content-type':'text/plain;charset=utf-8'},
      body:JSON.stringify({method:body.method,args:body.args,secret:env.RPC_SHARED_SECRET}),
      redirect:'follow',
    });
    if (!upstream.ok) return json({ok:false,error:'Backend ไม่ตอบสนองตามปกติ'},502);
    const content = await upstream.text();
    let result;
    try { result=JSON.parse(content); }
    catch { return json({ok:false,error:'Apps Script ส่งข้อมูลที่ไม่ใช่ JSON: ตรวจสิทธิ์ deployment และ URL'},502); }
    return json(result, result.code==='UNAUTHORIZED'?502:200);
  } catch (error) {
    console.error('Apps Script gateway error', String(error));
    return json({ok:false,error:'ไม่สามารถติดต่อ Apps Script Backend ได้'},502);
  }
}
export function onRequestGet() {
  return new Response(JSON.stringify({ok:false,error:'Method not allowed'}), {
    status:405, headers:{'content-type':'application/json','allow':'POST'}
  });
}
