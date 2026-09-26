// Exercise actual template function in a VM; no browser automation.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync('templates/base.html', 'utf8');
const start = html.indexOf('function aramaMetni(');
const end = html.indexOf('\nfunction toast(', start);
assert(start >= 0 && end > start);
const source = html.slice(start, end);
const context = { Headers, FormData, csrfToken: () => 'synthetic-token',
  fetch: async () => ({ok:true, redirected:true, status:200,
    headers:new Headers({'content-type':'text/html'}),
    text:async () => '<html>Giriş yapın</html>'}) };
vm.createContext(context);
vm.runInContext(source, context);
(async () => {
  let rejected=false, result;
  try { result=await context.pdgmFetch('/api/not',{method:'POST',body:'{}'}); }
  catch (_) { rejected=true; }
  const raw='TEST VERİSİ';
  const indexed=context.aramaMetni(raw);
  const query=context.aramaMetni('test verisi');
  const report={
    redirectedLogin:{rejected,returned:result,expected:'Reject non-JSON/redirected login response'},
    turkishSearch:{indexed,query,matches:indexed.includes(query),expected:true}
  };
  fs.writeFileSync('outputs/audit-fixes-20260922/frontend-probes.json',JSON.stringify(report,null,2));
  console.log(JSON.stringify(report,null,2));
  process.exitCode=(!rejected || !indexed.includes(query)) ? 1 : 0;
})();
