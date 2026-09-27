// Exercise the real shared helpers in static/js/ortak.js (fetch contract, Turkish search, dates).
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const js=fs.readFileSync('static/js/ortak.js','utf8');
const parca=(bas,son)=>js.slice(js.indexOf(bas),js.indexOf(son));
const code=parca('function csrfToken()','\nfunction depoOlustur(')
  + '\n' + parca('function ggAaYyyyGecerliMi(','\nfunction tarihAlaniniDogrula(');
const context={Headers,FormData,document:{querySelector:()=>null}};
vm.createContext(context);vm.runInContext(code,context);
(async()=>{
  for(const response of [
    {ok:true,status:200,redirected:true,type:'text/html',body:'<html>Login</html>'},
    {ok:true,status:200,redirected:false,type:'text/html',body:'unexpected'},
    {ok:false,status:401,redirected:false,type:'application/json',body:{hata:'Oturum sona erdi'}}
  ]) {
    context.fetch=async()=>({...response,headers:new Headers({'content-type':response.type}),
      json:async()=>response.body,text:async()=>response.body});
    await assert.rejects(()=>context.pdgmFetch('/api/not',{method:'POST',body:'{}'}));
  }
  context.fetch=async()=>({ok:true,status:200,redirected:false,
    headers:new Headers({'content-type':'application/json'}),json:async()=>({tamam:true})});
  assert.equal((await context.pdgmFetch('/api/not')).tamam,true);
  assert.equal(typeof context.aramaMetni,'function');
  for(const [source,query] of [['TEST VERİSİ','test verisi'],['IŞIK','ışık'],['ÇIĞ','çığ']]) {
    assert(context.aramaMetni(source).includes(context.aramaMetni(query)),source);
  }
  // gg.aa.yyyy <-> ISO: admin dialogları ve Pano dönem filtresi aynı biçimi kullanır.
  assert.equal(context.ggAaYyyyGecerliMi('29.02.2024'),true);
  assert.equal(context.ggAaYyyyGecerliMi('29.02.2026'),false);
  assert.equal(context.ggAaYyyyGecerliMi('2026-09-27'),false);
  assert.equal(context.isoyaCevir('07.08.2026'),'2026-08-07');
  assert.equal(context.isoyaCevir(''),'');
  assert.equal(context.isodanGoster('2026-08-07'),'07.08.2026');
  assert.equal(context.isodanGoster('2026-08-07 13:45:00'),'07.08.2026');
  assert.equal(context.isodanGoster(''),'');
  console.log('Frontend contract: HTML/401 rejection, valid JSON, Turkish/Unicode search, date format PASS');
})().catch(e=>{console.error(e);process.exitCode=1;});
