const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html=fs.readFileSync('templates/base.html','utf8');
const code=html.slice(html.indexOf('function csrfToken()'),html.indexOf('\nfunction toast('));
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
  for(const [source,query] of [['TEST VERİSİ','test verisi'],['IŞIK','ışık'],['C\u0327IG\u0306','çığ']]) {
    assert(context.aramaMetni(source).includes(context.aramaMetni(query)),source);
  }
  console.log('Frontend contract: HTML/401 rejection, valid JSON, Turkish/Unicode search PASS');
})().catch(e=>{console.error(e);process.exitCode=1;});
