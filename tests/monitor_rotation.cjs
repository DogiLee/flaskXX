// Exercise the real monitor script (static/js/monitor.js) with a deterministic clock and tiny DOM:
// rotation, resume after reload, and that a reload only happens when the server answers.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const script = fs.readFileSync('static/js/monitor.js','utf8');
const storage = new Map();
const tick = () => new Promise((r) => setImmediate(r));

function screen({server = () => ({surum: 's1'})} = {}) {
  let time=0,reloads=0,nextId=1;
  const tasks=[];
  const banner={hidden:true,textContent:''};
  const cards=Array.from({length:25},()=>({hidden:false,classList:{toggle(_,hide){this.owner.hidden=hide;}}}));
  cards.forEach(c=>c.classList.owner=c);
  const group={dataset:{monitorGrup:'dizgide'},querySelectorAll:()=>cards};
  const schedule=(fn,ms,repeat)=>{const t={id:nextId++,fn,at:time+ms,ms:repeat?ms:0};tasks.push(t);return t.id;};
  const window={matchMedia:()=>({matches:true}),addEventListener:()=>{},
    setInterval:(fn,ms)=>schedule(fn,ms,true),setTimeout:(fn,ms)=>schedule(fn,ms,false),
    clearTimeout(id){const i=tasks.findIndex(t=>t.id===id);if(i>=0)tasks.splice(i,1);},
    location:{reload(){reloads++;}}};
  const fetch=async()=>{const cevap=server();if(cevap instanceof Error)throw cevap;
    return {ok:true,status:200,redirected:false,json:async()=>cevap};};
  const oturumDeposu={al:k=>storage.get(k)??null,yaz:(k,v)=>storage.set(k,String(v))};
  const document={getElementById:(id)=>id==='monitor-baglanti'?banner:null,querySelectorAll:()=>[group],
    body:{dataset:{veriSurumu:'s1'}}};
  vm.runInNewContext(script,{window,document,oturumDeposu,fetch,AbortController,Intl,Date,Math,Number,String,Error});
  return {visible:()=>cards.flatMap((c,i)=>c.hidden?[]:[i]),reloads:()=>reloads,banner,
    async advance(target){while(true){tasks.sort((a,b)=>a.at-b.at); const t=tasks[0]; if(!t||t.at>target)break;
      tasks.shift();time=t.at;t.fn();if(t.ms){t.at+=t.ms;tasks.push(t);}await tick();}time=target;await tick();}};
}

(async()=>{
  // 1) Rotasyon: 12 sn'de bir sayfa, yenilemeden sonra kaldığı yerden, tüm sayfalar gösterilmeden yenileme yok.
  let s=screen();
  assert.deepEqual(s.visible(),[0,1,2,3]);
  await s.advance(11999);assert.deepEqual(s.visible(),[0,1,2,3]);
  await s.advance(12000);assert.deepEqual(s.visible(),[4,5,6,7]);
  const resumed=screen();assert.deepEqual(resumed.visible(),[4,5,6,7]);
  await s.advance(72000);assert.deepEqual(s.visible(),[24]);assert.equal(s.reloads(),0);
  await s.advance(84000);assert.equal(s.reloads(),1);

  // 2) Sunucu yokken yenileme yapılmaz (tarayıcı hata sayfasına düşmez), bant görünür;
  //    sunucu dönünce yenilenir.
  storage.clear();
  let ayakta=false;
  s=screen({server:()=>ayakta?{surum:'s1'}:new Error('ağ yok')});
  await s.advance(84000);
  assert.equal(s.reloads(),0);assert.equal(s.banner.hidden,false);
  assert.match(s.banner.textContent,/ulaşılamıyor/);
  ayakta=true;
  await s.advance(84000+60000);
  assert.equal(s.reloads(),1);assert.equal(s.banner.hidden,true);

  // 3) Veri başka ekranda değişince tam turu beklemeden yenilenir.
  storage.clear();
  s=screen({server:()=>({surum:'s2'})});
  await s.advance(29999);assert.equal(s.reloads(),0);
  await s.advance(30000);assert.equal(s.reloads(),1);

  console.log('monitor rotation: 12s dwell, resume, no reload while server is down, reload on data change: PASS');
})().catch(e=>{console.error(e);process.exitCode=1;});
