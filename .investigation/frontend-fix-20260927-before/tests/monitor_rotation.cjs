// Exercise the real inline monitor script with a deterministic clock and tiny DOM.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync('templates/monitor.html','utf8');
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].at(-1)[1];
const storage = new Map();
function screen() {
  let time=0,reloads=0;
  const tasks=[];
  const cards=Array.from({length:25},()=>({hidden:false,classList:{toggle(_,hide){this.owner.hidden=hide;}}}));
  cards.forEach(c=>c.classList.owner=c);
  const group={dataset:{monitorGrup:'dizgide'},querySelectorAll:()=>cards};
  const window={matchMedia:()=>({matches:true}),addEventListener:()=>{},
    setInterval(fn,ms){tasks.push({fn,at:time+ms,ms});},
    setTimeout(fn,ms){tasks.push({fn,at:time+ms,ms:0});},location:{reload(){reloads++;}}};
  const sessionStorage={getItem:k=>storage.get(k)??null,setItem:(k,v)=>storage.set(k,String(v))};
  vm.runInNewContext(script,{window,document:{getElementById:()=>null,querySelectorAll:()=>[group]},sessionStorage,Intl,Date,Math,Number});
  return {visible:()=>cards.flatMap((c,i)=>c.hidden?[]:[i]),reloads:()=>reloads,
    advance(target){while(true){tasks.sort((a,b)=>a.at-b.at); const t=tasks[0]; if(!t||t.at>target)break;
      tasks.shift();time=t.at;t.fn();if(t.ms){t.at+=t.ms;tasks.push(t);}}time=target;}};
}
let s=screen();
assert.deepEqual(s.visible(),[0,1,2,3]);
s.advance(11999);assert.deepEqual(s.visible(),[0,1,2,3]);
s.advance(12000);assert.deepEqual(s.visible(),[4,5,6,7]);
const resumed=screen();assert.deepEqual(resumed.visible(),[4,5,6,7]);
s.advance(72000);assert.deepEqual(s.visible(),[24]);assert.equal(s.reloads(),0);
s.advance(84000);assert.equal(s.reloads(),1);
console.log('monitor rotation: 12s dwell, reload continuity, all 25 cards visible before refresh: PASS');
