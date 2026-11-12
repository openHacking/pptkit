import assert from "node:assert/strict";
import test from "node:test";
import { compileTiming, sampleTiming } from "./research/timing-model.mjs";
const effect = (id, durationMs, extra = {}) => ({ id, kind: "effect", durationMs, ...extra });

test("nested sequence/parallel obey independent hand-calculated timestamps and event order", () => {
  const input = { id: "root", kind: "sequence", delayMs: 5, children: [effect("a", 10), { id: "pair", kind: "parallel", delayMs: 2, children: [effect("b", 4), effect("c", 3, { delayMs: 5 })] }, effect("zero", 0), effect("last", 2)] };
  const p = compileTiming(input);
  assert.equal(p.durationMs, 27);
  assert.deepEqual(p.leaves.map(x => [x.id,x.startMs,x.endMs]), [["a",5,15],["b",17,21],["c",22,25],["zero",25,25],["last",25,27]]);
  assert.deepEqual(p.events.filter(e => e.timeMs === 25).map(e => [e.id,e.type]), [["c","end"],["zero","start"],["zero","end"],["last","start"]]);
  input.children[0].durationMs = 999; assert.equal(p.leaves[0].durationMs,10);
  assert.ok(Object.isFrozen(p) && Object.isFrozen(p.leaves) && Object.isFrozen(p.leaves[0]));
});

test("absolute sampling at start/repeat/end boundaries is seek-order independent", () => {
  const p = compileTiming(effect("repeat", 10, { delayMs: 5, repeat: 3, fill: "hold" }));
  const expected = new Map([[4,["before",null,null]],[5,["active",0,0]],[14,["active",.9,0]],[15,["active",0,1]],[34,["active",.9,2]],[35,["ended",1,2]],[100,["ended",1,2]]]);
  for (const t of [100,15,4,35,14,5,34,15,5]) {
    const r = sampleTiming(p,t)[0]; assert.deepEqual([r.phase,r.progress,r.iteration],expected.get(t));
  }
  const zero = compileTiming(effect("zero",0,{repeat:3})); assert.equal(sampleTiming(zero,0)[0].phase,"ended"); assert.equal(sampleTiming(zero,10)[0].progress,null);
});

test("unknown nodes, bad scalars, overflow, accessors, duplicate IDs and cycles reject", () => {
  const cycle = { id: "cycle", kind: "sequence", children: [] }; cycle.children.push(cycle);
  const bad = [cycle, {kind:"future",id:"x"},effect("x",-1),effect("x",.5),effect("x",NaN),effect("x",Infinity),effect("x",1,{repeat:0}),effect("x",1,{repeat:Infinity}),effect("x",1,{delayMs:-1}),effect("x",1,{fill:"future"}),effect("x",1,{opaque:true}),{id:"x",kind:"parallel",children:[effect("same",1),effect("same",2)]}];
  for (const input of bad) assert.throws(() => compileTiming(input));
  assert.throws(() => compileTiming({ get id(){throw new Error("must not run");} }), /accessors/);
  assert.throws(() => compileTiming(effect("x",Number.MAX_SAFE_INTEGER,{delayMs:1}),{maxTimeMs:Number.MAX_SAFE_INTEGER}));
  assert.throws(() => sampleTiming(compileTiming(effect("x",1)),-1));
});

test("budgets bound expansion/depth without mutating input or poisoning retries", () => {
  const input = {id:"root",kind:"parallel",children:[effect("a",10),effect("b",20)]}, before=structuredClone(input);
  assert.throws(() => compileTiming(input,{maxNodes:2})); assert.throws(() => compileTiming(input,{maxDepth:1})); assert.throws(() => compileTiming(input,{maxTimeMs:19}));
  assert.deepEqual(input,before); assert.equal(compileTiming(input,{maxNodes:3,maxDepth:2,maxTimeMs:20}).durationMs,20);
  for (const options of [{maxDepth:0},{maxNodes:Infinity},{maxTimeMs:-1},{unknown:1}]) assert.throws(() => compileTiming(input,options));
});

test("seeded trees agree with a separate flattened interval oracle", () => {
  let seed=12541, id=0; const rand=n=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed%n;};
  const make=depth=>depth===0||rand(3)===0?effect(String(++id),rand(12),{delayMs:rand(5),repeat:1+rand(3)}):{id:String(++id),kind:rand(2)?"sequence":"parallel",delayMs:rand(5),children:Array.from({length:rand(4)},()=>make(depth-1))};
  function oracle(n){const delay=n.delayMs??0;if(n.kind==="effect")return {span:delay+n.durationMs*(n.repeat??1),rows:[[n.id,delay,delay+n.durationMs*(n.repeat??1)]]};let span=0,rows=[];for(const c of n.children){const r=oracle(c),offset=n.kind==="sequence"?span:0;rows.push(...r.rows.map(([id,a,b])=>[id,a+delay+offset,b+delay+offset]));span=n.kind==="sequence"?span+r.span:Math.max(span,r.span);}return {span:delay+span,rows};}
  for(let i=0;i<200;i++){const input=make(4),o=oracle(input),p=compileTiming(input);assert.equal(p.durationMs,o.span);assert.deepEqual(p.leaves.map(x=>[x.id,x.startMs,x.endMs]),o.rows);}
});
