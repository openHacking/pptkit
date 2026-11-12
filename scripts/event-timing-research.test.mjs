import assert from "node:assert/strict";
import test from "node:test";
import { bindEventTiming, replayEventTiming } from "./research/event-timing.mjs";
const stage=(id,trigger,duration=10)=>({id,target:{pageId:"p",elementId:"shape"},trigger,timing:{id:"leaf",kind:"effect",durationMs:duration}});
const click=(timeMs,pageId="p")=>({type:"page-click",timeMs,pageId});

test("stable targets and explicit with/after relations produce hand-calculated start/end times",()=>{
 const p=bindEventTiming("p",["shape"],[stage("a",{type:"enter"}),stage("b",{type:"with-previous"},20),stage("c",{type:"after-previous"},5),stage("d",{type:"page-click"})]);
 const events=[click(24),click(25)];
 assert.deepEqual(replayEventTiming(p,events,30).map(x=>[x.id,x.startMs,x.endMs]),[["a",0,10],["b",0,20],["c",20,25],["d",25,35]]);
 assert.deepEqual(replayEventTiming(p,events,0).map(x=>x.id),["a","b"]);
 assert.ok(Object.isFrozen(p.stages[0].target));
});

test("wrong-page/object/early clicks are ignored, exact boundaries and same-time log order are explicit",()=>{
 const p=bindEventTiming("p",["shape","button"],[stage("a",{type:"enter"}),stage("b",{type:"object-click",elementId:"button"},0),stage("c",{type:"page-click"},0),stage("d",{type:"page-click"},0)]);
 const events=[click(1),click(10,"other"),{type:"object-click",timeMs:10,pageId:"p",elementId:"shape"},{type:"object-click",timeMs:10,pageId:"p",elementId:"button"},click(10),click(10)];
 const result=replayEventTiming(p,events,10); assert.deepEqual(result.map(x=>[x.id,x.startMs]),[["a",0],["b",10],["c",10],["d",10]]);
 assert.deepEqual(replayEventTiming(p,events.slice(0,-1),10).map(x=>x.id),["a","b","c"]);
 for(const t of [0,10,5,100,10]) assert.deepEqual(replayEventTiming(p,events,t),replayEventTiming(p,structuredClone(events),t));
});

test("dangling/subtargets/unknown triggers/predecessors/IDs and malformed logs reject",()=>{
 const a=stage("a",{type:"page-click"});
 for(const mutated of [{...a,target:{pageId:"wrong",elementId:"shape"}},{...a,target:{...a.target,paragraphIndex:0}},{...a,target:{pageId:"p",elementId:"missing"}},{...a,trigger:{type:"object-click",elementId:"missing"}},{...a,trigger:{type:"media-end"}},{...a,trigger:{type:"after-previous"}}]) assert.throws(()=>bindEventTiming("p",["shape"],[mutated]));
 assert.throws(()=>bindEventTiming("p",["shape","shape"],[a])); assert.throws(()=>bindEventTiming("p",["shape"],[a,a]));
 const p=bindEventTiming("p",["shape"],[a]);
 for(const events of [[click(-1)],[click(2),click(1)],[{...click(1),type:"unknown"}],[{...click(1),extra:true}]]) assert.throws(()=>replayEventTiming(p,events,10));
 assert.throws(()=>replayEventTiming(p,[],Infinity));
});

test("binding snapshots objects/targets and replay neither mutates event logs nor leaks result mutation",()=>{
 const s=stage("a",{type:"page-click"}), objects=["shape"],p=bindEventTiming("p",objects,[s]);
 s.target.elementId="changed";s.timing.durationMs=99;objects[0]="gone";
 const events=[click(5)],before=structuredClone(events),r=replayEventTiming(p,events,10);
 assert.equal(r[0].endMs,15);r[0].samples[0].progress=.99;
 assert.equal(replayEventTiming(p,events,10)[0].samples[0].progress,.5);assert.deepEqual(events,before);
});


test("aggregate event research budgets count containers as well as leaves",()=>{
 const timing={id:"root",kind:"parallel",children:Array.from({length:6000},(_,i)=>({id:String(i),kind:"parallel",children:[]}))};
 const a={...stage("a",{type:"enter"}),timing},b={...stage("b",{type:"after-previous"}),timing};
 assert.throws(()=>bindEventTiming("p",["shape"],[a,b]),/budget|limit|children/i);
 assert.equal(bindEventTiming("p",["shape"],[a]).stages[0].timing.nodeCount,6001);
});
