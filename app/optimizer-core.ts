export type OptimizerCabinet={id:string;no:string;tons:number;au:number;ag:number};
export type OptimizerRules={auRates:{min:number;rate:number}[];agRates:{min:number;rate:number}[]};
export type OptimizerGoal={auMin:number;agMin:number;maxGroup:number};
export type OptimizerSolution={groups:string[][];payableAu:number;payableAg:number;covered:number;proven:boolean;upperAu?:number;upperAg?:number;message:string};

const EPS=1e-6;
type Totals={count:number;tons:number;au:number;ag:number};
type Candidate={groups:number[][];au:number;ag:number;covered:number};

const tier=(rows:{min:number;rate:number}[],value:number)=>rows.find(x=>value>=x.min)?.rate||0;
const add=(a:Totals,c:OptimizerCabinet):Totals=>({count:a.count+1,tons:a.tons+c.tons,au:a.au+c.tons*c.au,ag:a.ag+c.tons*c.ag});
const valid=(x:Totals,g:OptimizerGoal)=>x.count>=2&&x.count<=g.maxGroup&&x.au+1e-9>=g.auMin*x.tons&&x.ag+1e-9>=g.agMin*x.tons;
function score(x:Totals,r:OptimizerRules){const auGrade=x.au/x.tons,agGrade=x.ag/x.tons;return{au:x.au*tier(r.auRates,auGrade),ag:x.ag*(agGrade>=1000?tier(r.agRates,agGrade):agGrade>=100?.89-Math.ceil((1000-agGrade)/100)*.01:0)}}
function better(a:Candidate,b:Candidate|null){if(!b)return true;if(Math.abs(a.au-b.au)>EPS)return a.au>b.au;if(Math.abs(a.ag-b.ag)>EPS)return a.ag>b.ag;if(a.covered!==b.covered)return a.covered>b.covered;if(a.groups.length!==b.groups.length)return a.groups.length<b.groups.length;return a.groups.map(x=>x.join(',')).join('|')<b.groups.map(x=>x.join(',')).join('|')}
function candidate(groups:number[][],items:OptimizerCabinet[],r:OptimizerRules){let au=0,ag=0,covered=0;for(const group of groups){let t:Totals={count:0,tons:0,au:0,ag:0};for(const i of group)t=add(t,items[i]);const s=score(t,r);au+=s.au;ag+=s.ag;covered+=group.length}return{groups,au,ag,covered}}

/** Exact set-packing/partition solver used as the correctness oracle for small batches. */
export function solveExactSmall(items:OptimizerCabinet[],goal:OptimizerGoal,rules:OptimizerRules):Candidate{
 const n=items.length,size=1<<n,totals:Array<Totals>=Array(size);totals[0]={count:0,tons:0,au:0,ag:0};
 const feasible:number[][]=Array.from({length:n},()=>[]);
 for(let mask=1;mask<size;mask++){const bit=mask&-mask,i=31-Math.clz32(bit);totals[mask]=add(totals[mask^bit],items[i]);if(valid(totals[mask],goal))feasible[i].push(mask)}
 const memo=new Map<number,Candidate>();
 const visit=(mask:number):Candidate=>{const cached=memo.get(mask);if(cached)return cached;if(!mask)return{groups:[],au:0,ag:0,covered:0};const bit=mask&-mask,pivot=31-Math.clz32(bit);let best=visit(mask^bit);for(const groupMask of feasible[pivot]){if((groupMask&mask)!==groupMask)continue;const rest=visit(mask^groupMask),indices:number[]=[];for(let i=0;i<n;i++)if(groupMask&(1<<i))indices.push(i);const s=score(totals[groupMask],rules),next={groups:[indices,...rest.groups],au:s.au+rest.au,ag:s.ag+rest.ag,covered:indices.length+rest.covered};if(better(next,best))best=next}memo.set(mask,best);return best};
 return visit(size-1);
}
function bestSingleGroup(items:OptimizerCabinet[],goal:OptimizerGoal,rules:OptimizerRules,deadline:number):Candidate|null{
 const ordered=items.map((c,index)=>({c,index})).sort((a,b)=>b.c.tons*b.c.au-a.c.tons*a.c.au||b.c.tons*b.c.ag-a.c.tons*a.c.ag),n=ordered.length;
 const suffixAu=Array(n+1).fill(0),suffixAg=Array(n+1).fill(0),suffixAuHelp=Array(n+1).fill(0),suffixAgHelp=Array(n+1).fill(0),maxAgRate=Math.max(.89,...rules.agRates.map(x=>x.rate));
 for(let i=n-1;i>=0;i--){const c=ordered[i].c;suffixAu[i]=suffixAu[i+1]+c.tons*c.au;suffixAg[i]=suffixAg[i+1]+c.tons*c.ag;suffixAuHelp[i]=suffixAuHelp[i+1]+Math.max(0,c.tons*(c.au-goal.auMin));suffixAgHelp[i]=suffixAgHelp[i+1]+Math.max(0,c.tons*(c.ag-goal.agMin))}
 let best:Candidate|null=null,picked:number[]=[];let nodes=0,timedOut=false;
 const visit=(i:number,t:Totals)=>{if(timedOut)return;if((++nodes&8191)===0&&performance.now()>deadline){timedOut=true;return}if(picked.length>goal.maxGroup||picked.length+n-i<2)return;const auBal=t.au-goal.auMin*t.tons,agBal=t.ag-goal.agMin*t.tons;if(auBal+suffixAuHelp[i]<-1e-9||agBal+suffixAgHelp[i]<-1e-9)return;if(best){const maxAuRate=Math.max(0,...rules.auRates.map(x=>x.rate));if((t.au+suffixAu[i])*maxAuRate<best.au-EPS)return;if(Math.abs((t.au+suffixAu[i])*maxAuRate-best.au)<=EPS&&(t.ag+suffixAg[i])*maxAgRate<best.ag-EPS)return}if(valid(t,goal)){const s=score(t,rules),next={groups:[[...picked]],au:s.au,ag:s.ag,covered:picked.length};if(better(next,best))best=next}if(i===n||picked.length===goal.maxGroup)return;const row=ordered[i];picked.push(row.index);visit(i+1,add(t,row.c));picked.pop();visit(i+1,t)};
 visit(0,{count:0,tons:0,au:0,ag:0});return best as Candidate|null;
}

function exactTwoWay(indices:number[],items:OptimizerCabinet[],goal:OptimizerGoal,rules:OptimizerRules,deadline:number,onBest?:(x:Candidate)=>void){
 if(indices.length<4||indices.length>31)return null;const fixed=indices[0],rest=indices.slice(1),limit=2**rest.length;let total:Totals={count:0,tons:0,au:0,ag:0};for(const i of indices)total=add(total,items[i]);let left=add({count:0,tons:0,au:0,ag:0},items[fixed]),previous=0,best:Candidate|null=null;
 for(let grayIndex=0;grayIndex<limit;grayIndex++){if((grayIndex&16383)===0&&performance.now()>deadline)break;const gray=grayIndex^(grayIndex>>1);if(grayIndex){const changed=gray^previous,bit=31-Math.clz32(changed),c=items[rest[bit]];if(gray&changed)left=add(left,c);else left={count:left.count-1,tons:left.tons-c.tons,au:left.au-c.tons*c.au,ag:left.ag-c.tons*c.ag}}previous=gray;const right:Totals={count:total.count-left.count,tons:total.tons-left.tons,au:total.au-left.au,ag:total.ag-left.ag};if(!valid(left,goal)||!valid(right,goal))continue;const ls=score(left,rules),rs=score(right,rules),summary={groups:[] as number[][],au:ls.au+rs.au,ag:ls.ag+rs.ag,covered:indices.length};if(!better(summary,best))continue;const a:number[]=[fixed],b:number[]=[];for(let k=0;k<rest.length;k++)((gray&(2**k))?a:b).push(rest[k]);best={...summary,groups:[a,b]};onBest?.(best)}return best;
}

function improveBySplitting(start:Candidate,items:OptimizerCabinet[],goal:OptimizerGoal,rules:OptimizerRules,deadline:number,onBest?:(x:Candidate)=>void){let best=start,changed=true;while(changed&&performance.now()<deadline){changed=false;for(let i=0;i<best.groups.length;i++){const split=exactTwoWay(best.groups[i],items,goal,rules,deadline);if(!split)continue;const next=candidate([...best.groups.slice(0,i),...split.groups,...best.groups.slice(i+1)],items,rules);if(better(next,best)){best=next;changed=true;onBest?.(best);break}}}return best}

export function solveAnytime(items:OptimizerCabinet[],goal:OptimizerGoal,rules:OptimizerRules,timeLimitMs:number,onProgress?:(x:OptimizerSolution)=>void):OptimizerSolution{
 const started=performance.now(),deadline=started+timeLimitMs,emit=(x:Candidate,proven:boolean,message:string)=>{const solution:OptimizerSolution={groups:x.groups.map(g=>g.map(i=>items[i].id)),payableAu:x.au,payableAg:x.ag,covered:x.covered,proven,message};onProgress?.(solution);return solution};
 if(items.length<=14){const exact=solveExactSmall(items,goal,rules);return emit(exact,true,'已完成全部合法入组集合与分组的精确比较')}
 const single=bestSingleGroup(items,goal,rules,deadline);if(!single)return emit({groups:[],au:0,ag:0,covered:0},performance.now()<deadline,'没有找到满足门槛的组合');let current:Candidate=single;emit(current,false,'已找到初始合法方案，正在优化分组');
 const two=exactTwoWay(single.groups[0],items,goal,rules,deadline,x=>{if(better(x,current)){current=x;emit(current,false,'正在精确比较当前入组集合的全部两组拆分')}});if(two&&better(two,current))current=two;
 current=improveBySplitting(current,items,goal,rules,deadline,x=>emit(x,false,'正在继续检查多组拆分'));
 return emit(current,false,performance.now()>=deadline?'已达到本次搜索时限，保留当前最佳合法方案':'已完成当前精确搜索范围；40柜联合全局证明仍未完成');
}
